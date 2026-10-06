/** REST API 路由。 */
import express from "express";
import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { config } from "../config.js";
import { jobStore } from "../lib/jobs.js";
import { runJobInBackground, executeJob, OPS, archiveName } from "../lib/runner.js";
import { uploadFiles, collectUploads, parseOptions } from "./upload.js";
import { badRequest, notFound, unsupported, wrap } from "../lib/errors.js";
import { registerFile, getFile } from "../lib/registry.js";
import { ensureDir, uniquePath, mimeFor } from "../lib/paths.js";
import { detectEngines, invalidateEngineCache } from "../lib/engines/detect.js";
import { getPdfInfo } from "../lib/engines/render.js";
import { runProcess } from "../lib/proc.js";
import { log } from "../lib/logger.js";

const OP_SET = new Set(OPS);

function contentDisposition(name, inline = false) {
  const ascii = name.replace(/[^\x20-\x7e]/g, "_").replace(/"/g, "'");
  return `${inline ? "inline" : "attachment"}; filename="${ascii}"; filename*=UTF-8''${encodeURIComponent(name)}`;
}

async function sendFileTo(res, filePath, name, { inline = false, mime } = {}) {
  const stat = await fs.stat(filePath).catch(() => null);
  if (!stat || !stat.isFile()) throw notFound("文件不存在或已被移动。");
  res.setHeader("Content-Type", mime || mimeFor(name));
  res.setHeader("Content-Length", String(stat.size));
  res.setHeader("Content-Disposition", contentDisposition(name, inline));
  res.setHeader("Cache-Control", "no-store");
  await new Promise((resolve, reject) => {
    res.sendFile(path.resolve(filePath), (err) => (err ? reject(err) : resolve()));
  });
}

/** 只允许打开输出目录、临时目录或项目目录下的路径。 */
function assertOpenable(target) {
  const resolved = path.resolve(String(target || ""));
  const roots = [config.outputDir, config.tmpDir, config.root].map((p) => path.resolve(p));
  if (!roots.some((root) => resolved === root || resolved.startsWith(root + path.sep))) {
    throw badRequest("出于安全考虑，只能打开输出目录或项目目录中的路径。");
  }
  return resolved;
}

export function createApiRouter() {
  const router = express.Router();

  router.get("/health", (req, res) => {
    res.json({
      ok: true,
      name: config.name,
      version: config.version,
      uptime: Math.round(process.uptime()),
      jobs: jobStore.list().length,
    });
  });

  router.get("/engines", wrap(async (req, res) => {
    const engines = await detectEngines({ refresh: req.query.refresh === "1" });
    res.json({
      engines,
      limits: { maxUploadMb: config.maxUploadMb, convertTimeoutMs: config.convertTimeoutMs },
      outputDir: config.outputDir,
      ops: OPS,
    });
  }));

  router.get("/settings", (req, res) => {
    res.json({
      outputDir: config.outputDir,
      tmpDir: config.tmpDir,
      maxUploadMb: config.maxUploadMb,
      version: config.version,
      https: config.https,
      port: config.port,
    });
  });

  router.post("/settings", wrap(async (req, res) => {
    const next = req.body?.outputDir;
    if (next) {
      const resolved = path.resolve(String(next));
      await ensureDir(resolved);
      config.outputDir = resolved;
      log.info("output dir changed to", resolved);
    }
    res.json({ outputDir: config.outputDir });
  }));

  router.post("/system/open-folder", wrap(async (req, res) => {
    const target = assertOpenable(req.body?.path || config.outputDir);
    await fs.mkdir(target, { recursive: true }).catch(() => {});
    const [cmd, args] = process.platform === "win32"
      ? ["explorer.exe", [target]]
      : process.platform === "darwin" ? ["open", [target]] : ["xdg-open", [target]];
    await runProcess(cmd, args, { timeoutMs: 5000 });
    res.json({ ok: true, path: target });
  }));

  router.post("/system/open-file", wrap(async (req, res) => {
    const target = assertOpenable(req.body?.path);
    const stat = await fs.stat(target).catch(() => null);
    if (!stat) throw notFound("文件不存在。");
    let cmd;
    let args;
    if (process.platform === "win32") {
      // 用 cmd /c start 以便按扩展名调用默认程序
      cmd = "cmd.exe";
      args = ["/c", "start", "", target];
    } else if (process.platform === "darwin") {
      cmd = "open";
      args = [target];
    } else {
      cmd = "xdg-open";
      args = [target];
    }
    await runProcess(cmd, args, { timeoutMs: 8000 });
    res.json({ ok: true, path: target });
  }));

  /* ------------------------------------------------------------ 任务创建 */

  router.post("/jobs", uploadFiles, wrap(async (req, res) => {
    const { files, dir } = collectUploads(req);
    const op = String(req.body?.op || "").trim();
    if (!OP_SET.has(op)) throw badRequest(`不支持的操作: ${op || "(空)"}。可选: ${OPS.join(", ")}`);
    if (!files.length) throw badRequest("没有收到文件。");
    const options = parseOptions(req);
    const job = jobStore.create({ op, inputs: files, options, dir });
    log.info(`job ${job.id} ${op} (${files.length} file(s))`);
    if (req.query.sync === "1") {
      await executeJob(job);
      const status = jobStore.getPublic(job.id);
      if (status.status === "error") {
        res.status(500).json({ error: status.error });
        return;
      }
      const resultFiles = job.result?.files || [];
      if (resultFiles.length === 1) {
        await sendFileTo(res, resultFiles[0].path, resultFiles[0].name);
      } else {
        const zip = await zipResult(job);
        res.setHeader("Content-Type", "application/zip");
        res.setHeader("Content-Disposition", contentDisposition(archiveName(job)));
        res.send(zip);
      }
      return;
    }
    runJobInBackground(job);
    res.status(202).json({ jobId: job.id, status: job.status });
  }));

  /* -------------------------------------------------------------- 任务查询 */

  router.get("/jobs", (req, res) => {
    res.json({ jobs: jobStore.list().slice(0, 50) });
  });

  router.get("/jobs/:id", (req, res) => {
    const job = jobStore.getPublic(req.params.id);
    if (!job) throw notFound("任务不存在或已过期。");
    res.json(job);
  });

  router.post("/jobs/:id/cancel", (req, res) => {
    const job = jobStore.cancel(req.params.id);
    if (!job) throw notFound("任务不存在或已过期。");
    res.json({ ok: true, status: jobStore.getPublic(job.id).status });
  });

  router.delete("/jobs/:id", wrap(async (req, res) => {
    const ok = await jobStore.remove(req.params.id);
    if (!ok) throw notFound("任务不存在或已过期。");
    res.json({ ok: true });
  }));

  router.get("/jobs/:id/result/:index", wrap(async (req, res) => {
    const job = jobStore.get(req.params.id);
    if (!job) throw notFound("任务不存在或已过期。");
    if (job.status !== "done" || !job.result) throw badRequest("任务尚未完成。");
    const index = Number(req.params.index);
    const file = job.result.files[index];
    if (!file) throw notFound("结果文件不存在。");
    const inline = req.query.inline === "1";
    await sendFileTo(res, file.path, file.name, { inline });
  }));

  router.get("/jobs/:id/archive", wrap(async (req, res) => {
    const job = jobStore.get(req.params.id);
    if (!job) throw notFound("任务不存在或已过期。");
    if (job.status !== "done" || !job.result) throw badRequest("任务尚未完成。");
    const zip = await zipResult(job);
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", contentDisposition(archiveName(job)));
    res.send(zip);
  }));

  /* ------------------------------------------------------- 直接保存上传文件 */

  router.post("/files/save", uploadFiles, wrap(async (req, res) => {
    const { files, dir } = collectUploads(req);
    if (!files.length) throw badRequest("没有收到文件。");
    const options = parseOptions(req);
    await ensureDir(config.outputDir);
    const saved = [];
    const used = new Set();
    for (const file of files) {
      const desired = options.filename || file.name;
      const target = await uniquePath(config.outputDir, desired);
      if (used.has(target)) continue;
      used.add(target);
      // eslint-disable-next-line no-await-in-loop
      await fs.rename(file.path, target).catch(async () => {
        await fs.copyFile(file.path, target);
        await fs.rm(file.path, { force: true });
      });
      // eslint-disable-next-line no-await-in-loop
      const stat = await fs.stat(target);
      saved.push({
        name: path.basename(target),
        path: target,
        size: stat.size,
        downloadId: registerFile(target, { name: path.basename(target) }),
      });
    }
    await fs.rm(dir, { recursive: true, force: true }).catch(() => {});
    res.json({ files: saved, outputDir: config.outputDir });
  }));

  router.get("/files/:id/download", wrap(async (req, res) => {
    const entry = getFile(req.params.id);
    if (!entry) throw notFound("文件不存在或已过期。");
    await sendFileTo(res, entry.path, entry.name);
  }));

  /* ------------------------------------------------------------ 同步便捷接口 */

  router.post("/convert/:op", uploadFiles, wrap(async (req, res) => {
    const op = req.params.op;
    if (!OP_SET.has(op)) throw unsupported(`不支持的操作: ${op}`);
    const { files, dir } = collectUploads(req);
    if (!files.length) throw badRequest("没有收到文件。");
    const job = jobStore.create({ op, inputs: files, options: parseOptions(req), dir });
    await executeJob(job);
    const publicJob = jobStore.getPublic(job.id);
    if (publicJob.status === "error") {
      res.status(500).json({ error: publicJob.error });
      return;
    }
    const resultFiles = job.result?.files || [];
    if (resultFiles.length === 1) {
      await sendFileTo(res, resultFiles[0].path, resultFiles[0].name);
      return;
    }
    const zip = await zipResult(job);
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", contentDisposition(archiveName(job)));
    res.send(zip);
  }));

  router.post("/pdf/info", uploadFiles, wrap(async (req, res) => {
    const { files, dir } = collectUploads(req);
    if (!files.length) throw badRequest("没有收到文件。");
    const buffer = await fs.readFile(files[0].path);
    const info = await getPdfInfo(buffer);
    await fs.rm(dir, { recursive: true, force: true }).catch(() => {});
    res.json({ name: files[0].name, ...info });
  }));

  return router;
}

async function zipResult(job) {
  const zip = new JSZip();
  const used = new Set();
  for (const file of job.result.files) {
    let name = file.name;
    let i = 1;
    while (used.has(name)) {
      const ext = path.extname(file.name);
      name = `${path.basename(file.name, ext)}(${i++})${ext}`;
    }
    used.add(name);
    // eslint-disable-next-line no-await-in-loop
    const data = await fs.readFile(file.path).catch(() => null);
    if (data) zip.file(name, data);
  }
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } });
}

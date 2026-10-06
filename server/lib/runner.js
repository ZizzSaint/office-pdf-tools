/** 任务执行器：把上传的输入交给对应的转换/拆分合并实现。 */
import fs from "node:fs/promises";
import path from "node:path";
import { config } from "../config.js";
import { jobStore } from "./jobs.js";
import { log } from "./logger.js";
import { AppError, badRequest } from "./errors.js";
import { ensureDir, removeDir, baseName } from "./paths.js";
import { registerJobFiles } from "./registry.js";
import { officeToPdf } from "./convert/officeToPdf.js";
import { pdfToOffice } from "./convert/pdfToOffice.js";
import { pdfToImages } from "./convert/pdfToImages.js";
import { imagesToPdf } from "./convert/imagesToPdf.js";
import { mergeDocuments } from "./merge/index.js";
import { splitDocument } from "./split/index.js";

export const OPS = ["office-to-pdf", "pdf-to-office", "pdf-to-images", "images-to-pdf", "merge", "split"];

/**
 * @param {any} job 由 jobStore.create 创建的任务对象
 */
export async function executeJob(job) {
  const { id, op, options } = job;
  const progress = (ratio, step) => jobStore.progress(id, ratio, step);
  const warnings = [];
  const outputDir = config.outputDir;
  await ensureDir(outputDir);
  const inputPaths = job.inputs.map((f) => f.path);
  if (!inputPaths.length) throw badRequest("没有收到任何文件。");
  jobStore.update(id, { status: "running", startedAt: Date.now(), step: "开始处理", warnings });

  let files = [];
  const extra = {};

  switch (op) {
    case "office-to-pdf": {
      const engineReport = [];
      for (let i = 0; i < inputPaths.length; i++) {
        const inputPath = inputPaths[i];
        progress((i / inputPaths.length) * 0.9, `转换 ${path.basename(inputPath)}`);
        // eslint-disable-next-line no-await-in-loop
        const result = await officeToPdf({
          inputPath,
          outputDir,
          options,
          progress: (r, s) => progress(((i + r) / inputPaths.length) * 0.95, s),
          signal: job.abort.signal,
          warnings,
        });
        engineReport.push(result.engine);
        // eslint-disable-next-line no-await-in-loop
        const stat = await fs.stat(result.outputPath);
        files.push({ name: path.basename(result.outputPath), path: result.outputPath, kind: "pdf", size: stat.size });
      }
      extra.engines = [...new Set(engineReport)];
      break;
    }

    case "pdf-to-office": {
      const result = await pdfToOffice({
        inputPath: inputPaths[0],
        outputDir,
        options,
        progress,
        signal: job.abort.signal,
        warnings,
      });
      files = result.files;
      extra.mode = result.mode;
      extra.target = result.target;
      break;
    }

    case "pdf-to-images": {
      const result = await pdfToImages({
        inputPath: inputPaths[0],
        outputDir,
        options,
        progress,
        signal: job.abort.signal,
      });
      files = result.files;
      break;
    }

    case "images-to-pdf": {
      const result = await imagesToPdf({ inputPaths, outputDir, options, progress, warnings });
      files = result.files;
      break;
    }

    case "merge": {
      const result = await mergeDocuments({
        inputPaths,
        kind: options.kind,
        outputDir,
        options,
        progress,
        warnings,
      });
      files = result.files;
      break;
    }

    case "split": {
      const result = await splitDocument({
        inputPath: inputPaths[0],
        outputDir,
        kind: options.kind,
        strategy: options.strategy,
        param: options.param,
        progress,
        warnings,
      });
      files = result.files;
      break;
    }

    default:
      throw new AppError(`未知操作: ${op}`, { status: 400, code: "bad-request" });
  }

  if (!files.length) throw new AppError("处理完成但没有生成任何文件。", { status: 500, code: "empty-result" });
  const prepared = await jobStore.prepareResult(job, files, extra);
  registerJobFiles(job, prepared.files);
  prepared.warnings = warnings;
  job.result = prepared;
  job.warnings = warnings;
  jobStore.finish(id, prepared);
  return prepared;
}

/** 后台执行并处理异常与临时文件清理。 */
export function runJobInBackground(job) {
  executeJob(job)
    .catch((err) => {
      jobStore.fail(job.id, err);
      log.debug(`job ${job.id} error`, err?.stack || err);
    })
    .finally(async () => {
      // 结果文件已写入输出目录，上传的输入与中间产物可以立即清理。
      if (!config.keepTemp) await removeDir(job.dir);
    });
}

export function archiveName(job) {
  const first = job.inputs[0]?.name || "result";
  return `${baseName(first)}-${job.op}.zip`;
}

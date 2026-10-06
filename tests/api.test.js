import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { tempDir, makePng, makePdf } from "./helpers/samples.mjs";
import { createApp } from "../server/app.js";
import { config } from "../server/config.js";

async function withServer(fn) {
  const outDir = await tempDir("opt-api-out-");
  const tmpDir = await tempDir("opt-api-tmp-");
  const previousOut = config.outputDir;
  const previousTmp = config.tmpDir;
  config.outputDir = outDir;
  config.tmpDir = tmpDir;
  const app = createApp();
  const server = app.listen(0, "127.0.0.1");
  await new Promise((resolve) => server.once("listening", resolve));
  const port = server.address().port;
  const base = `http://127.0.0.1:${port}`;
  try {
    return await fn({ base, outDir, tmpDir });
  } finally {
    await new Promise((resolve) => server.close(resolve));
    config.outputDir = previousOut;
    config.tmpDir = previousTmp;
  }
}

async function postFiles(base, op, files, options = {}) {
  const form = new FormData();
  for (const [name, buffer, type] of files) {
    form.append("files", new Blob([buffer], { type }), name);
  }
  form.append("op", op);
  form.append("options", JSON.stringify(options));
  const res = await fetch(`${base}/api/jobs`, { method: "POST", body: form });
  return res;
}

async function waitDone(base, jobId, { tries = 200 } = {}) {
  for (let i = 0; i < tries; i++) {
    // eslint-disable-next-line no-await-in-loop
    const res = await fetch(`${base}/api/jobs/${jobId}`);
    // eslint-disable-next-line no-await-in-loop
    const job = await res.json();
    if (["done", "error", "canceled"].includes(job.status)) return job;
    // eslint-disable-next-line no-await-in-loop
    await new Promise((r) => setTimeout(r, 150));
  }
  throw new Error("任务超时");
}

test("API：健康检查/引擎/设置", async () => {
  await withServer(async ({ base }) => {
    const health = await (await fetch(`${base}/api/health`)).json();
    assert.equal(health.ok, true);
    assert.equal(health.name, "office-pdf-tools");
    const engines = await (await fetch(`${base}/api/engines`)).json();
    assert.ok(engines.engines.renderer.available, "PDF 渲染器可用");
    assert.ok(engines.limits.maxUploadMb > 0);
    const settings = await (await fetch(`${base}/api/settings`)).json();
    assert.ok(settings.outputDir);
    const updated = await (await fetch(`${base}/api/settings`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Office-Pdf-Tools": "1" },
      body: JSON.stringify({ outputDir: settings.outputDir }),
    })).json();
    assert.equal(updated.outputDir, settings.outputDir);
  });
});

test("API：写操作需要自定义头（防 CSRF）", async () => {
  await withServer(async ({ base }) => {
    const res = await fetch(`${base}/api/settings`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "User-Agent": "Mozilla/5.0 Chrome/120" },
      body: JSON.stringify({}),
    });
    assert.equal(res.status, 403);
    const body = await res.json();
    assert.equal(body.error.code, "forbidden");
  });
});

test("API：图片 → PDF 任务全流程（上传/轮询/下载/打包/清理）", async () => {
  await withServer(async ({ base, outDir }) => {
    const dir = await tempDir();
    const a = await makePng(path.join(dir, "p1.png"), { label: "ONE" });
    const b = await makePng(path.join(dir, "p2.png"), { label: "TWO" });
    const res = await postFiles(base, "images-to-pdf", [
      ["p1.png", await fs.readFile(a), "image/png"],
      ["p2.png", await fs.readFile(b), "image/png"],
    ], { pageSize: "a4", perPage: 1 });
    assert.equal(res.status, 202);
    const { jobId } = await res.json();
    const job = await waitDone(base, jobId);
    assert.equal(job.status, "done");
    assert.equal(job.progress, 1);
    assert.equal(job.result.files.length, 1);
    const file = job.result.files[0];
    assert.ok(file.path.startsWith(outDir), "文件写入配置的输出目录");

    const download = await fetch(`${base}/api/jobs/${jobId}/result/0`);
    assert.equal(download.status, 200);
    const bytes = Buffer.from(await download.arrayBuffer());
    assert.equal(bytes.slice(0, 5).toString(), "%PDF-");
    assert.match(download.headers.get("content-disposition") || "", /attachment/);

    const archive = await fetch(`${base}/api/jobs/${jobId}/archive`);
    assert.equal(archive.status, 200);
    assert.equal(archive.headers.get("content-type"), "application/zip");

    const list = await (await fetch(`${base}/api/jobs`)).json();
    assert.ok(list.jobs.some((j) => j.id === jobId));

    const del = await fetch(`${base}/api/jobs/${jobId}`, { method: "DELETE" });
    assert.equal(del.status, 200);
    const missing = await fetch(`${base}/api/jobs/${jobId}`);
    assert.equal(missing.status, 404);
  });
});

test("API：PDF → 图片（多文件结果）与同步接口", async () => {
  await withServer(async ({ base }) => {
    const dir = await tempDir();
    const pdf = await makePdf(path.join(dir, "doc.pdf"), { pages: 3 });
    const buffer = await fs.readFile(pdf);

    const res = await postFiles(base, "pdf-to-images", [["doc.pdf", buffer, "application/pdf"]], { dpi: 72, pages: "1-2" });
    const { jobId } = await res.json();
    const job = await waitDone(base, jobId);
    assert.equal(job.result.files.length, 2);
    const img = await fetch(`${base}/api/jobs/${jobId}/result/0?inline=1`);
    assert.equal(img.status, 200);
    assert.match(img.headers.get("content-disposition") || "", /inline/);
    assert.equal(Buffer.from(await img.arrayBuffer()).slice(1, 4).toString(), "PNG");

    // 同步接口直接返回文件
    const syncForm = new FormData();
    syncForm.append("files", new Blob([buffer], { type: "application/pdf" }), "doc.pdf");
    syncForm.append("op", "pdf-to-images");
    syncForm.append("options", JSON.stringify({ dpi: 72, pages: "1", format: "png" }));
    const sync = await fetch(`${base}/api/jobs?sync=1`, { method: "POST", body: syncForm });
    assert.equal(sync.status, 200);
    const syncBytes = Buffer.from(await sync.arrayBuffer());
    assert.equal(syncBytes.slice(1, 4).toString(), "PNG");
  });
});

test("API：文件保存接口写入输出目录并可下载", async () => {
  await withServer(async ({ base, outDir }) => {
    const dir = await tempDir();
    const pdf = await makePdf(path.join(dir, "native.pdf"), { pages: 1 });
    const form = new FormData();
    form.append("files", new Blob([await fs.readFile(pdf)], { type: "application/pdf" }), "当前文档.pdf");
    form.append("options", JSON.stringify({ filename: "当前文档.pdf" }));
    const res = await fetch(`${base}/api/files/save`, {
      method: "POST",
      headers: { "X-Office-Pdf-Tools": "1" },
      body: form,
    });
    assert.equal(res.status, 200);
    const data = await res.json();
    assert.equal(data.files.length, 1);
    assert.ok(data.files[0].name.endsWith(".pdf"));
    const onDisk = await fs.readdir(outDir);
    assert.equal(onDisk.length, 1);
    const dl = await fetch(`${base}/api/files/${data.files[0].downloadId}/download`);
    assert.equal(dl.status, 200);
  });
});

test("API：未知操作与缺失文件返回 4xx", async () => {
  await withServer(async ({ base }) => {
    const form = new FormData();
    form.append("op", "not-an-op");
    const res = await fetch(`${base}/api/jobs`, { method: "POST", body: form });
    assert.equal(res.status, 400);
    const bad = await fetch(`${base}/api/jobs/nope`);
    assert.equal(bad.status, 404);
  });
});

/** 内存任务队列：进度、结果、取消、过期清理。 */
import crypto from "node:crypto";
import path from "node:path";
import fs from "node:fs/promises";
import { EventEmitter } from "node:events";
import { config } from "../config.js";
import { log } from "./logger.js";
import { toAppError } from "./errors.js";
import { ensureDir, removeDir, mimeFor } from "./paths.js";

const jobs = new Map();
export const jobEvents = new EventEmitter();
jobEvents.setMaxListeners(0);

const ACTIVE = new Set(["queued", "running"]);

function publicJob(job) {
  return {
    id: job.id,
    op: job.op,
    status: job.status,
    progress: job.progress,
    step: job.step,
    createdAt: job.createdAt,
    startedAt: job.startedAt,
    finishedAt: job.finishedAt,
    inputs: job.inputs.map((f) => ({ name: f.name, size: f.size })),
    options: job.options,
    warnings: job.warnings,
    error: job.error,
    result: job.result,
  };
}

export const jobStore = {
  create({ op, inputs, options, dir }) {
    const id = crypto.randomUUID().slice(0, 8) + Date.now().toString(36).slice(-4);
    const job = {
      id,
      op,
      options: options || {},
      inputs,
      status: "queued",
      progress: 0,
      step: "排队中",
      warnings: [],
      createdAt: Date.now(),
      startedAt: null,
      finishedAt: null,
      error: null,
      result: null,
      abort: new AbortController(),
      dir: dir || path.join(config.tmpDir, id),
    };
    jobs.set(id, job);
    jobEvents.emit("job", publicJob(job));
    return job;
  },

  get(id) {
    return jobs.get(id);
  },

  getPublic(id) {
    const job = jobs.get(id);
    return job ? publicJob(job) : null;
  },

  list() {
    return [...jobs.values()].sort((a, b) => b.createdAt - a.createdAt).map(publicJob);
  },

  update(id, patch) {
    const job = jobs.get(id);
    if (!job) return null;
    Object.assign(job, patch);
    jobEvents.emit("job", publicJob(job));
    return job;
  },

  progress(id, ratio, step) {
    const job = jobs.get(id);
    if (!job || job.status === "canceled") return;
    job.status = "running";
    job.progress = Math.max(job.progress, Math.min(1, Number(ratio) || 0));
    if (step) job.step = step;
    jobEvents.emit("job", publicJob(job));
  },

  warn(id, message) {
    const job = jobs.get(id);
    if (!job || !message) return;
    job.warnings.push(String(message));
  },

  finish(id, result) {
    const job = jobs.get(id);
    if (!job) return null;
    job.status = "done";
    job.progress = 1;
    job.step = "完成";
    job.finishedAt = Date.now();
    job.result = result;
    jobEvents.emit("job", publicJob(job));
    log.info(`job ${id} ${job.op} done in ${job.finishedAt - (job.startedAt || job.createdAt)}ms`);
    return job;
  },

  fail(id, error) {
    const job = jobs.get(id);
    if (!job) return null;
    const appErr = toAppError(error);
    job.status = "error";
    job.finishedAt = Date.now();
    job.error = { code: appErr.code, message: appErr.message, hint: appErr.hint };
    jobEvents.emit("job", publicJob(job));
    log.warn(`job ${id} ${job.op} failed: ${appErr.message}`);
    return job;
  },

  cancel(id) {
    const job = jobs.get(id);
    if (!job) return null;
    if (!ACTIVE.has(job.status)) return job;
    job.status = "canceled";
    job.step = "已取消";
    job.finishedAt = Date.now();
    try {
      job.abort.abort();
    } catch { /* ignore */ }
    jobEvents.emit("job", publicJob(job));
    return job;
  },

  async remove(id) {
    const job = jobs.get(id);
    if (!job) return false;
    this.cancel(id);
    jobs.delete(id);
    if (!config.keepTemp) await removeDir(job.dir);
    return true;
  },

  isActive(id) {
    const job = jobs.get(id);
    return !!job && ACTIVE.has(job.status);
  },

  /** 打包任务结果文件（多文件时使用）。 */
  async prepareResult(job, files, extra = {}) {
    const out = [];
    for (const f of files) {
      const stat = await fs.stat(f.path).catch(() => null);
      out.push({
        name: f.name,
        path: f.path,
        size: stat?.size ?? 0,
        mime: f.mime || mimeFor(f.name),
        kind: f.kind || "file",
      });
    }
    return { files: out, outputDir: config.outputDir, ...extra };
  },

  async workDir(job) {
    await ensureDir(job.dir);
    return job.dir;
  },

  async cleanupExpired() {
    const ttl = config.jobTtlMinutes * 60 * 1000;
    const now = Date.now();
    for (const [id, job] of jobs) {
      if (ACTIVE.has(job.status)) continue;
      if (now - (job.finishedAt || job.createdAt) > ttl) {
        jobs.delete(id);
        if (!config.keepTemp) await removeDir(job.dir);
      }
    }
  },
};

export { publicJob };

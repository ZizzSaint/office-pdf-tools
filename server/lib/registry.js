/** 已生成文件的短期登记表，供 HTTP 下载使用（避免通过路径直接访问磁盘）。 */
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

const files = new Map();
const TTL = 6 * 60 * 60 * 1000;

export function registerFile(absPath, meta = {}) {
  const id = crypto.randomUUID().replace(/-/g, "").slice(0, 16);
  files.set(id, { id, path: absPath, name: meta.name || path.basename(absPath), at: Date.now() });
  return id;
}

export function getFile(id) {
  const entry = files.get(id);
  if (!entry) return null;
  if (Date.now() - entry.at > TTL) {
    files.delete(id);
    return null;
  }
  return entry;
}

export function registerJobFiles(job, resultFiles) {
  for (const file of resultFiles) {
    file.downloadId = registerFile(file.path, { name: file.name });
  }
  return resultFiles;
}

export async function cleanupRegistry() {
  const now = Date.now();
  for (const [id, entry] of files) {
    if (now - entry.at > TTL) files.delete(id);
  }
}

export async function fileStillExists(entry) {
  try {
    await fs.access(entry.path);
    return true;
  } catch {
    return false;
  }
}

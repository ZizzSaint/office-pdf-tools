/** 文件路径工具：安全文件名、唯一命名、目录准备。 */
import fs from "node:fs/promises";
import fssync from "node:fs";
import path from "node:path";
import crypto from "node:crypto";

/** 去掉路径分隔符与非法字符，保留可读的中文/空格。 */
export function safeName(name, fallback = "file") {
  const base = String(name ?? "").replace(/[\u0000-\u001f]/g, "");
  const cleaned = base
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/^\.+/, "")
    .replace(/[. ]+$/, "")
    .trim();
  return cleaned || fallback;
}

export function baseName(name) {
  const n = safeName(path.basename(String(name ?? "")));
  return n.replace(/\.[^.]+$/, "");
}

/** 目录 + 文件名组合：取文件名并保留扩展名。 */
export function fileName(name, fallback = "file") {
  return safeName(path.basename(String(name ?? "")), fallback);
}

export function extName(name) {
  const m = /\.([^.]+)$/.exec(String(name || ""));
  return m ? m[1].toLowerCase() : "";
}

export function withExt(name, ext) {
  const base = baseName(name);
  return ext ? `${base}.${ext.replace(/^\./, "")}` : base;
}

export async function ensureDir(dir) {
  await fs.mkdir(dir, { recursive: true });
  return dir;
}

export function ensureDirSync(dir) {
  fssync.mkdirSync(dir, { recursive: true });
  return dir;
}

/** 在 dir 下生成不冲突的文件名：report.pdf → report (1).pdf */
export async function uniquePath(dir, filename) {
  const safe = safeName(filename);
  const ext = path.extname(safe);
  const stem = path.basename(safe, ext);
  let candidate = path.join(dir, safe);
  let i = 1;
  // eslint-disable-next-line no-await-in-loop
  while (await exists(candidate)) {
    candidate = path.join(dir, `${stem} (${i})${ext}`);
    i++;
    if (i > 9999) {
      candidate = path.join(dir, `${stem}-${crypto.randomUUID().slice(0, 8)}${ext}`);
      break;
    }
  }
  return candidate;
}

export function uniquePathSync(dir, filename) {
  const safe = safeName(filename);
  const ext = path.extname(safe);
  const stem = path.basename(safe, ext);
  let candidate = path.join(dir, safe);
  let i = 1;
  while (fssync.existsSync(candidate)) {
    candidate = path.join(dir, `${stem} (${i})${ext}`);
    i++;
    if (i > 9999) {
      candidate = path.join(dir, `${stem}-${crypto.randomUUID().slice(0, 8)}${ext}`);
      break;
    }
  }
  return candidate;
}

export async function exists(p) {
  try {
    await fs.access(p);
    return true;
  } catch {
    return false;
  }
}

export async function dirSize(dir) {
  let total = 0;
  try {
    for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) total += await dirSize(full);
      else total += (await fs.stat(full)).size;
    }
  } catch { /* ignore */ }
  return total;
}

export async function removeDir(dir) {
  try {
    await fs.rm(dir, { recursive: true, force: true });
  } catch { /* ignore */ }
}

/** 把文件名转成 zip 内部安全的相对名。 */
export function zipEntryName(name) {
  return safeName(name).replace(/\\/g, "_");
}

export function mimeFor(fileName) {
  const ext = extName(fileName);
  const map = {
    pdf: "application/pdf",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    webp: "image/webp",
    gif: "image/gif",
    zip: "application/zip",
    txt: "text/plain; charset=utf-8",
    md: "text/markdown; charset=utf-8",
  };
  return map[ext] || "application/octet-stream";
}

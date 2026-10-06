/** 上传中间件（multer 磁盘存储）。 */
import crypto from "node:crypto";
import path from "node:path";
import multer from "multer";
import { config } from "../config.js";
import { ensureDirSync, safeName } from "../lib/paths.js";
import { badRequest } from "../lib/errors.js";

function sessionDir(req) {
  if (!req._uploadDir) {
    const id = crypto.randomUUID().replace(/-/g, "").slice(0, 12);
    req._uploadDir = path.join(config.tmpDir, "incoming", id);
    ensureDirSync(req._uploadDir);
  }
  return req._uploadDir;
}

const storage = multer.diskStorage({
  destination(req, file, cb) {
    try {
      cb(null, sessionDir(req));
    } catch (err) {
      cb(err);
    }
  },
  filename(req, file, cb) {
    const name = safeName(file.originalname, "upload");
    const unique = `${Date.now().toString(36)}-${crypto.randomUUID().slice(0, 4)}-${name}`;
    cb(null, unique.slice(-160));
  },
});

export const uploadFiles = multer({
  storage,
  limits: {
    fileSize: config.maxUploadBytes,
    files: 100,
    fieldSize: 2 * 1024 * 1024,
  },
  fileFilter(req, file, cb) {
    const name = (file.originalname || "").toLowerCase();
    if (name.endsWith(".zip") && !req.path.includes("archive")) {
      cb(badRequest("请不要上传 zip，请直接上传需要处理的文档。"));
      return;
    }
    cb(null, true);
  },
}).array("files", 100);

export function collectUploads(req) {
  const files = (req.files || []).map((f) => ({
    name: Buffer.from(f.originalname, "latin1").toString("utf8") || f.originalname,
    path: f.path,
    size: f.size,
  }));
  return { files, dir: req._uploadDir };
}

export function parseOptions(req) {
  const raw = req.body?.options;
  if (!raw) return {};
  if (typeof raw === "object") return raw;
  try {
    return JSON.parse(raw);
  } catch {
    throw badRequest("options 不是合法的 JSON。");
  }
}

/** 合并入口：按文件类型分派。 */
import path from "node:path";
import { AppError } from "../errors.js";
import { extName } from "../paths.js";
import { mergePdf } from "./pdf.js";
import { mergeDocx } from "./docx.js";
import { mergeXlsx } from "./xlsx.js";
import { mergePptx } from "./pptx.js";

const KIND_BY_EXT = {
  pdf: "pdf",
  docx: "docx", docm: "docx", doc: "docx",
  xlsx: "xlsx", xlsm: "xlsx", xls: "xlsx",
  pptx: "pptx", pptm: "pptx", ppt: "pptx",
};

export function detectKind(inputPaths) {
  const kinds = new Set(inputPaths.map((p) => KIND_BY_EXT[extName(p)]).filter(Boolean));
  if (!kinds.size) return null;
  if (kinds.size > 1) return null;
  return [...kinds][0];
}

export async function mergeDocuments({ inputPaths, kind, outputDir, options = {}, progress, warnings = [] }) {
  const resolvedKind = kind && kind !== "auto" ? kind : detectKind(inputPaths);
  if (!resolvedKind) {
    throw new AppError("无法识别文件类型（或类型不一致），请手动指定要合并的格式。", { status: 400, code: "bad-request" });
  }
  const cfg = { inputPaths, outputDir, options, progress, warnings };
  switch (resolvedKind) {
    case "pdf": return mergePdf(cfg);
    case "docx": return mergeDocx(cfg);
    case "xlsx": return mergeXlsx(cfg);
    case "pptx": return mergePptx(cfg);
    default:
      throw new AppError(`暂不支持的类型: ${resolvedKind}`, { status: 415, code: "unsupported-format" });
  }
}

export { mergePdf, mergeDocx, mergeXlsx, mergePptx };

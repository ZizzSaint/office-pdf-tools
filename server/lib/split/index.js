/** 拆分入口：按文件类型分派。 */
import { AppError } from "../errors.js";
import { extName } from "../paths.js";
import { splitPdf } from "./pdf.js";
import { splitDocx } from "./docx.js";
import { splitXlsx } from "./xlsx.js";
import { splitPptx } from "./pptx.js";
import { detectKind } from "../merge/index.js";

export const SPLIT_STRATEGIES = {
  pdf: ["every-n-pages", "ranges", "each-page"],
  docx: ["heading", "section"],
  xlsx: ["sheets", "every-n-rows"],
  pptx: ["every-n-slides", "ranges"],
};

export function normalizeStrategy(kind, strategy) {
  const allowed = SPLIT_STRATEGIES[kind] || [];
  if (!strategy) return allowed[0];
  if (allowed.includes(strategy)) return strategy;
  // 跨类型的同义策略
  const aliases = {
    "each-page": { pdf: "each-page", pptx: "every-n-slides" },
    "every-n-pages": { pdf: "every-n-pages", pptx: "every-n-slides", docx: "section" },
    "every-n-slides": { pptx: "every-n-slides", pdf: "every-n-pages" },
    sheets: { xlsx: "sheets" },
  };
  const mapped = aliases[strategy]?.[kind];
  if (mapped) return mapped;
  throw new AppError(`${kind} 不支持拆分方式 ${strategy}。`, { status: 400, code: "bad-request" });
}

export async function splitDocument({ inputPath, outputDir, kind, strategy, param, progress, warnings = [] }) {
  const resolvedKind = kind && kind !== "auto" ? kind : detectKind([inputPath]);
  if (!resolvedKind) {
    throw new AppError("无法识别文件类型，暂不支持拆分该格式。", { status: 415, code: "unsupported-format" });
  }
  const useStrategy = normalizeStrategy(resolvedKind, strategy);
  const cfg = { inputPath, outputDir, strategy: useStrategy, param, progress, warnings };
  switch (resolvedKind) {
    case "pdf": return splitPdf(cfg);
    case "docx": return splitDocx(cfg);
    case "xlsx": return splitXlsx(cfg);
    case "pptx": return splitPptx(cfg);
    default:
      throw new AppError(`暂不支持拆分 ${extName(inputPath)} 文件。`, { status: 415, code: "unsupported-format" });
  }
}

export { splitPdf, splitDocx, splitXlsx, splitPptx };

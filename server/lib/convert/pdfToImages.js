/** PDF → 图片（逐页渲染）。 */
import fs from "node:fs/promises";
import path from "node:path";
import { openPdf, closePdf, renderPageToBuffer } from "../engines/render.js";
import { parsePageList } from "../ranges.js";
import { AppError } from "../errors.js";
import { baseName, uniquePath } from "../paths.js";

const FORMAT_EXT = { jpg: "jpeg", jpeg: "jpeg", png: "png", webp: "webp" };

/**
 * @param {{inputPath:string, outputDir:string, options?:any,
 *          progress?:(r:number,s?:string)=>void, signal?:AbortSignal}} cfg
 */
export async function pdfToImages({ inputPath, outputDir, options = {}, progress, signal }) {
  const buffer = await fs.readFile(inputPath);
  const format = FORMAT_EXT[String(options.format || "png").toLowerCase()] || "png";
  const dpi = Math.max(36, Math.min(1200, Number(options.dpi) || 150));
  const doc = await openPdf(buffer);
  const files = [];
  try {
    const pages = parsePageList(options.pages, doc.numPages);
    if (!pages.length) throw new AppError("页码范围为空。", { status: 400, code: "bad-request" });
    const stem = options.outputName ? baseName(options.outputName) : baseName(inputPath);
    const pad = String(doc.numPages).length;
    for (let i = 0; i < pages.length; i++) {
      if (signal?.aborted) throw new AppError("任务已取消。", { code: "canceled", status: 499 });
      const pageNumber = pages[i] + 1;
      // eslint-disable-next-line no-await-in-loop
      const png = await renderPageToBuffer(doc, pageNumber, { dpi, format, quality: options.quality });
      const suffix = pages.length > 1 ? `-p${String(pageNumber).padStart(pad, "0")}` : "";
      // eslint-disable-next-line no-await-in-loop
      const target = await uniquePath(outputDir, `${stem}${suffix}.${format === "jpeg" ? "jpg" : format}`);
      // eslint-disable-next-line no-await-in-loop
      await fs.writeFile(target, png);
      files.push({ name: path.basename(target), path: target, kind: "image" });
      progress?.(0.05 + 0.9 * ((i + 1) / pages.length), `渲染第 ${pageNumber} 页 / 共 ${doc.numPages} 页`);
    }
  } finally {
    await closePdf(doc);
  }
  return { files };
}

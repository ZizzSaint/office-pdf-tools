/** PDF 渲染与文本抽取：pdf.js + @napi-rs/canvas。 */
import path from "node:path";
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";
import { AppError, engineMissing } from "../errors.js";

const require = createRequire(import.meta.url);

let pdfjsPromise;
let canvasPromise;

export async function loadPdfjs() {
  if (!pdfjsPromise) {
    pdfjsPromise = import("pdfjs-dist/legacy/build/pdf.mjs").catch((err) => {
      pdfjsPromise = null;
      throw engineMissing("缺少 pdfjs-dist，请运行 npm install。", { cause: err });
    });
  }
  return pdfjsPromise;
}

export async function loadCanvas() {
  if (!canvasPromise) {
    canvasPromise = import("@napi-rs/canvas").catch((err) => {
      canvasPromise = null;
      throw engineMissing("缺少 @napi-rs/canvas，无法渲染 PDF 页面。", { cause: err });
    });
  }
  return canvasPromise;
}

function standardFontDir() {
  try {
    const dir = path.join(path.dirname(require.resolve("pdfjs-dist/package.json")), "standard_fonts");
    // pdf.js 要求以斜杠结尾的 URL（Windows 盘符路径会校验失败）。
    return pathToFileURL(dir).href + "/";
  } catch {
    return undefined;
  }
}

/** 打开 PDF 文档。调用方负责 close()。 */
export async function openPdf(buffer) {
  const pdfjs = await loadPdfjs();
  // pdf.js 会把底层 ArrayBuffer 转移给 worker（原对象随即失效），所以这里始终复制一份。
  const src = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  const data = new Uint8Array(src);
  try {
    return await pdfjs.getDocument({
      data,
      useSystemFonts: true,
      isEvalSupported: false,
      disableFontFace: true,
      standardFontDataUrl: standardFontDir(),
      verbosity: 0,
    }).promise;
  } catch (err) {
    throw new AppError("PDF 文件无法解析，可能已损坏或被加密。", {
      code: "invalid-document",
      status: 400,
      cause: err,
    });
  }
}

/** 关闭文档（兼容 pdf.js v4/v6 的不同 API）。 */
export async function closePdf(doc) {
  try {
    if (typeof doc?.destroy === "function") return await doc.destroy();
    if (doc?.loadingTask && typeof doc.loadingTask.destroy === "function") return await doc.loadingTask.destroy();
    if (typeof doc?.cleanup === "function") return doc.cleanup();
  } catch { /* ignore */ }
  return undefined;
}

export async function getPdfInfo(buffer) {
  const doc = await openPdf(buffer);
  try {
    const sizes = [];
    for (let i = 1; i <= doc.numPages; i++) {
      const page = await doc.getPage(i);
      const viewport = page.getViewport({ scale: 1 });
      sizes.push({ width: viewport.width, height: viewport.height });
      page.cleanup();
    }
    return { pages: doc.numPages, pageSizes: sizes };
  } finally {
    await closePdf(doc);
  }
}

const FORMAT_ALIASES = { jpg: "jpeg", jpe: "jpeg", jpeg: "jpeg", png: "png", webp: "webp" };

/**
 * 渲染单页为图片 Buffer。
 * @param {any} doc pdf.js 文档
 * @param {number} pageNumber 1 基
 * @param {{dpi?:number, format?:string, quality?:number, background?:string}} [opts]
 */
export async function renderPageToBuffer(doc, pageNumber, opts = {}) {
  const { createCanvas } = await loadCanvas();
  const dpi = Math.max(36, Math.min(1200, Number(opts.dpi) || 150));
  const scale = dpi / 72;
  const format = FORMAT_ALIASES[String(opts.format || "png").toLowerCase()] || "png";
  const quality = Math.max(1, Math.min(100, Number(opts.quality) || 82));

  const page = await doc.getPage(pageNumber);
  const viewport = page.getViewport({ scale });
  const width = Math.max(1, Math.ceil(viewport.width));
  const height = Math.max(1, Math.ceil(viewport.height));
  const canvas = createCanvas(width, height);
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = opts.background || "#ffffff";
  ctx.fillRect(0, 0, width, height);

  await page.render({ canvasContext: ctx, viewport, intent: "display" }).promise;
  page.cleanup();

  if (format === "jpeg") return canvas.encode("jpeg", quality);
  if (format === "webp") return canvas.encode("webp", quality);
  return canvas.encode("png");
}

/**
 * 抽取每页文本（带坐标），用于 PDF → Word / Excel(文本模式)。
 * @returns {Promise<Array<{pageNumber:number,width:number,height:number,items:Array<any>}>>}
 */
export async function extractTextByPage(buffer, pageNumbers = null) {
  const doc = await openPdf(buffer);
  try {
    const pages = pageNumbers?.length
      ? pageNumbers.map((p) => p + 1)
      : Array.from({ length: doc.numPages }, (_, i) => i + 1);
    const out = [];
    for (const pageNumber of pages) {
      const page = await doc.getPage(pageNumber);
      const viewport = page.getViewport({ scale: 1 });
      const content = await page.getTextContent();
      const items = [];
      for (const item of content.items) {
        // 空字符串是 pdf.js 的换行标记；纯空白项保留为“这里有空格”的提示。
        if (!item.str) continue;
        const tr = item.transform || [1, 0, 0, 1, 0, 0];
        const fontSize = Math.hypot(tr[2], tr[3]) || Math.abs(tr[3]) || 10;
        items.push({
          text: item.str,
          x: tr[4],
          y: tr[5],
          width: item.width || 0,
          height: item.height || fontSize,
          fontSize,
          fontName: item.fontName || "",
          bold: /bold|black|heavy|semib/i.test(item.fontName || ""),
          italic: /italic|oblique/i.test(item.fontName || ""),
          isSpace: !item.str.trim().length,
        });
      }
      out.push({ pageNumber, width: viewport.width, height: viewport.height, items });
      page.cleanup();
    }
    return out;
  } finally {
    await closePdf(doc);
  }
}

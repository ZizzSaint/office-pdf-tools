/** 图片 → PDF：多图合并、N-up 排版、每图单独 PDF。 */
import fs from "node:fs/promises";
import path from "node:path";
import {
  PDFDocument, pushGraphicsState, popGraphicsState, moveTo, lineTo, closePath, clip, endPath,
} from "pdf-lib";
import { loadCanvas } from "../engines/render.js";
import { AppError } from "../errors.js";
import { baseName, extName, uniquePath } from "../paths.js";

const MM_TO_PT = 72 / 25.4;
export const PAGE_SIZES = {
  a3: [841.89, 1190.55],
  a4: [595.28, 841.89],
  a5: [419.53, 595.28],
  letter: [612, 792],
  legal: [612, 1008],
  tabloid: [792, 1224],
};

const GRID_PORTRAIT = {
  1: [1, 1], 2: [1, 2], 4: [2, 2], 6: [2, 3], 9: [3, 3],
};

const IMAGE_EXTS = new Set(["png", "jpg", "jpeg", "webp", "gif", "bmp", "tif", "tiff", "avif"]);

export function isImageFile(filePath) {
  return IMAGE_EXTS.has(extName(filePath));
}

/** 读入图片并转成 pdf-lib 可嵌入的对象。 */
async function embedImage(pdf, filePath) {
  const ext = extName(filePath);
  const bytes = await fs.readFile(filePath);
  try {
    if (ext === "jpg" || ext === "jpeg") return await pdf.embedJpg(bytes);
    if (ext === "png") return await pdf.embedPng(bytes);
  } catch {
    // 有些 .jpg/.png 实际是其它编码，继续走 canvas 兜底
  }
  const { loadImage } = await loadCanvas();
  const image = await loadImage(bytes);
  const { createCanvas } = await loadCanvas();
  const canvas = createCanvas(image.width, image.height);
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, image.width, image.height);
  ctx.drawImage(image, 0, 0);
  const png = await canvas.encode("png");
  return pdf.embedPng(png);
}

function resolvePageSize(opts, image, pageIndex) {
  const sizeName = String(opts.pageSize || "fit").toLowerCase();
  if (sizeName === "fit" || sizeName === "image") {
    const dpi = Math.max(36, Math.min(600, Number(opts.imageDpi) || 96));
    const factor = 72 / dpi;
    return [Math.max(1, image.width * factor), Math.max(1, image.height * factor)];
  }
  const base = PAGE_SIZES[sizeName] || PAGE_SIZES.a4;
  const [shortSide, longSide] = [Math.min(...base), Math.max(...base)];
  const orientation = String(opts.orientation || "auto");
  const isLandscape = orientation === "landscape"
    || (orientation === "auto" && opts.perPage <= 1 && image.width > image.height);
  return isLandscape ? [longSide, shortSide] : [shortSide, longSide];
}

function drawImageInBox(page, image, box, fitMode, radiusClip) {
  const { x, y, width, height } = box;
  const iw = image.width;
  const ih = image.height;
  let w = width;
  let h = height;
  if (fitMode !== "stretch") {
    const scale = fitMode === "cover"
      ? Math.max(width / iw, height / ih)
      : Math.min(width / iw, height / ih);
    w = iw * scale;
    h = ih * scale;
  }
  const drawX = x + (width - w) / 2;
  const drawY = y + (height - h) / 2;
  if (radiusClip && fitMode === "cover") {
    try {
      page.pushOperators(
        pushGraphicsState(),
        moveTo(x, y),
        lineTo(x + width, y),
        lineTo(x + width, y + height),
        lineTo(x, y + height),
        closePath(),
        clip(),
        endPath(),
      );
    } catch { /* 低层算子不可用时退化为不裁剪 */ }
  }
  page.drawImage(image, { x: drawX, y: drawY, width: w, height: h });
  if (radiusClip && fitMode === "cover") {
    try {
      page.pushOperators(popGraphicsState());
    } catch { /* ignore */ }
  }
}

/**
 * @param {{inputPaths:string[], outputDir:string, options?:any,
 *          progress?:(r:number, s?:string)=>void, warnings?:string[]}} cfg
 */
export async function imagesToPdf({ inputPaths, outputDir, options = {}, progress, warnings = [] }) {
  const opts = {
    pageSize: options.pageSize || "a4",
    orientation: options.orientation || "auto",
    margin: Number.isFinite(Number(options.margin)) ? Number(options.margin) : 10,
    perPage: Math.max(1, Math.min(16, Number(options.perPage) || 1)),
    fit: options.fit || "contain",
    separate: !!options.separate,
    imageDpi: Number(options.imageDpi) || 96,
    outputName: options.outputName,
  };
  if (!inputPaths.length) throw new AppError("没有可用的图片文件。", { status: 400, code: "bad-request" });

  const files = [];
  // separate = 每张图片一个 PDF；否则所有图片进入同一个 PDF，按 perPage 排版到多页。
  const groups = opts.separate
    ? inputPaths.map((p) => [p])
    : [inputPaths];

  for (let g = 0; g < groups.length; g++) {
    const group = groups[g];
    const pdf = await PDFDocument.create();
    pdf.setProducer("Office PDF Tools");
    pdf.setCreator("Office PDF Tools");
    if (opts.title) pdf.setTitle(String(opts.title));
    pdf.setCreationDate(new Date());

    const rowMap = [];
    for (const filePath of group) {
      // eslint-disable-next-line no-await-in-loop
      const image = await embedImage(pdf, filePath);
      rowMap.push({ filePath, image });
    }

    let pageIndex = 0;
    for (let i = 0; i < rowMap.length; i += opts.perPage) {
      const slice = rowMap.slice(i, i + opts.perPage);
      const first = slice[0].image;
      const [pageW, pageH] = resolvePageSize(opts, first, pageIndex);
      const page = pdf.addPage([pageW, pageH]);
      const margin = Math.max(0, opts.margin) * MM_TO_PT;
      const contentW = Math.max(1, pageW - margin * 2);
      const contentH = Math.max(1, pageH - margin * 2);
      const n = slice.length;
      let cols;
      let rows;
      if (n === 1) {
        cols = 1; rows = 1;
      } else {
        const portraitGrid = GRID_PORTRAIT[opts.perPage] || [Math.ceil(Math.sqrt(n)), Math.ceil(n / Math.ceil(Math.sqrt(n)))];
        const landscape = pageW > pageH;
        cols = landscape ? portraitGrid[1] : portraitGrid[0];
        rows = landscape ? portraitGrid[0] : portraitGrid[1];
      }
      const cellW = contentW / cols;
      const cellH = contentH / rows;
      slice.forEach((entry, idx) => {
        const r = Math.floor(idx / cols);
        const c = idx % cols;
        const box = {
          x: margin + c * cellW,
          y: pageH - margin - (r + 1) * cellH,
          width: cellW,
          height: cellH,
        };
        drawImageInBox(page, entry.image, box, opts.fit, true);
      });
      pageIndex++;
      progress?.(0.1 + 0.8 * ((g + i / Math.max(1, rowMap.length)) / groups.length), `排版第 ${g + 1} 组…`);
    }

    const stem = opts.outputName ? baseName(opts.outputName) : baseName(group[0]);
    const baseNameHint = opts.separate
      ? stem
      : g === 0 ? stem : `${stem}-${g + 1}`;
    const target = await uniquePath(outputDir, `${baseNameHint}.pdf`);
    const bytes = await pdf.save();
    // eslint-disable-next-line no-await-in-loop
    await fs.writeFile(target, bytes);
    files.push({ name: path.basename(target), path: target, kind: "pdf" });
  }

  progress?.(1, "完成");
  return { files };
}

function chunk(list, size) {
  if (size <= 1) return list.map((item) => [item]);
  const out = [];
  for (let i = 0; i < list.length; i += size) out.push(list.slice(i, i + size));
  return out;
}

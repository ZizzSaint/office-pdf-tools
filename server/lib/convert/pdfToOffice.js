/** PDF → Office：文本模式（可编辑）与图片模式（高保真）。 */
import fs from "node:fs/promises";
import path from "node:path";
import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ImageRun, PageBreak, AlignmentType,
} from "docx";
import ExcelJS from "exceljs";
import PptxGenJS from "pptxgenjs";
import { openPdf, closePdf, renderPageToBuffer, extractTextByPage, getPdfInfo } from "../engines/render.js";
import { buildLines, detectBlocks, linesToMatrix } from "../pdf/layout.js";
import { parsePageList } from "../ranges.js";
import { AppError } from "../errors.js";
import { baseName, uniquePath } from "../paths.js";

const PT_TO_TWIP = 20;
const PT_TO_INCH = 1 / 72;
const PT_TO_PX = 96 / 72;

const HEADINGS = [
  HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3,
  HeadingLevel.HEADING_4, HeadingLevel.HEADING_5, HeadingLevel.HEADING_6,
];

/**
 * @param {{inputPath:string, outputDir:string, options?:any,
 *          progress?:(r:number,s?:string)=>void, signal?:AbortSignal, warnings?:string[]}} cfg
 */
export async function pdfToOffice(cfg) {
  const { inputPath, outputDir, options = {}, progress, signal, warnings = [] } = cfg;
  const target = ["docx", "xlsx", "pptx"].includes(options.target) ? options.target : "docx";
  const requestedMode = options.mode === "raster" ? "raster" : "text";
  const dpi = Math.max(72, Math.min(600, Number(options.dpi) || 150));

  const buffer = await fs.readFile(inputPath);
  const info = await getPdfInfo(buffer);
  const pages = parsePageList(options.pages, info.pages);

  let mode = requestedMode;
  if (target === "xlsx" && mode === "raster") {
    warnings.push("Excel 不支持整页图片模式，已改用文本表格模式。");
    mode = "text";
  }

  const stem = options.outputName ? baseName(options.outputName) : baseName(inputPath);
  const ext = target;
  const outPath = await uniquePath(outputDir, `${stem}.${ext}`);

  if (mode === "raster") {
    progress?.(0.1, "渲染页面为图片…");
    const images = await renderPages(buffer, pages, dpi, progress, signal);
    if (target === "docx") await writeRasterDocx(outPath, images, info, pages);
    else await writeRasterPptx(outPath, images, info, pages);
  } else {
    progress?.(0.1, "提取文本与版面…");
    const pageItems = await extractTextByPage(buffer, pages);
    const pageBlocks = pageItems.map((page) => {
      const lines = buildLines(page.items);
      return { page, lines, blocks: detectBlocks(lines) };
    });
    const totalChars = pageBlocks.reduce((sum, p) => sum + p.lines.reduce((s, l) => s + l.text.length, 0), 0);
    if (totalChars < 8) {
      warnings.push("PDF 中没有可提取的文本（可能是扫描件）。建议改用“图片（高保真）”模式。");
    }
    if (target === "docx") await writeTextDocx(outPath, pageBlocks);
    else if (target === "xlsx") await writeTextXlsx(outPath, pageBlocks);
    else await writeTextPptx(outPath, pageBlocks);
  }

  progress?.(1, "完成");
  return { files: [{ name: path.basename(outPath), path: outPath, kind: ext }], mode, target };
}

async function renderPages(buffer, pages, dpi, progress, signal) {
  const doc = await openPdf(buffer);
  const out = [];
  try {
    for (let i = 0; i < pages.length; i++) {
      if (signal?.aborted) throw new AppError("任务已取消。", { code: "canceled", status: 499 });
      const pageNumber = pages[i] + 1;
      // eslint-disable-next-line no-await-in-loop
      const png = await renderPageToBuffer(doc, pageNumber, { dpi, format: "png" });
      const page = await doc.getPage(pageNumber);
      const viewport = page.getViewport({ scale: 1 });
      page.cleanup();
      out.push({ pageNumber, png, widthPt: viewport.width, heightPt: viewport.height });
      progress?.(0.1 + 0.85 * ((i + 1) / pages.length), `渲染第 ${pageNumber} 页…`);
    }
  } finally {
    await closePdf(doc);
  }
  return out;
}

/* --------------------------------------------------------------- 图片模式 */

async function writeRasterDocx(outPath, images, info, pages) {
  const sections = images.map((img) => ({
    properties: {
      page: {
        size: {
          width: Math.round(img.widthPt * PT_TO_TWIP),
          height: Math.round(img.heightPt * PT_TO_TWIP),
        },
        margin: { top: 0, right: 0, bottom: 0, left: 0 },
      },
    },
    children: [
      new Paragraph({
        spacing: { before: 0, after: 0, line: 240 },
        children: [
          new ImageRun({
            type: "png",
            data: img.png,
            transformation: {
              width: Math.round(img.widthPt * PT_TO_PX),
              height: Math.round(img.heightPt * PT_TO_PX),
            },
          }),
        ],
      }),
    ],
  }));
  const doc = new Document({ creator: "Office PDF Tools", title: baseName(outPath), sections });
  const buffer = await Packer.toBuffer(doc);
  await fs.writeFile(outPath, buffer);
}

async function writeRasterPptx(outPath, images, info) {
  const pptx = new PptxGenJS();
  const width = images[0].widthPt * PT_TO_INCH;
  const height = images[0].heightPt * PT_TO_INCH;
  pptx.defineLayout({ name: "PDF_PAGE", width, height });
  pptx.layout = "PDF_PAGE";
  pptx.author = "Office PDF Tools";
  for (const img of images) {
    const slide = pptx.addSlide();
    const w = img.widthPt * PT_TO_INCH;
    const h = img.heightPt * PT_TO_INCH;
    slide.addImage({ data: `image/png;base64,${img.png.toString("base64")}`, x: 0, y: 0, w, h });
  }
  const data = await pptx.write({ outputType: "nodebuffer" });
  await fs.writeFile(outPath, data);
}

/* --------------------------------------------------------------- 文本模式 */

function blocksToDocxChildren(blocks, { pageBreakAfter = false } = {}) {
  const children = [];
  for (const block of blocks) {
    if (block.type === "heading") {
      children.push(new Paragraph({
        heading: HEADINGS[Math.min(6, Math.max(1, block.level)) - 1],
        spacing: { before: 160, after: 80 },
        children: [new TextRun({ text: block.text, bold: true })],
      }));
      continue;
    }
    if (block.type === "table") {
      const rows = block.rows.map((cells) => new TableRow({
        children: cells.map((cellText) => new TableCell({
          children: [new Paragraph({ children: [new TextRun({ text: cellText })] })],
        })),
      }));
      children.push(new Table({
        rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
      }));
      children.push(new Paragraph({ text: "" }));
      continue;
    }
    children.push(new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { after: 100 },
      children: [new TextRun({ text: block.text || "" })],
    }));
  }
  if (pageBreakAfter) children.push(new Paragraph({ children: [new PageBreak()] }));
  return children;
}

async function writeTextDocx(outPath, pageBlocks) {
  const sections = pageBlocks.map(({ page, blocks }, index) => ({
    properties: {
      page: {
        size: {
          width: Math.round(page.width * PT_TO_TWIP),
          height: Math.round(page.height * PT_TO_TWIP),
        },
        margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 },
      },
    },
    children: blocksToDocxChildren(blocks, { pageBreakAfter: false }),
  }));
  const doc = new Document({
    creator: "Office PDF Tools",
    title: baseName(outPath),
    sections: sections.length ? sections : undefined,
  });
  await fs.writeFile(outPath, await Packer.toBuffer(doc));
}

async function writeTextXlsx(outPath, pageBlocks) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Office PDF Tools";
  wb.created = new Date();
  pageBlocks.forEach(({ page, lines }, index) => {
    const name = `Page ${index + 1}`.slice(0, 31);
    const ws = wb.addWorksheet(name);
    const { columns, rows } = linesToMatrix(lines, { tolerance: Math.max(3, page.width * 0.016) });
    rows.forEach((row) => ws.addRow(row));
    columns.forEach((x, i) => {
      const col = ws.getColumn(i + 1);
      let width = 10;
      for (const row of rows) {
        const value = String(row[i] ?? "");
        width = Math.max(width, Math.min(60, value.length * 1.15 + 2));
      }
      col.width = width;
      col.alignment = { vertical: "top", horizontal: Number.isFinite(x) && i === 0 ? "left" : "left" };
    });
    ws.views = [{ state: "frozen", ySplit: 0 }];
  });
  if (!wb.worksheets.length) wb.addWorksheet("Page 1");
  await fs.writeFile(outPath, await wb.xlsx.writeBuffer());
}

async function writeTextPptx(outPath, pageBlocks) {
  const pptx = new PptxGenJS();
  const first = pageBlocks[0]?.page;
  const width = (first?.width || 595) * PT_TO_INCH;
  const height = (first?.height || 842) * PT_TO_INCH;
  pptx.defineLayout({ name: "PDF_TEXT", width, height });
  pptx.layout = "PDF_TEXT";
  pptx.author = "Office PDF Tools";
  for (const { page, lines, blocks } of pageBlocks) {
    const slide = pptx.addSlide();
    const heading = blocks.find((b) => b.type === "heading");
    const title = heading?.text || lines[0]?.text || "";
    const bodyLines = lines.filter((l) => l.text !== title).map((l) => l.text);
    const margin = 0.4;
    const w = width - margin * 2;
    let y = margin;
    if (title) {
      slide.addText(title, {
        x: margin, y, w, h: 0.7, fontSize: 22, bold: true, color: "1F2329",
        valign: "top", autoFit: true,
      });
      y += 0.75;
    }
    if (bodyLines.length) {
      slide.addText(bodyLines.join("\n"), {
        x: margin, y, w, h: Math.max(0.5, height - y - margin),
        fontSize: 12, color: "333333", valign: "top", shrinkText: true,
      });
    }
  }
  const data = await pptx.write({ outputType: "nodebuffer" });
  await fs.writeFile(outPath, data);
}

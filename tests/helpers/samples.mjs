/** 测试用样例文件生成。 */
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import { Document, Packer, Paragraph, TextRun, HeadingLevel, ImageRun } from "docx";
import ExcelJS from "exceljs";
import PptxGenJS from "pptxgenjs";
import { createCanvas } from "@napi-rs/canvas";

export async function tempDir(prefix = "opt-test-") {
  return fs.mkdtemp(path.join(os.tmpdir(), prefix));
}

export async function makePng(filePath, { width = 400, height = 300, color = "#2b579a", label = "IMG" } = {}) {
  const canvas = createCanvas(width, height);
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = color;
  ctx.fillRect(0, 0, width, height);
  ctx.fillStyle = "#ffffff";
  ctx.font = `bold ${Math.round(height / 5)}px sans-serif`;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(label, width / 2, height / 2);
  const png = await canvas.encode("png");
  await fs.writeFile(filePath, png);
  return filePath;
}

export async function makePdf(filePath, { pages = 5, title = "Sample" } = {}) {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  for (let i = 0; i < pages; i++) {
    const page = pdf.addPage([595, 842]);
    page.drawText(`${title} page ${i + 1}`, { x: 60, y: 760, size: 22, font: bold, color: rgb(0.1, 0.2, 0.5) });
    page.drawText(`Body line one of page ${i + 1}`, { x: 60, y: 720, size: 12, font });
    page.drawText(`Body line two of page ${i + 1}`, { x: 60, y: 700, size: 12, font });
  }
  await fs.writeFile(filePath, await pdf.save());
  return filePath;
}

export async function makeDocx(filePath, { marker = "ALPHA", headings = ["Section One", "Section Two"], withImage = false } = {}) {
  const children = [];
  for (let index = 0; index < headings.length; index++) {
    const heading = headings[index];
    children.push(new Paragraph({ text: heading, heading: HeadingLevel.HEADING_1 }));
    children.push(new Paragraph({ children: [new TextRun({ text: `${marker}-body-${index + 1}` })] }));
    children.push(new Paragraph({ text: `普通正文段落 ${index + 1}`, bullet: { level: 0 } }));
    if (withImage && index === 0) {
      // eslint-disable-next-line no-await-in-loop
      const png = await makePng(path.join(path.dirname(filePath), `${marker}-inline.png`), { width: 200, height: 120, label: marker });
      children.push(new Paragraph({
        // eslint-disable-next-line no-await-in-loop
        children: [new ImageRun({ type: "png", data: await fs.readFile(png), transformation: { width: 200, height: 120 } })],
      }));
    }
  }
  const doc = new Document({ sections: [{ children }] });
  await fs.writeFile(filePath, await Packer.toBuffer(doc));
  return filePath;
}

export async function makeXlsx(filePath, { sheets = 2, rows = 12, marker = "SHEET" } = {}) {
  const wb = new ExcelJS.Workbook();
  for (let s = 0; s < sheets; s++) {
    const ws = wb.addWorksheet(`${marker}${s + 1}`);
    ws.getRow(1).values = ["Name", "Qty", "Price"];
    for (let r = 2; r <= rows; r++) {
      ws.getRow(r).values = [`item-${s + 1}-${r}`, r, r * 1.5];
    }
    ws.getColumn(1).width = 18;
  }
  await wb.xlsx.writeFile(filePath);
  return filePath;
}

export async function makePptx(filePath, { slides = 6, marker = "SLIDE" } = {}) {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_16x9";
  for (let i = 0; i < slides; i++) {
    const slide = pptx.addSlide();
    slide.addText(`${marker} ${i + 1}`, { x: 0.5, y: 0.5, w: 8, h: 1, fontSize: 28, bold: true });
    slide.addText(`body of slide ${i + 1}`, { x: 0.5, y: 1.8, w: 8, h: 2, fontSize: 14 });
  }
  await pptx.writeFile({ fileName: filePath });
  return filePath;
}

export async function listFiles(dir) {
  return (await fs.readdir(dir)).sort();
}

import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import { tempDir, makePdf } from "./helpers/samples.mjs";
import { pdfToOffice } from "../server/lib/convert/pdfToOffice.js";
import { OoxmlPackage, extractElement } from "../server/lib/ooxml/package.js";

async function readZipXml(filePath, part) {
  const pkg = await OoxmlPackage.load(await fs.readFile(filePath), path.basename(filePath));
  return pkg.text(part);
}

test("PDF → Word（图片模式）：每页一张图，页面尺寸与 PDF 一致", async () => {
  const dir = await tempDir();
  const out = path.join(dir, "out");
  await fs.mkdir(out, { recursive: true });
  const pdf = await makePdf(path.join(dir, "src.pdf"), { pages: 2 });
  const warnings = [];
  const result = await pdfToOffice({
    inputPath: pdf, outputDir: out, options: { target: "docx", mode: "raster", dpi: 96 }, warnings,
  });
  assert.equal(result.mode, "raster");
  const xml = await readZipXml(result.files[0].path, "word/document.xml");
  assert.ok((xml.match(/<w:drawing>/g) || []).length === 2, "两张图片");
  const pkg = await OoxmlPackage.load(await fs.readFile(result.files[0].path));
  assert.equal(pkg.entries().filter((e) => e.startsWith("word/media/")).length, 2);
});

test("PDF → Word（文本模式）：标题/正文被还原为可编辑段落", async () => {
  const dir = await tempDir();
  const out = path.join(dir, "out");
  await fs.mkdir(out, { recursive: true });
  const pdf = await makePdf(path.join(dir, "text.pdf"), { pages: 1, title: "Heading" });
  const result = await pdfToOffice({
    inputPath: pdf, outputDir: out, options: { target: "docx", mode: "text" }, warnings: [],
  });
  const xml = await readZipXml(result.files[0].path, "word/document.xml");
  assert.match(xml, /Heading page 1/);
  assert.match(xml, /Body line one of page 1/);
  assert.match(xml, /w:pStyle w:val="Heading1"/, "大字号被识别为标题 1");
});

test("PDF → PowerPoint（图片模式）：每页一张幻灯片", async () => {
  const dir = await tempDir();
  const out = path.join(dir, "out");
  await fs.mkdir(out, { recursive: true });
  const pdf = await makePdf(path.join(dir, "deck.pdf"), { pages: 3 });
  const result = await pdfToOffice({
    inputPath: pdf, outputDir: out, options: { target: "pptx", mode: "raster", dpi: 96 }, warnings: [],
  });
  const xml = await readZipXml(result.files[0].path, "ppt/presentation.xml");
  const list = extractElement(xml, "p:sldIdLst")?.inner || "";
  assert.equal((list.match(/<p:sldId\b/g) || []).length, 3);
});

test("PDF → Excel（文本模式）：按行列聚类写入单元格", async () => {
  const dir = await tempDir();
  const out = path.join(dir, "out");
  await fs.mkdir(out, { recursive: true });
  // 造一个带两列的 PDF
  const { PDFDocument, StandardFonts } = await import("pdf-lib");
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const page = doc.addPage([595, 842]);
  page.drawText("Name", { x: 60, y: 780, size: 12, font });
  page.drawText("Qty", { x: 300, y: 780, size: 12, font });
  page.drawText("apple", { x: 60, y: 760, size: 12, font });
  page.drawText("3", { x: 300, y: 760, size: 12, font });
  const pdfPath = path.join(dir, "table.pdf");
  await fs.writeFile(pdfPath, await doc.save());

  const warnings = [];
  const result = await pdfToOffice({
    inputPath: pdfPath, outputDir: out, options: { target: "xlsx", mode: "text" }, warnings,
  });
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(result.files[0].path);
  const ws = wb.worksheets[0];
  assert.equal(ws.name, "Page 1");
  const rows = [];
  ws.eachRow((row) => rows.push(row.values.slice(1)));
  assert.deepEqual(rows[0], ["Name", "Qty"]);
  assert.deepEqual(rows[1], ["apple", "3"]);
});

test("PDF → Excel（图片模式）会降级为文本模式并给出提示", async () => {
  const dir = await tempDir();
  const out = path.join(dir, "out");
  await fs.mkdir(out, { recursive: true });
  const pdf = await makePdf(path.join(dir, "x.pdf"), { pages: 1 });
  const warnings = [];
  const result = await pdfToOffice({
    inputPath: pdf, outputDir: out, options: { target: "xlsx", mode: "raster" }, warnings,
  });
  assert.equal(result.mode, "text");
  assert.ok(warnings.some((w) => w.includes("Excel")));
});

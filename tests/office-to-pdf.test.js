import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { tempDir, makeDocx, makeXlsx, makePptx } from "./helpers/samples.mjs";
import { detectEngines } from "../server/lib/engines/detect.js";
import { officeToPdf } from "../server/lib/convert/officeToPdf.js";
import { getPdfInfo, extractTextByPage } from "../server/lib/engines/render.js";
import { buildLines } from "../server/lib/pdf/layout.js";

const engines = await detectEngines();
const hasEngine = !!(engines.libreoffice?.available || engines.msoffice?.available);
const skip = hasEngine ? false : "未安装 Microsoft Office 或 LibreOffice，跳过真实转换测试";

test("Office → PDF：Word / Excel / PowerPoint", { skip, timeout: 300000 }, async () => {
  const dir = await tempDir();
  const out = path.join(dir, "out");
  await fs.mkdir(out, { recursive: true });
  const warnings = [];

  const docx = await makeDocx(path.join(dir, "report.docx"), { marker: "CONVERT", headings: ["季度报告"] });
  const pdfFromDocx = await officeToPdf({ inputPath: docx, outputDir: out, options: {}, warnings });
  const docxInfo = await getPdfInfo(await fs.readFile(pdfFromDocx.outputPath));
  assert.ok(docxInfo.pages >= 1);
  assert.ok(["msoffice", "libreoffice"].includes(pdfFromDocx.engine));
  const text = await extractTextByPage(await fs.readFile(pdfFromDocx.outputPath));
  const all = text.flatMap((p) => buildLines(p.items).map((l) => l.text)).join("\n");
  assert.match(all, /CONVERT-body-1/, "PDF 中包含 Word 正文");
  assert.match(all, /季度报告/, "中文标题被提取");

  const xlsx = await makeXlsx(path.join(dir, "book.xlsx"), { sheets: 2, marker: "CALC" });
  const pdfFromXlsx = await officeToPdf({ inputPath: xlsx, outputDir: out, options: {}, warnings });
  const xlsxInfo = await getPdfInfo(await fs.readFile(pdfFromXlsx.outputPath));
  assert.ok(xlsxInfo.pages >= 1);

  const pptx = await makePptx(path.join(dir, "deck.pptx"), { slides: 3, marker: "DECK" });
  const pdfFromPptx = await officeToPdf({ inputPath: pptx, outputDir: out, options: {}, warnings });
  const pptxInfo = await getPdfInfo(await fs.readFile(pdfFromPptx.outputPath));
  assert.equal(pptxInfo.pages, 3, "3 张幻灯片 → 3 页 PDF");

  const files = await fs.readdir(out);
  assert.equal(files.filter((f) => f.endsWith(".pdf")).length, 3);
});

#!/usr/bin/env node
/**
 * 端到端冒烟测试（不依赖 Office 界面）：
 * 启动本地服务 → 通过 HTTP API 依次跑通 图片→PDF、PDF→图片、PDF→Office、
 * Office→PDF（需要本机 Office 或 LibreOffice）、Word 合并与拆分，并打印结果表。
 *
 * 用法：
 *   npm run smoke                 # 使用自动生成的样例文件
 *   npm run smoke -- a.docx b.pdf # 使用指定文件（依次尝试各操作）
 */
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { createApp } from "../server/app.js";
import { config } from "../server/config.js";
import { detectEngines } from "../server/lib/engines/detect.js";

const results = [];

async function main() {
  const workDir = await fs.mkdtemp(path.join(os.tmpdir(), "office-pdf-smoke-"));
  const outDir = path.join(workDir, "out");
  await fs.mkdir(outDir, { recursive: true });
  config.outputDir = outDir;
  config.tmpDir = path.join(workDir, "tmp");

  const inputs = process.argv.slice(2).filter((a) => !a.startsWith("-"));
  const samples = inputs.length ? await describeInputs(inputs) : await makeSamples(workDir);

  const app = createApp();
  const server = app.listen(0, "127.0.0.1");
  await new Promise((resolve) => server.once("listening", resolve));
  const base = `http://127.0.0.1:${server.address().port}`;
  console.log(`本地服务: ${base}`);
  console.log(`输出目录: ${outDir}\n`);

  const engines = await detectEngines();
  const hasOfficeEngine = !!(engines.libreoffice?.available || engines.msoffice?.available);
  console.log(`引擎: LibreOffice=${engines.libreoffice?.available ? "可用" : "不可用"}, ` +
    `MS Office=${engines.msoffice?.available ? "可用" : "不可用"}, ` +
    `PDF 渲染=${engines.renderer?.available ? "可用" : "不可用"}`);
  console.log("");

  try {
    await step("健康检查", () => get(base, "/api/health"));

    if (samples.images.length >= 2) {
      await step("图片 → PDF（2 张，A4 每页 1 张）", () => job(base, "images-to-pdf", samples.images, { pageSize: "a4", perPage: 1 }));
    }
    if (samples.pdf[0]) {
      await step("PDF → 图片（PNG, 150dpi, 前 2 页）", () => job(base, "pdf-to-images", [samples.pdf[0]], { format: "png", dpi: 150, pages: "1-2" }));
      await step("PDF → Word（图片高保真模式）", () => job(base, "pdf-to-office", [samples.pdf[0]], { target: "docx", mode: "raster", dpi: 120 }));
      await step("PDF → Word（文本可编辑模式）", () => job(base, "pdf-to-office", [samples.pdf[0]], { target: "docx", mode: "text" }));
      await step("PDF → PowerPoint（图片模式）", () => job(base, "pdf-to-office", [samples.pdf[0]], { target: "pptx", mode: "raster", dpi: 120 }));
      await step("PDF → Excel（文本表格模式）", () => job(base, "pdf-to-office", [samples.pdf[0]], { target: "xlsx", mode: "text" }));
    }
    if (samples.office.length) {
      if (hasOfficeEngine) {
        await step("Office → PDF（Word/Excel/PPT 全部转换）", () => job(base, "office-to-pdf", samples.office, { engine: "auto" }));
      } else {
        results.push({ name: "Office → PDF", status: "跳过", detail: "本机没有 Microsoft Office 或 LibreOffice" });
      }
    }
    if (samples.docs.length >= 2) {
      await step("Word 合并", () => job(base, "merge", samples.docs, { kind: "docx", pageBreak: true }));
      await step("Word 拆分（按标题 1）", () => job(base, "split", [samples.docs[0]], { kind: "docx", strategy: "heading", param: "1" }));
    }
    if (samples.pdf.length >= 2) {
      await step("PDF 合并", () => job(base, "merge", samples.pdf, { kind: "pdf" }));
    }
    if (samples.pdf[0]) {
      await step("PDF 拆分（每 2 页）", () => job(base, "split", [samples.pdf[0]], { kind: "pdf", strategy: "every-n-pages", param: "2" }));
    }
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }

  console.log("\n================ 结果 ================");
  for (const row of results) {
    const mark = row.status === "通过" ? "✔" : row.status === "跳过" ? "-" : "✘";
    console.log(`${mark} ${row.name.padEnd(34)} ${row.status.padEnd(4)} ${row.detail}`);
  }
  const failed = results.filter((r) => r.status === "失败");
  console.log(`\n共 ${results.length} 项，失败 ${failed.length} 项。文件位于: ${outDir}`);
  process.exit(failed.length ? 1 : 0);
}

async function describeInputs(paths) {
  const list = { images: [], pdf: [], office: [], docs: [] };
  for (const p of paths) {
    const abs = path.resolve(p);
    const ext = path.extname(abs).toLowerCase();
    if ([".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"].includes(ext)) list.images.push(abs);
    else if (ext === ".pdf") list.pdf.push(abs);
    else if ([".docx", ".xlsx", ".pptx"].includes(ext)) {
      list.office.push(abs);
      if (ext === ".docx") list.docs.push(abs);
    }
  }
  return list;
}

async function makeSamples(dir) {
  const { PDFDocument, StandardFonts } = await import("pdf-lib");
  const { Document, Packer, Paragraph, HeadingLevel } = await import("docx");
  const { createCanvas } = await import("@napi-rs/canvas");
  const { default: ExcelJS } = await import("exceljs");
  const { default: PptxGenJS } = await import("pptxgenjs");
  const samples = { images: [], pdf: [], office: [], docs: [] };

  for (const [name, color, label] of [["img-a.png", "#c43e1c", "A"], ["img-b.png", "#2b579a", "B"]]) {
    const canvas = createCanvas(600, 400);
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = color;
    ctx.fillRect(0, 0, 600, 400);
    ctx.fillStyle = "#fff";
    ctx.font = "bold 90px sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(label, 300, 200);
    const file = path.join(dir, name);
    await fs.writeFile(file, await canvas.encode("png"));
    samples.images.push(file);
  }

  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  for (let i = 1; i <= 4; i++) {
    const page = pdf.addPage([595, 842]);
    page.drawText(`Smoke test heading ${i}`, { x: 60, y: 770, size: 22, font: bold });
    page.drawText(`Body line for page ${i}`, { x: 60, y: 730, size: 12, font });
    page.drawText("Name", { x: 60, y: 690, size: 12, font });
    page.drawText("Qty", { x: 300, y: 690, size: 12, font });
    page.drawText(`row-${i}`, { x: 60, y: 670, size: 12, font });
    page.drawText(String(i * 3), { x: 300, y: 670, size: 12, font });
  }
  const pdfFile = path.join(dir, "sample.pdf");
  await fs.writeFile(pdfFile, await pdf.save());
  samples.pdf.push(pdfFile, pdfFile);

  const mkDocx = async (file, marker, headings) => {
    const children = [];
    for (const h of headings) {
      children.push(new Paragraph({ text: h, heading: HeadingLevel.HEADING_1 }));
      children.push(new Paragraph(`${marker} 正文段落 —— 用于验证合并与拆分。`));
    }
    await fs.writeFile(file, await Packer.toBuffer(new Document({ sections: [{ children }] })));
  };
  const docA = path.join(dir, "doc-a.docx");
  const docB = path.join(dir, "doc-b.docx");
  await mkDocx(docA, "AAA", ["第一章", "第二章"]);
  await mkDocx(docB, "BBB", ["附录 A"]);
  samples.docs.push(docA, docB);
  samples.office.push(docA);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("数据");
  ws.getRow(1).values = ["项目", "数量"];
  ws.getRow(2).values = ["示例", 42];
  const xlsx = path.join(dir, "book.xlsx");
  await wb.xlsx.writeFile(xlsx);
  samples.office.push(xlsx);

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_16x9";
  pptx.addSlide().addText("冒烟测试", { x: 0.5, y: 0.5, w: 8, h: 1, fontSize: 32, bold: true });
  pptx.addSlide().addText("第二页", { x: 0.5, y: 0.5, w: 8, h: 1, fontSize: 28 });
  const pptxFile = path.join(dir, "deck.pptx");
  await pptx.writeFile({ fileName: pptxFile });
  samples.office.push(pptxFile);

  console.log(`已生成样例文件: ${samples.images.length} 图片 / ${samples.pdf.length} PDF / ${samples.office.length} Office 文档\n`);
  return samples;
}

async function get(base, apiPath) {
  const res = await fetch(base + apiPath);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
}

async function job(base, op, files, options) {
  const form = new FormData();
  for (const file of files) {
    const buffer = await fs.readFile(file);
    form.append("files", new Blob([buffer]), path.basename(file));
  }
  form.append("op", op);
  form.append("options", JSON.stringify(options || {}));
  const res = await fetch(`${base}/api/jobs`, { method: "POST", body: form });
  const data = await res.json();
  if (!res.ok) throw new Error(data?.error?.message || `HTTP ${res.status}`);
  for (let i = 0; i < 600; i++) {
    // eslint-disable-next-line no-await-in-loop
    const current = await get(base, `/api/jobs/${data.jobId}`);
    if (current.status === "done") return current;
    if (current.status === "error") throw new Error(current.error?.message || "任务失败");
    if (current.status === "canceled") throw new Error("任务被取消");
    // eslint-disable-next-line no-await-in-loop
    await new Promise((r) => setTimeout(r, 250));
  }
  throw new Error("任务超时");
}

async function step(name, fn) {
  const started = Date.now();
  try {
    const result = await fn();
    const files = result?.result?.files || [];
    const detail = files.length
      ? `${files.length} 个文件 · ${files.map((f) => f.name).slice(0, 3).join(", ")}${files.length > 3 ? " …" : ""}`
      : "OK";
    results.push({ name, status: "通过", detail: `${detail} (${Date.now() - started}ms)` });
    console.log(`✔ ${name} — ${detail}`);
  } catch (err) {
    results.push({ name, status: "失败", detail: err.message });
    console.log(`✘ ${name} — ${err.message}`);
  }
}

main().catch((err) => {
  console.error("冒烟测试异常:", err);
  process.exit(1);
});

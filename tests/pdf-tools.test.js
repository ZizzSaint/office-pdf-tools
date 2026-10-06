import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { tempDir, makePdf, makePng } from "./helpers/samples.mjs";
import { getPdfInfo } from "../server/lib/engines/render.js";
import { imagesToPdf } from "../server/lib/convert/imagesToPdf.js";
import { pdfToImages } from "../server/lib/convert/pdfToImages.js";
import { mergePdf } from "../server/lib/merge/pdf.js";
import { splitPdf } from "../server/lib/split/pdf.js";

test("图片 → PDF：多图合并、N-up、每图单独 PDF", async () => {
  const dir = await tempDir();
  const out = path.join(dir, "out");
  await fs.mkdir(out, { recursive: true });
  const a = await makePng(path.join(dir, "a.png"), { label: "A", color: "#c43e1c" });
  const b = await makePng(path.join(dir, "b.png"), { width: 300, height: 500, label: "B", color: "#2b579a" });

  const merged = await imagesToPdf({ inputPaths: [a, b], outputDir: out, options: { pageSize: "a4", perPage: 1 } });
  assert.equal(merged.files.length, 1);
  const info = await getPdfInfo(await fs.readFile(merged.files[0].path));
  assert.equal(info.pages, 2);
  // 首图是横向的，orientation=auto 时应使用横向 A4
  assert.ok(Math.abs(info.pageSizes[0].width - 841.89) < 3, "横向 A4 宽度");
  assert.ok(Math.abs(info.pageSizes[0].height - 595.28) < 3, "横向 A4 高度");

  const nup = await imagesToPdf({ inputPaths: [a, b], outputDir: out, options: { pageSize: "a4", perPage: 4 } });
  const nupInfo = await getPdfInfo(await fs.readFile(nup.files[0].path));
  assert.equal(nupInfo.pages, 1, "两张图放在一页上");

  const separate = await imagesToPdf({ inputPaths: [a, b], outputDir: out, options: { separate: true, pageSize: "fit" } });
  assert.equal(separate.files.length, 2);
  const fitInfo = await getPdfInfo(await fs.readFile(separate.files[0].path));
  assert.ok(Math.abs(fitInfo.pageSizes[0].width - 400 * 0.75) < 2, "fit 模式使用图片尺寸(96dpi)");
});

test("PDF → 图片：页码范围与命名", async () => {
  const dir = await tempDir();
  const pdfPath = await makePdf(path.join(dir, "src.pdf"), { pages: 4 });
  const out = path.join(dir, "out");
  await fs.mkdir(out, { recursive: true });
  const result = await pdfToImages({ inputPath: pdfPath, outputDir: out, options: { dpi: 72, format: "png", pages: "2-3" } });
  assert.equal(result.files.length, 2);
  const png = await fs.readFile(result.files[0].path);
  assert.equal(png.slice(1, 4).toString(), "PNG");
  assert.match(result.files[0].name, /-p2\.png$/);
});

test("PDF 合并与拆分", async () => {
  const dir = await tempDir();
  const out = path.join(dir, "out");
  await fs.mkdir(out, { recursive: true });
  const p1 = await makePdf(path.join(dir, "one.pdf"), { pages: 2, title: "One" });
  const p2 = await makePdf(path.join(dir, "two.pdf"), { pages: 3, title: "Two" });
  const merged = await mergePdf({ inputPaths: [p1, p2], outputDir: out, options: {} });
  const mergedInfo = await getPdfInfo(await fs.readFile(merged.files[0].path));
  assert.equal(mergedInfo.pages, 5);

  const split = await splitPdf({ inputPath: merged.files[0].path, outputDir: out, strategy: "every-n-pages", param: "2" });
  assert.equal(split.files.length, 3);
  const firstInfo = await getPdfInfo(await fs.readFile(split.files[0].path));
  assert.equal(firstInfo.pages, 2);

  const ranges = await splitPdf({ inputPath: merged.files[0].path, outputDir: out, strategy: "ranges", param: "1-2,4-5" });
  assert.equal(ranges.files.length, 2);

  const each = await splitPdf({ inputPath: merged.files[0].path, outputDir: out, strategy: "each-page" });
  assert.equal(each.files.length, 5);
});

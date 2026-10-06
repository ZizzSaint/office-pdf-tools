import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import { tempDir, makeXlsx } from "./helpers/samples.mjs";
import { mergeXlsx } from "../server/lib/merge/xlsx.js";
import { splitXlsx } from "../server/lib/split/xlsx.js";

test("Excel 合并：工作表汇总并保留数据", async () => {
  const dir = await tempDir();
  const out = path.join(dir, "out");
  await fs.mkdir(out, { recursive: true });
  const a = await makeXlsx(path.join(dir, "a.xlsx"), { sheets: 2, marker: "AAA" });
  const b = await makeXlsx(path.join(dir, "b.xlsx"), { sheets: 1, marker: "BBB" });
  const result = await mergeXlsx({ inputPaths: [a, b], outputDir: out, options: {}, warnings: [] });
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(result.files[0].path);
  assert.equal(wb.worksheets.length, 3);
  assert.deepEqual(wb.worksheets.map((ws) => ws.name), ["AAA1", "AAA2", "BBB1"]);
  assert.equal(wb.getWorksheet("AAA1").getCell("A2").value, "item-1-2");
  assert.equal(wb.getWorksheet("BBB1").getCell("B2").value, 2);
});

test("Excel 拆分：按工作表 / 每 N 行", async () => {
  const dir = await tempDir();
  const out = path.join(dir, "out");
  await fs.mkdir(out, { recursive: true });
  const src = await makeXlsx(path.join(dir, "book.xlsx"), { sheets: 3, rows: 10, marker: "WS" });
  const bySheet = await splitXlsx({ inputPath: src, outputDir: out, strategy: "sheets" });
  assert.equal(bySheet.files.length, 3);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(bySheet.files[0].path);
  assert.equal(wb.worksheets.length, 1);
  assert.equal(wb.worksheets[0].name, "WS1");

  const byRows = await splitXlsx({ inputPath: src, outputDir: out, strategy: "every-n-rows", param: "5" });
  assert.equal(byRows.files.length, 6, "3 个工作表，每个 10 行数据 + 表头，每 5 行一个文件");
  const chunk = new ExcelJS.Workbook();
  await chunk.xlsx.readFile(byRows.files[0].path);
  const ws = chunk.worksheets[0];
  assert.equal(ws.getCell("A1").value, "Name", "每个分片都带表头");
  assert.equal(ws.rowCount, 6, "表头 + 5 行");
});

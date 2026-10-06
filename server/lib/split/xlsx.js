/** Excel 拆分：每个工作表一个文件 / 每 N 行一个文件。 */
import fs from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import { AppError } from "../errors.js";
import { baseName, safeName, uniquePath } from "../paths.js";
import { copyCell, loadWorkbook } from "../merge/xlsx.js";

function newWorkbook() {
  const out = new ExcelJS.Workbook();
  out.creator = "Office PDF Tools";
  out.created = new Date();
  return out;
}

function copySheetRows(srcSheet, dstWorkbook, name, { from = 1, to = null, withHeader = null } = {}) {
  const ws = dstWorkbook.addWorksheet(name);
  srcSheet.columns?.forEach((col, index) => {
    const target = ws.getColumn(index + 1);
    if (col?.width) target.width = col.width;
  });
  const last = to ?? srcSheet.rowCount;
  const rows = [];
  if (withHeader && withHeader >= 1 && withHeader < from) rows.push(withHeader);
  for (let r = from; r <= last; r++) rows.push(r);
  let outRow = 1;
  for (const r of rows) {
    const srcRow = srcSheet.getRow(r);
    const target = ws.getRow(outRow++);
    if (srcRow.height) target.height = srcRow.height;
    srcRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      copyCell(cell, target.getCell(colNumber));
    });
  }
  for (const range of Object.values(srcSheet._mergedCells || {})) {
    if (typeof range === "string") ws.mergeCells(range);
  }
  return ws;
}

export async function splitXlsx({ inputPath, outputDir, strategy = "sheets", param, progress }) {
  const wb = await loadWorkbook(inputPath);
  if (!wb.worksheets.length) throw new AppError("工作簿中没有工作表。", { status: 400, code: "bad-request" });
  const stem = baseName(inputPath);
  const files = [];

  const emit = async (sheetName, build) => {
    const out = newWorkbook();
    if (!build(out)) return;
    // eslint-disable-next-line no-await-in-loop
    const target = await uniquePath(outputDir, `${stem}-${safeName(sheetName, "sheet")}.xlsx`);
    // eslint-disable-next-line no-await-in-loop
    await out.xlsx.writeFile(target);
    files.push({ name: path.basename(target), path: target, kind: "xlsx" });
  };

  if (strategy === "every-n-rows") {
    const n = Math.max(1, Number(param) || 100);
    for (const sheet of wb.worksheets) {
      if (!sheet.rowCount) continue;
      const headerRow = sheet.getRow(1);
      const hasHeader = headerRow && headerRow.values && headerRow.values.filter((v) => v !== undefined && v !== null).length > 0;
      for (let start = hasHeader ? 2 : 1; start <= sheet.rowCount; start += n) {
        const end = Math.min(sheet.rowCount, start + n - 1);
        // eslint-disable-next-line no-await-in-loop
        await emit(`${sheet.name}-rows${start}-${end}`, (out) => {
          copySheetRows(sheet, out, sheet.name.slice(0, 31), { from: start, to: end, withHeader: hasHeader ? 1 : null });
          return true;
        });
      }
    }
  } else {
    for (const sheet of wb.worksheets) {
      // eslint-disable-next-line no-await-in-loop
      await emit(sheet.name, (out) => {
        copySheetRows(sheet, out, sheet.name.slice(0, 31));
        return true;
      });
    }
  }

  if (!files.length) throw new AppError("没有可拆分的内容。", { status: 400, code: "bad-request" });
  progress?.(1, "完成");
  return { files };
}

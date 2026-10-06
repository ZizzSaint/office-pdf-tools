/** Excel 合并：把每个工作簿的工作表汇总到一个新工作簿（exceljs）。 */
import fs from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import { AppError } from "../errors.js";
import { baseName, uniquePath } from "../paths.js";

/** 深拷贝单元格内容与样式。 */
export function copyCell(srcCell, dstCell, { withStyle = true } = {}) {
  if (srcCell.value !== null && srcCell.value !== undefined) dstCell.value = srcCell.value;
  if (srcCell.formula && srcCell.value === null) dstCell.value = { formula: srcCell.formula, result: srcCell.result };
  if (!withStyle) return;
  if (srcCell.numFmt) dstCell.numFmt = srcCell.numFmt;
  if (srcCell.font) dstCell.font = { ...srcCell.font };
  if (srcCell.fill) dstCell.fill = JSON.parse(JSON.stringify(srcCell.fill));
  if (srcCell.border) dstCell.border = JSON.parse(JSON.stringify(srcCell.border));
  if (srcCell.alignment) dstCell.alignment = { ...srcCell.alignment };
}

export function copyWorksheet(srcSheet, dstWorkbook, name) {
  const ws = dstWorkbook.addWorksheet(name);
  // 列宽 / 隐藏
  srcSheet.columns?.forEach((col, index) => {
    const target = ws.getColumn(index + 1);
    if (col?.width) target.width = col.width;
    if (col?.hidden) target.hidden = true;
    if (col?.outlineLevel) target.outlineLevel = col.outlineLevel;
  });
  // 行与单元格
  srcSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const target = ws.getRow(rowNumber);
    if (row.height) target.height = row.height;
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      copyCell(cell, target.getCell(colNumber));
    });
    if (row.hidden) target.hidden = true;
  });
  // 合并单元格
  for (const range of Object.values(srcSheet._mergedCells || {})) {
    if (typeof range === "string") ws.mergeCells(range);
    else if (range?.model) ws.mergeCells(range.model.left, range.model.top, range.model.right, range.model.bottom);
  }
  ws.views = srcSheet.views?.map((v) => ({ ...v })) || [];
  if (srcSheet.autoFilter) ws.autoFilter = srcSheet.autoFilter;
  ws.properties.defaultRowHeight = srcSheet.properties?.defaultRowHeight || 15;
  return ws;
}

export async function loadWorkbook(inputPath) {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.readFile(inputPath);
  } catch (err) {
    throw new AppError(`${path.basename(inputPath)} 无法读取（仅支持 .xlsx/.xlsm，且不能加密）。`, {
      code: "invalid-document",
      status: 400,
      cause: err,
    });
  }
  return wb;
}

export function uniqueSheetName(workbook, name) {
  const base = String(name || "Sheet").replace(/[\\/?*\[\]:]/g, "_").slice(0, 31) || "Sheet";
  let candidate = base;
  let i = 2;
  while (workbook.worksheets.some((ws) => ws.name.toLowerCase() === candidate.toLowerCase())) {
    candidate = `${base.slice(0, 28)}(${i})`;
    i++;
  }
  return candidate;
}

export async function mergeXlsx({ inputPaths, outputDir, options = {}, progress, warnings = [] }) {
  if (inputPaths.length < 2) throw new AppError("至少需要两个 Excel 文件。", { status: 400, code: "bad-request" });
  const out = new ExcelJS.Workbook();
  out.creator = "Office PDF Tools";
  out.created = new Date();

  for (let i = 0; i < inputPaths.length; i++) {
    const inputPath = inputPaths[i];
    // eslint-disable-next-line no-await-in-loop
    const wb = await loadWorkbook(inputPath);
    const prefix = options.sheetPrefix ? `${baseName(inputPath)}-` : "";
    for (const sheet of wb.worksheets) {
      if (sheet.state === "veryHidden") continue;
      const name = uniqueSheetName(out, `${prefix}${sheet.name}`);
      copyWorksheet(sheet, out, name);
    }
    progress?.(0.1 + 0.8 * ((i + 1) / inputPaths.length), `合并 ${path.basename(inputPath)}`);
  }
  if (!out.worksheets.length) throw new AppError("没有可合并的工作表。", { status: 400, code: "bad-request" });
  const name = options.outputName ? baseName(options.outputName) : `${baseName(inputPaths[0])}-merged`;
  const target = await uniquePath(outputDir, `${name}.xlsx`);
  await out.xlsx.writeFile(target);
  warnings.push("Excel 合并会保留单元格内容与常用样式，但图表、图片、数据透视表、宏不会被复制。");
  return { files: [{ name: path.basename(target), path: target, kind: "xlsx" }] };
}

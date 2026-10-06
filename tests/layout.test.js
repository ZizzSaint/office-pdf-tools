import test from "node:test";
import assert from "node:assert/strict";
import { buildLines, detectBlocks, linesToMatrix } from "../server/lib/pdf/layout.js";

const item = (text, x, y, fontSize = 12, extra = {}) => ({
  text, x, y, width: text.length * fontSize * 0.5, height: fontSize, fontSize, bold: false, italic: false, ...extra,
});

test("buildLines 按 y 聚合、按间距切分单元格", () => {
  const lines = buildLines([
    item("Hello", 60, 700),
    item("World", 110, 700),
    item("Column2", 400, 700),
    item("Next line", 60, 680),
  ]);
  assert.equal(lines.length, 2);
  assert.equal(lines[0].text.startsWith("Hello"), true);
  assert.equal(lines[0].cells.length, 2, "大间距应切成两个单元格");
  assert.equal(lines[1].cells.length, 1);
});

test("detectBlocks 识别标题、段落与表格", () => {
  const lines = buildLines([
    item("Big Title", 60, 780, 26, { bold: true }),
    item("This is a paragraph line one", 60, 740, 12),
    item("and it continues here", 60, 724, 12),
    item("Name", 60, 680, 12),
    item("Qty", 300, 680, 12),
    item("apple", 60, 664, 12),
    item("3", 300, 664, 12),
  ]);
  const blocks = detectBlocks(lines);
  assert.equal(blocks[0].type, "heading");
  assert.equal(blocks[0].level, 1);
  assert.ok(blocks.some((b) => b.type === "paragraph"));
  const table = blocks.find((b) => b.type === "table");
  assert.ok(table, "应识别出表格");
  assert.equal(table.rows.length, 2);
  assert.deepEqual(table.rows[0], ["Name", "Qty"]);
});

test("linesToMatrix 按列聚类", () => {
  const lines = buildLines([
    item("a", 60, 700),
    item("b", 200, 700),
    item("c", 60, 680),
    item("d", 200, 680),
  ]);
  const { columns, rows } = linesToMatrix(lines);
  assert.equal(columns.length, 2);
  assert.deepEqual(rows[0], ["a", "b"]);
  assert.deepEqual(rows[1], ["c", "d"]);
});

import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { tempDir, makeDocx } from "./helpers/samples.mjs";
import { mergeDocx } from "../server/lib/merge/docx.js";
import { splitDocx } from "../server/lib/split/docx.js";
import { OoxmlPackage } from "../server/lib/ooxml/package.js";

async function docText(filePath) {
  const pkg = await OoxmlPackage.load(await fs.readFile(filePath), path.basename(filePath));
  const xml = await pkg.text("word/document.xml");
  return xml;
}

test("Word 合并：正文、样式与图片关系都搬迁", async () => {
  const dir = await tempDir();
  const out = path.join(dir, "out");
  await fs.mkdir(out, { recursive: true });
  const a = await makeDocx(path.join(dir, "a.docx"), { marker: "ALPHA-MARKER", headings: ["Alpha Head"], withImage: true });
  const b = await makeDocx(path.join(dir, "b.docx"), { marker: "BETA-MARKER", headings: ["Beta Head"] });
  const warnings = [];
  const result = await mergeDocx({ inputPaths: [a, b], outputDir: out, options: { pageBreak: true }, warnings });
  assert.equal(result.files.length, 1);

  const xml = await docText(result.files[0].path);
  assert.ok(xml.includes("ALPHA-MARKER"), "保留基准文档内容");
  assert.ok(xml.includes("BETA-MARKER"), "追加了第二个文档");
  assert.ok(xml.includes('w:type="page"'), "插入了分页符");

  const pkg = await OoxmlPackage.load(await fs.readFile(result.files[0].path));
  const rels = await pkg.readRels("word/document.xml");
  const images = rels.filter((r) => /\/image$/.test(r.Type));
  assert.equal(images.length, 1, "图片关系应被搬迁");
  const media = pkg.entries().filter((e) => e.startsWith("word/media/"));
  assert.equal(media.length, 1);
  // 被引用的图片必须真实存在
  const target = images[0].Target.replace(/^\.\//, "");
  assert.ok(pkg.has(`word/${target}`), "图片部件存在");
});

test("Word 拆分：按标题 1 拆分，每个文件只包含对应章节", async () => {
  const dir = await tempDir();
  const out = path.join(dir, "out");
  await fs.mkdir(out, { recursive: true });
  const src = await makeDocx(path.join(dir, "book.docx"), {
    marker: "SPLIT",
    headings: ["Chapter One", "Chapter Two", "Chapter Three"],
  });
  const result = await splitDocx({ inputPath: src, outputDir: out, strategy: "heading", param: "1" });
  assert.equal(result.files.length, 3);
  for (let i = 0; i < 3; i++) {
    const xml = await docText(result.files[i].path);
    assert.ok(xml.includes(`SPLIT-body-${i + 1}`), `第 ${i + 1} 个文件包含对应正文`);
    const others = [0, 1, 2].filter((k) => k !== i);
    for (const other of others) {
      assert.ok(!xml.includes(`SPLIT-body-${other + 1}`), `第 ${i + 1} 个文件不应包含第 ${other + 1} 章正文`);
    }
  }
  assert.match(result.files[0].name, /Chapter One\.docx$/);
});

test("Word 拆分：按分节符", async () => {
  const dir = await tempDir();
  const out = path.join(dir, "out");
  await fs.mkdir(out, { recursive: true });
  // 用两个 section 生成带分节符的文档
  const { Document, Packer, Paragraph } = await import("docx");
  const doc = new Document({
    sections: [
      { children: [new Paragraph("FIRST-SECTION")] },
      { children: [new Paragraph("SECOND-SECTION")] },
    ],
  });
  const src = path.join(dir, "sections.docx");
  await fs.writeFile(src, await Packer.toBuffer(doc));
  const result = await splitDocx({ inputPath: src, outputDir: out, strategy: "section" });
  assert.equal(result.files.length, 2);
  const first = await docText(result.files[0].path);
  const second = await docText(result.files[1].path);
  assert.ok(first.includes("FIRST-SECTION") && !first.includes("SECOND-SECTION"));
  assert.ok(second.includes("SECOND-SECTION"));
});

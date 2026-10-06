import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { tempDir, makePptx } from "./helpers/samples.mjs";
import { mergePptx } from "../server/lib/merge/pptx.js";
import { splitPptx } from "../server/lib/split/pptx.js";
import { OoxmlPackage, extractElement } from "../server/lib/ooxml/package.js";

async function slideCount(filePath) {
  const pkg = await OoxmlPackage.load(await fs.readFile(filePath), path.basename(filePath));
  const xml = await pkg.text("ppt/presentation.xml");
  const list = extractElement(xml, "p:sldIdLst")?.inner || "";
  return (list.match(/<p:sldId\b/g) || []).length;
}

test("PowerPoint 合并：幻灯片、母版与版式都被正确搬迁", async () => {
  const dir = await tempDir();
  const out = path.join(dir, "out");
  await fs.mkdir(out, { recursive: true });
  const a = await makePptx(path.join(dir, "a.pptx"), { slides: 3, marker: "AAA" });
  const b = await makePptx(path.join(dir, "b.pptx"), { slides: 4, marker: "BBB" });
  const result = await mergePptx({ inputPaths: [a, b], outputDir: out, options: {}, warnings: [] });
  assert.equal(await slideCount(result.files[0].path), 7);

  const pkg = await OoxmlPackage.load(await fs.readFile(result.files[0].path));
  const xml = await pkg.text("ppt/presentation.xml");
  // 基准 + 追加的母版都应在 sldMasterIdLst 中注册
  const masters = extractElement(xml, "p:sldMasterIdLst")?.inner || "";
  assert.ok((masters.match(/<p:sldMasterId\b/g) || []).length >= 2, "两个演示文稿的母版均已注册");
  const rels = await pkg.readRels("ppt/presentation.xml");
  for (const rel of rels.filter((r) => /\/slide$/.test(r.Type))) {
    const target = rel.Target.replace(/^\.\.\//, "ppt/").replace(/^\//, "");
    assert.ok(pkg.has(target) || pkg.has(`ppt/${rel.Target.replace(/^\.\.\//, "")}`), `幻灯片部件存在: ${rel.Target}`);
  }
  // 每张幻灯片引用的版式必须存在
  for (const entry of pkg.entries().filter((e) => /^ppt\/slides\/slide\d+\.xml$/.test(e))) {
    const slideRels = await pkg.readRels(entry);
    const layoutRel = slideRels.find((r) => /\/slideLayout$/.test(r.Type));
    assert.ok(layoutRel, "幻灯片有版式关系");
    const layoutPath = path.posix.normalize(path.posix.join("ppt/slides", layoutRel.Target));
    assert.ok(pkg.has(layoutPath), `版式存在: ${layoutPath}`);
  }
});

test("PowerPoint 拆分：按每 N 张 / 范围", async () => {
  const dir = await tempDir();
  const out = path.join(dir, "out");
  await fs.mkdir(out, { recursive: true });
  const src = await makePptx(path.join(dir, "deck.pptx"), { slides: 7, marker: "DDD" });
  const everyN = await splitPptx({ inputPath: src, outputDir: out, strategy: "every-n-slides", param: "3" });
  assert.equal(everyN.files.length, 3);
  assert.equal(await slideCount(everyN.files[0].path), 3);
  assert.equal(await slideCount(everyN.files[2].path), 1);
  assert.match(everyN.files[0].name, /slides1-3\.pptx$/);

  const ranges = await splitPptx({ inputPath: src, outputDir: out, strategy: "ranges", param: "1-2,6-7" });
  assert.equal(ranges.files.length, 2);
  assert.equal(await slideCount(ranges.files[1].path), 2);
});

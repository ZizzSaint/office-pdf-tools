/** Word (.docx) 合并：把源文档正文追加到基准文档，并搬迁图片/图表等关系部件。 */
import fs from "node:fs/promises";
import path from "node:path";
import {
  OoxmlPackage, extractElement, splitTopLevelElements, resolvePartTarget, relativeTarget,
  nextRelId, rewriteRelRefs, serializeContentTypes, registerPart, contentTypeFor,
} from "../ooxml/package.js";
import { createContext, copyPartDeep } from "../ooxml/partCopy.js";
import { AppError } from "../errors.js";
import { baseName, uniquePath } from "../paths.js";

const PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

/** 这些关系由目标文档自身提供，不需要从源文档搬迁。 */
const SKIP_REL_TYPES = [
  /\/(styles|stylesWithEffects|settings|webSettings|fontTable|theme|themeOverride|numbering|header|footer|footnotes|endnotes|comments|people|glossaryDocument)$/,
];

const NUMBERING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";

function bodyChildren(documentXml, label) {
  const body = extractElement(documentXml, "w:body");
  if (!body) {
    throw new AppError(`${label} 缺少 w:body 节点，可能不是有效的 Word 文档。`, { code: "invalid-document", status: 400 });
  }
  const kids = splitTopLevelElements(body.inner);
  let sectPr = null;
  const last = kids[kids.length - 1];
  if (last && last.name === "w:sectPr") {
    sectPr = last.xml;
    kids.pop();
  }
  return { kids, sectPr, body };
}

function makeAllocator(dst, ct, srcCt) {
  const used = new Set(dst.entries());
  return (srcPart, srcType) => {
    const dir = path.posix.dirname(srcPart) === "." ? "" : path.posix.dirname(srcPart);
    const ext = path.posix.extname(srcPart);
    const stem = path.posix.basename(srcPart, ext);
    const prefix = stem.replace(/\d+$/, "") || "part";
    let n = 1;
    let candidate;
    do {
      candidate = `${dir ? `${dir}/` : ""}${prefix}${n}${ext}`;
      n++;
    } while (used.has(candidate));
    used.add(candidate);
    registerPart(ct, candidate, srcType || contentTypeFor(srcCt, srcPart));
    return candidate;
  };
}

/** 合并 styles.xml 中目标不存在的样式定义。 */
async function mergeStyles(dst, src, warnings) {
  const dstXml = await dst.text("word/styles.xml");
  const srcXml = await src.text("word/styles.xml");
  if (!dstXml || !srcXml) return;
  const dstInner = extractElement(dstXml, "w:styles")?.inner || "";
  const srcInner = extractElement(srcXml, "w:styles")?.inner || "";
  const existing = new Set();
  for (const el of splitTopLevelElements(dstInner)) {
    if (el.name !== "w:style") continue;
    const id = /w:styleId\s*=\s*"([^"]+)"/.exec(el.xml)?.[1];
    if (id) existing.add(id);
  }
  const additions = [];
  for (const el of splitTopLevelElements(srcInner)) {
    if (el.name !== "w:style") continue;
    const id = /w:styleId\s*=\s*"([^"]+)"/.exec(el.xml)?.[1];
    if (!id || existing.has(id)) continue;
    existing.add(id);
    additions.push(el.xml);
  }
  if (!additions.length) return;
  const styles = extractElement(dstXml, "w:styles");
  const newXml = `${dstXml.slice(0, styles.end - "</w:styles>".length)}${additions.join("")}</w:styles>${dstXml.slice(styles.end)}`;
  dst.set("word/styles.xml", newXml);
  warnings.push(`已合并 ${additions.length} 个样式定义（同名样式保留基准文档的定义）。`);
}

/** 合并 numbering.xml，返回 源 numId → 目标 numId 的映射。 */
async function mergeNumbering(dst, ct, src, docRels, warnings) {
  const srcXml = await src.text("word/numbering.xml");
  if (!srcXml) return new Map();
  let dstXml = await dst.text("word/numbering.xml");
  const hasRel = docRels.some((r) => r.Type === NUMBERING_REL);

  if (!dstXml) {
    dst.set("word/numbering.xml", srcXml);
    registerPart(ct, "word/numbering.xml", contentTypeFor(await src.contentTypes(), "word/numbering.xml")
      || "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml");
    if (!hasRel) docRels.push({ Id: nextRelId(docRels), Type: NUMBERING_REL, Target: "numbering.xml" });
    return new Map();
  }

  const dstInner = extractElement(dstXml, "w:numbering")?.inner || "";
  const srcInner = extractElement(srcXml, "w:numbering")?.inner || "";
  let maxAbstract = 0;
  let maxNum = 0;
  for (const el of splitTopLevelElements(dstInner)) {
    if (el.name === "w:abstractNum") {
      const id = Number(/w:abstractNumId\s*=\s*"(\d+)"/.exec(el.xml)?.[1] || 0);
      maxAbstract = Math.max(maxAbstract, id);
    } else if (el.name === "w:num") {
      const id = Number(/w:numId\s*=\s*"(\d+)"/.exec(el.xml)?.[1] || 0);
      maxNum = Math.max(maxNum, id);
    }
  }

  const abstractMap = new Map();
  const numMap = new Map();
  const srcChildren = splitTopLevelElements(srcInner);
  for (const el of srcChildren) {
    if (el.name === "w:abstractNum") {
      const id = Number(/w:abstractNumId\s*=\s*"(\d+)"/.exec(el.xml)?.[1]);
      if (Number.isFinite(id)) abstractMap.set(id, ++maxAbstract);
    }
  }
  for (const el of srcChildren) {
    if (el.name === "w:num") {
      const id = Number(/w:numId\s*=\s*"(\d+)"/.exec(el.xml)?.[1]);
      if (Number.isFinite(id)) numMap.set(id, ++maxNum);
    }
  }

  const additions = [];
  let pictureBullets = 0;
  for (const el of srcChildren) {
    if (el.name === "w:abstractNum") {
      const id = Number(/w:abstractNumId\s*=\s*"(\d+)"/.exec(el.xml)?.[1]);
      let xml = el.xml.replace(/(w:abstractNumId\s*=\s*")\d+(")/, `$1${abstractMap.get(id)}$2`);
      xml = xml.replace(/(w:nsid\s+w:val\s*=\s*")([0-9A-Fa-f]{8})(")/, (m, a, v, c) =>
        `${a}${Math.floor(Math.random() * 0xfffffff + 0x10000000).toString(16).toUpperCase().padStart(8, "0")}${c}`);
      if (/w:lvlPicBulletId/.test(xml)) pictureBullets++;
      additions.push(xml);
    } else if (el.name === "w:num") {
      const id = Number(/w:numId\s*=\s*"(\d+)"/.exec(el.xml)?.[1]);
      let xml = el.xml.replace(/(w:numId\s*=\s*")\d+(")/, `$1${numMap.get(id)}$2`);
      xml = xml.replace(/(w:abstractNumId\s+w:val\s*=\s*")(\d+)(")/, (m, a, v, c) =>
        `${a}${abstractMap.get(Number(v)) ?? v}${c}`);
      additions.push(xml);
    }
  }
  if (pictureBullets) warnings.push("源文档中的图片项目符号未随编号定义迁移。");

  const numbering = extractElement(dstXml, "w:numbering");
  dstXml = `${dstXml.slice(0, numbering.end - "</w:numbering>".length)}${additions.join("")}</w:numbering>${dstXml.slice(numbering.end)}`;
  dst.set("word/numbering.xml", dstXml);
  if (!hasRel) docRels.push({ Id: nextRelId(docRels), Type: NUMBERING_REL, Target: "numbering.xml" });
  return numMap;
}

/**
 * @param {{inputPaths:string[], outputDir:string, options?:any,
 *          progress?:(r:number,s?:string)=>void, warnings?:string[]}} cfg
 */
export async function mergeDocx({ inputPaths, outputDir, options = {}, progress, warnings = [] }) {
  const basePath = inputPaths[0];
  const dst = await OoxmlPackage.load(await fs.readFile(basePath), path.basename(basePath));
  const ct = await dst.contentTypes();
  const baseDocXml = await dst.text("word/document.xml");
  if (!baseDocXml) throw new AppError("基准文件不是有效的 Word 文档。", { code: "invalid-document", status: 400 });

  const { kids, sectPr } = bodyChildren(baseDocXml, path.basename(basePath));
  const mergedKids = kids.map((k) => k.xml);
  const docRels = await dst.readRels("word/document.xml");

  for (let i = 1; i < inputPaths.length; i++) {
    const inputPath = inputPaths[i];
    const label = path.basename(inputPath);
    progress?.(0.1 + 0.8 * (i / inputPaths.length), `追加 ${label}`);
    // eslint-disable-next-line no-await-in-loop
    const src = await OoxmlPackage.load(await fs.readFile(inputPath), label);
    // eslint-disable-next-line no-await-in-loop
    const srcDocXml = await src.text("word/document.xml");
    if (!srcDocXml) {
      warnings.push(`${label} 不是 Word 文档，已跳过。`);
      continue;
    }
    const srcBody = bodyChildren(srcDocXml, label);
    // eslint-disable-next-line no-await-in-loop
    const srcCt = await src.contentTypes();
    // eslint-disable-next-line no-await-in-loop
    const srcDocRels = await src.readRels("word/document.xml");
    const ctx = createContext({ src, dst, srcContentTypes: srcCt, allocName: makeAllocator(dst, ct, srcCt) });
    const relMap = new Map();

    for (const rel of srcDocRels) {
      if (SKIP_REL_TYPES.some((re) => re.test(rel.Type))) continue;
      if (rel.TargetMode === "External" || /^[a-z]+:/i.test(rel.Target)) {
        const id = nextRelId(docRels);
        docRels.push({ Id: id, Type: rel.Type, Target: rel.Target, TargetMode: "External" });
        relMap.set(rel.Id, id);
        continue;
      }
      const target = resolvePartTarget("word/document.xml", rel.Target);
      if (!target) continue;
      // eslint-disable-next-line no-await-in-loop
      const copied = await copyPartDeep(target, ctx);
      if (!copied) continue;
      const id = nextRelId(docRels);
      docRels.push({ Id: id, Type: rel.Type, Target: relativeTarget("word/document.xml", copied.part) });
      relMap.set(rel.Id, id);
    }

    // eslint-disable-next-line no-await-in-loop
    const numMap = await mergeNumbering(dst, ct, src, docRels, warnings);
    // eslint-disable-next-line no-await-in-loop
    await mergeStyles(dst, src, warnings);

    let appended = srcBody.kids.map((k) => k.xml).join("");
    appended = rewriteRelRefs(appended, relMap);
    if (numMap.size) {
      appended = appended.replace(/(<w:numId[^>]*w:val=")(\d+)(")/g, (full, a, v, c) =>
        numMap.has(Number(v)) ? `${a}${numMap.get(Number(v))}${c}` : full);
    }
    if (/<w:(footnoteReference|endnoteReference|commentReference)\b/.test(appended)) {
      warnings.push(`${label} 含脚注/尾注/批注引用，合并后可能丢失对应内容。`);
    }
    if (options.pageBreak !== false && mergedKids.length) mergedKids.push(PAGE_BREAK);
    mergedKids.push(appended);
  }

  const body = extractElement(baseDocXml, "w:body");
  const newInner = `${mergedKids.join("")}${sectPr || ""}`;
  const newDocXml = `${baseDocXml.slice(0, body.start)}${newInner}${baseDocXml.slice(body.end)}`;
  dst.set("word/document.xml", newDocXml);
  await dst.writeRels("word/document.xml", docRels);
  dst.set("[Content_Types].xml", serializeContentTypes(ct));

  const name = options.outputName ? baseName(options.outputName) : `${baseName(basePath)}-merged`;
  const target = await uniquePath(outputDir, `${name}.docx`);
  await fs.writeFile(target, await dst.save());
  progress?.(1, "完成");
  return { files: [{ name: path.basename(target), path: target, kind: "docx" }] };
}

/** Word 拆分：按标题级别或分节符。 */
import fs from "node:fs/promises";
import path from "node:path";
import {
  OoxmlPackage, extractElement, splitTopLevelElements, serializeContentTypes,
} from "../ooxml/package.js";
import { garbageCollect } from "../ooxml/gc.js";
import { AppError } from "../errors.js";
import { baseName, safeName, uniquePath } from "../paths.js";

const HEADING_STYLE_PATTERNS = (level) => [
  `Heading${level}`, `heading ${level}`, `标题 ${level}`, `标题${level}`, `berschrift${level}`,
];

function headingLevelOf(paragraphXml) {
  const styleMatch = /<w:pStyle\b[^>]*w:val="([^"]+)"/.exec(paragraphXml);
  const style = styleMatch?.[1] || "";
  for (let level = 1; level <= 6; level++) {
    if (HEADING_STYLE_PATTERNS(level).some((p) => p.toLowerCase() === style.toLowerCase())) return level;
  }
  const outline = /<w:outlineLvl\b[^>]*w:val="(\d+)"/.exec(paragraphXml);
  if (outline) return Number(outline[1]) + 1;
  return 0;
}

function paragraphText(paragraphXml) {
  const parts = [...paragraphXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => m[1]);
  return parts.join("").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").trim();
}

function isSectionBreak(paragraphXml) {
  return /<w:pPr>[\s\S]*<w:sectPr\b/.test(paragraphXml) || /<w:sectPr\b[^>]*w:type="(nextPage|continuous|evenPage|oddPage)"/.test(paragraphXml);
}

export async function splitDocx({ inputPath, outputDir, strategy = "heading", param, progress, warnings = [] }) {
  const buffer = await fs.readFile(inputPath);
  const pkg = await OoxmlPackage.load(buffer, path.basename(inputPath));
  const ct = await pkg.contentTypes();
  const documentXml = await pkg.text("word/document.xml");
  if (!documentXml) throw new AppError("不是有效的 Word 文档。", { code: "invalid-document", status: 400 });

  const body = extractElement(documentXml, "w:body");
  const children = splitTopLevelElements(body.inner);
  let trailingSectPr = null;
  if (children.length && children[children.length - 1].name === "w:sectPr") {
    trailingSectPr = children.pop().xml;
  }

  /** @type {Array<{title:string, nodes:string[]}>} */
  const groups = [];
  const level = Math.max(1, Math.min(6, Number(param) || 1));
  let current = { title: "part", nodes: [] };

  for (const child of children) {
    if (strategy === "heading" && child.name === "w:p") {
      const headingLevel = headingLevelOf(child.xml);
      if (headingLevel > 0 && headingLevel <= level) {
        if (current.nodes.length) groups.push(current);
        const text = paragraphText(child.xml) || `part${groups.length + 1}`;
        current = { title: text.slice(0, 40), nodes: [child.xml] };
        continue;
      }
    }
    current.nodes.push(child.xml);
    if (strategy === "section" && child.name === "w:p" && isSectionBreak(child.xml)) {
      groups.push(current);
      current = { title: `part${groups.length + 1}`, nodes: [] };
    }
  }
  if (current.nodes.length) groups.push(current);
  if (!groups.length) {
    throw new AppError("没有找到可拆分的分界点。", {
      status: 400,
      code: "bad-request",
      hint: strategy === "heading" ? `文档中没有“标题 ${level}”样式的段落。` : "文档中没有分节符。",
    });
  }
  const meaningful = groups.filter((g) => g.nodes.length);
  if (meaningful.length !== groups.length) groups.length = meaningful.length;
  const stem = baseName(inputPath);
  const files = [];

  for (let i = 0; i < groups.length; i++) {
    const group = groups[i];
    const hasSectPr = group.nodes.some((n) => /<w:sectPr\b/.test(n));
    const inner = `${group.nodes.join("")}${hasSectPr ? "" : trailingSectPr || ""}`;
    // eslint-disable-next-line no-await-in-loop
    const outPkg = await OoxmlPackage.load(buffer, path.basename(inputPath));
    // eslint-disable-next-line no-await-in-loop
    const outCt = await outPkg.contentTypes();
    const newXml = `${documentXml.slice(0, body.start)}${inner}${documentXml.slice(body.end)}`;
    outPkg.set("word/document.xml", newXml);
    // eslint-disable-next-line no-await-in-loop
    await garbageCollect(outPkg, outCt);
    outPkg.set("[Content_Types].xml", serializeContentTypes(outCt));
    const suffix = `${String(i + 1).padStart(2, "0")}-${safeName(group.title, "part")}`.slice(0, 60);
    // eslint-disable-next-line no-await-in-loop
    const target = await uniquePath(outputDir, `${stem}-${suffix}.docx`);
    // eslint-disable-next-line no-await-in-loop
    await fs.writeFile(target, await outPkg.save());
    files.push({ name: path.basename(target), path: target, kind: "docx" });
    progress?.(0.1 + 0.85 * ((i + 1) / groups.length), `生成 ${path.basename(target)}`);
  }
  return { files };
}

export { headingLevelOf, paragraphText, isSectionBreak };

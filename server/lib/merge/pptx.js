/** PowerPoint (.pptx) 合并：把源演示文稿的幻灯片（含版式/母版/媒体）搬到基准演示文稿。 */
import fs from "node:fs/promises";
import path from "node:path";
import {
  OoxmlPackage, relativeTarget, nextRelId, serializeContentTypes, registerPart,
  contentTypeFor, resolvePartTarget,
} from "../ooxml/package.js";
import { createContext, copyPartDeep } from "../ooxml/partCopy.js";
import { AppError } from "../errors.js";
import { baseName, uniquePath } from "../paths.js";

const NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SLIDE_REL = `${NS}/slide`;
const LAYOUT_REL = `${NS}/slideLayout`;
const MASTER_REL = `${NS}/slideMaster`;

const PRESENTATION = "ppt/presentation.xml";

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

function listSlideRelIds(presentationXml) {
  const ids = [];
  const re = /<p:sldId\b[^>]*r:id="([^"]+)"[^>]*\/?>/g;
  let m;
  while ((m = re.exec(presentationXml))) ids.push(m[1]);
  return ids;
}

function maxNumericAttr(xml, pattern, fallback) {
  let max = fallback;
  const re = new RegExp(pattern, "g");
  let m;
  while ((m = re.exec(xml))) {
    const value = Number(m[1]);
    if (Number.isFinite(value)) max = Math.max(max, value);
  }
  return max;
}

function appendIntoList(xml, tag, items) {
  if (!items.length) return xml;
  const selfClosing = new RegExp(`<${tag}\\/>`, "i");
  if (selfClosing.test(xml)) return xml.replace(selfClosing, `<${tag}>${items.join("")}</${tag}>`);
  const close = `</${tag}>`;
  const index = xml.indexOf(close);
  if (index < 0) return xml;
  return `${xml.slice(0, index)}${items.join("")}${xml.slice(index)}`;
}

/**
 * @param {{inputPaths:string[], outputDir:string, options?:any,
 *          progress?:(r:number,s?:string)=>void, warnings?:string[]}} cfg
 */
export async function mergePptx({ inputPaths, outputDir, options = {}, progress, warnings = [] }) {
  const basePath = inputPaths[0];
  const dst = await OoxmlPackage.load(await fs.readFile(basePath), path.basename(basePath));
  const ct = await dst.contentTypes();
  let presentationXml = await dst.text(PRESENTATION);
  if (!presentationXml) throw new AppError("基准文件不是有效的 PowerPoint 演示文稿。", { code: "invalid-document", status: 400 });
  const presRels = await dst.readRels(PRESENTATION);

  let nextSlideId = maxNumericAttr(presentationXml, /<p:sldId\b[^>]*\bid="(\d+)"/g, 255) + 1;
  let nextMasterId = maxNumericAttr(presentationXml, /<p:sldMasterId\b[^>]*\bid="(\d+)"/g, 2147483647) + 1;
  const registeredMasters = new Set();
  const newSlideIds = [];
  const newMasterIds = [];

  for (let i = 1; i < inputPaths.length; i++) {
    const inputPath = inputPaths[i];
    const label = path.basename(inputPath);
    progress?.(0.1 + 0.8 * (i / inputPaths.length), `追加 ${label}`);
    // eslint-disable-next-line no-await-in-loop
    const src = await OoxmlPackage.load(await fs.readFile(inputPath), label);
    // eslint-disable-next-line no-await-in-loop
    const srcPresentation = await src.text(PRESENTATION);
    if (!srcPresentation) {
      warnings.push(`${label} 不是 PowerPoint 演示文稿，已跳过。`);
      continue;
    }
    // eslint-disable-next-line no-await-in-loop
    const srcRels = await src.readRels(PRESENTATION);
    // eslint-disable-next-line no-await-in-loop
    const srcCt = await src.contentTypes();
    const ctx = createContext({ src, dst, srcContentTypes: srcCt, allocName: makeAllocator(dst, ct, srcCt) });

    const slideRelIds = listSlideRelIds(srcPresentation);
    let added = 0;
    for (const relId of slideRelIds) {
      const rel = srcRels.find((r) => r.Id === relId);
      if (!rel) continue;
      const slidePart = resolvePartTarget(PRESENTATION, rel.Target);
      if (!slidePart) continue;
      // eslint-disable-next-line no-await-in-loop
      const copied = await copyPartDeep(slidePart, ctx);
      if (!copied) continue;
      const newRelId = nextRelId(presRels);
      presRels.push({ Id: newRelId, Type: SLIDE_REL, Target: relativeTarget(PRESENTATION, copied.part) });
      newSlideIds.push(`<p:sldId id="${nextSlideId++}" r:id="${newRelId}"/>`);
      added++;

      // 确保其版式所属母版已在 presentation.xml 中注册
      const srcSlideRels = await src.readRels(slidePart);
      const layoutRel = srcSlideRels.find((r) => r.Type === LAYOUT_REL);
      if (layoutRel) {
        const srcLayout = resolvePartTarget(slidePart, layoutRel.Target);
        // eslint-disable-next-line no-await-in-loop
        const layoutRels = srcLayout ? await src.readRels(srcLayout) : [];
        const masterRel = layoutRels.find((r) => r.Type === MASTER_REL);
        if (masterRel && srcLayout) {
          const srcMaster = resolvePartTarget(srcLayout, masterRel.Target);
          const copiedMaster = srcMaster ? ctx.map.get(srcMaster) : null;
          if (copiedMaster && !registeredMasters.has(copiedMaster)) {
            registeredMasters.add(copiedMaster);
            const masterRelId = nextRelId(presRels);
            presRels.push({ Id: masterRelId, Type: MASTER_REL, Target: relativeTarget(PRESENTATION, copiedMaster) });
            newMasterIds.push(`<p:sldMasterId id="${nextMasterId++}" r:id="${masterRelId}"/>`);
          }
        }
      }
    }
    if (!added) warnings.push(`${label} 中没有可追加的幻灯片。`);
  }

  presentationXml = appendIntoList(presentationXml, "p:sldIdLst", newSlideIds);
  presentationXml = appendIntoList(presentationXml, "p:sldMasterIdLst", newMasterIds);
  dst.set(PRESENTATION, presentationXml);
  await dst.writeRels(PRESENTATION, presRels);
  dst.set("[Content_Types].xml", serializeContentTypes(ct));

  const name = options.outputName ? baseName(options.outputName) : `${baseName(basePath)}-merged`;
  const target = await uniquePath(outputDir, `${name}.pptx`);
  await fs.writeFile(target, await dst.save());
  progress?.(1, "完成");
  return { files: [{ name: path.basename(target), path: target, kind: "pptx" }] };
}

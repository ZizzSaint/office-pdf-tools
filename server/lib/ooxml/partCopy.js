/** 关系部件深拷贝：把一个 part 及其全部传递关系复制到目标包，并重编号 rId。 */
import {
  relsPathFor, resolvePartTarget, relativeTarget, parseRels, serializeRels, rewriteRelRefs, stripSlash,
} from "./package.js";

const SKIP_CONTENT_TYPE_SUFFIXES = ["/notesSlide"];

/**
 * @typedef {Object} PartCopyContext
 * @property {import('./package.js').OoxmlPackage} src
 * @property {import('./package.js').OoxmlPackage} dst
 * @property {Map<string,string>} map        源 part → 目标 part
 * @property {Map<string,Map<string,string>>} relMaps 源 part → (旧 rId → 新 rId)
 * @property {(srcPart:string, srcType:string)=>string} allocName
 * @property {(srcPart:string, srcType:string)=>boolean} [skipPart]
 * @property {{defaults:Map<string,string>,overrides:Map<string,string>}} [srcContentTypes]
 */

/** 清空上下文（每次合并/拆分复用同一目标包时调用）。 */
export function createContext(init) {
  return {
    map: new Map(),
    relMaps: new Map(),
    ...init,
  };
}

function shouldSkip(ctx, srcPart) {
  const type = ctx.srcContentTypes?.overrides?.get(stripSlash(srcPart)) || "";
  if (SKIP_CONTENT_TYPE_SUFFIXES.some((s) => type.endsWith(s))) return true;
  return ctx.skipPart ? ctx.skipPart(srcPart, type) : false;
}

/**
 * 深拷贝一个 part。
 * @returns {Promise<{part:string, relMap:Map<string,string>}|null>}
 */
export async function copyPartDeep(srcPart, ctx) {
  const norm = stripSlash(srcPart);
  if (ctx.map.has(norm)) {
    return { part: ctx.map.get(norm), relMap: ctx.relMaps.get(norm) || new Map() };
  }
  if (!ctx.src.has(norm) || shouldSkip(ctx, norm)) return null;

  const data = await ctx.src.read(norm);
  if (!data) return null;

  const type = ctx.srcContentTypes?.overrides?.get(norm) || "";
  const dstPart = ctx.allocName(norm, type);
  ctx.map.set(norm, dstPart);
  ctx.relMaps.set(norm, new Map());
  ctx.dst.set(dstPart, data);

  const srcRelsXml = await ctx.src.text(relsPathFor(norm));
  const srcRels = srcRelsXml ? parseRels(srcRelsXml) : [];
  const newRels = [];
  const relMap = new Map();
  let counter = 1;

  for (const rel of srcRels) {
    const newId = `rId${counter}`;
    if (rel.TargetMode === "External" || /^[a-z]+:/i.test(rel.Target)) {
      counter++;
      relMap.set(rel.Id, newId);
      newRels.push({ Id: newId, Type: rel.Type, Target: rel.Target, TargetMode: rel.TargetMode || "External" });
      continue;
    }
    const targetPart = resolvePartTarget(norm, rel.Target);
    if (!targetPart) continue;
    // eslint-disable-next-line no-await-in-loop
    const child = await copyPartDeep(targetPart, ctx);
    if (!child) continue;
    counter++;
    relMap.set(rel.Id, newId);
    newRels.push({ Id: newId, Type: rel.Type, Target: relativeTarget(dstPart, child.part) });
  }

  ctx.relMaps.set(norm, relMap);
  if (newRels.length) {
    ctx.dst.set(relsPathFor(dstPart), serializeRels(newRels));
  } else {
    ctx.dst.remove(relsPathFor(dstPart));
  }

  if (/\.xml$/i.test(norm) && relMap.size) {
    const text = data.toString("utf8");
    const rewritten = rewriteRelRefs(text, relMap);
    if (rewritten !== text) ctx.dst.set(dstPart, Buffer.from(rewritten, "utf8"));
  }

  return { part: dstPart, relMap };
}

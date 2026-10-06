/** OOXML 包垃圾回收：从包根关系出发做可达性分析，删除孤立部件。 */
import { relsPathFor, parseRels, resolvePartTarget, unregisterPart, stripSlash } from "./package.js";

/**
 * @param {import('./package.js').OoxmlPackage} pkg
 * @param {{overrides:Map<string,string>,defaults:Map<string,string>}} ct
 * @param {{extraRoots?:string[]}} [opts]
 * @returns {Promise<number>} 删除的部件数量
 */
export async function garbageCollect(pkg, ct, opts = {}) {
  const reachable = new Set();
  const queue = ["_rels/.rels", ...(opts.extraRoots || [])];
  while (queue.length) {
    const part = stripSlash(queue.pop());
    if (reachable.has(part)) continue;
    reachable.add(part);
    // eslint-disable-next-line no-await-in-loop
    const relsXml = part === "_rels/.rels"
      ? await pkg.text("_rels/.rels")
      : await pkg.text(relsPathFor(part));
    if (!relsXml) continue;
    const fromPart = part === "_rels/.rels" ? "" : part;
    for (const rel of parseRels(relsXml)) {
      if (rel.TargetMode === "External" || /^[a-z]+:/i.test(rel.Target)) continue;
      const target = resolvePartTarget(fromPart, rel.Target);
      if (target && !reachable.has(target)) queue.push(target);
    }
  }

  let removed = 0;
  for (const part of pkg.entries()) {
    if (part.endsWith(".rels")) continue;
    if (reachable.has(part)) continue;
    pkg.remove(part);
    pkg.remove(relsPathFor(part));
    unregisterPart(ct, part);
    removed++;
  }
  return removed;
}

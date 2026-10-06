/** PowerPoint 拆分：每 N 张 / 幻灯片范围。 */
import fs from "node:fs/promises";
import path from "node:path";
import { OoxmlPackage, serializeContentTypes } from "../ooxml/package.js";
import { garbageCollect } from "../ooxml/gc.js";
import { AppError } from "../errors.js";
import { baseName, uniquePath } from "../paths.js";
import { parseRangeGroups, chunkBySize } from "../ranges.js";

const PRESENTATION = "ppt/presentation.xml";

function slideEntries(presentationXml) {
  const entries = [];
  const re = /<p:sldId\b[^>]*\bid="(\d+)"[^>]*r:id="([^"]+)"[^>]*\/?>/g;
  let m;
  while ((m = re.exec(presentationXml))) entries.push({ id: m[1], relId: m[2], raw: m[0] });
  return entries;
}

function groupsFor(strategy, param, total) {
  const all = Array.from({ length: total }, (_, i) => i);
  if (strategy === "ranges") return parseRangeGroups(param, total);
  return chunkBySize(all, Math.max(1, Number(param) || 10));
}

export async function splitPptx({ inputPath, outputDir, strategy = "every-n-slides", param, progress }) {
  const buffer = await fs.readFile(inputPath);
  const probe = await OoxmlPackage.load(buffer, path.basename(inputPath));
  const presentationXml = await probe.text(PRESENTATION);
  if (!presentationXml) throw new AppError("不是有效的 PowerPoint 演示文稿。", { code: "invalid-document", status: 400 });
  const entries = slideEntries(presentationXml);
  if (!entries.length) throw new AppError("演示文稿中没有幻灯片。", { status: 400, code: "bad-request" });

  const groups = groupsFor(strategy, param, entries.length);
  const stem = baseName(inputPath);
  const pad = String(entries.length).length;
  const files = [];

  for (let i = 0; i < groups.length; i++) {
    const group = groups[i];
    const keepRelIds = new Set(group.map((index) => entries[index].relId));
    // eslint-disable-next-line no-await-in-loop
    const pkg = await OoxmlPackage.load(buffer, path.basename(inputPath));
    // eslint-disable-next-line no-await-in-loop
    const ct = await pkg.contentTypes();
    const xml = await pkg.text(PRESENTATION);
    const all = slideEntries(xml);
    const keptRels = new Set();
    const newEntries = [];
    for (const entry of all) {
      if (keepRelIds.has(entry.relId)) {
        keptRels.add(entry.relId);
        newEntries.push(entry.raw);
      }
    }
    let newXml = xml;
    // 删除不需要的 sldId 元素
    for (const entry of all) {
      if (keepRelIds.has(entry.relId)) continue;
      newXml = newXml.replace(entry.raw, "");
    }
    // 重写 sldIdLst
    const listMatch = /<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/.exec(newXml);
    if (listMatch) {
      newXml = `${newXml.slice(0, listMatch.index)}<p:sldIdLst>${newEntries.join("")}</p:sldIdLst>${newXml.slice(listMatch.index + listMatch[0].length)}`;
    }
    pkg.set(PRESENTATION, newXml);
    // 删除未保留幻灯片的演示文稿级关系
    const rels = await pkg.readRels(PRESENTATION);
    const filtered = rels.filter((rel) => !/\/slide$/.test(rel.Type) || keptRels.has(rel.Id));
    await pkg.writeRels(PRESENTATION, filtered);
    // 删除孤立部件（被剔除的幻灯片及其专属媒体）
    // eslint-disable-next-line no-await-in-loop
    await garbageCollect(pkg, ct);
    pkg.set("[Content_Types].xml", serializeContentTypes(ct));

    const suffix = group.length === 1
      ? `slide${String(group[0] + 1).padStart(pad, "0")}`
      : `slides${group[0] + 1}-${group[group.length - 1] + 1}`;
    // eslint-disable-next-line no-await-in-loop
    const target = await uniquePath(outputDir, `${stem}-${suffix}.pptx`);
    // eslint-disable-next-line no-await-in-loop
    await fs.writeFile(target, await pkg.save());
    files.push({ name: path.basename(target), path: target, kind: "pptx" });
    progress?.(0.1 + 0.85 * ((i + 1) / groups.length), `生成 ${path.basename(target)}`);
  }
  return { files };
}

/** OOXML (docx/xlsx/pptx) 包操作：zip 读写、关系(rels)、内容类型、XML 工具。 */
import path from "node:path";
import JSZip from "jszip";
import { AppError } from "../errors.js";

export function stripSlash(p) {
  return String(p || "").replace(/^\/+/, "").replace(/\\/g, "/");
}

export function partDir(partPath) {
  const dir = path.posix.dirname(stripSlash(partPath));
  return dir === "." ? "" : dir;
}

export function partName(partPath) {
  return path.posix.basename(stripSlash(partPath));
}

export function relsPathFor(partPath) {
  const norm = stripSlash(partPath);
  const dir = partDir(norm);
  return `${dir ? `${dir}/` : ""}_rels/${path.posix.basename(norm)}.rels`;
}

export function resolvePartTarget(fromPart, target) {
  if (/^[a-z]+:/i.test(target)) return null;
  if (target.startsWith("/")) return stripSlash(target);
  const dir = partDir(fromPart);
  return path.posix.normalize(path.posix.join(dir || ".", target));
}

export function relativeTarget(fromPart, toPart) {
  const fromDir = partDir(fromPart) || ".";
  const rel = path.posix.relative(fromDir, stripSlash(toPart));
  return rel || path.posix.basename(stripSlash(toPart));
}

export function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

export function unescapeXml(value) {
  return String(value)
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

export class OoxmlPackage {
  constructor(zip, label = "") {
    this.zip = zip;
    this.label = label;
  }

  static async load(buffer, label = "") {
    try {
      const zip = await JSZip.loadAsync(buffer);
      return new OoxmlPackage(zip, label);
    } catch (err) {
      throw new AppError(`${label || "文件"} 不是有效的 Office (OOXML) 文件。`, {
        code: "invalid-document",
        status: 400,
        cause: err,
      });
    }
  }

  entries() {
    return Object.keys(this.zip.files).filter((n) => !this.zip.files[n].dir);
  }

  has(part) {
    return !!this.zip.file(stripSlash(part));
  }

  async read(part) {
    const file = this.zip.file(stripSlash(part));
    if (!file) return null;
    return Buffer.from(await file.async("nodebuffer"));
  }

  async text(part) {
    const file = this.zip.file(stripSlash(part));
    if (!file) return null;
    return file.async("string");
  }

  set(part, data) {
    this.zip.file(stripSlash(part), data);
  }

  remove(part) {
    this.zip.remove(stripSlash(part));
  }

  async readRels(part) {
    const relsPart = relsPathFor(part);
    const xml = await this.text(relsPart);
    return xml ? parseRels(xml) : [];
  }

  async writeRels(part, rels) {
    const relsPart = relsPathFor(part);
    if (!rels.length) {
      this.remove(relsPart);
      return;
    }
    this.set(relsPart, serializeRels(rels));
  }

  async contentTypes() {
    const xml = await this.text("[Content_Types].xml");
    if (!xml) {
      throw new AppError("文件缺少 [Content_Types].xml，可能已损坏。", { code: "invalid-document", status: 400 });
    }
    return parseContentTypes(xml);
  }

  async save() {
    return this.zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
  }

  async cloneFrom(buffer) {
    return OoxmlPackage.load(buffer, this.label);
  }
}

/* ------------------------------------------------------------------ rels */

export function parseRels(xml) {
  const rels = [];
  const re = /<Relationship\b([^>]*?)\/?>/g;
  let match;
  while ((match = re.exec(xml))) {
    const attrs = {};
    const attrRe = /([\w:.-]+)\s*=\s*"([^"]*)"/g;
    let a;
    while ((a = attrRe.exec(match[1]))) attrs[a[1]] = unescapeXml(a[2]);
    if (!attrs.Id) continue;
    rels.push({ Id: attrs.Id, Type: attrs.Type || "", Target: attrs.Target || "", TargetMode: attrs.TargetMode });
  }
  return rels;
}

export function serializeRels(rels) {
  const body = rels
    .map((r) => `<Relationship Id="${escapeXml(r.Id)}" Type="${escapeXml(r.Type)}" Target="${escapeXml(r.Target)}"${r.TargetMode ? ` TargetMode="${escapeXml(r.TargetMode)}"` : ""}/>`)
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
}

export function nextRelId(rels) {
  let max = 0;
  for (const rel of rels) {
    const m = /^rId(\d+)$/i.exec(rel.Id);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return `rId${max + 1}`;
}

/** 重写 XML 中引用 rId 的属性。 */
export function rewriteRelRefs(xml, relMap) {
  if (!relMap || !relMap.size) return xml;
  return xml.replace(
    /(\s[\w:.-]*(?::id|:embed|:link|:relid|:pict|:dm|:lo|:qs|:cs|:href)\s*=\s*")(rId\d+)(")/gi,
    (full, prefix, id, suffix) => {
      const mapped = relMap.get(id) || relMap.get(id.toLowerCase());
      return mapped ? `${prefix}${mapped}${suffix}` : full;
    },
  );
}

/** 收集 XML 中出现的所有 rId。 */
export function collectRelIds(xml) {
  const ids = new Set();
  const re = /\s[\w:.-]*(?::id|:embed|:link|:relid|:pict|:dm|:lo|:qs|:cs|:href)\s*=\s*"(rId\d+)"/gi;
  let m;
  while ((m = re.exec(xml))) ids.add(m[1]);
  return ids;
}

/* -------------------------------------------------------- content types */

export function parseContentTypes(xml) {
  const defaults = new Map();
  const overrides = new Map();
  const defaultsRe = /<Default\b([^>]*)\/?>/g;
  let m;
  while ((m = defaultsRe.exec(xml))) {
    const ext = /Extension\s*=\s*"([^"]*)"/i.exec(m[1]);
    const type = /ContentType\s*=\s*"([^"]*)"/i.exec(m[1]);
    if (ext && type) defaults.set(ext[1].toLowerCase(), unescapeXml(type[1]));
  }
  const overridesRe = /<Override\b([^>]*)\/?>/g;
  while ((m = overridesRe.exec(xml))) {
    const name = /PartName\s*=\s*"([^"]*)"/i.exec(m[1]);
    const type = /ContentType\s*=\s*"([^"]*)"/i.exec(m[1]);
    if (name && type) overrides.set(stripSlash(unescapeXml(name[1])), unescapeXml(type[1]));
  }
  return { xml, defaults, overrides };
}

export function contentTypeFor(ct, partPath) {
  const norm = stripSlash(partPath);
  if (ct.overrides.has(norm)) return ct.overrides.get(norm);
  const ext = path.posix.extname(norm).replace(".", "").toLowerCase();
  return ct.defaults.get(ext) || "";
}

export function serializeContentTypes(ct) {
  const defaults = [...ct.defaults.entries()]
    .map(([ext, type]) => `<Default Extension="${escapeXml(ext)}" ContentType="${escapeXml(type)}"/>`)
    .join("");
  const overrides = [...ct.overrides.entries()]
    .map(([name, type]) => `<Override PartName="/${escapeXml(name)}" ContentType="${escapeXml(type)}"/>`)
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}${overrides}</Types>`;
}

export function registerPart(ct, partPath, type) {
  const norm = stripSlash(partPath);
  const ext = path.posix.extname(norm).replace(".", "").toLowerCase();
  if (type) ct.overrides.set(norm, type);
  if (ext && !ct.defaults.has(ext) && type) {
    // 只有明确是二进制扩展名时才用 Default；XML 部分统一用 Override。
    if (!["xml", "rels"].includes(ext)) ct.defaults.set(ext, type);
  }
}

export function unregisterPart(ct, partPath) {
  ct.overrides.delete(stripSlash(partPath));
}

/* --------------------------------------------------------------- XML 工具 */

/**
 * 把 XML 的顶层子元素切成字符串数组（用于 Word body / slide 列表等）。
 * @returns {Array<{name:string, xml:string, selfClosing:boolean, inner:string}>}
 */
export function splitTopLevelElements(xml) {
  const out = [];
  const re = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\/([\w:.-]+)\s*>|<([\w:.-]+)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/g;
  let depth = 0;
  let start = -1;
  let name = "";
  let match;
  while ((match = re.exec(xml))) {
    const full = match[0];
    if (full.startsWith("<!--") || full.startsWith("<![")) continue;
    const isClosing = match[1] !== undefined;
    if (!isClosing) {
      const tagName = match[2];
      const selfClosing = match[4] === "/";
      if (depth === 0) {
        start = match.index;
        name = tagName;
        if (selfClosing) {
          // 顶层的自闭合元素（如 <w:bookmarkStart/>）直接产出一个元素
          out.push({ name: tagName, xml: full, selfClosing: true, inner: "" });
          start = -1;
          continue;
        }
      }
      // 自闭合标签不改变嵌套深度，否则后续顶层元素会被吞掉
      if (!selfClosing) depth++;
      continue;
    }
    depth--;
    if (depth <= 0 && start >= 0) {
      const endIndex = match.index + full.length;
      const fullXml = xml.slice(start, endIndex);
      const innerStart = fullXml.indexOf(">") + 1;
      const inner = fullXml.slice(innerStart, fullXml.length - full.length);
      out.push({ name, xml: fullXml, selfClosing: false, inner });
      start = -1;
      depth = 0;
    }
  }
  return out;
}

/** 取 <tag ...>...</tag> 内的内容（不校验嵌套，用于结构固定的 OOXML 容器）。 */
export function extractElement(xml, tag) {
  const open = new RegExp(`<${tag}(\\s[^>]*)?>`, "i");
  const m = open.exec(xml);
  if (!m) return null;
  const start = m.index + m[0].length;
  const closeTag = `</${tag}>`;
  const end = xml.indexOf(closeTag, start);
  if (end < 0) return null;
  return { inner: xml.slice(start, end), start, end: end + closeTag.length, openTag: m[0] };
}

export function replaceElement(xml, tag, newInner) {
  const found = extractElement(xml, tag);
  if (!found) return xml;
  return xml.slice(0, found.start) + newInner + xml.slice(found.end);
}

export function escapeAttr(value) {
  return escapeXml(value).replace(/[\r\n]/g, " ");
}

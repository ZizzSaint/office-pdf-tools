/** 页码范围解析：支持 "1-3,5,8-"、"all"、单页。 */

/** 解析为 0 基页码数组（去重、升序）。 */
export function parsePageList(spec, total) {
  const text = String(spec ?? "").trim().toLowerCase();
  if (!text || text === "all" || text === "*") {
    return Array.from({ length: total }, (_, i) => i);
  }
  const out = new Set();
  let matched = 0;
  for (const chunk of text.split(/[,;\s]+/).filter(Boolean)) {
    const m = /^(\d+)?\s*(?:-|~|—|到)\s*(\d+)?$/.exec(chunk);
    if (m) {
      matched++;
      const start = m[1] ? Number(m[1]) : 1;
      const end = m[2] ? Number(m[2]) : total;
      for (let p = Math.max(1, start); p <= Math.min(total, end); p++) out.add(p - 1);
      continue;
    }
    if (/^\d+$/.test(chunk)) {
      matched++;
      const p = Number(chunk);
      if (p >= 1 && p <= total) out.add(p - 1);
    }
  }
  if (!matched) throw new Error(`无法解析页码范围: ${spec}`);
  return [...out].sort((a, b) => a - b);
}

/** 解析为“分组”页码范围（用于拆分）：1-3,4-6,7- → [[0,1,2],[3,4,5],[6..]] 。 */
export function parseRangeGroups(spec, total) {
  const text = String(spec ?? "").trim();
  if (!text) throw new Error("请填写页码范围，例如 1-3,4-6");
  const groups = [];
  for (const chunk of text.split(/[,;]/).map((s) => s.trim()).filter(Boolean)) {
    const m = /^(\d+)?\s*(?:-|~|—|到)\s*(\d+)?$/.exec(chunk);
    if (m) {
      const start = m[1] ? Number(m[1]) : 1;
      const end = m[2] ? Number(m[2]) : total;
      const list = [];
      for (let p = Math.max(1, start); p <= Math.min(total, end); p++) list.push(p - 1);
      if (list.length) groups.push(list);
      continue;
    }
    if (/^\d+$/.test(chunk)) {
      const p = Number(chunk);
      if (p >= 1 && p <= total) groups.push([p - 1]);
    }
  }
  if (!groups.length) throw new Error(`无法解析页码范围: ${spec}`);
  return groups;
}

/** 每 n 个一组。 */
export function chunkBySize(list, n) {
  const size = Math.max(1, Number(n) || 1);
  const out = [];
  for (let i = 0; i < list.length; i += size) out.push(list.slice(i, i + size));
  return out;
}

/** [0,1,2,4] → "1-3,5" */
export function formatPageList(pages) {
  const sorted = [...new Set(pages)].sort((a, b) => a - b);
  const parts = [];
  let start = null;
  let prev = null;
  for (const p of sorted) {
    if (start === null) { start = p; prev = p; continue; }
    if (p === prev + 1) { prev = p; continue; }
    parts.push(start === prev ? `${start + 1}` : `${start + 1}-${prev + 1}`);
    start = p;
    prev = p;
  }
  if (start !== null) parts.push(start === prev ? `${start + 1}` : `${start + 1}-${prev + 1}`);
  return parts.join(",");
}

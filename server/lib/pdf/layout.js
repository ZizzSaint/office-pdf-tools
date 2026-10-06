/** PDF 文本版面重建：把带坐标的文本片段还原为行、段落、标题与表格。 */

const CJK = /[\u3000-\u9fff\uff00-\uffef]/;

function hasCJK(text) {
  return CJK.test(text || "");
}

/**
 * 把一页的文本片段聚合成行。
 * @param {Array<{text:string,x:number,y:number,width:number,height:number,fontSize:number,bold:boolean,italic:boolean}>} items
 */
export function buildLines(items) {
  if (!items?.length) return [];
  const sorted = [...items].sort((a, b) => (b.y - a.y) || (a.x - b.x));
  const lines = [];
  for (const item of sorted) {
    const tol = Math.max(1.5, item.fontSize * 0.55);
    let line = lines.find((l) => Math.abs(l.y - item.y) <= tol && Math.abs(l.y - item.y) <= Math.max(tol, l.fontSize * 0.55));
    if (!line) {
      line = { y: item.y, fontSize: item.fontSize, items: [] };
      lines.push(line);
    }
    line.items.push(item);
    line.y = line.items.reduce((sum, it) => sum + it.y, 0) / line.items.length;
    line.fontSize = Math.max(line.fontSize, item.fontSize);
  }

  const result = [];
  for (const line of lines) {
    const parts = [...line.items].sort((a, b) => a.x - b.x);
    const cells = [];
    let current = null;
    let pendingSpace = false;
    for (const part of parts) {
      if (part.isSpace) {
        // pdf.js 会把文字定位的“跳跃”也变成一个很宽的空白项（跨列时宽达上百 pt），
        // 因此空白项只在不宽于正常空格时作为“有空格”的提示，且绝不扩展当前单元格的覆盖范围。
        const normalSpace = part.width <= Math.max(6, part.fontSize * 3);
        if (current && normalSpace) pendingSpace = true;
        continue;
      }
      const text = part.text;
      if (!current) {
        current = { text, x: part.x, width: part.width, fontSize: part.fontSize };
        continue;
      }
      const gap = part.x - (current.x + current.width);
      const spaceThreshold = Math.max(1.2, part.fontSize * 0.33);
      const cellThreshold = Math.max(8, part.fontSize * 1.8);
      if (gap > cellThreshold) {
        cells.push(current);
        current = { text, x: part.x, width: part.width, fontSize: part.fontSize };
        pendingSpace = false;
        continue;
      }
      const shouldSpace = pendingSpace || gap > spaceThreshold;
      if (shouldSpace) {
        current.text = `${current.text} ${text}`;
      } else {
        const needsSpace = !hasCJK(current.text) && !hasCJK(text)
          && /[A-Za-z0-9]$/.test(current.text) && /^[A-Za-z0-9]/.test(text);
        current.text = needsSpace ? `${current.text} ${text}` : `${current.text}${text}`;
      }
      current.width = part.x + part.width - current.x;
      current.fontSize = Math.max(current.fontSize, part.fontSize);
      pendingSpace = false;
    }
    if (current) cells.push(current);

    const boldRatio = line.items.filter((it) => it.bold).length / Math.max(1, line.items.length);
    const text = cells.map((c) => c.text).join(hasCJK(cells[0]?.text || "") ? "  " : "    ").trim();
    if (!text) continue;
    result.push({
      y: line.y,
      fontSize: line.fontSize,
      bold: boldRatio >= 0.5,
      italic: line.items.filter((it) => it.italic).length / line.items.length >= 0.5,
      x: Math.min(...cells.map((c) => c.x)),
      x2: Math.max(...cells.map((c) => c.x + c.width)),
      cells,
      text,
    });
  }
  return result.sort((a, b) => (b.y - a.y) || (a.x - b.x));
}

function weightedMedianFontSize(lines) {
  const buckets = new Map();
  for (const line of lines) {
    const key = Math.round(line.fontSize);
    buckets.set(key, (buckets.get(key) || 0) + line.text.length);
  }
  const entries = [...buckets.entries()].sort((a, b) => a[0] - b[0]);
  const total = entries.reduce((s, [, n]) => s + n, 0);
  let acc = 0;
  for (const [size, n] of entries) {
    acc += n;
    if (acc >= total / 2) return size || 10;
  }
  return 10;
}

function looksLikeTable(lines) {
  if (lines.length < 2) return false;
  const counts = lines.map((l) => l.cells.length);
  const min = Math.min(...counts);
  if (min < 2) return false;
  const consistent = counts.filter((c) => c === counts[0]).length >= Math.max(2, lines.length - 1);
  return consistent;
}

/**
 * 把行序列转换成文档块（标题 / 段落 / 表格）。
 * @returns {Array<{type:'heading'|'paragraph'|'table'|'pagebreak', level?:number, text?:string, rows?:string[][], y?:number}>}
 */
export function detectBlocks(lines) {
  const blocks = [];
  if (!lines.length) return blocks;
  const body = weightedMedianFontSize(lines);

  let i = 0;
  let paragraph = [];

  const flush = () => {
    if (paragraph.length) {
      blocks.push({ type: "paragraph", text: paragraph.join(hasCJK(paragraph.join("")) ? "" : " ") });
      paragraph = [];
    }
  };

  while (i < lines.length) {
    const line = lines[i];
    const size = line.fontSize;
    const ratio = size / (body || 10);
    const short = line.text.length <= 80;

    // 表格：连续多行且列数一致
    if (line.cells.length >= 2) {
      let j = i;
      while (j < lines.length && lines[j].cells.length >= 2) j++;
      const run = lines.slice(i, j);
      if (looksLikeTable(run)) {
        flush();
        blocks.push({ type: "table", rows: run.map((l) => l.cells.map((c) => c.text.trim())) });
        i = j;
        continue;
      }
    }

    const numbered = /^(\d+(\.\d+)*)[\s.、]\s*/.exec(line.text);
    let level = 0;
    if (short) {
      if (ratio >= 1.5) level = 1;
      else if (ratio >= 1.28) level = 2;
      else if (ratio >= 1.12 && (line.bold || numbered)) level = 3;
      else if (line.bold && ratio >= 1.02 && line.text.length <= 40) level = 4;
      else if (numbered && line.text.length <= 60) {
        const depth = numbered[1].split(".").length;
        level = Math.min(4, depth + 1);
      }
    }

    if (level > 0) {
      flush();
      blocks.push({ type: "heading", level, text: line.text });
      i++;
      continue;
    }

    // 段落合并：下一行与当前行垂直间距合理且都不是标题/表格时继续累积
    paragraph.push(line.text);
    const next = lines[i + 1];
    if (!next) {
      flush();
      i++;
      continue;
    }
    const gap = line.y - next.y;
    const expected = Math.max(line.fontSize, next.fontSize) * 1.9;
    const nextRatio = next.fontSize / (body || 10);
    const nextNumbered = /^(\d+(\.\d+)*)[\s.、]\s*/.test(next.text);
    const endsSentence = /[。！？.!?;；:：]$/.test(line.text);
    if (gap > expected || nextRatio >= 1.12 || next.cells.length >= 2 || (endsSentence && next.text.length > 0 && next.x <= line.x + 2)) {
      flush();
    }
    i++;
  }
  flush();
  return blocks;
}

/** 把行按 x 起点聚成列，返回矩阵（用于 PDF → Excel）。 */
export function linesToMatrix(lines, { tolerance = 3 } = {}) {
  const starts = [];
  for (const line of lines) {
    for (const cell of line.cells) {
      const found = starts.find((s) => Math.abs(s.x - cell.x) <= tolerance);
      if (found) found.x = (found.x + cell.x) / 2;
      else starts.push({ x: cell.x });
    }
  }
  starts.sort((a, b) => a.x - b.x);
  if (!starts.length) return { columns: [], rows: [] };
  const columns = starts.map((s) => s.x);
  const rows = lines.map((line) => {
    const row = new Array(columns.length).fill("");
    for (const cell of line.cells) {
      let idx = 0;
      let best = Infinity;
      columns.forEach((x, k) => {
        const d = Math.abs(x - cell.x);
        if (d < best) { best = d; idx = k; }
      });
      row[idx] = row[idx] ? `${row[idx]} ${cell.text}` : cell.text;
    }
    return row;
  });
  return { columns, rows };
}

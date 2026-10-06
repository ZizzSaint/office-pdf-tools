/** PDF 拆分：按每 N 页 / 页码范围 / 每页。 */
import fs from "node:fs/promises";
import path from "node:path";
import { PDFDocument } from "pdf-lib";
import { AppError } from "../errors.js";
import { baseName, uniquePath } from "../paths.js";
import { parseRangeGroups, chunkBySize, formatPageList } from "../ranges.js";

function groupsFor(strategy, param, totalPages) {
  const all = Array.from({ length: totalPages }, (_, i) => i);
  switch (strategy) {
    case "each-page":
      return all.map((p) => [p]);
    case "ranges":
      return parseRangeGroups(param, totalPages);
    case "every-n-pages":
    default: {
      const n = Math.max(1, Number(param) || 1);
      return chunkBySize(all, n);
    }
  }
}

export async function splitPdf({ inputPath, outputDir, strategy = "every-n-pages", param, progress }) {
  const bytes = await fs.readFile(inputPath);
  const src = await PDFDocument.load(bytes, { ignoreEncryption: true }).catch((err) => {
    throw new AppError("PDF 无法读取（可能已加密或损坏）。", { code: "invalid-document", status: 400, cause: err });
  });
  const total = src.getPageCount();
  const groups = groupsFor(strategy, param, total);
  const stem = baseName(inputPath);
  const pad = String(total).length;
  const files = [];

  for (let i = 0; i < groups.length; i++) {
    const group = groups[i];
    const doc = await PDFDocument.create();
    doc.setProducer("Office PDF Tools");
    // eslint-disable-next-line no-await-in-loop
    const pages = await doc.copyPages(src, group);
    pages.forEach((page) => doc.addPage(page));
    const suffix = group.length === 1
      ? `p${String(group[0] + 1).padStart(pad, "0")}`
      : `p${formatPageList(group).replace(/,/g, "_")}`;
    // eslint-disable-next-line no-await-in-loop
    const target = await uniquePath(outputDir, `${stem}-${suffix}.pdf`);
    // eslint-disable-next-line no-await-in-loop
    await fs.writeFile(target, await doc.save());
    files.push({ name: path.basename(target), path: target, kind: "pdf" });
    progress?.(0.1 + 0.85 * ((i + 1) / groups.length), `生成 ${path.basename(target)}`);
  }
  return { files };
}

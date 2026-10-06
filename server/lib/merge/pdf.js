/** PDF 合并（pdf-lib 逐页拷贝）。 */
import fs from "node:fs/promises";
import path from "node:path";
import { PDFDocument } from "pdf-lib";
import { AppError } from "../errors.js";
import { baseName, uniquePath } from "../paths.js";

export async function mergePdf({ inputPaths, outputDir, options = {}, progress }) {
  if (inputPaths.length < 2) throw new AppError("至少需要两个 PDF 文件。", { status: 400, code: "bad-request" });
  const out = await PDFDocument.create();
  out.setProducer("Office PDF Tools");
  out.setCreator("Office PDF Tools");
  const total = inputPaths.length;
  for (let i = 0; i < total; i++) {
    const inputPath = inputPaths[i];
    // eslint-disable-next-line no-await-in-loop
    const bytes = await fs.readFile(inputPath);
    // eslint-disable-next-line no-await-in-loop
    const src = await PDFDocument.load(bytes, { ignoreEncryption: true }).catch((err) => {
      throw new AppError(`${path.basename(inputPath)} 无法读取（可能已加密）。`, {
        code: "invalid-document",
        status: 400,
        cause: err,
      });
    });
    // eslint-disable-next-line no-await-in-loop
    const pages = await out.copyPages(src, src.getPageIndices());
    pages.forEach((page) => out.addPage(page));
    progress?.(0.1 + 0.8 * ((i + 1) / total), `合并 ${path.basename(inputPath)}`);
  }
  const name = options.outputName ? baseName(options.outputName) : `${baseName(inputPaths[0])}-merged`;
  const target = await uniquePath(outputDir, `${name}.pdf`);
  await fs.writeFile(target, await out.save());
  return { files: [{ name: path.basename(target), path: target, kind: "pdf" }] };
}

/** LibreOffice headless 转换引擎。 */
import path from "node:path";
import fs from "node:fs/promises";
import { pathToFileURL } from "node:url";
import { config } from "../../config.js";
import { runProcess } from "../proc.js";
import { AppError, engineMissing, timeoutError } from "../errors.js";
import { detectLibreOffice } from "./detect.js";
import { ensureDir, exists, extName } from "../paths.js";
import { log } from "../logger.js";

/** LibreOffice 用户配置目录不能被并发进程共用，这里串行化所有转换。 */
let queue = Promise.resolve();
function withLock(fn) {
  const run = queue.then(fn, fn);
  queue = run.then(() => undefined, () => undefined);
  return run;
}

const WRITER = new Set(["doc", "docx", "docm", "dot", "dotx", "dotm", "rtf", "odt", "ott", "txt", "html", "htm", "xml", "fodt"]);
const CALC = new Set(["xls", "xlsx", "xlsm", "xlsb", "ods", "ots", "csv", "tsv", "fods"]);
const IMPRESS = new Set(["ppt", "pptx", "pptm", "pot", "potx", "odp", "otp", "fodp"]);

export function libreOfficeFilter(inputPath) {
  const ext = extName(inputPath);
  if (WRITER.has(ext)) return "pdf:writer_pdf_Export";
  if (CALC.has(ext)) return "pdf:calc_pdf_Export";
  if (IMPRESS.has(ext)) return "pdf:impress_pdf_Export";
  return "pdf";
}

async function profileDir() {
  const dir = path.join(config.tmpDir, "lo-profile");
  await ensureDir(dir);
  return dir;
}

/**
 * 用 LibreOffice 把文件转换为 PDF。
 * @param {{inputPath:string, outputDir:string, filter?:string, timeoutMs?:number, signal?:AbortSignal}} opts
 * @returns {Promise<{outputPath:string, engine:string, stdout:string}>}
 */
export async function convertWithLibreOffice(opts) {
  const { inputPath, outputDir, filter = libreOfficeFilter(inputPath), signal } = opts;
  const timeoutMs = opts.timeoutMs || config.convertTimeoutMs;
  const engine = await detectLibreOffice();
  if (!engine.available) {
    throw engineMissing("未找到 LibreOffice（soffice），无法使用该引擎转换。", { hint: engine.detail });
  }
  await ensureDir(outputDir);
  const profile = await profileDir();

  const args = [
    "--headless",
    "--invisible",
    "--nologo",
    "--nodefault",
    "--nolockcheck",
    "--norestore",
    "--nofirststartwizard",
    `-env:UserInstallation=${pathToFileURL(profile).href}`,
    "--convert-to",
    filter,
    "--outdir",
    outputDir,
    inputPath,
  ];

  return withLock(async () => {
    const started = Date.now();
    const { code, stdout, stderr, timedOut, aborted } = await runProcess(engine.path, args, {
      timeoutMs,
      signal,
      cwd: outputDir,
    });
    if (timedOut) throw timeoutError(`LibreOffice 转换超时（>${Math.round(timeoutMs / 1000)}s）。`);
    if (aborted) throw new AppError("任务已取消。", { code: "canceled", status: 499 });
    if (code !== 0) {
      throw new AppError("LibreOffice 转换失败。", {
        code: "conversion-failed",
        status: 500,
        hint: (stderr || stdout).split(/\r?\n/).filter(Boolean).slice(-3).join(" | "),
      });
    }

    // LibreOffice 按输入文件的主名生成输出：<dir>/<base>.pdf
    const base = path.basename(inputPath).replace(/\.[^.]+$/, "");
    const expected = path.join(outputDir, `${base}.pdf`);
    if (await exists(expected)) {
      log.debug(`libreoffice ok in ${Date.now() - started}ms -> ${expected}`);
      return { outputPath: expected, engine: "libreoffice", stdout };
    }
    // 兜底：找目录里最新生成的 pdf
    const entries = await fs.readdir(outputDir, { withFileTypes: true }).catch(() => []);
    const pdfs = entries.filter((e) => e.isFile() && e.name.toLowerCase().endsWith(".pdf"));
    if (pdfs.length === 1) {
      return { outputPath: path.join(outputDir, pdfs[0].name), engine: "libreoffice", stdout };
    }
    throw new AppError("LibreOffice 未生成预期的 PDF 文件。", {
      code: "conversion-failed",
      hint: (stderr || stdout).slice(-400),
    });
  });
}

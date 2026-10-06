/** Office 文档 → PDF（自动选择引擎）。 */
import path from "node:path";
import { config } from "../../config.js";
import { detectEngines, pickOfficeEngine } from "../engines/detect.js";
import { convertWithLibreOffice } from "../engines/libreoffice.js";
import { convertWithMsOffice, comSupports } from "../engines/msoffice.js";
import { AppError, engineMissing } from "../errors.js";
import { baseName, extName, mimeFor } from "../paths.js";

const MS_EXT = new Set(["doc", "docx", "docm", "dot", "dotx", "rtf", "txt", "odt", "xls", "xlsx", "xlsm", "xlsb", "ods", "csv", "ppt", "pptx", "pptm", "odp"]);
const LO_ONLY = new Set(["fodt", "fods", "fodp", "ott", "ots", "otp", "tsv", "xml", "html", "htm", "wps", "wpd", "svg", "bmp", "png", "jpg", "jpeg", "tif", "tiff", "gif", "webp"]);

/**
 * @param {{inputPath:string, outputDir:string, options?:{engine?:string, sheet?:string, recalculate?:boolean},
 *          progress?:(ratio:number, step?:string)=>void, signal?:AbortSignal, warnings?:string[]}} cfg
 */
export async function officeToPdf({ inputPath, outputDir, options = {}, progress, signal, warnings = [] }) {
  const engines = await detectEngines();
  const preferred = options.engine || "auto";
  const ext = extName(inputPath);

  let engineName = pickOfficeEngine(preferred, engines);
  if (engineName === "msoffice" && !comSupports(inputPath)) {
    if (engines.libreoffice?.available) {
      warnings.push(`.${ext} 不在 Microsoft Office 的处理范围内，已改用 LibreOffice。`);
      engineName = "libreoffice";
    } else {
      engineName = null;
    }
  }
  if (!engineName) {
    const hint = [
      engines.libreoffice?.available ? null : `LibreOffice: ${engines.libreoffice?.detail || "未安装"}`,
      engines.msoffice?.available ? null : `Microsoft Office: ${engines.msoffice?.detail || "未安装"}`,
    ].filter(Boolean).join(" / ");
    throw engineMissing("没有可用的 Office → PDF 转换引擎，请安装 Microsoft Office 或 LibreOffice。", { hint });
  }

  progress?.(0.15, `使用 ${engineName === "msoffice" ? "Microsoft Office" : "LibreOffice"} 转换…`);
  const target = path.join(outputDir, `${baseName(inputPath)}.pdf`);

  if (engineName === "msoffice") {
    const result = await convertWithMsOffice({
      inputPath,
      outputPath: target,
      sheet: options.sheet,
      recalculate: options.recalculate !== false,
      signal,
    });
    progress?.(0.9, "转换完成");
    return { outputPath: result.outputPath, engine: result.engine };
  }

  if (options.sheet) {
    warnings.push("LibreOffice 引擎不支持只导出指定工作表，已导出全部工作表。");
  }
  const result = await convertWithLibreOffice({ inputPath, outputDir, signal });
  progress?.(0.9, "转换完成");
  return { outputPath: result.outputPath, engine: result.engine };
}

export function officeToPdfMeta(inputPath) {
  return { mime: mimeFor(".pdf"), kind: "pdf" };
}

export { MS_EXT, LO_ONLY };

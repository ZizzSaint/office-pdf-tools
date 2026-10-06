/** Microsoft Office COM 转换引擎（仅 Windows，需要本机安装 Office）。 */
import path from "node:path";
import fs from "node:fs/promises";
import { config } from "../../config.js";
import { runProcess } from "../proc.js";
import { AppError, engineMissing, timeoutError } from "../errors.js";
import { detectMsOffice } from "./detect.js";
import { ensureDir, exists, extName } from "../paths.js";

const PS = "powershell.exe";

export const COM_SUPPORTED = new Set([
  "doc", "docx", "docm", "dot", "dotx", "rtf", "txt", "odt",
  "xls", "xlsx", "xlsm", "xlsb", "ods", "csv",
  "ppt", "pptx", "pptm", "odp",
]);

export function comAppFor(inputPath) {
  const ext = extName(inputPath);
  if (["doc", "docx", "docm", "dot", "dotx", "rtf", "txt", "odt"].includes(ext)) return "word";
  if (["xls", "xlsx", "xlsm", "xlsb", "ods", "csv"].includes(ext)) return "excel";
  if (["ppt", "pptx", "pptm", "odp"].includes(ext)) return "powerpoint";
  return "auto";
}

export function comSupports(inputPath) {
  return COM_SUPPORTED.has(extName(inputPath));
}

/**
 * 用本机 Office（COM 自动化）导出 PDF。
 * @param {{inputPath:string, outputPath:string, app?:string, sheet?:string, recalculate?:boolean,
 *          timeoutMs?:number, signal?:AbortSignal}} opts
 */
export async function convertWithMsOffice(opts) {
  const { inputPath, outputPath, sheet, recalculate, signal } = opts;
  const timeoutMs = opts.timeoutMs || config.convertTimeoutMs;
  const engine = await detectMsOffice();
  if (!engine.available) {
    throw engineMissing("未检测到可用的 Microsoft Office（Word/Excel/PowerPoint），无法使用该引擎。", { hint: engine.detail });
  }
  const app = opts.app && opts.app !== "auto" ? opts.app : comAppFor(inputPath);
  if (app !== "auto" && engine.apps && !engine.apps[app]) {
    throw engineMissing(`本机未安装 ${app}，无法转换该文件。`);
  }
  await ensureDir(path.dirname(outputPath));
  const pidFile = `${outputPath}.pid`;

  const args = [
    "-NoProfile",
    "-NonInteractive",
    "-ExecutionPolicy", "Bypass",
    "-File", config.officeComScript,
    "-InputPath", inputPath,
    "-OutputPath", outputPath,
    "-App", app,
    "-PidFile", pidFile,
  ];
  if (sheet) args.push("-Sheet", sheet);
  if (recalculate) args.push("-Recalculate");

  const { code, stdout, stderr, timedOut, aborted } = await runProcess(PS, args, {
    timeoutMs,
    signal,
    pidFile,
  });
  await fs.rm(pidFile, { force: true }).catch(() => {});

  if (timedOut) {
    throw timeoutError(`Microsoft Office 转换超时（>${Math.round(timeoutMs / 1000)}s），已尝试结束挂起的 Office 进程。`);
  }
  if (aborted) throw new AppError("任务已取消。", { code: "canceled", status: 499 });

  const line = stdout.split(/\r?\n/).map((s) => s.trim()).filter(Boolean).pop() || "";
  let parsed = null;
  try {
    parsed = JSON.parse(line);
  } catch { /* 可能是 Office 弹窗输出，忽略 */ }

  if (code !== 0 || !parsed?.ok || !(await exists(outputPath))) {
    const detail = parsed?.error || stderr.split(/\r?\n/).filter(Boolean).slice(-3).join(" | ") || `exit ${code}`;
    const hint = String(detail).slice(0, 400);
    const stuck = /80080005|CO_E_SERVER_EXEC_FAILURE/i.test(String(detail));
    throw new AppError("Microsoft Office 转换失败。", {
      code: "conversion-failed",
      status: 500,
      hint: stuck
        ? `${hint}｜提示：检测到 Office 进程启动失败，通常是有一个卡住的 Office 实例（任务管理器中的 WINWORD/EXCEL/POWERPNT）阻塞了自动化，结束它后重试即可。`
        : hint,
    });
  }
  return { outputPath, engine: "msoffice", app: parsed.app };
}

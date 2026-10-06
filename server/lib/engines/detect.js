/** 引擎探测：LibreOffice、Microsoft Office (COM)、PDF 渲染器。 */
import fs from "node:fs/promises";
import path from "node:path";
import { config } from "../../config.js";
import { runProcess } from "../proc.js";
import { log } from "../logger.js";

const cache = new Map();
const TTL = 60 * 1000;

function cached(key, fn) {
  const hit = cache.get(key);
  const now = Date.now();
  if (hit && now - hit.at < TTL) return hit.value;
  const value = Promise.resolve()
    .then(fn)
    .catch((err) => ({ available: false, detail: String(err?.message || err) }));
  cache.set(key, { at: now, value });
  return value;
}

export function invalidateEngineCache() {
  cache.clear();
}

const LO_CANDIDATES = {
  win32: [
    "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
    "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    "C:\\Program Files\\LibreOffice 7\\program\\soffice.exe",
    "C:\\Program Files\\LibreOffice 24.2\\program\\soffice.exe",
  ],
  darwin: ["/Applications/LibreOffice.app/Contents/MacOS/soffice"],
  linux: ["/usr/bin/soffice", "/usr/local/bin/soffice", "/usr/bin/libreoffice", "/snap/bin/libreoffice", "/opt/libreoffice/program/soffice"],
};

async function firstExisting(paths) {
  for (const p of paths) {
    try {
      await fs.access(p);
      return p;
    } catch { /* keep looking */ }
  }
  return null;
}

async function findSoffice() {
  if (config.sofficePath) {
    try {
      await fs.access(config.sofficePath);
      return config.sofficePath;
    } catch { /* fall through */ }
  }
  // DSH / 其它集成环境可能自带 LibreOffice kit
  const kitHints = [process.env.DSH_LIBREOFFICE_KIT, process.env.LIBREOFFICE_PATH].filter(Boolean);
  for (const hint of kitHints) {
    const candidates = [
      path.join(hint, "program", "soffice.exe"),
      path.join(hint, "program", "soffice"),
      path.join(hint, "soffice.exe"),
      path.join(hint, "soffice"),
      path.join(hint, "LibreOffice", "program", "soffice.exe"),
    ];
    const found = await firstExisting(candidates);
    if (found) return found;
  }
  const local = await firstExisting(LO_CANDIDATES[process.platform] || LO_CANDIDATES.linux);
  if (local) return local;
  // PATH
  const probe = process.platform === "win32" ? "where" : "which";
  const { code, stdout } = await runProcess(probe, ["soffice"], { timeoutMs: 8000 });
  if (code === 0) {
    const line = stdout.split(/\r?\n/).map((s) => s.trim()).filter(Boolean)[0];
    if (line) return line;
  }
  return null;
}

export function detectLibreOffice() {
  return cached("libreoffice", async () => {
    const soffice = await findSoffice();
    if (!soffice) {
      return {
        available: false,
        detail: "未找到 soffice；可安装 LibreOffice，或设置环境变量 SOFFICE_PATH 指向 soffice 可执行文件。",
      };
    }
    const { code, stdout, stderr } = await runProcess(soffice, ["--version"], { timeoutMs: 20000 });
    const version = (stdout || stderr).split(/\r?\n/)[0] || "";
    return {
      available: code === 0,
      path: soffice,
      version: version.replace(/^LibreOffice\s*/i, "").trim() || "(未知版本)",
      detail: version.trim(),
    };
  });
}

const PS = process.platform === "win32" ? "powershell.exe" : null;

export function detectMsOffice() {
  return cached("msoffice", async () => {
    if (process.platform !== "win32") {
      return { available: false, detail: "Microsoft Office COM 仅在 Windows 上可用。" };
    }
    const script = [
      "$ErrorActionPreference='SilentlyContinue'",
      "$map=[ordered]@{word='Word.Application';excel='Excel.Application';powerpoint='PowerPoint.Application'}",
      "$out=[ordered]@{}",
      "foreach($k in $map.Keys){ $t=[Type]::GetTypeFromProgID($map[$k]); $out[$k]= [bool]$t }",
      "$ver=(Get-ItemProperty 'HKLM:\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration' -Name Version -ErrorAction SilentlyContinue).Version",
      "$out['version']=$ver",
      "$out | ConvertTo-Json -Compress",
    ].join("; ");
    const { code, stdout } = await runProcess(PS, ["-NoProfile", "-NonInteractive", "-Command", script], { timeoutMs: 30000 });
    if (code !== 0 || !stdout.trim()) {
      return { available: false, detail: "无法查询 Microsoft Office COM 组件（未安装或权限不足）。" };
    }
    let parsed = {};
    try {
      parsed = JSON.parse(stdout.trim().split(/\r?\n/).pop());
    } catch {
      return { available: false, detail: "解析 Office 探测结果失败。" };
    }
    const apps = {
      word: !!parsed.word,
      excel: !!parsed.excel,
      powerpoint: !!parsed.powerpoint,
    };
    const available = apps.word || apps.excel || apps.powerpoint;
    const names = Object.entries(apps).filter(([, v]) => v).map(([k]) => k).join(", ");
    return {
      available,
      apps,
      version: parsed.version || "",
      detail: available ? `已启用: ${names}${parsed.version ? " · " + parsed.version : ""}` : "未检测到 Word/Excel/PowerPoint。",
    };
  });
}

export function detectRenderer() {
  return cached("renderer", async () => {
    const canvas = await import("@napi-rs/canvas");
    const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
    return {
      available: typeof canvas.createCanvas === "function" && typeof pdfjs.getDocument === "function",
      detail: `pdf.js ${pdfjs.version || ""} + @napi-rs/canvas`.trim(),
    };
  });
}

export async function detectEngines({ refresh = false } = {}) {
  if (refresh) invalidateEngineCache();
  const [libreoffice, msoffice, renderer] = await Promise.all([
    detectLibreOffice(),
    detectMsOffice(),
    detectRenderer(),
  ]);
  return { libreoffice, msoffice, renderer };
}

/** 按优先级选择 Office → PDF 引擎。 */
export function pickOfficeEngine(preferred = "auto", engines) {
  const list = preferred === "auto"
    ? (process.platform === "win32" ? ["msoffice", "libreoffice"] : ["libreoffice", "msoffice"])
    : [preferred];
  for (const name of list) {
    if (engines?.[name]?.available) return name;
  }
  return null;
}

export function logEngineSummary(engines) {
  log.info("engines:", {
    libreoffice: engines.libreoffice?.available ? engines.libreoffice.path : "no",
    msoffice: engines.msoffice?.available ? engines.msoffice.detail : "no",
    renderer: engines.renderer?.available ? engines.renderer.detail : "no",
  });
}

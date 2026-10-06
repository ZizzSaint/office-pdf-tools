/** 运行配置：命令行参数 + 环境变量 + 默认值。 */
import path from "node:path";
import os from "node:os";
import fs from "node:fs";
import { fileURLToPath } from "node:url";

export const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

/** 版本号统一从 package.json 读取，避免与安装包版本不一致。 */
function readVersion() {
  try {
    const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, "package.json"), "utf8"));
    return pkg.version || "0.0.0";
  } catch {
    return "0.0.0";
  }
}

function argValue(argv, name, short) {
  const i = argv.findIndex((a) => a === name || (short && a === short));
  if (i >= 0) {
    const next = argv[i + 1];
    return next && !next.startsWith("-") ? next : true;
  }
  const inline = argv.find((a) => a.startsWith(name + "="));
  return inline ? inline.slice(name.length + 1) : undefined;
}

const argv = process.argv.slice(2);

export const config = {
  name: "office-pdf-tools",
  version: readVersion(),
  host: String(argValue(argv, "--host") ?? process.env.HOST ?? "127.0.0.1"),
  port: Number(argValue(argv, "--port", "-p") ?? process.env.PORT ?? 3000),
  https: argv.includes("--https") || process.env.HTTPS === "1",
  dev: argv.includes("--dev"),
  open: argv.includes("--open"),

  root: ROOT,
  addinDir: path.join(ROOT, "addin"),
  outputDir: path.resolve(String(argValue(argv, "--output", "-o") ?? process.env.OUTPUT_DIR ?? path.join(ROOT, "output"))),
  tmpDir: path.resolve(String(process.env.TMP_DIR ?? path.join(ROOT, "tmp"))),
  certDir: path.resolve(String(process.env.CERT_DIR ?? path.join(ROOT, ".certs"))),

  maxUploadMb: Number(process.env.MAX_UPLOAD_MB ?? 300),
  jobTtlMinutes: Number(process.env.JOB_TTL_MIN ?? 120),
  keepTemp: process.env.KEEP_TEMP === "1",
  convertTimeoutMs: Number(process.env.CONVERT_TIMEOUT_MS ?? 10 * 60 * 1000),
  sofficePath: process.env.SOFFICE_PATH || "",
  // 安装版（桌面软件）用固定位置的证书；开发时留空则自动使用 office-addin-dev-certs
  tlsKey: String(argValue(argv, "--tls-key") ?? process.env.TLS_KEY ?? ""),
  tlsCert: String(argValue(argv, "--tls-cert") ?? process.env.TLS_CERT ?? ""),
  tlsPfx: String(argValue(argv, "--tls-pfx") ?? process.env.TLS_PFX ?? ""),
  tlsPassphrase: String(process.env.TLS_PFX_PASSPHRASE ?? ""),
  officeComScript: path.join(ROOT, "scripts", "office-com.ps1"),
  cpus: Math.max(1, Math.min(4, os.cpus()?.length || 2)),

  get maxUploadBytes() {
    return this.maxUploadMb * 1024 * 1024;
  },
};

export function describeConfig() {
  return {
    host: config.host,
    port: config.port,
    https: config.https,
    outputDir: config.outputDir,
    tmpDir: config.tmpDir,
    maxUploadMb: config.maxUploadMb,
    version: config.version,
  };
}

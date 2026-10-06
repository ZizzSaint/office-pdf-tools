/** 运行配置：命令行参数 + 环境变量 + 默认值。 */
import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";

export const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

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
  version: "1.0.0",
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

#!/usr/bin/env node
/** 服务入口。 */
import http from "node:http";
import https from "node:https";
import fs from "node:fs/promises";
import path from "node:path";
import { config } from "./config.js";
import { createApp } from "./app.js";
import { log } from "./lib/logger.js";
import { ensureDir } from "./lib/paths.js";
import { detectEngines, logEngineSummary } from "./lib/engines/detect.js";
import { jobStore } from "./lib/jobs.js";
import { cleanupRegistry } from "./lib/registry.js";

async function httpsOptions() {
  // 1) 安装版：直接读取安装目录里的证书
  if (config.tlsKey && config.tlsCert) {
    try {
      const [key, cert] = await Promise.all([fs.readFile(config.tlsKey), fs.readFile(config.tlsCert)]);
      log.info(`已加载安装目录证书: ${path.basename(config.tlsCert)}`);
      return { key, cert };
    } catch (err) {
      log.warn("读取安装目录证书失败，回退到开发证书:", err?.message || err);
    }
  }
  if (config.tlsPfx) {
    try {
      return { pfx: await fs.readFile(config.tlsPfx), passphrase: config.tlsPassphrase || undefined };
    } catch (err) {
      log.warn("读取 PFX 失败，回退到开发证书:", err?.message || err);
    }
  }
  // 2) 开发模式：office-addin-dev-certs（自动生成并信任 localhost 证书）
  try {
    const { getHttpsServerOptions } = await import("office-addin-dev-certs");
    const options = await getHttpsServerOptions();
    log.info("已加载受信任的开发证书 (office-addin-dev-certs)");
    return options;
  } catch (err) {
    log.error("无法获取 HTTPS 证书:", err?.message || err);
    log.error("请运行: npx office-addin-dev-certs install  （或改用 npm start 以 http 启动，仅限 API 调试）");
    throw err;
  }
}

/** pdfjs-dist v6 要求 Node >= 22.13，提前给出明确提示而不是让 PDF 解析莫名失败。 */
function assertNodeVersion() {
  const [major, minor] = process.versions.node.split(".").map(Number);
  if (major > 22 || (major === 22 && minor >= 13)) return;
  log.error(`当前 Node.js 版本为 ${process.versions.node}，本项目需要 Node >= 22.13（pdf.js 6 的要求）。`);
  log.error("请升级 Node.js：https://nodejs.org/  或用 nvm/fnm 安装 LTS 版本后重试。");
  process.exit(1);
}

async function main() {
  assertNodeVersion();
  await ensureDir(config.outputDir);
  await ensureDir(config.tmpDir);

  const app = createApp();
  const server = config.https
    ? https.createServer(await httpsOptions(), app)
    : http.createServer(app);

  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(config.port, config.host, resolve);
  });

  const scheme = config.https ? "https" : "http";
  const base = `${scheme}://localhost:${config.port}`;
  const shown = `${scheme}://${config.host === "0.0.0.0" ? "localhost" : config.host}:${config.port}`;
  log.info(`${config.name} v${config.version} 已启动`);
  log.info(`  任务窗格: ${base}/taskpane.html`);
  log.info(`  API:      ${base}/api/health`);
  log.info(`  输出目录: ${config.outputDir}`);
  if (!config.https) {
    log.warn("当前为 HTTP 模式；Office 加载项要求 HTTPS，请在 Office 中使用 npm run https。");
  }

  const engines = await detectEngines();
  logEngineSummary(engines);
  if (!engines.libreoffice?.available && !engines.msoffice?.available) {
    log.warn("未检测到 Office → PDF 引擎；加载项内的“当前文档导出 PDF”仍可用（由 Office 本体完成）。");
  }

  const timer = setInterval(() => {
    jobStore.cleanupExpired().catch(() => {});
    cleanupRegistry().catch(() => {});
  }, 5 * 60 * 1000);
  timer.unref?.();

  const shutdown = () => {
    log.info("正在关闭服务…");
    server.close(() => process.exit(0));
    setTimeout(() => process.exit(0), 3000).unref();
  };
  process.on("SIGINT", shutdown);
  process.on("SIGTERM", shutdown);

  if (config.open) {
    const { runProcess } = await import("./lib/proc.js");
    runProcess(process.platform === "win32" ? "cmd.exe" : process.platform === "darwin" ? "open" : "xdg-open",
      process.platform === "win32" ? ["/c", "start", "", base] : [base], { timeoutMs: 5000 }).catch(() => {});
  }

  return server;
}

main().catch((err) => {
  log.error("启动失败:", err?.message || err);
  process.exit(1);
});

#!/usr/bin/env node
/** 服务入口。 */
import http from "node:http";
import https from "node:https";
import { config } from "./config.js";
import { createApp } from "./app.js";
import { log } from "./lib/logger.js";
import { ensureDir } from "./lib/paths.js";
import { detectEngines, logEngineSummary } from "./lib/engines/detect.js";
import { jobStore } from "./lib/jobs.js";
import { cleanupRegistry } from "./lib/registry.js";

async function httpsOptions() {
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

async function main() {
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

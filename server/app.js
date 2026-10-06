/** Express 应用装配：安全头、静态资源、API、错误处理。 */
import path from "node:path";
import fs from "node:fs";
import express from "express";
import { config } from "./config.js";
import { createApiRouter } from "./routes/api.js";
import { toAppError } from "./lib/errors.js";
import { log } from "./lib/logger.js";

const LOCAL_ORIGIN = /^https?:\/\/(localhost|127\.0\.0\.1|\[::1\])(:\d+)?$/i;

function securityMiddleware() {
  return (req, res, next) => {
    const origin = req.headers.origin;
    if (origin && LOCAL_ORIGIN.test(origin)) {
      res.setHeader("Access-Control-Allow-Origin", origin);
      res.setHeader("Vary", "Origin");
      res.setHeader("Access-Control-Allow-Methods", "GET,POST,DELETE,OPTIONS");
      res.setHeader("Access-Control-Allow-Headers", "Content-Type, X-Office-Pdf-Tools");
      res.setHeader("Access-Control-Max-Age", "600");
    }
    if (req.method === "OPTIONS") {
      res.status(origin && LOCAL_ORIGIN.test(origin) ? 204 : 403).end();
      return;
    }

    // 写操作防 CSRF：要求自定义头（跨站无法在预检通过前发送），CLI 客户端无 Origin 时豁免。
    if (req.method !== "GET" && req.path.startsWith("/api")) {
      const hasHeader = req.headers["x-office-pdf-tools"] === "1";
      const fromBrowser = !!origin || /mozilla|chrome|safari|edg\//i.test(req.headers["user-agent"] || "");
      if (!hasHeader && fromBrowser) {
        res.status(403).json({ error: { code: "forbidden", message: "缺少 X-Office-Pdf-Tools 请求头（防止跨站请求伪造）。" } });
        return;
      }
    }

    res.setHeader("X-Content-Type-Options", "nosniff");
    res.setHeader("Referrer-Policy", "no-referrer");
    if (req.path.startsWith("/api")) res.setHeader("Cache-Control", "no-store");
    next();
  };
}

export function createApp() {
  const app = express();
  app.disable("x-powered-by");
  app.use(express.json({ limit: "2mb" }));
  app.use(express.urlencoded({ extended: false, limit: "2mb" }));
  app.use(securityMiddleware());

  app.use((req, res, next) => {
    const started = Date.now();
    res.on("finish", () => {
      if (req.path.startsWith("/api")) {
        log.debug(`${req.method} ${req.originalUrl} → ${res.statusCode} (${Date.now() - started}ms)`);
      }
    });
    next();
  });

  app.use("/api", createApiRouter());

  const staticOptions = {
    etag: !config.dev,
    lastModified: !config.dev,
    setHeaders(res, filePath) {
      if (config.dev && /\.(js|css|html)$/i.test(filePath)) res.setHeader("Cache-Control", "no-store");
    },
  };
  app.use(express.static(config.addinDir, staticOptions));
  app.get("/", (req, res) => res.sendFile(path.join(config.addinDir, "taskpane.html"), staticOptions));

  app.get("/manifest.xml", (req, res) => {
    const manifest = path.join(config.root, "manifest", "manifest.xml");
    if (!fs.existsSync(manifest)) {
      res.status(404).type("text/plain").send("manifest/manifest.xml 不存在");
      return;
    }
    res.type("application/xml").send(fs.readFileSync(manifest, "utf8"));
  });

  app.use((req, res) => {
    res.status(404).json({ error: { code: "not-found", message: `未找到 ${req.method} ${req.path}` } });
  });

  // eslint-disable-next-line no-unused-vars
  app.use((err, req, res, next) => {
    const appErr = toAppError(err);
    if (appErr.status >= 500) log.error(`${req.method} ${req.path}:`, appErr.message, appErr.cause?.stack || "");
    else log.debug(`${req.method} ${req.path} → ${appErr.status}: ${appErr.message}`);
    if (res.headersSent) {
      res.end();
      return;
    }
    res.status(appErr.status).json({ error: appErr.toJSON() });
  });

  return app;
}

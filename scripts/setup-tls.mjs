#!/usr/bin/env node
/**
 * 生成并信任 localhost 开发证书，然后把 PEM 复制到安装目录（安装版运行时使用固定路径）。
 * 用法: node scripts/setup-tls.mjs "<安装目录>\tls"
 */
import fs from "node:fs/promises";
import path from "node:path";

const targetDir = process.argv[2];
if (!targetDir) {
  console.error("用法: node scripts/setup-tls.mjs <目标目录>");
  process.exit(1);
}

try {
  const { getHttpsServerOptions } = await import("office-addin-dev-certs");
  const options = await getHttpsServerOptions();
  await fs.mkdir(targetDir, { recursive: true });
  const keyPath = path.join(targetDir, "localhost.key");
  const certPath = path.join(targetDir, "localhost.crt");
  await fs.writeFile(keyPath, options.key);
  await fs.writeFile(certPath, options.cert);
  if (options.ca) await fs.writeFile(path.join(targetDir, "ca.crt"), options.ca);
  console.log(JSON.stringify({ ok: true, key: keyPath, cert: certPath }));
} catch (err) {
  console.log(JSON.stringify({ ok: false, error: String(err?.message || err) }));
  process.exit(1);
}

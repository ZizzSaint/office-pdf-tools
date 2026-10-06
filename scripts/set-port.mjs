#!/usr/bin/env node
/**
 * 批量改写加载项地址：端口 或 完整域名。
 *
 *   npm run set:port -- 3001                       本地换端口
 *   npm run set:port -- https://pdf.example.com    改成公网地址（集中部署用）
 *   npm run set:port -- https://pdf.example.com:8443
 *
 * 会同步更新清单、文档、脚本中的 https://localhost:<旧端口> 引用。
 */
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const arg = process.argv[2];

if (!arg) {
  console.error("用法: npm run set:port -- <端口|https://域名[:端口]>");
  process.exit(1);
}

const FILES = [
  "manifest/manifest.xml",
  "README.md",
  "docs/INSTALL.md",
  "docs/DESKTOP-APP.md",
  "docs/API.md",
  "docs/TROUBLESHOOTING.md",
  "addin/js/commands.js",
  "scripts/sideload.ps1",
  "scripts/install.ps1",
];

const isPort = /^\d{2,5}$/.test(arg);
let next = "";
if (isPort) {
  next = arg;
} else if (/^https?:\/\/[^\s/]+$/i.test(arg)) {
  next = arg.replace(/\/+$/, "");
} else {
  console.error("参数必须是端口号（如 3001）或形如 https://pdf.example.com 的地址。");
  process.exit(1);
}

const basePattern = /https:\/\/localhost:(\d+)/g;
let changed = 0;
const touched = [];

for (const rel of FILES) {
  const file = path.join(ROOT, rel);
  let text;
  try {
    text = await fs.readFile(file, "utf8");
  } catch {
    continue;
  }
  const before = text;
  const updated = text.replace(basePattern, (full, port) => (isPort ? `https://localhost:${next}` : next));
  if (updated !== before) {
    await fs.writeFile(file, updated);
    changed++;
    touched.push(rel);
  }
}

if (!changed) {
  console.log("没有需要更新的文件（可能已经是目标地址）。");
  process.exit(0);
}

if (isPort) {
  console.log(`端口已更新为 ${next}：`);
  console.log("  " + touched.join("\n  "));
  console.log("\n记得：npm run https -- --port " + next + " 然后重新侧载（npm run sideload）或让 Office 重启。");
} else {
  console.log(`加载项地址已更新为 ${next}：`);
  console.log("  " + touched.join("\n"));
  console.log("\n提示：集中部署时清单需托管在公网 HTTPS，且 AppDomains 已包含该域名。");
}

/** 修改加载项服务端口：npm run set:port -- 3001 */
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const target = process.argv[2];
if (!target || !/^\d{2,5}$/.test(target)) {
  console.error("用法: npm run set:port -- <端口，例如 3001>");
  process.exit(1);
}
const current = process.env.PORT || "3000";
const files = ["manifest/manifest.xml", "README.md", "docs/INSTALL.md", "addin/js/commands.js", "scripts/sideload.ps1"];
let changed = 0;
for (const rel of files) {
  const file = path.join(ROOT, rel);
  let text;
  try {
    text = await fs.readFile(file, "utf8");
  } catch {
    continue;
  }
  const next = text.replace(new RegExp(`localhost:${current}`, "g"), `localhost:${target}`);
  if (next !== text) {
    await fs.writeFile(file, next);
    changed++;
    console.log(`updated ${rel}`);
  }
}
console.log(changed ? `端口已更新为 ${target}（共 ${changed} 个文件）` : "没有需要更新的文件");

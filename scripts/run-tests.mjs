#!/usr/bin/env node
/**
 * 跨版本测试入口：自己发现 tests/*.test.js 再交给 node --test。
 * 这样在 Node 18/20（不支持 --test 通配符）与 Windows 上都能直接 npm test。
 */
import { readdirSync, statSync } from "node:fs";
import { spawnSync } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const testsDir = path.join(ROOT, "tests");

function collect(dir, base = "") {
  const out = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const rel = path.posix.join(base, entry.name);
    if (entry.isDirectory()) out.push(...collect(path.join(dir, entry.name), rel));
    else if (entry.name.endsWith(".test.js") || entry.name.endsWith(".test.mjs")) out.push(`tests/${rel}`);
  }
  return out;
}

const files = collect(testsDir).sort();
if (!files.length) {
  console.error("没有找到测试文件（tests/*.test.js）");
  process.exit(1);
}

const [major, minor] = process.versions.node.split(".").map(Number);
const supportsConcurrency = major > 21 || (major === 21) || (major === 20 && minor >= 10);

const args = ["--test", ...(supportsConcurrency ? ["--test-concurrency=1"] : []), ...files];
console.log(`node ${process.versions.node} · ${files.length} 个测试文件`);
const result = spawnSync(process.execPath, args, { stdio: "inherit", cwd: ROOT });
process.exit(result.status ?? 1);

void statSync;

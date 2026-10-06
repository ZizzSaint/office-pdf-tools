import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { i18nKeys, hasI18nKey } from "../addin/js/i18n.js";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const read = (rel) => fs.readFile(path.join(ROOT, rel), "utf8");

/** 去掉注释，避免注释里的示例代码影响静态检查。 */
function stripComments(code) {
  return code.replace(/\/\*[\s\S]*?\*\//g, "").replace(/(^|[^:])\/\/[^\n]*/g, "$1");
}

async function jsFiles() {
  const dir = path.join(ROOT, "addin/js");
  const out = [];
  async function walk(current) {
    for (const entry of await fs.readdir(current, { withFileTypes: true })) {
      const full = path.join(current, entry.name);
      if (entry.isDirectory()) await walk(full);
      else if (entry.name.endsWith(".js")) out.push(full);
    }
  }
  await walk(dir);
  return out;
}

test("加载项 HTML/JS：引用的资源与 DOM id 都存在", async () => {
  const html = await read("addin/taskpane.html");
  const ids = new Set([...html.matchAll(/\sid="([^"]+)"/g)].map((m) => m[1]));
  assert.ok(ids.size > 20, "taskpane.html 应包含完整界面");

  // 静态资源存在
  for (const rel of ["addin/css/style.css", "addin/js/app.js", "addin/commands.html", "addin/js/commands.js"]) {
    await fs.access(path.join(ROOT, rel));
  }
  for (const match of html.matchAll(/(?:src|href)="((?!https?:)[^"]+)"/g)) {
    const rel = path.join("addin", match[1]);
    // eslint-disable-next-line no-await-in-loop
    await fs.access(path.join(ROOT, rel));
  }

  // JS 里选择的 id 必须存在于 HTML
  const missing = [];
  for (const file of await jsFiles()) {
    const code = await fs.readFile(file, "utf8");
    for (const match of code.matchAll(/\$\("#([A-Za-z0-9_-]+)"\)/g)) {
      if (!ids.has(match[1])) missing.push(`${path.relative(ROOT, file)} → #${match[1]}`);
    }
  }
  assert.deepEqual(missing, [], "存在引用了不存在的元素 id：\n" + missing.join("\n"));
});

test("加载项 i18n：中英词典键一致且被引用", async () => {
  const zh = new Set(i18nKeys("zh-CN"));
  const en = new Set(i18nKeys("en"));
  assert.ok(zh.size > 80, "词典规模");
  const onlyZh = [...zh].filter((k) => !en.has(k));
  const onlyEn = [...en].filter((k) => !zh.has(k));
  assert.deepEqual(onlyZh, [], "英文词典缺少");
  assert.deepEqual(onlyEn, [], "中文词典缺少");

  const html = await read("addin/taskpane.html");
  const used = new Set([...html.matchAll(/data-i18n(?:-placeholder|-title)?="([^"]+)"/g)].map((m) => m[1]));
  for (const file of await jsFiles()) {
    const code = stripComments(await fs.readFile(file, "utf8"));
    for (const match of code.matchAll(/\bt\("([a-z0-9_.]+)"\)/gi)) used.add(match[1]);
  }
  const undefinedKeys = [...used].filter((k) => !zh.has(k));
  assert.deepEqual(undefinedKeys, [], "以下 key 未定义：\n" + undefinedKeys.join("\n"));

  // 动态拼接的 key（拆分策略 / 提示）
  const kinds = ["pdf", "docx", "xlsx", "pptx"];
  const labels = ["everyN", "ranges", "eachPage", "heading", "section", "sheets", "everyNRows"];
  for (const kind of kinds) {
    for (const label of labels) {
      for (const prefix of ["split.strategy", "split.hint"]) {
        const key = `${prefix}.${kind}.${label}`;
        const isRelevant = prefix === "split.hint" ? true : true;
        if (!isRelevant) continue;
        if (hasI18nKey(key, "zh-CN")) assert.ok(hasI18nKey(key, "en"), `英文缺少 ${key}`);
      }
    }
  }
});

test("清单：三端宿主、图标资源、CustomTab 结构与资源引用", async () => {
  const xml = await read("manifest/manifest.xml");
  for (const host of ['Name="Document"', 'Name="Workbook"', 'Name="Presentation"']) {
    assert.ok(xml.includes(host), `清单缺少宿主 ${host}`);
  }
  assert.ok(xml.includes("VersionOverrides"), "清单缺少 VersionOverrides（ribbon 按钮）");

  // 图标文件真实存在
  for (const size of [16, 32, 80]) {
    await fs.access(path.join(ROOT, `manifest/assets/icon-${size}.png`));
  }

  // 每个 CustomTab 结尾必须有 Label（Office 校验网关的硬性要求）
  const tabs = [...xml.matchAll(/<CustomTab\b[\s\S]*?<\/CustomTab>/g)].map((m) => m[0]);
  assert.equal(tabs.length, 3, "三个宿主各有一个自定义选项卡");
  for (const tab of tabs) {
    const tail = tab.slice(-200);
    assert.match(tail, /<Label resid="Tab\.OfficePdfTools\.Label"\/>/, "CustomTab 必须以 Label 结尾");
    assert.ok((tab.match(/<Group\b/g) || []).length >= 1, "CustomTab 至少包含一个 Group");
    assert.ok((tab.match(/<Control\b/g) || []).length >= 1, "Group 至少包含一个按钮");
  }

  // 所有 resid 都要在 Resources 中定义
  const resourceIds = new Set([
    ...[...xml.matchAll(/<bt:(?:Image|Url|String)\s+id="([^"]+)"/g)].map((m) => m[1]),
  ]);
  const usedResIds = new Set([
    ...[...xml.matchAll(/resid="([^"]+)"/g)].map((m) => m[1]),
  ]);
  const dangling = [...usedResIds].filter((id) => !resourceIds.has(id));
  assert.deepEqual(dangling, [], "以下 resid 未在 Resources 中定义：\n" + dangling.join("\n"));

  // 按钮动作只使用 ShowTaskpane（保证三个宿主行为一致）
  const actions = [...xml.matchAll(/<Action\s+xsi:type="([^"]+)"/g)].map((m) => m[1]);
  assert.ok(actions.length >= 9, "至少 9 个 ribbon 按钮");
  assert.ok(actions.every((a) => a === "ShowTaskpane"), "全部使用 ShowTaskpane 动作");
});

test("服务端与加载项的接口约定一致", async () => {
  const runner = await read("server/lib/runner.js");
  const ops = /export const OPS = \[([^\]]+)\]/.exec(runner.replace(/\s+/g, " "));
  assert.ok(ops, "OPS 定义可解析");
  const serverOps = ops[1].split(",").map((s) => s.trim().replace(/"/g, "")).filter(Boolean);
  const clientOps = new Set();
  for (const file of await jsFiles()) {
    const code = await fs.readFile(file, "utf8");
    for (const match of code.matchAll(/op:\s*"([a-z-]+)"/g)) clientOps.add(match[1]);
    for (const match of code.matchAll(/\?\s*"([a-z-]+)"\s*:\s*"pdf-to-images"/g)) clientOps.add(match[1]);
  }
  for (const op of clientOps) {
    assert.ok(serverOps.includes(op), `前端使用了服务端不支持的操作: ${op}（支持：${serverOps.join(", ")}）`);
  }
  assert.ok(serverOps.includes("office-to-pdf") && serverOps.includes("images-to-pdf"));
});

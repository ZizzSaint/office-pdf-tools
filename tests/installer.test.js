import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const isWindows = process.platform === "win32";
const PS = "powershell.exe";

function ps(args, opts = {}) {
  return spawnSync(PS, ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", ...args], {
    encoding: "utf8",
    cwd: ROOT,
    timeout: opts.timeout || 120000,
  });
}

/** 最小化安装包夹具：只保留安装脚本真正需要触碰的文件。 */
async function makeFixture(dir) {
  const payload = path.join(dir, "payload");
  await fs.mkdir(path.join(payload, "server"), { recursive: true });
  await fs.mkdir(path.join(payload, "manifest"), { recursive: true });
  await fs.mkdir(path.join(payload, "scripts"), { recursive: true });
  await fs.writeFile(path.join(payload, "server", "index.js"), "// fixture\n");
  await fs.writeFile(path.join(payload, "package.json"), JSON.stringify({ name: "office-pdf-tools", version: "1.1.0" }));
  await fs.copyFile(path.join(ROOT, "manifest", "manifest.xml"), path.join(payload, "manifest", "manifest.xml"));
  for (const file of ["tray.ps1", "install.ps1", "uninstall.ps1", "start-hidden.vbs"]) {
    await fs.copyFile(path.join(ROOT, "scripts", file), path.join(payload, "scripts", file));
  }
  await fs.copyFile(path.join(ROOT, "scripts", "install.ps1"), path.join(dir, "install.ps1"));
  await fs.copyFile(path.join(ROOT, "scripts", "uninstall.ps1"), path.join(dir, "uninstall.ps1"));
  return payload;
}

test("安装脚本：PowerShell 语法检查", { skip: !isWindows }, () => {
  const files = ["install.ps1", "uninstall.ps1", "tray.ps1"];
  for (const file of files) {
    const checker = [
      "$errors=$null; $tokens=$null",
      `[System.Management.Automation.Language.Parser]::ParseFile('${path.join(ROOT, "scripts", file)}', [ref]$tokens, [ref]$errors) | Out-Null`,
      "if ($errors) { $errors | ForEach-Object { Write-Output (\"line \" + $_.Extent.StartLineNumber + \": \" + $_.Message) }; exit 1 }",
      "Write-Output 'OK'",
    ].join("; ");
    const result = ps(["-Command", checker]);
    assert.equal(result.status, 0, `${file} 解析失败:\n${result.stdout}${result.stderr}`);
  }
});

test("安装脚本：DryRun 不产生副作用", { skip: !isWindows }, async () => {
  const dir = await fs.mkdtemp(path.join(process.env.TEMP || "/tmp", "opt-install-dry-"));
  await makeFixture(dir);
  const installDir = path.join(dir, "not-created");
  const result = ps(["-File", path.join(dir, "install.ps1"), "-SourceDir", dir, "-InstallDir", installDir, "-Port", "3456", "-DryRun", "-Silent"]);
  assert.equal(result.status, 0, "DryRun 应正常退出");
  const output = result.stdout + result.stderr;
  assert.match(output, /DryRun/);
  assert.match(output, /3456/);
  const exists = await fs.stat(installDir).then(() => true).catch(() => false);
  assert.equal(exists, false, "DryRun 不应创建安装目录");
});

test("安装脚本：真实复制文件并改写清单端口", { skip: !isWindows }, async () => {
  const dir = await fs.mkdtemp(path.join(process.env.TEMP || "/tmp", "opt-install-real-"));
  await makeFixture(dir);
  const installDir = path.join(dir, "OfficePdfTools");
  const result = ps([
    "-File", path.join(dir, "install.ps1"),
    "-SourceDir", dir,
    "-InstallDir", installDir,
    "-Port", "3456",
    "-SkipTls", "-SkipRegistry", "-SkipShortcuts", "-NoStart", "-Silent",
  ]);
  assert.equal(result.status, 0, `安装失败:\n${result.stdout}\n${result.stderr}`);

  assert.equal((await fs.readFile(path.join(installDir, "port.txt"), "utf8")).trim(), "3456");
  const manifest = await fs.readFile(path.join(installDir, "manifest.xml"), "utf8");
  assert.match(manifest, /https:\/\/localhost:3456\/taskpane\.html/, "清单中的端口已改写");
  assert.doesNotMatch(manifest, /localhost:3000/, "旧端口不应残留");
  // 结构与工具脚本
  for (const rel of ["app/server/index.js", "tools/tray.ps1", "tools/uninstall.ps1", "start-hidden.vbs", "output"]) {
    await fs.access(path.join(installDir, rel));
  }
  const vbs = await fs.readFile(path.join(installDir, "start-hidden.vbs"), "utf8");
  assert.match(vbs, /tray\.ps1/);
  assert.match(vbs, /powershell\.exe/);
});

test("卸载脚本：DryRun 与静态约定", { skip: !isWindows }, async () => {
  const dir = await fs.mkdtemp(path.join(process.env.TEMP || "/tmp", "opt-uninstall-"));
  const installDir = path.join(dir, "OfficePdfTools");
  await fs.mkdir(path.join(installDir, "tools"), { recursive: true });
  await fs.copyFile(path.join(ROOT, "scripts", "uninstall.ps1"), path.join(installDir, "tools", "uninstall.ps1"));
  const result = ps(["-File", path.join(installDir, "tools", "uninstall.ps1"), "-InstallDir", installDir, "-DryRun", "-Silent"]);
  assert.equal(result.status, 0, result.stdout + result.stderr);
  const output = result.stdout + result.stderr;
  assert.match(output, /DryRun/);
  assert.ok(await fs.stat(installDir).then(() => true), "DryRun 不应删除目录");

  const script = await fs.readFile(path.join(ROOT, "scripts", "uninstall.ps1"), "utf8");
  // 卸载必须清掉：加载项侧载项、登录自启、卸载项登记
  assert.match(script, /Wef\\Developer/);
  assert.match(script, /CurrentVersion\\Run/);
  assert.match(script, /Uninstall\\\$AppId/);
});

test("托盘程序：关键行为与注册约定", async () => {
  const tray = await fs.readFile(path.join(ROOT, "scripts", "tray.ps1"), "utf8");
  for (const expected of ["打开任务窗格", "打开输出目录", "重启服务", "查看日志", "退出"]) {
    assert.ok(tray.includes(expected), `托盘菜单缺少“${expected}”`);
  }
  assert.match(tray, /--tls-key/, "托盘启动服务时使用安装目录证书");
  assert.match(tray, /service\.pid/);
  assert.match(tray, /api\/health/, "启动后做健康检查");

  const install = await fs.readFile(path.join(ROOT, "scripts", "install.ps1"), "utf8");
  assert.match(install, /setup-tls\.mjs/);
  assert.match(install, /Wef\\Developer/, "安装脚本必须注册开发者侧载项");
  assert.match(install, /CurrentVersion\\Run/, "安装脚本注册登录自启");
  assert.match(install, /start-hidden\.vbs/);
  assert.match(install, /Uninstall/);
});

test("安装脚本：-File 调用且省略 -SourceDir 时能自行定位", { skip: !isWindows }, async () => {
  // 回归测试：PowerShell 5.1 在 [CmdletBinding()] + -File 时 $PSScriptRoot 于参数默认值中为空，
  // 脚本必须自己在函数体里兜底解析，否则真实安装会直接失败。
  const dir = await fs.mkdtemp(path.join(process.env.TEMP || "/tmp", "opt-install-noroot-"));
  await makeFixture(dir);
  const installDir = path.join(dir, "installed");
  const result = ps([
    "-File", path.join(dir, "install.ps1"),
    "-InstallDir", installDir,
    "-Port", "3457",
    "-SkipTls", "-SkipRegistry", "-SkipShortcuts", "-NoStart",
  ]);
  assert.equal(result.status, 0, `安装失败:\n${result.stdout}\n${result.stderr}`);
  assert.match(result.stdout, /安装完成/, "应走到安装完成");
  await fs.access(path.join(installDir, "app", "server", "index.js"));
  assert.equal((await fs.readFile(path.join(installDir, "port.txt"), "utf8")).trim(), "3457");
});

test("卸载脚本：省略 -InstallDir 时从自身位置推断安装目录", { skip: !isWindows }, async () => {
  const dir = await fs.mkdtemp(path.join(process.env.TEMP || "/tmp", "opt-uninstall-noroot-"));
  const installDir = path.join(dir, "OfficePdfTools");
  const toolsDir = path.join(installDir, "tools");
  await fs.mkdir(toolsDir, { recursive: true });
  await fs.writeFile(path.join(installDir, "port.txt"), "3458");
  await fs.copyFile(path.join(ROOT, "scripts", "uninstall.ps1"), path.join(toolsDir, "uninstall.ps1"));
  const result = ps([
    "-File", path.join(toolsDir, "uninstall.ps1"),
    "-DryRun",
  ]);
  assert.equal(result.status, 0, `卸载脚本失败:\n${result.stdout}\n${result.stderr}`);
  const output = result.stdout + result.stderr;
  assert.ok(output.includes(installDir), `应推断出安装目录 ${installDir}，实际输出:\n${output}`);
  assert.ok(await fs.stat(installDir).then(() => true), "DryRun 不应删除目录");
});

test("安装包构建脚本：参数与产物约定", async () => {
  const build = await fs.readFile(path.join(ROOT, "scripts", "build-installer.mjs"), "utf8");
  assert.match(build, /npm ci|npm\.cmd/);
  assert.match(build, /--omit=dev/);
  assert.match(build, /nodejs\.org\/dist/, "默认下载官方 Node 运行时");
  assert.match(build, /--no-download/);
  assert.match(build, /iexpress/i, "生成自解压 EXE");
  assert.match(build, /payload\.zip/);
  assert.match(build, /install\.cmd/);

  const pkg = JSON.parse(await fs.readFile(path.join(ROOT, "package.json"), "utf8"));
  assert.ok(pkg.scripts["build:installer"], "缺少 build:installer 脚本");
  assert.ok(pkg.scripts["install:app"] && pkg.scripts["uninstall:app"]);

  const gitignore = await fs.readFile(path.join(ROOT, ".gitignore"), "utf8");
  assert.match(gitignore, /^dist\/$/m, "dist/ 不应提交");

  await fs.access(path.join(ROOT, "addin", "assets", "app.ico"), fs.constants.R_OK);
});

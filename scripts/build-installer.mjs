#!/usr/bin/env node
/**
 * 生成 Windows 安装包：
 *   dist/office-pdf-tools-<版本>-setup.zip   解压后运行 install.cmd
 *   dist/office-pdf-tools-<版本>-setup.exe   自解压安装程序（IExpress 生成，双击即装）
 *
 * 用法:
 *   npm run build:installer                   # 下载官方 Node 运行时（约 30MB）
 *   npm run build:installer -- --no-download  # 离线：使用本机 node.exe
 *   npm run build:installer -- --no-exe       # 只生成 ZIP
 */
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawn } from "node:child_process";
import JSZip from "jszip";
import { pngToIco } from "./gen-ico.mjs";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const argv = process.argv.slice(2);
const hasFlag = (name) => argv.includes(name);

const pkg = JSON.parse(await fs.readFile(path.join(ROOT, "package.json"), "utf8"));
const VERSION = pkg.version;
const DIST = path.join(ROOT, "dist");
const STAGE = path.join(DIST, "office-pdf-tools-" + VERSION);
const PAYLOAD = path.join(STAGE, "payload");

function run(command, args, opts = {}) {
  return new Promise((resolve, reject) => {
    // Windows 上 npm 是 npm.cmd，Node 20+ 不允许直接 spawn .cmd（EINVAL），需经 shell
    const child = spawn(command, args, {
      cwd: opts.cwd || ROOT,
      stdio: "inherit",
      windowsHide: true,
      shell: process.platform === "win32",
    });
    child.on("error", reject);
    child.on("close", (code) => (code === 0 ? resolve() : reject(new Error(command + " 退出码 " + code))));
  });
}

async function copyTree(from, to, skip = []) {
  await fs.mkdir(to, { recursive: true });
  for (const entry of await fs.readdir(from, { withFileTypes: true })) {
    if (skip.includes(entry.name)) continue;
    const src = path.join(from, entry.name);
    const dst = path.join(to, entry.name);
    if (entry.isDirectory()) await copyTree(src, dst, skip);
    else await fs.copyFile(src, dst);
  }
}

async function dirSize(dir) {
  let total = 0;
  for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) total += await dirSize(full);
    else total += (await fs.stat(full)).size;
  }
  return total;
}

async function zipDir(dir, prefix) {
  const zip = new JSZip();
  async function walk(current, rel) {
    for (const entry of await fs.readdir(current, { withFileTypes: true })) {
      const full = path.join(current, entry.name);
      const name = rel ? rel + "/" + entry.name : entry.name;
      if (entry.isDirectory()) await walk(full, name);
      else zip.file(name, await fs.readFile(full));
    }
  }
  await walk(dir, prefix || "");
  return zip;
}

async function downloadNodeRuntime(targetDir) {
  const index = await (await fetch("https://nodejs.org/dist/index.json")).json();
  const lts = index.find((v) => v.lts && /^v\d+\.\d+\.\d+$/.test(v.version));
  if (!lts) throw new Error("无法获取 Node LTS 列表");
  const version = lts.version;
  const url = "https://nodejs.org/dist/" + version + "/node-" + version + "-win-x64.zip";
  console.log("==> 下载 Node 运行时 " + version + " (win-x64)");
  const res = await fetch(url);
  if (!res.ok) throw new Error("下载失败: HTTP " + res.status);
  const zip = await JSZip.loadAsync(Buffer.from(await res.arrayBuffer()));
  const exe = Object.keys(zip.files).find((n) => /\/node\.exe$/.test(n));
  const license = Object.keys(zip.files).find((n) => /\/LICENSE$/.test(n));
  if (!exe) throw new Error("压缩包中没有 node.exe");
  await fs.mkdir(targetDir, { recursive: true });
  await fs.writeFile(path.join(targetDir, "node.exe"), await zip.file(exe).async("nodebuffer"));
  if (license) await fs.writeFile(path.join(targetDir, "LICENSE-node.txt"), await zip.file(license).async("nodebuffer"));
  return version;
}

async function copyLocalNode(targetDir) {
  await fs.mkdir(targetDir, { recursive: true });
  await fs.copyFile(process.execPath, path.join(targetDir, "node.exe"));
  console.log("==> 使用本机 Node 运行时: " + process.execPath + " (" + process.version + ")");
  return process.version;
}

async function pruneNodeModules(appDir) {
  const at = path.join(appDir, "node_modules", "@napi-rs");
  let removed = 0;
  try {
    for (const entry of await fs.readdir(at, { withFileTypes: true })) {
      if (!entry.isDirectory()) continue;
      if (entry.name === "canvas" || entry.name === "canvas-win32-x64-msvc") continue;
      await fs.rm(path.join(at, entry.name), { recursive: true, force: true });
      removed++;
    }
  } catch { /* 目录不存在 */ }
  return removed;
}

const SETUP_CMD = [
  "@echo off",
  "chcp 65001 >nul",
  "setlocal",
  "title Office PDF 工具 - 安装",
  "set STAGE=%~dp0stage",
  "echo.",
  "echo ============================================",
  "echo   Office PDF 工具 - 安装程序",
  "echo ============================================",
  "echo.",
  "powershell -NoProfile -ExecutionPolicy Bypass -Command \"Expand-Archive -LiteralPath '%~dp0payload.zip' -DestinationPath '%STAGE%' -Force\"",
  "if errorlevel 1 (",
  "  echo [错误] 解压安装包失败。",
  "  pause",
  "  exit /b 1",
  ")",
  "powershell -NoProfile -ExecutionPolicy Bypass -File \"%STAGE%\\install.ps1\" %*",
  "set RC=%ERRORLEVEL%",
  "echo.",
  "if not \"%RC%\"==\"0\" echo [提示] 安装脚本返回代码 %RC%",
  "pause",
  "exit /b %RC%",
  "",
].join("\r\n");

const INSTALL_CMD = [
  "@echo off",
  "chcp 65001 >nul",
  "title Office PDF 工具 - 安装",
  "powershell -NoProfile -ExecutionPolicy Bypass -File \"%~dp0install.ps1\" %*",
  "pause",
  "",
].join("\r\n");

const README_TXT = [
  "Office PDF 工具 v" + VERSION + " —— 安装说明",
  "============================================",
  "",
  "【一键安装】",
  "  双击 setup.exe（自解压安装程序）按提示完成；",
  "  或解压本压缩包后双击 install.cmd（不需要管理员权限，装到当前用户）。",
  "",
  "【安装做了什么】",
  "  1. 程序复制到 %LOCALAPPDATA%\\OfficePdfTools",
  "  2. 生成并信任 localhost 证书（Office 任务窗格必须走 HTTPS）",
  "  3. 把加载项注册到 Office 开发者侧载项 → 之后每次打开 Word/Excel/PowerPoint",
  "     都会自动出现“PDF 工具”选项卡，无需手动加载",
  "  4. 注册登录自启（托盘程序自动拉起本地服务）",
  "  5. 创建开始菜单快捷方式与“应用和功能”卸载项",
  "",
  "【安装之后】",
  "  * 可能弹出一次证书确认框，点“是”",
  "  * 完全退出并重新打开 Word / Excel / PowerPoint，功能区即出现“PDF 工具”",
  "  * 通知区域的托盘图标可打开任务窗格 / 输出目录 / 查看日志 / 退出",
  "",
  "【常用参数】",
  "  install.cmd -Port 3001                       指定端口（默认 3000，占用则自动顺延）",
  "  install.cmd -NoAutoStart                     不注册登录自启",
  "  install.cmd -Silent                          静默安装",
  "  install.cmd -InstallDir D:\\OfficePdfTools   指定安装目录",
  "",
  "【卸载】",
  "  开始菜单 → Office PDF 工具 → 卸载 Office PDF 工具",
  "  或运行 %LOCALAPPDATA%\\OfficePdfTools\\tools\\uninstall.ps1（加 -KeepOutput 可保留已转换文件）",
  "",
  "项目主页: https://github.com/ZizzSaint/office-pdf-tools",
  "",
].join("\r\n");

async function buildExe() {
  const iexpress = path.join(process.env.SystemRoot || "C:\\Windows", "System32", "iexpress.exe");
  try {
    await fs.access(iexpress);
  } catch {
    console.log("==> 未找到 iexpress.exe，跳过自解压 EXE");
    return null;
  }
  const sfxDir = path.join(DIST, "_sfx");
  await fs.rm(sfxDir, { recursive: true, force: true });
  await fs.mkdir(sfxDir, { recursive: true });
  await fs.copyFile(path.join(STAGE, "payload.zip"), path.join(sfxDir, "payload.zip"));
  await fs.writeFile(path.join(sfxDir, "setup.cmd"), SETUP_CMD, "utf8");

  const exePath = path.join(DIST, "office-pdf-tools-" + VERSION + "-setup.exe");
  await fs.rm(exePath, { force: true });
  const sedPath = path.join(sfxDir, "package.sed");
  const sed = [
    "[Version]",
    "Class=IEXPRESS",
    "SEDVersion=3",
    "[Options]",
    "PackagePurpose=InstallApp",
    "ShowInstallProgramWindow=1",
    "HideExtractAnimation=0",
    "UseLongFileName=1",
    "InsideCompressed=0",
    "CAB_FixedSize=0",
    "CAB_ResvCodeSigning=0",
    "RebootMode=N",
    "InstallPrompt=",
    "DisplayLicense=",
    "FinishMessage=",
    "TargetName=" + exePath,
    "FriendlyName=Office PDF Tools Setup",
    "AppLaunched=setup.cmd",
    "PostInstallCmd=<None>",
    "AdminQuietInstCmd=",
    "UserQuietInstCmd=",
    "SourceFiles=SourceFiles",
    "[Strings]",
    'FILE0="payload.zip"',
    'FILE1="setup.cmd"',
    "[SourceFiles]",
    "SourceFiles0=" + sfxDir + "\\",
    "[SourceFiles0]",
    "%FILE0%=",
    "%FILE1%=",
    "",
  ].join("\r\n");
  await fs.writeFile(sedPath, sed, "utf8");
  console.log("==> 生成自解压安装程序 (IExpress)…");
  try {
    await run(iexpress, ["/N", "/Q", sedPath]);
  } catch (err) {
    console.warn("    IExpress 失败: " + err.message + "（保留 ZIP 安装包）");
    return null;
  }
  return (await fs.stat(exePath).catch(() => null)) ? exePath : null;
}

async function main() {
  console.log("==> 构建 Office PDF 工具 安装包 v" + VERSION);
  await fs.rm(STAGE, { recursive: true, force: true });
  await fs.mkdir(PAYLOAD, { recursive: true });

  console.log("==> 复制程序文件");
  for (const dir of ["server", "addin", "scripts", "manifest"]) {
    await copyTree(path.join(ROOT, dir), path.join(PAYLOAD, dir));
  }
  for (const file of ["package.json", "package-lock.json", "README.md", "LICENSE"]) {
    await fs.copyFile(path.join(ROOT, file), path.join(PAYLOAD, file));
  }

  console.log("==> 安装生产依赖 (npm ci --omit=dev)");
  const npm = process.platform === "win32" ? "npm.cmd" : "npm";
  await run(npm, ["ci", "--omit=dev", "--no-audit", "--no-fund"], { cwd: PAYLOAD });
  const pruned = await pruneNodeModules(PAYLOAD);
  if (pruned) console.log("    已移除 " + pruned + " 个非 Windows 平台原生依赖");

  const nodeDir = path.join(PAYLOAD, "node");
  const nodeVersion = hasFlag("--no-download")
    ? await copyLocalNode(nodeDir)
    : await downloadNodeRuntime(nodeDir).catch(async (err) => {
        console.warn("    下载失败(" + err.message + ")，回退本机 Node");
        return copyLocalNode(nodeDir);
      });

  await pngToIco(path.join(ROOT, "addin/assets/icon-128.png"), path.join(PAYLOAD, "scripts/app.ico"));
  await fs.copyFile(path.join(PAYLOAD, "scripts/app.ico"), path.join(ROOT, "addin/assets/app.ico")).catch(() => {});

  console.log("==> 写入安装脚本");
  await fs.copyFile(path.join(ROOT, "scripts/install.ps1"), path.join(STAGE, "install.ps1"));
  await fs.copyFile(path.join(ROOT, "scripts/uninstall.ps1"), path.join(STAGE, "uninstall.ps1"));
  await fs.writeFile(path.join(STAGE, "install.cmd"), INSTALL_CMD, "utf8");
  await fs.writeFile(path.join(STAGE, "安装说明.txt"), README_TXT, "utf8");

  console.log("==> 打包 ZIP");
  const zip = await zipDir(STAGE, "");
  const zipBuffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } });
  const zipPath = path.join(DIST, "office-pdf-tools-" + VERSION + "-setup.zip");
  await fs.writeFile(zipPath, zipBuffer);

  const payloadZip = await zipDir(PAYLOAD, "payload");
  payloadZip.file("install.ps1", await fs.readFile(path.join(STAGE, "install.ps1")));
  payloadZip.file("uninstall.ps1", await fs.readFile(path.join(STAGE, "uninstall.ps1")));
  payloadZip.file("安装说明.txt", await fs.readFile(path.join(STAGE, "安装说明.txt")));
  const payloadBuffer = await payloadZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 9 } });
  await fs.writeFile(path.join(STAGE, "payload.zip"), payloadBuffer);

  const exePath = hasFlag("--no-exe") ? null : await buildExe();

  const payloadSize = await dirSize(PAYLOAD);
  console.log("");
  console.log("=== 构建完成 ===");
  console.log("  Node 运行时 : " + nodeVersion);
  console.log("  程序文件    : " + (payloadSize / 1024 / 1024).toFixed(1) + " MB");
  console.log("  ZIP         : " + path.relative(ROOT, zipPath) + " (" + (zipBuffer.length / 1024 / 1024).toFixed(1) + " MB)");
  if (exePath) {
    const stat = await fs.stat(exePath);
    console.log("  EXE         : " + path.relative(ROOT, exePath) + " (" + (stat.size / 1024 / 1024).toFixed(1) + " MB)");
  }
  console.log("  目录        : " + path.relative(ROOT, STAGE));
}

main().catch((err) => {
  console.error("构建失败:", err);
  process.exit(1);
});

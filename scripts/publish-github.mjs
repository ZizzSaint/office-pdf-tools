#!/usr/bin/env node
/**
 * 把本项目发布到 GitHub。
 *
 * 用法：
 *   npm run publish:github                      # 使用已登录的 git 凭据 / 环境变量
 *   npm run publish:github -- --name my-repo --private
 *   GITHUB_TOKEN=ghp_xxx npm run publish:github # 用 token 创建仓库
 *
 * 凭据获取顺序：--token 参数 → GITHUB_TOKEN / GH_TOKEN 环境变量 → git 已保存的凭据。
 * 本脚本不会打印任何令牌内容。
 */
import { spawn } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const argv = process.argv.slice(2);
const flag = (name, short) => argv.includes(name) || (short && argv.includes(short));
const value = (name, fallback = "") => {
  const i = argv.indexOf(name);
  if (i >= 0 && argv[i + 1] && !argv[i + 1].startsWith("-")) return argv[i + 1];
  const inline = argv.find((a) => a.startsWith(name + "="));
  return inline ? inline.slice(name.length + 1) : fallback;
};

const repoName = value("--name", "office-pdf-tools");
const isPrivate = flag("--private");
const description = value("--description", "Office 加载项 + 本地服务：Word/Excel/PowerPoint ⇄ PDF、图片↔PDF、文档拆分合并");

function run(command, args, { input, env, cwd = ROOT, allowFail = false } = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd,
      env: { ...process.env, ...(env || {}) },
      windowsHide: true,
      stdio: [input ? "pipe" : "ignore", "pipe", "pipe"],
    });
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (c) => { stdout += c.toString(); });
    child.stderr.on("data", (c) => { stderr += c.toString(); });
    if (input) {
      child.stdin.write(input);
      child.stdin.end();
    }
    child.on("error", (err) => (allowFail ? resolve({ code: -1, stdout, stderr: String(err) }) : reject(err)));
    child.on("close", (code) => {
      if (code !== 0 && !allowFail) reject(new Error(`${command} ${args.join(" ")} 失败 (${code}): ${(stderr || stdout).trim()}`));
      else resolve({ code, stdout, stderr });
    });
  });
}

async function gitCredentialToken() {
  // 通过 git 的凭据助手读取已保存的 GitHub 凭据（不回显内容）
  const { code, stdout } = await run("git", ["credential", "fill"], {
    input: "protocol=https\nhost=github.com\n\n",
    allowFail: true,
  });
  if (code !== 0) return null;
  const password = /password=(.*)/.exec(stdout)?.[1]?.trim();
  return password || null;
}

async function api(url, token, { method = "GET", body } = {}) {
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/vnd.github+json",
      "User-Agent": "office-pdf-tools-publisher",
      ...(body ? { "Content-Type": "application/json" } : {}),
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  let data = null;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    data = { raw: text };
  }
  return { ok: res.ok, status: res.status, data };
}

async function main() {
  console.log("==> 准备本地仓库");
  const gitDir = path.join(ROOT, ".git");
  const hasGit = await fs.stat(gitDir).then(() => true).catch(() => false);
  if (!hasGit) {
    await run("git", ["init", "-b", "main"]);
    console.log("    已初始化 git 仓库 (main)");
  } else {
    console.log("    已存在 git 仓库");
  }

  await run("git", ["add", "-A"]);
  const status = await run("git", ["status", "--porcelain"]);
  if (status.stdout.trim()) {
    await run("git", ["commit", "-m", "chore: publish office-pdf-tools"]);
    console.log("    已提交本地改动");
  } else {
    console.log("    没有需要提交的改动");
  }
  const branch = (await run("git", ["rev-parse", "--abbrev-ref", "HEAD"])).stdout.trim() || "main";

  let token = value("--token", "") || process.env.GITHUB_TOKEN || process.env.GH_TOKEN || "";
  if (!token) {
    token = (await gitCredentialToken()) || "";
    if (token) console.log("==> 使用 git 凭据助手中保存的 GitHub 凭据（内容不显示）");
  }
  if (!token) {
    console.error("未找到可用的 GitHub 凭据。请设置 GITHUB_TOKEN 环境变量（需要 repo 权限），或先执行 git 登录。");
    process.exit(2);
  }

  const me = await api("https://api.github.com/user", token);
  if (!me.ok) {
    console.error(`GitHub 凭据无效（HTTP ${me.status}）。`);
    process.exit(2);
  }
  const owner = me.data.login;
  console.log(`==> 已认证为 ${owner}`);

  let repo = await api(`https://api.github.com/repos/${owner}/${repoName}`, token);
  if (!repo.ok) {
    const created = await api("https://api.github.com/user/repos", token, {
      method: "POST",
      body: {
        name: repoName,
        description,
        private: isPrivate,
        has_issues: true,
        has_wiki: false,
        auto_init: false,
      },
    });
    if (!created.ok) {
      console.error(`创建仓库失败 (HTTP ${created.status}): ${created.data?.message || ""}`);
      process.exit(1);
    }
    repo = created;
    console.log(`==> 已创建仓库 ${owner}/${repoName}`);
  } else {
    console.log(`==> 使用已存在的仓库 ${owner}/${repoName}`);
  }

  const remoteUrl = `https://github.com/${owner}/${repoName}.git`;
  const remotes = await run("git", ["remote"], { allowFail: true });
  if (remotes.stdout.split(/\s+/).includes("origin")) {
    await run("git", ["remote", "set-url", "origin", remoteUrl]);
  } else {
    await run("git", ["remote", "add", "origin", remoteUrl]);
  }

  console.log("==> 推送代码");
  // 用一次性的凭据参数推送，避免把令牌写进 .git/config
  const pushArgs = ["-c", "credential.helper=", "-c", `http.https://github.com/.extraheader=Authorization: Basic ${Buffer.from(`x-access-token:${token}`).toString("base64")}`, "push", "-u", "origin", branch];
  const pushed = await run("git", pushArgs, { allowFail: true, env: { GIT_TERMINAL_PROMPT: "0" } });
  if (pushed.code !== 0) {
    console.error("推送失败:", (pushed.stderr || pushed.stdout).trim());
    process.exit(1);
  }
  console.log(`\n完成：${repo.data.html_url || remoteUrl}`);
}

main().catch((err) => {
  console.error("发布失败:", err.message);
  process.exit(1);
});

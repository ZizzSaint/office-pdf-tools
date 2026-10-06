# 发布到 GitHub

项目自带一键发布脚本 `scripts/publish-github.mjs`（`npm run publish:github`），它会：

1. 初始化/复用本地 git 仓库并提交改动
2. 用 GitHub API 创建仓库（已存在则复用）
3. 添加 `origin` 远程并推送当前分支

## 凭据从哪来

按顺序尝试，**脚本不会打印任何令牌内容**：

1. `--token <token>` 参数
2. `GITHUB_TOKEN` / `GH_TOKEN` 环境变量
3. git 已保存的 GitHub 凭据（通过 `git credential fill` 读取，即 Git Credential Manager / `git config credential.helper` 中保存的账号）

> 需要的令牌权限：`repo`（私有仓库）或 `public_repo`（公开仓库）。在 GitHub → Settings → Developer settings → Personal access tokens 生成。

## 用法

```bash
# 用已登录的 git 凭据，创建/推送到 ZizzSaint/office-pdf-tools
npm run publish:github

# 指定仓库名与可见性
npm run publish:github -- --name my-office-tools --private

# 直接用令牌
GITHUB_TOKEN=ghp_xxx npm run publish:github
```

推送时脚本使用一次性的 `http.extraheader` 参数传递凭据，**不会把令牌写入 `.git/config`**。

## 手动发布（不使用脚本）

```bash
git init -b main
git add -A
git commit -m "feat: office-pdf-tools v1.0.0"
# 在 GitHub 网页上新建空仓库（不要勾选 README）
git remote add origin https://github.com/<user>/<repo>.git
git push -u origin main
```

## 发布前自检

```bash
npm test                     # 32 项测试
npm run smoke                # 端到端 12 条链路
npm run validate:manifest    # 微软官方清单校验（需要联网）
```

## 目录/文件约定

- 不提交：`node_modules/`、`output/`、`tmp/`、`.certs/`、日志（见 `.gitignore`）
- `package-lock.json` 建议提交，保证依赖可复现
- 截图与图标放在 `docs/images/`、`addin/assets/`、`manifest/assets/`

# 桌面版安装包（一键安装 · Office 自动加载）

> 目标：**装一次，之后每次打开 Word / Excel / PowerPoint，"PDF 工具" 选项卡自动出现**，不需要手动侧载、不需要记命令。

## 1. 它是什么

把「Office 加载项 + 本地转换服务」打包成一个 Windows 桌面软件：

```
office-pdf-tools-1.1.0-setup.exe      ← 双击安装（自解压）
office-pdf-tools-1.1.0-setup.zip      ← 免安装包，解压后运行 install.cmd
   ├─ payload/          程序本体（含 Node 运行时、生产依赖、加载项页面）
   ├─ install.ps1       安装脚本
   ├─ uninstall.ps1     卸载脚本
   └─ 安装说明.txt
```

构建方式（在项目目录）：

```bash
npm run build:installer                # 下载官方 Node LTS 运行时后打包
npm run build:installer -- --no-download   # 离线：用本机 node.exe
npm run build:installer -- --no-exe        # 只出 ZIP，不出自解压 EXE
```

## 2. 安装后会变成什么样

| 项目 | 位置 / 值 |
| --- | --- |
| 程序目录 | `%LOCALAPPDATA%\OfficePdfTools\app`（含自带 `node\node.exe`，无需另装 Node） |
| 输出目录 | `%LOCALAPPDATA%\OfficePdfTools\output` |
| 证书 | `%LOCALAPPDATA%\OfficePdfTools\tls\localhost.crt|key`（安装时生成并加入“当前用户 → 受信任的根证书”） |
| 端口 | `port.txt`，默认 3000；被占用会自动顺延并同步改写清单 |
| 加载项注册 | `HKCU\Software\Microsoft\Office\16.0\Wef\Developer`（同时写 15.0） |
| 登录自启 | `HKCU\Software\Microsoft\Windows\CurrentVersion\Run → OfficePdfTools`（运行托盘程序） |
| 卸载登记 | `HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\OfficePdfTools` |
| 快捷方式 | 开始菜单 → `Office PDF 工具`（打开任务窗格 / 打开输出目录 / 查看日志 / 卸载） |
| 托盘图标 | 通知区域，可打开面板、重启服务、看日志、退出 |

**全程写入当前用户范围，不需要管理员权限。**

## 3. 为什么“打开 Office 就自动加载”

Office 桌面版识别加载项有三条路径，本安装包用的是第 1 条：

| 方式 | 是否自动出现在功能区 | 需要什么 | 适用场景 |
| --- | --- | --- | --- |
| **① 开发者侧载项（本安装包）** | ✅ 每次启动 Office 自动出现 | 当前用户注册表一项 | 个人 / 小团队，任何版本 Office 2016+ |
| ② 受信任的加载项目录（共享文件夹） | ✅ | 需要配置信任中心 + 网络共享 | 局域网内团队分发 |
| ③ 集中部署（Microsoft 365 管理中心） | ✅ | 管理员 + 用户以工作/学校账号登录 + 清单托管在公网 HTTPS（或上架 AppSource） | 企业统一分发；个人版/永久版 Office 不支持 |

安装脚本写入的就是 ①：

```
HKCU\Software\Microsoft\Office\16.0\Wef\Developer
    名称 = <安装目录>\manifest.xml
    数据 = <安装目录>\manifest.xml
```

Office 启动时会扫描该键，把清单里的 Add-in Commands（我们的「PDF 工具」自定义选项卡 + 3 个按钮）直接渲染到功能区。**这项注册长期有效**，所以不需要每次开机重新侧载——安装包做的正是把手动执行的 `npm run sideload` 自动化。

## 4. 本地服务为什么必须常驻

Office 加载项的任务窗格必须从一个 HTTPS 地址加载，而 PDF 转换（LibreOffice/Office COM、pdf.js 渲染）也跑在本机进程里。所以安装包注册了**登录自启的托盘进程**：

```
用户登录 → start-hidden.vbs（隐藏窗口）→ powershell tray.ps1 → node server/index.js --https --port <N>
                                                              ↑ 使用安装目录证书，随托盘退出而结束
```

托盘常驻只在通知区域显示一个图标，占用极小；如果你不希望自启，可以用 `-NoAutoStart` 安装，之后手动从开始菜单启动。

## 5. 安装步骤

1. 关闭正在运行的 Word / Excel / PowerPoint（可选，但推荐）
2. 双击 `office-pdf-tools-1.1.0-setup.exe`
3. 若出现“Windows 已保护你的电脑”（SmartScreen，因为安装包未做代码签名）→ 更多信息 → 仍要运行
4. 若弹出证书确认框 → 点是（这是本地 HTTPS 证书，仅用于 `localhost`）
5. 安装脚本会启动服务并做健康检查，最后提示“安装完成”
6. 打开 Word / Excel / PowerPoint → 功能区出现 **PDF 工具** 选项卡

> 已有文档打开时也可以安装，但加载项要等 Office 下次启动才会出现。

## 6. 命令行 / 静默安装

```powershell
# 默认安装
powershell -ExecutionPolicy Bypass -File install.ps1

# 指定端口、静默
powershell -ExecutionPolicy Bypass -File install.ps1 -Port 3001 -Silent

# 装到 D 盘、不自启
powershell -ExecutionPolicy Bypass -File install.ps1 -InstallDir D:\OfficePdfTools -NoAutoStart

# 只装文件不碰注册表/证书（排查用）
powershell -ExecutionPolicy Bypass -File install.ps1 -SkipRegistry -SkipTls -SkipShortcuts -NoStart

# 预演（不写任何东西）
powershell -ExecutionPolicy Bypass -File install.ps1 -DryRun
```

| 参数 | 说明 |
| --- | --- |
| `-Port <n>` | 服务端口，默认 3000 |
| `-InstallDir <path>` | 安装目录，默认 `%LOCALAPPDATA%\OfficePdfTools` |
| `-NoAutoStart` | 不写登录自启 |
| `-SkipRegistry` / `-SkipTls` / `-SkipShortcuts` | 跳过对应步骤 |
| `-NoStart` | 安装后不立即启动服务 |
| `-Silent` | 少输出 |
| `-DryRun` | 只打印计划 |

## 7. 升级与卸载

- **升级**：直接运行新版本安装包即可，会覆盖 `app`、重建工具脚本、复用已有证书与 `output` 目录；端口保持不变。
- **卸载**：开始菜单 → `Office PDF 工具` → `卸载 Office PDF 工具`，或
  ```powershell
  powershell -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\OfficePdfTools\tools\uninstall.ps1"
  # 保留已转换的文件：
  powershell -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\OfficePdfTools\tools\uninstall.ps1" -KeepOutput
  ```
  卸载会停止服务、移除加载项注册（Office 下次启动不再显示选项卡）、删除自启项/快捷方式/程序文件。
- 证书默认保留（可能被其它 Office 加载项开发共用）；彻底清除：
  `certutil -user -delstore Root "Developer CA for Microsoft Office Add-ins"`

## 8. 常见问题

**装完 Office 里没有“PDF 工具”？**
1. 任务管理器确认没有残留的 WINWORD/EXCEL/POWERPNT，再重新打开
2. 检查注册表项是否存在：`reg query "HKCU\Software\Microsoft\Office\16.0\Wef\Developer"`
3. Excel 里：文件 → 选项 → 加载项 → 管理(COM 加载项/我的加载项) 看是否被禁用
4. 托盘图标是否在？（不在说明服务/托盘没起来，右键查看日志，或运行 `tools\tray.ps1` 看报错）

**任务窗格白屏 / 显示“无法访问此页面”？**
- 托盘右键 → 打开任务窗格，先确认 `https://localhost:<端口>/taskpane.html` 能在浏览器打开
- 浏览器报证书错误 → 重新安装（会重建证书），或手动执行 `app\scripts\setup-tls.mjs`

**端口冲突？**
- 安装时自动顺延；安装后想改：编辑 `%LOCALAPPDATA%\OfficePdfTools\port.txt` 并重新运行安装包（`-Port <n>`），它会改写 manifest 与端口文件

**杀软/EDR 报“脚本行为可疑”？**
- 安装脚本会写 HKCU 注册表、生成证书、注册自启——这些都是正常安装行为；企业环境建议先在测试机放行，或改用第 3 种集中部署方式

**想让整个公司自动装、自动更新？**
- 用 Microsoft 365 管理中心的**集中部署**：把清单托管到公网 HTTPS（`npm run set:url -- https://pdf.example.com` 可批量改写清单地址），上传清单后全员 Word/Excel/PowerPoint 自动出现该加载项，且支持统一撤下

## 9. 与“开发者手动安装”的对比

| | 手动（npm run sideload） | 桌面安装包 |
| --- | --- | --- |
| 需要 Node.js | ✅ 需要 | ❌ 自带运行时 |
| 需要手动启动服务 | ✅ 每次开机 `npm run https` | ❌ 托盘随登录自启 |
| 证书 | 首次运行自动生成 | 安装时生成 |
| Office 自动加载 | ✅（写同一注册表项） | ✅ |
| 适合人群 | 开发者、二次开发 | 普通用户 |

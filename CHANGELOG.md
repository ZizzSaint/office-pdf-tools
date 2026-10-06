# 更新日志

## v1.1.0

**新增：Windows 桌面安装包（一键安装 · Office 自动加载）**

- `npm run build:installer` 构建安装包：产出 `office-pdf-tools-<版本>-setup.exe`（IExpress 自解压）与 `setup.zip`，
  内置官方 Node LTS 运行时（约 72 MB / 232 MB 安装体积），用户**无需安装 Node.js**
- `scripts/install.ps1`：程序复制到 `%LOCALAPPDATA%\OfficePdfTools`、生成并信任 localhost 证书、
  注册 Office 开发者侧载项（16.0 / 15.0，**装完即自动加载**）、注册登录自启、
  创建开始菜单快捷方式与「应用和功能」卸载项；全程当前用户权限，不需要管理员
- `scripts/tray.ps1`：托盘守护进程，随登录启动并拉起本地 HTTPS 服务，
  右键可打开任务窗格 / 输出目录 / 日志、重启服务、退出
- `scripts/uninstall.ps1`：停止服务、移除加载项注册与自启项、删除程序文件（`-KeepOutput` 可保留转换结果）
- 服务端支持从安装目录读取证书：`--tls-key` / `--tls-cert` / `--tls-pfx`
- `npm run set:url -- https://pdf.example.com`：一键把清单地址改成公网域名，便于企业集中部署
- 新增 8 项安装包测试：PowerShell 语法校验、`-DryRun` 无副作用、真实复制与清单端口改写、
  卸载约定、托盘行为、构建产物约定，以及两条真实调用路径的回归测试
  （`-File` 调用且省略 `-SourceDir`、省略 `-InstallDir` 的卸载）

## v1.0.0

首个可用版本。

**基础功能**

- 一份清单同时支持 Word / Excel / PowerPoint，功能区新增“PDF 工具”选项卡（导出 PDF / PDF 转换 / 拆分合并）
- 当前文档 → PDF：调用 Office 原生导出，无需上传、完全保真
- 转换结果支持下载、写入输出目录、打开文件/文件夹，`.docx`/`.pptx` 可插入回当前文档

**附加功能 1：格式互转**

- Office → PDF：Microsoft Office COM（Windows）/ LibreOffice（跨平台）双引擎自动选择
- PDF → Word / Excel / PowerPoint：文本（可编辑重排）与图片（逐页高保真）两种模式
- PDF → 图片：PNG / JPEG / WebP，36–1200 DPI，页码范围，多页打包 ZIP
- 图片 → PDF：多图排序、页面尺寸/方向/边距、1/2/4/6/9 宫格、每图单独成 PDF

**附加功能 2：拆分与合并**

- PDF：每 N 页 / 页码范围 / 每页拆分；逐页拼接合并
- Word：按标题级别 / 分节符拆分；合并正文并迁移图片、样式、编号定义
- Excel：按工作表 / 每 N 行拆分；工作表汇总合并
- PowerPoint：每 N 张 / 幻灯片范围拆分；追加幻灯片并连同版式母版搬迁

**工程**

- 自研 OOXML 关系深拷贝与孤立部件回收引擎
- 内存任务队列：进度轮询、取消、TTL 清理
- REST API（异步任务 + 同步直出）与双语任务窗格
- 32 项自动化测试（含真实 Office → PDF 转换）、端到端冒烟脚本、微软官方清单校验
- 侧载/取消侧载脚本、端口同步脚本、图标生成脚本、一键发布脚本

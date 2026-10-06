# 安装与侧载指南

本加载项是标准的 Office Web Add-in（任务窗格 + Add-in Commands），通过 `manifest/manifest.xml` 载入；
一份清单同时声明了 **Word（Document）/ Excel（Workbook）/ PowerPoint（Presentation）** 三个宿主。

> **普通用户请看 [桌面版安装包](DESKTOP-APP.md)**：双击 `setup.exe` 即可，
> 自带 Node 运行时、自动生成证书、注册 Office 自动加载与登录自启，装完不用做任何配置。
> 下面这份文档面向**开发者**（从源码运行、手动侧载）。

## 0. 前置条件

1. 安装 Node.js ≥ 18.17
2. `npm install`
3. `npm run https` 启动本地服务（**必须是 HTTPS**，Office 不允许加载 http 的任务窗格）
   - 首次运行会调用 `office-addin-dev-certs` 生成并信任本地证书（Windows 上会写入用户根证书存储）
   - 如果证书安装失败：`npx office-addin-dev-certs install`

确认服务可用：

```bash
curl -k https://localhost:3000/api/health
# {"ok":true,"name":"office-pdf-tools","version":"1.0.0",...}
```

## 1. Windows 桌面版（Word / Excel / PowerPoint）

### 方式 A：一键脚本（推荐）

```powershell
npm run sideload
```

脚本会检查清单端口与服务连通性，然后写入开发者侧载注册表项：

```
HKCU\Software\Microsoft\Office\16.0\Wef\Developer
  名称 = <manifest.xml 的绝对路径>
  数据 = <同上>
```

**完全退出** Office（检查任务管理器中无 WINWORD/EXCEL/POWERPNT 残留）后重新打开即可。

取消侧载：

```powershell
npm run unsideload
```

### 方式 B：共享文件夹 + 受信任目录（不使用注册表）

```powershell
npm run sideload -- -SharedFolder
```

然后在 Office 中：

1. 文件 → 选项 → 信任中心 → 信任中心设置 → **受信任的加载项目录**
2. 添加 `%USERPROFILE%\OfficeAddinShare`，勾选“显示在菜单中”
3. 重启 Office → 插入 → 我的加载项 → **共享文件夹** → 选择“Office PDF 工具”

### 方式 C：手动复制

把 `manifest/manifest.xml` 复制到：

```
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

重启 Office 后，在 插入 → 我的加载项 → 上方下拉选择“文件夹”即可看到。

## 2. macOS 桌面版

```bash
# 先确认本地服务已在 https://localhost:3000 运行
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
cp manifest/manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
# Excel / PowerPoint 同理：
#   ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
#   ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
```

若使用 LibreOffice 作为转换引擎，可安装：`brew install --cask libreoffice`。

## 3. Office 网页版

在浏览器打开 Word/Excel/PowerPoint 网页版 → 插入 → 加载项 → 上传我的加载项 → 选择 `manifest/manifest.xml`
（网页版只能访问 HTTPS 地址，本地 `https://localhost:3000` 需要浏览器信任该证书）。

## 4. 验证安装

1. 打开任一文档，功能区应出现 **“PDF 工具”** 选项卡
2. 点击「导出 PDF」→ 右侧任务窗格打开，顶部显示宿主名称（Word/Excel/PowerPoint）与绿色服务指示灯
3. 若指示灯为红色：本地服务未启动或端口不一致（见 [TROUBLESHOOTING.md](TROUBLESHOOTING.md)）

## 5. 修改端口

默认端口 3000。修改时**清单和脚本要一起改**，项目提供了同步脚本：

```bash
npm run set:port -- 3001
npm run https -- --port 3001
npm run sideload
```

## 6. 生产/团队分发提示

- 把 `manifest.xml` 里的 `https://localhost:3000` 换成真实的 HTTPS 域名（需在 `AppDomains` 中声明），即可用 **集中部署**（Microsoft 365 管理中心 → 设置 → 集成应用 → 上传自定义应用）分发给整个组织。
- 服务端可部署到任意 Node 环境；跨平台转换建议安装 LibreOffice，Windows 上也可直接使用已装的 Office。

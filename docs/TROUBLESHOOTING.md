# 常见问题排查

## 任务窗格相关

### 功能区没有“PDF 工具”选项卡

1. 确认加载项已侧载：`npm run sideload`（或见 [INSTALL.md](INSTALL.md) 的其它方式）
2. **完全退出** Office：任务管理器里不能有 WINWORD/EXCEL/POWERPNT 残留进程
3. 文件 → 选项 → 加载项 → 下方“管理”选 **COM 加载项**/“我的加载项”，确认没有被禁用
4. Excel 中若提示“此加载项不再可用”，检查 信任中心 → 受信任的加载项目录 是否允许加载

### 任务窗格一片空白 / 一直转圈

- 任务窗格必须走 HTTPS：确认用 `npm run https` 启动，而不是 `npm start`
- 打开 `https://localhost:3000/taskpane.html` 看是否报证书错误；浏览器提示不安全说明开发证书未受信任：
  ```bash
  npx office-addin-dev-certs install
  ```
- 端口被改过：`npm run set:port -- 3000` 让清单与脚本回到同一端口，然后重新侧载
- Office 桌面端缓存了旧页面：在窗格内右键 → 重新加载，或清空 `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`

### 右上角指示灯是红色

说明任务窗格连不上本地服务：

```bash
curl -k https://localhost:3000/api/health     # 应返回 {"ok":true,...}
```

- 服务没启动 → `npm run https`
- 端口不一致 → 设置页里改“服务地址”，或 `npm run set:port`
- 公司代理拦截 localhost → 在代理例外中加入 `localhost;127.0.0.1`

## 转换相关

### 提示“没有可用的 Office → PDF 转换引擎”

服务启动横幅里的 `engines` 会显示探测结果：

- **Windows**：安装 Microsoft Office 桌面版即可自动启用 COM 引擎（Word/Excel/PowerPoint 任一存在即可处理对应格式）
- **任意平台**：安装 LibreOffice；若装在非默认路径，设置环境变量后重启服务：
  ```bash
  set SOFFICE_PATH=C:\Program Files\LibreOffice\program\soffice.exe
  ```
  自检：`curl -k https://localhost:3000/api/engines`

### Microsoft Office 转换失败，提示 CO_E_SERVER_EXEC_FAILURE / 80080005

通常是**有卡死的 Office 进程**占住了自动化接口（PowerPoint 尤其常见，它只允许单实例）：

1. 任务管理器结束残留的 WINWORD.EXE / EXCEL.EXE / POWERPNT.EXE
2. 重试；服务在超时后也会按 PID 精确清理自己创建的实例
3. Word 首次运行弹出激活/登录对话框也会导致失败，手动打开一次 Word 完成初始化再试

### PDF → Word 结果版式“跑版”

文本模式是**版面重建**而非逆向还原：

- 需要 100% 一致的外观 → 选 **图片（高保真）模式**
- 扫描件（没有文本层）→ 必须用图片模式，文本模式只能得到空白
- 复杂多栏、文本框、公式、艺术字不保证还原

### PDF → Excel 只有一列 / 多列被合并

列是按文字横坐标聚类得到的：

- 表格线清晰、列对齐的 PDF 效果最好
- 纯图片型表格（扫描件）无法抽取，请改用图片模式或先 OCR

### Word 合并后脚注、批注不见了

已知限制：合并只迁移正文、表格、图片、样式与编号定义；脚注/尾注/批注/修订不迁移（服务会在 `warnings` 中提醒）。

### Excel 合并后图表/透视表消失

ExcelJS 能力边界：单元格值、公式、常用样式、合并单元格、列宽、冻结窗格会保留，图表/图片/透视表/宏不保留。

### PowerPoint 合并后文件变大

为了不让新幻灯片引用别的文件里的母版（那样 PowerPoint 会判定文件损坏），合并会把源演示文稿的版式与母版一并复制，因此体积增加属于正常现象；备注页不迁移。

### 超大文件或超时

- 默认上传上限 300 MB、单次转换超时 10 分钟，可用环境变量调整：
  ```bash
  set MAX_UPLOAD_MB=800
  set CONVERT_TIMEOUT_MS=1800000
  ```
- 上千页 PDF 转图片建议分批（页码范围）并降低 DPI

## 其它

### 输出文件在哪里？

默认 `<项目目录>/output`，可用 `npm run https -- --output D:\pdf-out` 或“设置”页修改；结果面板里也有「打开所在文件夹」。

### 端口被占用

```bash
netstat -ano | findstr :3000
npm run set:port -- 3001
npm run https -- --port 3001
```

### 想连非本机的服务

```bash
npm run https -- --host 0.0.0.0 --port 3000
```
然后在任务窗格“设置”里把服务地址改成 `https://<局域网IP>:3000`。注意：此时 API 会暴露在局域网中，请自行加防火墙/反向代理与鉴权。

### 日志

服务端日志级别：`LOG_LEVEL=debug`（显示每个 API 调用、LibreOffice 输出等）。

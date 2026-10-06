# 架构说明

## 分层

```
addin/            纯前端（无打包步骤）
  taskpane.html   四个面板的静态结构，用 data-i18n 标注文案
  js/i18n.js      中英词典 + t()/applyI18n()
  js/api.js       本地服务客户端：XHR 上传进度、任务轮询、下载
  js/office-bridge.js  Office.js 封装：宿主识别 / getFileAsync / insertFileFromBase64 / insertSlidesFromBase64
  js/components.js    文件选择器（拖拽+排序+缩略图）、进度条、结果面板、toast
  js/jobflow.js       统一任务流：上传 → 轮询 → 渲染结果 → 错误提示
  js/panels/*.js      各面板业务逻辑

server/           本地服务（Express 5）
  index.js        入口：HTTP/HTTPS、启动横幅、引擎探测、定时清理、优雅退出
  app.js          中间件：CORS 白名单、CSRF 头校验、静态资源、统一错误处理
  routes/api.js   REST 路由；routes/upload.js 负责 multer 磁盘上传
  lib/jobs.js     内存任务表：进度、结果、取消（AbortController）、TTL 清理
  lib/runner.js   任务执行器：把 op 分派给具体实现
  lib/engines/    libreoffice.js / msoffice.js / render.js / detect.js
  lib/convert/    officeToPdf · pdfToOffice · pdfToImages · imagesToPdf
  lib/pdf/        layout.js（PDF 版面重建）
  lib/ooxml/      package.js（zip/rels/ContentTypes/XML 工具）· partCopy.js（关系深拷贝）· gc.js（孤立部件回收）
  lib/merge|split/  pdf / docx / xlsx / pptx 四种格式的合并与拆分
```

## 请求时序

```
任务窗格               本地服务
   │  POST /api/jobs (multipart)  │
   ├─────────────────────────────►│ multer 落盘 tmp/incoming/<id>/
   │                              │ jobStore.create() → 202 {jobId}
   │                              │ runJobInBackground(): executeJob()
   │  GET /api/jobs/:id (600ms)   │   ├─ engine.detect → 选引擎
   ├─────────────────────────────►│   ├─ 转换 → 写入 output/
   │  ◄── {status, progress, step}│   └─ finish(files[])
   │  GET .../result/0            │
   ├─────────────────────────────►│ sendFile(output/xxx)
   │  POST /api/system/open-folder│ explorer.exe <output>
   └─────────────────────────────►│
```

## 引擎选择

| 能力 | 首选 | 备选 | 说明 |
| --- | --- | --- | --- |
| 当前文档 → PDF | Office 原生 `getFileAsync(FileType.Pdf)` | 取 Compressed 原文件交服务端 | 原生导出保真且不占用服务，故不设开关 |
| Office → PDF | Windows：Microsoft Office COM | LibreOffice headless | `pickOfficeEngine()`；扩展名不支持 COM 时自动回退 LibreOffice |
| PDF 渲染/取字 | pdf.js（`@napi-rs/canvas` 绘制） | — | 同一渲染器服务 PDF→图片、PDF→Office(图片模式) 与缩略图 |
| PDF 读写 | pdf-lib | — | 合并、拆分、图片排版 |
| docx/pptx 合并拆分 | 自研 OOXML 引擎 | — | 关系深拷贝 + 可达性回收 |
| xlsx 合并拆分 | ExcelJS | — | 单元格/公式/样式级复制 |

LibreOffice 调用要点：`--headless`、独立 `-env:UserInstallation` 配置目录，并对同一进程内的转换做**串行化**（LibreOffice 用户配置目录不能并发共享）。
Microsoft Office 调用要点：PowerShell 脚本 `scripts/office-com.ps1`，`GetActiveObject` 优先复用已运行实例（PowerPoint 单实例限制），只在自己创建实例时才 `Quit`，并把新建进程的 PID 写入 pid 文件，超时后按 PID 精确清理。

## OOXML 引擎（合并 / 拆分 / 关系搬迁）

Office 文档本质是一个 zip 包 + 大量 `_rels` 关系。合并/拆分时最容易出问题的就是“部件复制了一半、rId 对不上”。

1. **`OoxmlPackage`**：zip 读写、`[Content_Types].xml` 解析/序列化、rels 解析/序列化、XML 顶层元素切分（`splitTopLevelElements`，正确处理自闭合与注释）。
2. **`copyPartDeep`**：递归复制一个部件及其全部传递关系（幻灯片→版式→母版→主题→媒体…），为目标包分配不冲突的新部件名并重编号 rId，最后按映射表改写 XML 中的 `r:id`/`r:embed`/`r:link` 等引用。带 `map` 去重，天然抗环。
3. **`garbageCollect`**：从 `_rels/.rels` 出发做可达性分析，删除不可达部件及其 rels 与内容类型声明——拆分后自动丢掉不用的图片/幻灯片，合并时也能顺带瘦身。
4. **Word 合并**：基准文档的 `w:body` 保留，源文档正文子节点（去掉 `w:sectPr`）追加进去；同时合并 `styles.xml`（补齐缺失样式）、`numbering.xml`（重编号 `w:abstractNumId`/`w:numId` 并改写正文引用，随机化 `w:nsid` 防冲突），图片等关系通过 `copyPartDeep` 搬迁。
5. **PowerPoint 合并**：按源 `p:sldIdLst` 顺序复制幻灯片，为其版式所属母版在目标 `p:sldMasterIdLst` 注册新条目（否则 PowerPoint 会认为版式无主），最后重写 `presentation.xml` 与 rels。
6. **Word 拆分**：按「标题 N 样式 / outlineLvl」或「段落内的 `w:sectPr`」切分 `w:body` 子节点，逐份重建 document.xml 并做回收（自动删掉本节不再引用的图片）。
7. **PowerPoint 拆分**：删掉不在范围内的 `p:sldId` 与对应 presentation 级关系，再交给回收器清理幻灯片及其专属媒体。
8. **Excel**：用 ExcelJS 做工作表/行的复制（值、公式、数字格式、字体、填充、边框、对齐、列宽、合并单元格、冻结窗格）。

## PDF 版面重建（PDF → Word/Excel 文本模式）

```
pdf.js getTextContent()            每段文字 + 变换矩阵 + 宽度
        │
        ├─ buildLines()             按 y 聚类成行；按 x 间距切单元格；空白项作为“有空格”提示（跨列的宽空白不参与）
        ├─ detectBlocks()           正文字号 = 字符数加权中位数；按字号/加粗/编号推断标题层级；
        │                           连续多行且列数一致 → 表格；行距与句末标点决定段落合并
        ├─ linesToMatrix()          按 x 起点聚类成列，输出二维数组（PDF → Excel）
        └─ 输出                      docx（标题样式 + 段落 + 表格）· xlsx（每页一个工作表，列宽自适应）· pptx（标题 + 正文文本框）
```

实现细节见 `server/lib/pdf/layout.js`：所有启发式都以“尽量少产生错误空格/错误分列”为优先，宁可保守。

## 可扩展点

- **新增 op**：在 `server/lib/runner.js` 的 `OPS` 与 `switch` 中加一个分支 + 一个实现模块，前端在面板里调用即可。
- **新增引擎**：实现 `{ outputPath }` 契约并在 `detect.js` 中注册，`pickOfficeEngine()` 里调整优先级。
- **新增语言**：在 `addin/js/i18n.js` 加一份词典（HTML 中的 `data-i18n` 会自动套用）。
- **打包分发**：把 `manifest.xml` 的 URL 换成正式域名 → 用管理中心的集中部署即可。

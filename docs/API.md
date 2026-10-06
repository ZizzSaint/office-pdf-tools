# HTTP API 参考

基础地址：`https://localhost:3000/api`（`npm start` 时是 http）。
除 GET 外，请求都必须带自定义头 `X-Office-Pdf-Tools: 1`（防 CSRF），跨域只允许 localhost 来源。

## 通用约定

### 异步任务（推荐）

```
POST /api/jobs                       multipart: files[], op, options(JSON 字符串)
  → 202 { "jobId": "e7674cca0q6e" }

GET  /api/jobs/:id                   → 任务状态
GET  /api/jobs/:id/result/:index     → 下载第 index 个结果文件（?inline=1 可直接预览）
GET  /api/jobs/:id/archive           → 把所有结果打包成 ZIP
POST /api/jobs/:id/cancel            → 取消
DELETE /api/jobs/:id                 → 删除任务与临时文件
GET  /api/jobs                       → 最近 50 个任务
```

任务对象：

```json
{
  "id": "e7674cca0q6e",
  "op": "office-to-pdf",
  "status": "done",            // queued | running | done | error | canceled
  "progress": 1,
  "step": "完成",
  "createdAt": 1760000000000,
  "finishedAt": 1760000040000,
  "warnings": ["..."],
  "error": null,
  "result": {
    "files": [{ "name": "report.pdf", "path": "D:\\out\\report.pdf", "size": 12345, "mime": "application/pdf", "downloadId": "ab12..." }],
    "outputDir": "D:\\out",
    "engines": ["msoffice"]
  }
}
```

### 同步形态

`POST /api/jobs?sync=1`：阻塞到转换完成，直接返回文件本体（单文件）或 ZIP（多文件）。
`POST /api/convert/:op`：同上，路径更短。

### 其它

| 接口 | 说明 |
| --- | --- |
| `GET /api/health` | 存活检测 |
| `GET /api/engines` | 引擎探测结果（LibreOffice / Microsoft Office / pdf.js）、上传上限、输出目录 |
| `GET /api/settings`、`POST /api/settings` | 读取/修改输出目录 |
| `POST /api/system/open-folder` `{path}` | 在资源管理器中打开（仅限输出/临时/项目目录） |
| `POST /api/system/open-file` `{path}` | 用默认程序打开结果文件 |
| `POST /api/files/save` | 保存上传的文件到输出目录（加载项“当前文档 → PDF”走这里） |
| `GET /api/files/:id/download` | 下载上面保存的文件 |
| `POST /api/pdf/info` | 返回 PDF 页数与每页尺寸 |

---

## op 与 options

### 1. `office-to-pdf`

把 Office/其它办公文档转成 PDF。

| 字段 | 类型 | 默认 | 说明 |
| --- | --- | --- | --- |
| `engine` | `auto\|msoffice\|libreoffice` | `auto` | Windows 优先 Microsoft Office COM，其它平台用 LibreOffice |
| `sheet` | string | — | 只导出指定工作表（仅 Microsoft Office 引擎） |
| `recalculate` | boolean | `true` | Excel 转换前重新计算（仅 Microsoft Office 引擎） |

支持扩展名：doc/docx/docm/dot/dotx/rtf/txt/odt、xls/xlsx/xlsm/xlsb/ods/csv、ppt/pptx/pptm/odp、html/xml/tsv 等（LibreOffice 引擎覆盖更广）。

### 2. `pdf-to-office`

| 字段 | 类型 | 默认 | 说明 |
| --- | --- | --- | --- |
| `target` | `docx\|xlsx\|pptx` | `docx` | 目标格式 |
| `mode` | `text\|raster` | `text` | text=可编辑重排；raster=逐页图片（高保真）。xlsx 不支持 raster，会自动降级并给出 warning |
| `dpi` | number | 150 | raster 模式的渲染分辨率 |
| `pages` | string | 全部 | 如 `1-3,5,8-` |
| `outputName` | string | 源文件名 | 输出主名 |

### 3. `pdf-to-images`

| 字段 | 类型 | 默认 | 说明 |
| --- | --- | --- | --- |
| `format` | `png\|jpeg\|webp` | `png` | 输出格式 |
| `dpi` | number | 150 | 36–1200 |
| `quality` | number | 82 | jpeg/webp 质量 |
| `pages` | string | 全部 | 页码范围 |

### 4. `images-to-pdf`

| 字段 | 类型 | 默认 | 说明 |
| --- | --- | --- | --- |
| `pageSize` | `fit\|a3\|a4\|a5\|letter\|legal\|tabloid` | `a4` | `fit` = 页面等于图片（按 96dpi 换算） |
| `orientation` | `auto\|portrait\|landscape` | `auto` | 自动按图片/页面比例选择 |
| `margin` | number(mm) | 10 | 页面边距 |
| `perPage` | `1\|2\|4\|6\|9` | 1 | 每页图片数（宫格排版） |
| `fit` | `contain\|cover\|stretch` | `contain` | cover 会按单元格裁剪 |
| `separate` | boolean | false | 每张图片生成单独的 PDF（返回多个文件） |
| `imageDpi` | number | 96 | `pageSize=fit` 时的像素→点换算基准 |
| `outputName` | string | 首图名 | 输出主名 |

### 5. `merge`

| 字段 | 类型 | 默认 | 说明 |
| --- | --- | --- | --- |
| `kind` | `auto\|pdf\|docx\|xlsx\|pptx` | `auto` | 自动按扩展名判断，混类型会报错 |
| `outputName` | string | `<首个文件>-merged` | 输出主名 |
| `pageBreak` | boolean | true | Word：文档之间插入分页符 |
| `sheetPrefix` | boolean | false | Excel：工作表名加上来源文件名前缀 |

### 6. `split`

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| `kind` | `auto\|pdf\|docx\|xlsx\|pptx` | 自动判断 |
| `strategy` | 见下表 | 拆分方式 |
| `param` | string | 策略参数 |

| kind | strategy | param |
| --- | --- | --- |
| pdf | `every-n-pages` | 每份页数，如 `2` |
| pdf | `ranges` | `1-3,4-6,7-` |
| pdf | `each-page` | — |
| docx | `heading` | 标题级别 1–6 |
| docx | `section` | — |
| xlsx | `sheets` | — |
| xlsx | `every-n-rows` | 每份数据行数（自动重复表头） |
| pptx | `every-n-slides` | 每份幻灯片数 |
| pptx | `ranges` | `1-5,6-10` |

---

## 示例

```bash
# 1) Office → PDF（同步，直接落盘）
curl -o report.pdf -F "files=@report.docx" -F "op=office-to-pdf" \
     -F 'options={"engine":"auto"}' "http://localhost:3000/api/jobs?sync=1"

# 2) PDF → Word（文本模式）
curl -o report.docx -F "files=@report.pdf" -F "op=pdf-to-office" \
     -F 'options={"target":"docx","mode":"text"}' "http://localhost:3000/api/jobs?sync=1"

# 3) 图片 → PDF（每页 4 张，A4，横向自动）
curl -o album.pdf -F "files=@1.png" -F "files=@2.png" -F "files=@3.png" -F "files=@4.png" \
     -F "op=images-to-pdf" -F 'options={"pageSize":"a4","perPage":4}' \
     "http://localhost:3000/api/jobs?sync=1"

# 4) 合并 PDF
curl -o merged.pdf -F "files=@a.pdf" -F "files=@b.pdf" -F "op=merge" \
     -F 'options={"kind":"pdf"}' "http://localhost:3000/api/jobs?sync=1"

# 5) Word 按标题 1 拆分（多文件 → ZIP）
curl -o parts.zip -F "files=@book.docx" -F "op=split" \
     -F 'options={"kind":"docx","strategy":"heading","param":"1"}' \
     "http://localhost:3000/api/jobs?sync=1"
```

## 错误格式

```json
{ "error": { "code": "engine-unavailable", "message": "没有可用的 Office → PDF 转换引擎…", "hint": "LibreOffice: 未找到 soffice…" } }
```

常见 code：`bad-request`、`unsupported-format`、`invalid-document`、`engine-unavailable`、`conversion-failed`、`timeout`、`not-found`、`forbidden`、`canceled`。

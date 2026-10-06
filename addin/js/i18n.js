/**
 * 轻量 i18n：扁平 key，中英双语。HTML 中用 data-i18n="key" 标注，JS 中用 t("key")。
 */
const DICT = {
  "zh-CN": {
    "app.title": "Office PDF 工具",
    "app.hostDetecting": "检测中…",
    "app.host.word": "Word",
    "app.host.excel": "Excel",
    "app.host.powerpoint": "PowerPoint",
    "app.host.browser": "浏览器预览模式",

    "tab.convert": "转换",
    "tab.images": "图片 / PDF",
    "tab.splitmerge": "拆分合并",
    "tab.settings": "设置",

    "common.advanced": "高级选项",
    "common.pages": "页码范围",
    "common.dpi": "分辨率 DPI",
    "common.auto": "自动识别",
    "common.loading": "加载中…",
    "common.remove": "移除",
    "common.clear": "清空",
    "common.moveUp": "上移",
    "common.moveDown": "下移",
    "common.unknown": "未知",

    "dropzone.hint": "点击选择文件，或把文件拖到这里",
    "dropzone.hintPdf": "点击选择 PDF，或拖拽到这里",
    "dropzone.hintImages": "点击选择图片，或拖拽到这里",
    "dropzone.hintMerge": "点击选择多个同类型文件，或拖拽到这里",
    "dropzone.hintSplit": "点击选择要拆分的文件，或拖拽到这里",

    "convert.current.title": "当前文档 → PDF",
    "convert.current.desc": "使用 Office 原生导出，保留字体与排版，无需上传。",
    "convert.current.action": "导出当前文档为 PDF",
    "convert.current.openAfter": "导出后自动打开 PDF",
    "convert.current.working": "正在从 Office 导出 PDF…",
    "convert.current.saved": "PDF 已生成",
    "convert.current.unsupported": "当前 Office 版本不支持原生导出 PDF，请改用下方“本地文件 → PDF”。",

    "convert.file.title": "本地文件 → PDF",
    "convert.file.desc": "支持 Word / Excel / PowerPoint / ODF / RTF / TXT / CSV / HTML 等，由服务端引擎批量转换。",
    "convert.file.action": "开始转换",
    "convert.engine": "转换引擎",
    "convert.engine.auto": "自动",
    "convert.engine.ms": "Microsoft Office",
    "convert.engine.lo": "LibreOffice",
    "convert.sheet": "仅导出工作表(Excel)",

    "convert.pdf.title": "PDF → Office / 图片",
    "convert.pdf.desc": "文本模式尽量还原可编辑内容；图片模式逐页渲染，版式 100% 保真。",
    "convert.pdf.target": "目标格式",
    "convert.pdf.mode": "转换模式",
    "convert.pdf.mode.text": "文本（可编辑）",
    "convert.pdf.mode.raster": "图片（高保真）",
    "convert.pdf.action": "开始转换",
    "fmt.docx": "Word (.docx)",
    "fmt.xlsx": "Excel (.xlsx)",
    "fmt.pptx": "PowerPoint (.pptx)",

    "images.toPdf.title": "图片 → PDF",
    "images.toPdf.desc": "支持 PNG / JPG / WebP / GIF / BMP / TIFF，可调整顺序、页面尺寸与每页张数。",
    "images.toPdf.action": "生成 PDF",
    "images.pageSize": "页面尺寸",
    "images.pageSize.fit": "适应图片",
    "images.orientation": "方向",
    "images.orientation.auto": "自动",
    "images.orientation.portrait": "纵向",
    "images.orientation.landscape": "横向",
    "images.margin": "边距 (mm)",
    "images.perPage": "每页张数",
    "images.fit": "适配方式",
    "images.fit.contain": "等比缩放(完整)",
    "images.fit.cover": "等比裁剪(铺满)",
    "images.fit.stretch": "拉伸",
    "images.separate": "每张图片生成单独的 PDF（打包 ZIP）",

    "pdf.toImages.title": "PDF → 图片",
    "pdf.toImages.desc": "按页渲染为图片，可指定页码范围与分辨率，多页自动打包 ZIP。",
    "pdf.toImages.format": "格式",
    "pdf.toImages.action": "导出图片",

    "merge.title": "合并文档",
    "merge.desc": "合并同类型文件：PDF 直接拼页；Word 追加正文；Excel 汇总工作表；PowerPoint 追加幻灯片。",
    "merge.kind": "文件类型",
    "merge.outputName": "输出文件名",
    "merge.pageBreak": "Word：每个文档之间插入分页符",
    "merge.action": "合并",
    "merge.needTwo": "至少选择两个文件才能合并。",
    "merge.mixed": "所选文件类型不一致，请选择同类型文件或手动指定类型。",
    "merge.unsupported": "暂不支持该类型的合并：",

    "split.title": "拆分文档",
    "split.desc": "PDF 按页；Word 按标题或分节符；Excel 按工作表或行数；PowerPoint 按幻灯片。",
    "split.strategy": "拆分方式",
    "split.param": "参数",
    "split.action": "拆分（ZIP 打包）",
    "split.strategy.pdf.everyN": "每 N 页一个文件",
    "split.strategy.pdf.ranges": "按页码范围拆分",
    "split.strategy.pdf.eachPage": "每页一个文件",
    "split.strategy.docx.heading": "按标题级别拆分",
    "split.strategy.docx.section": "按分节符拆分",
    "split.strategy.xlsx.sheets": "每个工作表一个文件",
    "split.strategy.xlsx.everyNRows": "每 N 行一个文件",
    "split.strategy.pptx.everyN": "每 N 张幻灯片一个文件",
    "split.strategy.pptx.ranges": "按幻灯片范围拆分",
    "split.hint.pdf.everyN": "参数：数字，例如 2 表示每 2 页生成一个 PDF。",
    "split.hint.pdf.ranges": "参数：页码范围，例如 1-3,4-6,7- 表示三段。",
    "split.hint.pdf.eachPage": "参数：无需填写。",
    "split.hint.docx.heading": "参数：标题级别 1-6，例如 1 表示按“标题 1”拆分。",
    "split.hint.docx.section": "参数：无需填写，按分节符拆分。",
    "split.hint.xlsx.sheets": "参数：无需填写，每个工作表另存为一个工作簿。",
    "split.hint.xlsx.everyNRows": "参数：数字，例如 100 表示每 100 行（含表头）生成一个文件。",
    "split.hint.pptx.everyN": "参数：数字，例如 10 表示每 10 张幻灯片一个 PPT。",
    "split.hint.pptx.ranges": "参数：幻灯片范围，例如 1-5,6-10。",

    "settings.service": "本地服务",
    "settings.address": "服务地址",
    "settings.outputDir": "输出目录",
    "settings.open": "打开",
    "settings.save": "保存",
    "settings.check": "检测连接",
    "settings.engines": "转换引擎",
    "settings.help": "帮助",
    "settings.helpText": "首次使用请在项目目录运行 npm run https 启动本地服务，然后用 npm run sideload 载入加载项。",
    "settings.docs": "打开使用说明",
    "settings.saved": "设置已保存",
    "settings.engineOk": "可用",
    "settings.engineNo": "不可用",
    "settings.reconnected": "连接正常",
    "settings.reconnectFail": "无法连接本地服务",

    "res.title": "转换完成",
    "res.savedTo": "已保存到：",
    "res.download": "下载",
    "res.downloadZip": "下载 ZIP",
    "res.openFolder": "打开所在文件夹",
    "res.open": "打开文件",
    "res.insert": "插入到当前文档",
    "res.inserted": "已插入到当前文档",
    "res.insertUnsupported": "当前 Office 版本不支持自动插入，请使用“打开文件”。",
    "res.files": "个文件",

    "msg.uploading": "上传中…",
    "msg.processing": "服务端处理中…",
    "msg.done": "完成",
    "msg.failed": "失败",
    "msg.cancel": "取消",
    "msg.canceled": "已取消",
    "msg.noFile": "请先选择文件。",
    "msg.needServer": "本地服务未连接，请先运行 npm run https。",
    "msg.queued": "排队中…",
    "msg.exporting": "正在导出…",
  },

  "en": {
    "app.title": "Office PDF Tools",
    "app.hostDetecting": "Detecting…",
    "app.host.word": "Word",
    "app.host.excel": "Excel",
    "app.host.powerpoint": "PowerPoint",
    "app.host.browser": "Browser preview mode",

    "tab.convert": "Convert",
    "tab.images": "Images / PDF",
    "tab.splitmerge": "Split & Merge",
    "tab.settings": "Settings",

    "common.advanced": "Advanced options",
    "common.pages": "Page range",
    "common.dpi": "DPI",
    "common.auto": "Auto",
    "common.loading": "Loading…",
    "common.remove": "Remove",
    "common.clear": "Clear",
    "common.moveUp": "Move up",
    "common.moveDown": "Move down",
    "common.unknown": "unknown",

    "dropzone.hint": "Click to choose files, or drop them here",
    "dropzone.hintPdf": "Click to choose a PDF, or drop it here",
    "dropzone.hintImages": "Click to choose images, or drop them here",
    "dropzone.hintMerge": "Click to choose several files of the same type",
    "dropzone.hintSplit": "Click to choose a file to split",

    "convert.current.title": "Current document → PDF",
    "convert.current.desc": "Native Office export keeps fonts and layout. No upload needed.",
    "convert.current.action": "Export current document to PDF",
    "convert.current.openAfter": "Open the PDF when finished",
    "convert.current.working": "Exporting PDF from Office…",
    "convert.current.saved": "PDF ready",
    "convert.current.unsupported": "This Office version cannot export PDF natively. Use “Local files → PDF” below instead.",

    "convert.file.title": "Local files → PDF",
    "convert.file.desc": "Word / Excel / PowerPoint / ODF / RTF / TXT / CSV / HTML, converted by the local service.",
    "convert.file.action": "Convert",
    "convert.engine": "Engine",
    "convert.engine.auto": "Auto",
    "convert.engine.ms": "Microsoft Office",
    "convert.engine.lo": "LibreOffice",
    "convert.sheet": "Single sheet (Excel)",

    "convert.pdf.title": "PDF → Office / images",
    "convert.pdf.desc": "Text mode keeps content editable; image mode renders pages pixel-perfect.",
    "convert.pdf.target": "Target format",
    "convert.pdf.mode": "Mode",
    "convert.pdf.mode.text": "Text (editable)",
    "convert.pdf.mode.raster": "Image (faithful)",
    "convert.pdf.action": "Convert",
    "fmt.docx": "Word (.docx)",
    "fmt.xlsx": "Excel (.xlsx)",
    "fmt.pptx": "PowerPoint (.pptx)",

    "images.toPdf.title": "Images → PDF",
    "images.toPdf.desc": "PNG / JPG / WebP / GIF / BMP / TIFF with ordering, page size and N-up layout.",
    "images.toPdf.action": "Create PDF",
    "images.pageSize": "Page size",
    "images.pageSize.fit": "Fit to image",
    "images.orientation": "Orientation",
    "images.orientation.auto": "Auto",
    "images.orientation.portrait": "Portrait",
    "images.orientation.landscape": "Landscape",
    "images.margin": "Margin (mm)",
    "images.perPage": "Images per page",
    "images.fit": "Fit mode",
    "images.fit.contain": "Contain",
    "images.fit.cover": "Cover (crop)",
    "images.fit.stretch": "Stretch",
    "images.separate": "One PDF per image (ZIP)",

    "pdf.toImages.title": "PDF → images",
    "pdf.toImages.desc": "Render pages to images with page range and DPI. Multiple pages are zipped.",
    "pdf.toImages.format": "Format",
    "pdf.toImages.action": "Export images",

    "merge.title": "Merge documents",
    "merge.desc": "PDF: concatenate pages. Word: append bodies. Excel: gather worksheets. PowerPoint: append slides.",
    "merge.kind": "File type",
    "merge.outputName": "Output name",
    "merge.pageBreak": "Word: insert page break between documents",
    "merge.action": "Merge",
    "merge.needTwo": "Select at least two files to merge.",
    "merge.mixed": "Selected files are of different types. Pick same-type files or set the type manually.",
    "merge.unsupported": "Merging is not supported for: ",

    "split.title": "Split document",
    "split.desc": "PDF by pages; Word by heading or section break; Excel by sheet or rows; PowerPoint by slides.",
    "split.strategy": "Strategy",
    "split.param": "Parameter",
    "split.action": "Split (ZIP)",
    "split.strategy.pdf.everyN": "Every N pages",
    "split.strategy.pdf.ranges": "Page ranges",
    "split.strategy.pdf.eachPage": "One file per page",
    "split.strategy.docx.heading": "By heading level",
    "split.strategy.docx.section": "By section break",
    "split.strategy.xlsx.sheets": "One file per worksheet",
    "split.strategy.xlsx.everyNRows": "Every N rows",
    "split.strategy.pptx.everyN": "Every N slides",
    "split.strategy.pptx.ranges": "Slide ranges",
    "split.hint.pdf.everyN": "Parameter: a number, e.g. 2 = one PDF per 2 pages.",
    "split.hint.pdf.ranges": "Parameter: ranges, e.g. 1-3,4-6,7-",
    "split.hint.pdf.eachPage": "Parameter: not required.",
    "split.hint.docx.heading": "Parameter: heading level 1-6, e.g. 1 = split on Heading 1.",
    "split.hint.docx.section": "Parameter: not required.",
    "split.hint.xlsx.sheets": "Parameter: not required.",
    "split.hint.xlsx.everyNRows": "Parameter: rows per file, e.g. 100.",
    "split.hint.pptx.everyN": "Parameter: slides per file, e.g. 10.",
    "split.hint.pptx.ranges": "Parameter: slide ranges, e.g. 1-5,6-10.",

    "settings.service": "Local service",
    "settings.address": "Service URL",
    "settings.outputDir": "Output folder",
    "settings.open": "Open",
    "settings.save": "Save",
    "settings.check": "Test connection",
    "settings.engines": "Conversion engines",
    "settings.help": "Help",
    "settings.helpText": "Run npm run https in the project folder to start the service, then npm run sideload to load the add-in.",
    "settings.docs": "Open the guide",
    "settings.saved": "Settings saved",
    "settings.engineOk": "available",
    "settings.engineNo": "not available",
    "settings.reconnected": "Connected",
    "settings.reconnectFail": "Cannot reach the local service",

    "res.title": "Finished",
    "res.savedTo": "Saved to: ",
    "res.download": "Download",
    "res.downloadZip": "Download ZIP",
    "res.openFolder": "Open folder",
    "res.open": "Open file",
    "res.insert": "Insert into document",
    "res.inserted": "Inserted into the current document",
    "res.insertUnsupported": "This Office version cannot insert automatically. Use “Open file”.",
    "res.files": " file(s)",

    "msg.uploading": "Uploading…",
    "msg.processing": "Processing on the server…",
    "msg.done": "Done",
    "msg.failed": "Failed",
    "msg.cancel": "Cancel",
    "msg.canceled": "Canceled",
    "msg.noFile": "Choose a file first.",
    "msg.needServer": "Local service is offline. Start it with npm run https.",
    "msg.queued": "Queued…",
    "msg.exporting": "Exporting…",
  },
};

let current = "zh-CN";

export function detectLang() {
  const saved = localStorage.getItem("opt.lang");
  if (saved && DICT[saved]) return saved;
  const nav = (navigator.language || "zh-CN").toLowerCase();
  return nav.startsWith("zh") ? "zh-CN" : "en";
}

export function setLang(lang) {
  current = DICT[lang] ? lang : "zh-CN";
  localStorage.setItem("opt.lang", current);
  document.documentElement.lang = current;
  applyI18n();
}

export function getLang() {
  return current;
}

/** 供测试使用：列出所有已定义的 key。 */
export function i18nKeys(lang = "zh-CN") {
  return Object.keys(DICT[lang] || {});
}

export function hasI18nKey(key, lang = "zh-CN") {
  return Object.prototype.hasOwnProperty.call(DICT[lang] || {}, key);
}

export function t(key, ...args) {
  const s = DICT[current]?.[key] ?? DICT["zh-CN"][key] ?? key;
  return args.length ? s.replace(/\{(\d+)\}/g, (_, i) => String(args[Number(i)] ?? "")) : s;
}

export function applyI18n(root = document) {
  root.querySelectorAll("[data-i18n]").forEach((el) => {
    el.textContent = t(el.dataset.i18n);
  });
  root.querySelectorAll("[data-i18n-placeholder]").forEach((el) => {
    el.placeholder = t(el.dataset.i18nPlaceholder);
  });
  root.querySelectorAll("[data-i18n-title]").forEach((el) => {
    el.title = t(el.dataset.i18nTitle);
  });
}

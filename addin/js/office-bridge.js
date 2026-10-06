/**
 * Office.js 桥接层：宿主识别、当前文档导出、结果回插。
 * 在普通浏览器里打开 taskpane.html 时自动进入“预览模式”，服务端功能依旧可用。
 */

const state = {
  ready: false,
  host: "Browser", // Word | Excel | PowerPoint | Browser
  platform: "unknown",
  documentName: "document",
};

export function officeState() {
  return state;
}

export function hasOffice() {
  return typeof Office !== "undefined" && !!Office.context;
}

export function isSetSupported(name, version) {
  try {
    return !!(Office?.context?.requirements?.isSetSupported?.(name, version));
  } catch {
    return false;
  }
}

/** 等待 Office.js 就绪（浏览器中直接返回预览模式）。 */
export function initOffice() {
  if (state.ready) return Promise.resolve(state);
  return new Promise((resolve) => {
    if (!hasOffice()) {
      state.ready = true;
      resolve(state);
      return;
    }
    const finish = (info) => {
      state.host = info?.host || "Browser";
      state.platform = info?.platform || "unknown";
      state.documentName = guessDocumentName();
      state.ready = true;
      resolve(state);
    };
    const timer = setTimeout(() => {
      if (!state.ready) {
        state.host = "Browser";
        state.ready = true;
        resolve(state);
      }
    }, 8000);
    try {
      Office.onReady((info) => {
        clearTimeout(timer);
        finish(info);
      });
    } catch {
      clearTimeout(timer);
      finish(null);
    }
  });
}

export function hostKey() {
  switch (state.host) {
    case "Word": return "word";
    case "Excel": return "excel";
    case "PowerPoint": return "powerpoint";
    default: return "browser";
  }
}

export function guessDocumentName() {
  try {
    const url = Office?.context?.document?.url || "";
    if (url) {
      const base = decodeURIComponent(url.split(/[\\/]/).pop() || "");
      if (base) return base.replace(/\.[^.]+$/, "");
    }
  } catch { /* ignore */ }
  return "document";
}

/** 用 getFileAsync 把当前文档拉成 Blob。 */
function getCurrentFileBlob(fileType, sliceSize = 4 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(fileType, { sliceSize }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(result.error?.message || "getFileAsync failed"));
        return;
      }
      const file = result.value;
      const slices = [];
      const readSlice = (i) =>
        new Promise((res, rej) => {
          file.getSliceAsync(i, (r) => {
            if (r.status === Office.AsyncResultStatus.Succeeded) res(r.value.data);
            else rej(new Error(r.error?.message || "getSliceAsync failed"));
          });
        });
      (async () => {
        try {
          for (let i = 0; i < file.sliceCount; i++) slices.push(await readSlice(i));
          resolve(new Blob(slices, { type: "application/octet-stream" }));
        } catch (err) {
          reject(err);
        } finally {
          file.closeAsync(() => { /* noop */ });
        }
      })();
    });
  });
}

/**
 * 导出当前文档为 PDF。
 * 优先使用 Office 原生 PDF 导出；不支持时退回“取原始文件 + 服务端转换”。
 * @returns {Promise<{blob: Blob, native: boolean, fallbackName: string}>}
 */
export async function exportCurrentToPdf() {
  if (state.host === "Browser") {
    throw new Error("preview-mode");
  }
  const name = guessDocumentName();
  try {
    const blob = await getCurrentFileBlob(Office.FileType.Pdf);
    if (blob.size > 0) return { blob, native: true, fallbackName: name + ".pdf" };
  } catch { /* fall through to server-side conversion */ }
  const ext = state.host === "Word" ? ".docx" : state.host === "Excel" ? ".xlsx" : ".pptx";
  const blob = await getCurrentFileBlob(Office.FileType.Compressed);
  return { blob, native: false, fallbackName: name + ext };
}

export function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const s = String(reader.result);
      resolve(s.slice(s.indexOf(",") + 1));
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

/** 把转换结果插回当前文档（能力允许时）。 */
export async function insertIntoCurrentDocument(blob, filename = "") {
  if (state.host === "Browser") return { ok: false, reason: "preview" };
  const base64 = await blobToBase64(blob);
  if (state.host === "Word") {
    if (!isSetSupported("WordApi", "1.5")) return { ok: false, reason: "requirement" };
    return Word.run(async (context) => {
      context.document.body.insertFileFromBase64(base64, "End");
      await context.sync();
      return { ok: true };
    });
  }
  if (state.host === "PowerPoint") {
    if (!isSetSupported("PowerPointApi", "1.2")) return { ok: false, reason: "requirement" };
    return PowerPoint.run(async (context) => {
      context.presentation.insertSlidesFromBase64(base64, { formatting: "UseDestinationTheme" });
      await context.sync();
      return { ok: true };
    });
  }
  if (state.host === "Excel") {
    // Excel.js 不支持插入完整工作簿；这里只提示用户打开文件。
    return { ok: false, reason: "unsupported-host" };
  }
  return { ok: false, reason: "unsupported-host" };
}

/** 打开链接（桌面端用系统浏览器）。 */
export function openExternal(url) {
  try {
    if (Office?.context?.ui?.openBrowserWindow) {
      Office.context.ui.openBrowserWindow(url);
      return true;
    }
  } catch { /* ignore */ }
  window.open(url, "_blank", "noopener");
  return true;
}

/** 重新载入宿主文档（插入内容后可选）。 */
export function hostSupportsNativePdf() {
  return state.host !== "Browser";
}

export { getCurrentFileBlob };

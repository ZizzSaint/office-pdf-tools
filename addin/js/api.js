/**
 * 本地服务 API 客户端。
 * 默认与服务同源（taskpane.html 由服务托管），也可在设置里改成其它地址。
 */
import { t } from "./i18n.js";

const STORAGE_KEY = "opt.serverUrl";
const CLIENT_HEADER = "X-Office-Pdf-Tools";

export function defaultBaseUrl() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) return saved.replace(/\/+$/, "");
  if (location.protocol === "http:" || location.protocol === "https:") {
    if (location.port && !["5500", "8080"].includes(location.port)) return location.origin;
    return location.origin;
  }
  return "https://localhost:3000";
}

let baseUrl = defaultBaseUrl();

export function getBaseUrl() {
  return baseUrl;
}

export function setBaseUrl(url) {
  baseUrl = String(url || "").trim().replace(/\/+$/, "");
  localStorage.setItem(STORAGE_KEY, baseUrl);
}

async function request(path, { method = "GET", body, headers = {}, timeout = 30000, raw = false } = {}) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeout);
  try {
    const res = await fetch(baseUrl + path, {
      method,
      headers: {
        [CLIENT_HEADER]: "1",
        ...(body && !(body instanceof FormData) ? { "Content-Type": "application/json" } : {}),
        ...headers,
      },
      body: body instanceof FormData ? body : body ? JSON.stringify(body) : undefined,
      signal: controller.signal,
    });
    if (raw) return res;
    const text = await res.text();
    let data = null;
    try {
      data = text ? JSON.parse(text) : null;
    } catch {
      data = { raw: text };
    }
    if (!res.ok) {
      const message = data?.error?.message || data?.message || `HTTP ${res.status}`;
      const err = new Error(message);
      err.status = res.status;
      err.detail = data?.error;
      throw err;
    }
    return data;
  } finally {
    clearTimeout(timer);
  }
}

export const api = {
  health: () => request("/api/health", { timeout: 4000 }),
  engines: () => request("/api/engines", { timeout: 8000 }),
  settings: () => request("/api/settings", { timeout: 6000 }),
  saveSettings: (patch) => request("/api/settings", { method: "POST", body: patch }),
  openFolder: (path) => request("/api/system/open-folder", { method: "POST", body: { path } }),
  openFile: (path) => request("/api/system/open-file", { method: "POST", body: { path } }),
  job: (id) => request(`/api/jobs/${id}`, { timeout: 8000 }),
  cancelJob: (id) => request(`/api/jobs/${id}/cancel`, { method: "POST" }),
  deleteJob: (id) => request(`/api/jobs/${id}`, { method: "DELETE" }),
};

/**
 * 上传文件并创建一个转换任务。
 * @returns {Promise<{jobId:string}>}
 */
export function startJob(op, files, options = {}, { onProgress } = {}) {
  const form = new FormData();
  for (const f of files) form.append("files", f, f.name);
  form.append("op", op);
  form.append("options", JSON.stringify(options));

  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open("POST", `${baseUrl}/api/jobs`);
    xhr.setRequestHeader(CLIENT_HEADER, "1");
    xhr.upload.onprogress = (e) => {
      if (e.lengthComputable && onProgress) onProgress(e.loaded / e.total);
    };
    xhr.onload = () => {
      let data = null;
      try {
        data = JSON.parse(xhr.responseText);
      } catch {
        data = null;
      }
      if (xhr.status >= 200 && xhr.status < 300 && data?.jobId) resolve(data);
      else reject(new Error(data?.error?.message || `HTTP ${xhr.status}`));
    };
    xhr.onerror = () => reject(new Error(t("msg.needServer")));
    xhr.ontimeout = () => reject(new Error("timeout"));
    xhr.send(form);
  });
}

/** 轮询任务直到结束。 */
export async function waitJob(jobId, { onUpdate, intervalMs = 600 } = {}) {
  for (;;) {
    const job = await api.job(jobId);
    onUpdate?.(job);
    if (["done", "error", "canceled"].includes(job.status)) return job;
    await new Promise((r) => setTimeout(r, intervalMs));
  }
}

export function fileUrl(jobId, index = 0) {
  return `${baseUrl}/api/jobs/${jobId}/result/${index}`;
}

/** /api/files/save 保存后的下载地址。 */
export function savedFileUrl(downloadId) {
  return `${baseUrl}/api/files/${downloadId}/download`;
}

export function zipUrl(jobId) {
  return `${baseUrl}/api/jobs/${jobId}/archive`;
}

/** 把 Office 导出的 Blob 交给本地服务保存到输出目录。 */
export async function saveBlob(blob, filename, extra = {}) {
  const form = new FormData();
  form.append("files", blob, filename);
  form.append("options", JSON.stringify({ filename, ...extra }));
  const res = await fetch(`${baseUrl}/api/files/save`, {
    method: "POST",
    headers: { [CLIENT_HEADER]: "1" },
    body: form,
  });
  const data = await res.json().catch(() => null);
  if (!res.ok) throw new Error(data?.error?.message || `HTTP ${res.status}`);
  return data;
}

export function triggerDownload(url, filename) {
  const a = document.createElement("a");
  a.href = url;
  a.download = filename || "";
  a.target = "_blank";
  a.rel = "noopener";
  document.body.appendChild(a);
  a.click();
  a.remove();
}

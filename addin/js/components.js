/** 通用 UI 组件：文件选择、可排序列表、进度条、结果面板、提示条。 */
import { t } from "./i18n.js";
import { fileUrl, zipUrl, triggerDownload, api } from "./api.js";
import { insertIntoCurrentDocument, officeState } from "./office-bridge.js";

export const $ = (sel, root = document) => root.querySelector(sel);
export const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

export function el(tag, attrs = {}, children = []) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "class") node.className = v;
    else if (k === "text") node.textContent = v;
    else if (k === "html") node.innerHTML = v;
    else if (k.startsWith("on") && typeof v === "function") node.addEventListener(k.slice(2), v);
    else if (v !== undefined && v !== null) node.setAttribute(k, v);
  }
  for (const c of [].concat(children)) if (c) node.appendChild(c);
  return node;
}

export function fmtSize(bytes) {
  if (!Number.isFinite(bytes)) return "";
  const units = ["B", "KB", "MB", "GB"];
  let i = 0;
  let n = bytes;
  while (n >= 1024 && i < units.length - 1) { n /= 1024; i++; }
  return `${n < 10 && i > 0 ? n.toFixed(1) : Math.round(n)} ${units[i]}`;
}

/* ------------------------------------------------------------------ toast */
export function toast(message, type = "info", ms = 3600) {
  const wrap = $("#toastWrap");
  if (!wrap) return;
  const node = el("div", { class: `toast ${type}`, text: message });
  wrap.appendChild(node);
  setTimeout(() => {
    node.style.opacity = "0";
    node.style.transition = "opacity .3s";
    setTimeout(() => node.remove(), 320);
  }, ms);
  return node;
}

/* -------------------------------------------------------------- file list */
/**
 * 可拖拽/可排序的文件选择器。
 * @param {{dropzone:string, input:string, list:string, multiple?:boolean, thumbnails?:boolean, sortable?:boolean, accept?:string}} cfg
 */
export function createFilePicker(cfg) {
  const dz = $(cfg.dropzone);
  const input = $(cfg.input);
  const list = $(cfg.list);
  const multiple = cfg.multiple !== false;
  const thumbnails = !!cfg.thumbnails;
  const sortable = !!cfg.sortable;
  /** @type {File[]} */
  let files = [];
  const listeners = [];

  function notify() {
    render();
    listeners.forEach((fn) => fn(files.slice()));
  }

  function render() {
    list.textContent = "";
    files.forEach((file, index) => {
      const item = el("div", { class: "file-item", draggable: sortable ? "true" : "false" });
      item.dataset.index = String(index);
      if (thumbnails && file.type?.startsWith("image/")) {
        const img = el("img", { class: "thumb", alt: "" });
        img.src = URL.createObjectURL(file);
        img.onload = () => URL.revokeObjectURL(img.src);
        item.appendChild(img);
      }
      item.appendChild(el("span", { class: "name", text: file.name, title: file.name }));
      item.appendChild(el("span", { class: "size", text: fmtSize(file.size) }));
      if (sortable) {
        item.appendChild(el("button", { class: "btn btn-small btn-ghost", text: "↑", title: t("common.moveUp"), onclick: (e) => { e.stopPropagation(); move(index, -1); } }));
        item.appendChild(el("button", { class: "btn btn-small btn-ghost", text: "↓", title: t("common.moveDown"), onclick: (e) => { e.stopPropagation(); move(index, 1); } }));
      }
      item.appendChild(el("button", { class: "rm", text: "✕", title: t("common.remove"), onclick: (e) => { e.stopPropagation(); remove(index); } }));

      if (sortable) {
        item.addEventListener("dragstart", (e) => {
          item.classList.add("dragging");
          e.dataTransfer.setData("text/plain", String(index));
          e.dataTransfer.effectAllowed = "move";
        });
        item.addEventListener("dragend", () => item.classList.remove("dragging"));
        item.addEventListener("dragover", (e) => { e.preventDefault(); e.dataTransfer.dropEffect = "move"; });
        item.addEventListener("drop", (e) => {
          e.preventDefault();
          const from = Number(e.dataTransfer.getData("text/plain"));
          if (Number.isInteger(from) && from !== index) {
            const [moved] = files.splice(from, 1);
            files.splice(index, 0, moved);
            notify();
          }
        });
      }
      list.appendChild(item);
    });
  }

  function move(index, delta) {
    const target = index + delta;
    if (target < 0 || target >= files.length) return;
    const [moved] = files.splice(index, 1);
    files.splice(target, 0, moved);
    notify();
  }

  function remove(index) {
    files.splice(index, 1);
    notify();
  }

  function add(newFiles, { append = true } = {}) {
    let incoming = Array.from(newFiles || []);
    if (cfg.accept) {
      const patterns = cfg.accept.split(",").map((s) => s.trim().toLowerCase()).filter(Boolean);
      incoming = incoming.filter((f) => {
        const name = f.name.toLowerCase();
        return patterns.some((p) => (p.startsWith(".") ? name.endsWith(p) : f.type.toLowerCase().startsWith(p.replace("/*", "/")) || name.includes(p)));
      });
    }
    files = append ? files.concat(incoming) : incoming;
    notify();
  }

  if (dz) {
    dz.addEventListener("click", () => input?.click());
    dz.addEventListener("dragover", (e) => { e.preventDefault(); dz.classList.add("hover"); });
    dz.addEventListener("dragleave", () => dz.classList.remove("hover"));
    dz.addEventListener("drop", (e) => {
      e.preventDefault();
      dz.classList.remove("hover");
      add(e.dataTransfer?.files);
    });
  }
  if (input) {
    input.multiple = multiple;
    if (cfg.accept) input.accept = cfg.accept;
    input.addEventListener("change", () => {
      add(input.files, { append: multiple ? true : false });
      input.value = "";
    });
  }

  return {
    get files() { return files.slice(); },
    add: (f) => add(f),
    set: (f) => add(f, { append: false }),
    clear: () => { files = []; notify(); },
    onChange: (fn) => listeners.push(fn),
  };
}

/* --------------------------------------------------------------- progress */
export function progressFor(root) {
  const bar = $(".bar", root);
  const text = $(".progress-text", root);
  return {
    show(label = "") {
      root.classList.remove("hidden");
      this.set(0, label);
    },
    set(ratio, label) {
      const pct = Math.max(0, Math.min(1, ratio || 0));
      bar?.style.setProperty("--p", String(pct));
      if (text) text.textContent = `${Math.round(pct * 100)}%  ${label || ""}`;
    },
    hide() { root.classList.add("hidden"); },
  };
}

/* ----------------------------------------------------------------- result */
/**
 * 渲染任务结果：文件列表 + 操作按钮。
 * @param {HTMLElement} root
 * @param {{job:any, allowInsert?:boolean, gallery?:HTMLElement}} opts
 */
export function renderResult(root, { job, allowInsert = false, gallery = null } = {}) {
  root.textContent = "";
  root.classList.remove("hidden", "err", "ok");
  const files = job.result?.files || [];
  const hasErr = job.status === "error";
  root.classList.add(hasErr ? "err" : "ok");

  if (hasErr) {
    root.appendChild(el("div", { class: "result-title", text: `✗ ${t("msg.failed")}` }));
    root.appendChild(el("div", { class: "muted", text: job.error?.message || "" }));
    if (job.error?.hint) root.appendChild(el("div", { class: "muted small", text: job.error.hint }));
    return;
  }

  root.appendChild(el("div", { class: "result-title", text: `✓ ${t("res.title")} — ${files.length}${t("res.files")}` }));
  if (job.result?.outputDir) {
    root.appendChild(el("div", { class: "muted small", text: `${t("res.savedTo")}${job.result.outputDir}` }));
  }

  files.forEach((f, i) => {
    const line = el("div", { class: "result-file" });
    line.appendChild(el("span", { class: "name", text: f.name, title: f.path }));
    line.appendChild(el("span", { class: "size muted", text: fmtSize(f.size) }));
    line.appendChild(el("button", {
      class: "btn btn-small",
      text: t("res.download"),
      onclick: () => triggerDownload(fileUrl(job.id, i), f.name),
    }));
    if (f.path) {
      line.appendChild(el("button", {
        class: "btn btn-small btn-ghost",
        text: t("res.open"),
        onclick: () => api.openFile(f.path).catch((e) => toast(e.message, "err")),
      }));
    }
    root.appendChild(line);
  });

  const actions = el("div", { class: "actions" });
  if (files.length > 1) {
    actions.appendChild(el("button", {
      class: "btn btn-small",
      text: t("res.downloadZip"),
      onclick: () => triggerDownload(zipUrl(job.id), `${job.op}-${job.id}.zip`),
    }));
  }
  if (job.result?.outputDir) {
    actions.appendChild(el("button", {
      class: "btn btn-small btn-ghost",
      text: t("res.openFolder"),
      onclick: () => api.openFolder(job.result.outputDir).catch((e) => toast(e.message, "err")),
    }));
  }
  if (allowInsert && files.length === 1 && officeState().host !== "Browser") {
    const target = files[0];
    const canMaybeInsert = target.name.endsWith(".docx") || target.name.endsWith(".pptx");
    if (canMaybeInsert) {
      actions.appendChild(el("button", {
        class: "btn btn-small btn-ghost",
        text: t("res.insert"),
        onclick: async () => {
          try {
            const res = await fetch(fileUrl(job.id, 0));
            const blob = await res.blob();
            const out = await insertIntoCurrentDocument(blob, target.name);
            toast(out.ok ? t("res.inserted") : t("res.insertUnsupported"), out.ok ? "ok" : "warn");
          } catch (e) {
            toast(e.message, "err");
          }
        },
      }));
    }
  }
  if (actions.childElementCount) root.appendChild(actions);
}

/** 渲染图片结果画廊（仅当结果是图片时）。 */
export function renderGallery(root, job) {
  root.textContent = "";
  const files = (job.result?.files || []).filter((f) => /\.(png|jpe?g|webp|gif)$/i.test(f.name));
  files.slice(0, 24).forEach((f, i) => {
    const img = el("img", { alt: f.name, title: f.name });
    img.loading = "lazy";
    img.src = fileUrl(job.id, (job.result.files || []).indexOf(f));
    root.appendChild(img);
  });
}

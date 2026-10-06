/** 「设置」面板。 */
import { api, getBaseUrl, setBaseUrl } from "../api.js";
import { $, toast } from "../components.js";
import { t, setLang, getLang } from "../i18n.js";
import { officeState, openExternal } from "../office-bridge.js";

let serverOk = false;

export function isServerOk() {
  return serverOk;
}

export async function refreshServerStatus({ silent = true } = {}) {
  const dot = $("#serverDot");
  try {
    const health = await api.health();
    serverOk = true;
    dot?.classList.remove("dot-unknown", "dot-bad");
    dot?.classList.add("dot-ok");
    dot.title = `${health.name} ${health.version}`;
    if (!silent) toast(t("settings.reconnected"), "ok");
    return health;
  } catch (err) {
    serverOk = false;
    dot?.classList.remove("dot-unknown", "dot-ok");
    dot?.classList.add("dot-bad");
    dot.title = err.message;
    if (!silent) toast(t("settings.reconnectFail") + " — " + err.message, "err", 6000);
    return null;
  }
}

export async function loadEngines() {
  const list = $("#engineList");
  try {
    const info = await api.engines();
    list.textContent = "";
    const rows = [
      ["Microsoft Office (COM)", info.engines?.msoffice, info.engines?.msoffice?.detail],
      ["LibreOffice", info.engines?.libreoffice, info.engines?.libreoffice?.path || info.engines?.libreoffice?.detail],
      ["PDF 渲染 (pdf.js)", info.engines?.renderer, info.engines?.renderer?.detail],
      ["PDF 读写 (pdf-lib)", { available: true }, "pdf-lib"],
    ];
    for (const [name, engine, detail] of rows) {
      const li = document.createElement("li");
      const left = document.createElement("span");
      left.textContent = name;
      const right = document.createElement("span");
      right.className = engine?.available ? "ok" : "no";
      right.textContent = engine?.available ? t("settings.engineOk") + (detail ? ` · ${detail}` : "") : t("settings.engineNo");
      li.append(left, right);
      list.appendChild(li);
    }
    if (info.limits?.maxUploadMb) {
      const li = document.createElement("li");
      const left = document.createElement("span");
      left.textContent = "Max upload";
      const right = document.createElement("span");
      right.textContent = `${info.limits.maxUploadMb} MB`;
      li.append(left, right);
      list.appendChild(li);
    }
  } catch (err) {
    list.textContent = "";
    const li = document.createElement("li");
    li.className = "muted";
    li.textContent = err.message;
    list.appendChild(li);
  }
}

export function initSettingsPanel() {
  $("#inServerUrl").value = getBaseUrl();

  $("#btnSaveSettings").addEventListener("click", async () => {
    setBaseUrl($("#inServerUrl").value);
    toast(t("settings.saved"), "ok", 2000);
    await refreshServerStatus({ silent: true });
    await loadEngines();
    await loadServerSettings();
  });

  $("#btnCheckServer").addEventListener("click", () => refreshServerStatus({ silent: false }));

  $("#btnOpenOutput").addEventListener("click", async () => {
    try {
      await api.openFolder($("#inOutputDir").value);
    } catch (err) {
      toast(err.message, "err");
    }
  });

  $("#btnOpenDocs").addEventListener("click", () => openExternal("https://github.com/ZizzSaint/office-pdf-tools#readme"));

  const about = $("#aboutLine");
  if (about) {
    const st = officeState();
    about.textContent = `v1.0.0 · host=${st.host} · platform=${st.platform} · lang=${getLang()}`;
  }
}

export async function loadServerSettings() {
  try {
    const info = await api.settings();
    if (info.outputDir) $("#inOutputDir").value = info.outputDir;
  } catch { /* ignore */ }
}

export function initLangToggle() {
  $("#langToggle").addEventListener("click", () => {
    setLang(getLang() === "zh-CN" ? "en" : "zh-CN");
    loadEngines();
  });
}



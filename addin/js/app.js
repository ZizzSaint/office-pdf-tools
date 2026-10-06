/** 任务窗格入口：宿主识别、标签页路由、面板初始化。 */
import { applyI18n, detectLang, setLang, t } from "./i18n.js";
import { initOffice, officeState, hostKey } from "./office-bridge.js";
import { $, $$, toast } from "./components.js";
import { initConvertPanel } from "./panels/convert.js";
import { initImagesPanel } from "./panels/images.js";
import { initSplitMergePanel } from "./panels/splitmerge.js";
import {
  initSettingsPanel, initLangToggle, refreshServerStatus, loadEngines, loadServerSettings,
} from "./panels/settings.js";

const HASH_TAB = {
  "#to-pdf": "convert",
  "#pdf-to-office": "convert",
  "#split-merge": "splitmerge",
  "#images": "images",
  "#settings": "settings",
};

function activateTab(name, { scrollTo = null } = {}) {
  $$(".tab").forEach((tab) => tab.classList.toggle("active", tab.dataset.tab === name));
  $$(".panel").forEach((panel) => panel.classList.toggle("active", panel.id === `panel-${name}`));
  if (scrollTo) {
    const node = document.querySelector(scrollTo);
    node?.scrollIntoView({ behavior: "smooth", block: "start" });
    node?.classList.add("flash");
    setTimeout(() => node?.classList.remove("flash"), 1200);
  }
}

function bindTabs() {
  $$(".tab").forEach((tab) => {
    tab.addEventListener("click", () => activateTab(tab.dataset.tab));
  });
}

function applyHash() {
  const hash = location.hash || "";
  const tab = HASH_TAB[hash] || "convert";
  const scrollTo = hash === "#to-pdf" ? "#cardFileToPdf"
    : hash === "#pdf-to-office" ? "#cardPdfToOffice"
    : null;
  activateTab(tab, { scrollTo });
}

function renderHostBadge() {
  const badge = $("#hostBadge");
  const st = officeState();
  const key = hostKey();
  badge.textContent = t(`app.host.${key}`);
  badge.title = `${st.host} · ${st.platform} · ${st.documentName}`;
}

async function boot() {
  setLang(detectLang());
  applyI18n();
  bindTabs();
  initLangToggle();
  initConvertPanel();
  initImagesPanel();
  initSplitMergePanel();
  initSettingsPanel();

  await initOffice();
  renderHostBadge();
  applyHash();
  window.addEventListener("hashchange", applyHash);

  await refreshServerStatus({ silent: true });
  await loadServerSettings();
  loadEngines();
  setInterval(() => refreshServerStatus({ silent: true }), 30000);

  if (officeState().host === "Browser") {
    toast(t("app.host.browser"), "warn", 4000);
  }
}

window.addEventListener("DOMContentLoaded", () => {
  boot().catch((err) => toast(String(err?.message || err), "err", 8000));
});

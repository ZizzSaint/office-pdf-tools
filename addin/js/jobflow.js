/** 任务流：上传 → 轮询 → 渲染结果，并统一处理按钮禁用与错误提示。 */
import { startJob, waitJob } from "./api.js";
import { progressFor, renderResult, renderGallery, toast, $ } from "./components.js"; // eslint-disable-line no-unused-vars
import { t } from "./i18n.js";

let running = 0;

export function isBusy() {
  return running > 0;
}

/**
 * @param {{op:string, files:File[], options?:object, progressEl:HTMLElement,
 *          resultEl:HTMLElement, galleryEl?:HTMLElement|null, allowInsert?:boolean,
 *          buttons?:HTMLElement[], label?:string}} cfg
 */
export async function runJob(cfg) {
  const { op, files, options = {}, progressEl, resultEl, galleryEl = null, allowInsert = false } = cfg;
  const buttons = (cfg.buttons || []).filter(Boolean);
  const prog = progressFor(progressEl);
  buttons.forEach((b) => (b.disabled = true));
  running++;
  resultEl?.classList.add("hidden");
  if (galleryEl) galleryEl.textContent = "";
  try {
    if (!files.length) throw new Error(t("msg.noFile"));
    prog.show(cfg.label || t("msg.uploading"));
    const uploaded = await startJob(op, files, options, {
      onProgress: (r) => prog.set(r * 0.3, t("msg.uploading")),
    });
    prog.set(0.35, t("msg.processing"));
    const job = await waitJob(uploaded.jobId, {
      onUpdate: (j) => prog.set(0.35 + (Number(j.progress) || 0) * 0.65, j.step || t("msg.processing")),
    });
    prog.hide();
    if (resultEl) renderResult(resultEl, { job, allowInsert });
    if (galleryEl && job.status === "done") renderGallery(galleryEl, job);
    if (job.status === "error") toast(job.error?.message || t("msg.failed"), "err", 6000);
    else toast(t("msg.done"), "ok", 2200);
    return job;
  } catch (err) {
    prog.hide();
    if (resultEl) {
      resultEl.textContent = "";
      resultEl.classList.remove("hidden");
      resultEl.classList.add("err");
      const title = document.createElement("div");
      title.className = "result-title";
      title.textContent = "✗ " + t("msg.failed");
      const body = document.createElement("div");
      body.className = "muted";
      body.textContent = err.message || String(err);
      resultEl.append(title, body);
    }
    toast(err.message || String(err), "err", 6000);
    return { status: "error", error: { message: err.message } };
  } finally {
    running--;
    buttons.forEach((b) => (b.disabled = false));
  }
}

/** 从选择器里取出文件、校验并跑任务。 */
export async function runPickerJob(cfg) {
  const files = cfg.picker.files;
  if (!files.length) {
    toast(t("msg.noFile"), "warn");
    return null;
  }
  return runJob({ ...cfg, files });
}

export { $ };

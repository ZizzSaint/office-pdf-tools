/** 「转换」面板：当前文档导出、本地文件转 PDF、PDF 转 Office/图片。 */
import { runPickerJob, runJob } from "../jobflow.js";
import { createFilePicker, $, toast } from "../components.js";
import { exportCurrentToPdf, officeState, hostKey } from "../office-bridge.js";
import { saveBlob, api, triggerDownload } from "../api.js";
import { t } from "../i18n.js";

export function initConvertPanel() {
  const officePicker = createFilePicker({
    dropzone: "#dzOfficeToPdf",
    input: "#inOfficeToPdf",
    list: "#listOfficeToPdf",
    multiple: true,
  });

  const pdfPicker = createFilePicker({
    dropzone: "#dzPdfToOffice",
    input: "#inPdfToOffice",
    list: "#listPdfToOffice",
    multiple: false,
    accept: ".pdf",
  });

  /* ------------------------- 当前文档 → PDF ------------------------- */
  $("#btnCurrentToPdf").addEventListener("click", async (e) => {
    const btn = e.currentTarget;
    const result = $("#currentResult");
    const host = officeState().host;
    if (host === "Browser") {
      toast(t("convert.current.unsupported"), "warn", 5000);
      return;
    }
    btn.disabled = true;
    const old = btn.textContent;
    btn.textContent = t("convert.current.working");
    try {
      const { blob, native, fallbackName } = await exportCurrentToPdf();
      if (!native) {
        // 宿主不支持原生 PDF 导出：把源文件交给服务端引擎转换。
        const file = new File([blob], fallbackName, { type: blob.type });
        const job = await runJob({
          op: "office-to-pdf",
          files: [file],
          options: { engine: $("#selEngine")?.value || "auto" },
          progressEl: $("#progOfficeToPdf"),
          resultEl: $("#resOfficeToPdf"),
          buttons: [btn],
        });
        if (job?.status === "done" && $("#chkOpenAfter").checked) {
          const f = job.result.files[0];
          if (f?.path) api.openFile(f.path).catch(() => {});
        }
        return;
      }
      const name = (officeState().documentName || "document") + ".pdf";
      const saved = await saveBlob(blob, name, { source: "office-native" });
      const file = saved.files?.[0];
      result.textContent = "";
      result.classList.remove("hidden", "err");
      result.classList.add("ok");
      const title = document.createElement("div");
      title.className = "result-title";
      title.textContent = "✓ " + t("convert.current.saved");
      const info = document.createElement("div");
      info.className = "muted small";
      info.textContent = file?.path ? t("res.savedTo") + file.path : name;
      const actions = document.createElement("div");
      actions.className = "actions";
      const dl = document.createElement("button");
      dl.className = "btn btn-small";
      dl.textContent = t("res.download");
      dl.onclick = () => triggerDownload(saved.downloadUrl || api.downloadFallback?.(file), name);
      const op = document.createElement("button");
      op.className = "btn btn-small btn-ghost";
      op.textContent = t("res.open");
      op.onclick = () => api.openFile(file?.path).catch((err) => toast(err.message, "err"));
      const fo = document.createElement("button");
      fo.className = "btn btn-small btn-ghost";
      fo.textContent = t("res.openFolder");
      fo.onclick = () => api.openFolder(saved.outputDir).catch((err) => toast(err.message, "err"));
      actions.append(dl, op, fo);
      result.append(title, info, actions);
      toast(t("convert.current.saved") + " · " + (file?.path || name), "ok", 5000);
      if ($("#chkOpenAfter").checked && file?.path) api.openFile(file.path).catch(() => {});
    } catch (err) {
      toast(err.message || String(err), "err", 6000);
    } finally {
      btn.disabled = false;
      btn.textContent = old;
    }
  });

  /* ------------------------- 本地文件 → PDF ------------------------- */
  $("#btnOfficeToPdf").addEventListener("click", () =>
    runPickerJob({
      picker: officePicker,
      op: "office-to-pdf",
      options: {
        engine: $("#selEngine").value,
        sheet: $("#inSheet").value.trim() || undefined,
      },
      progressEl: $("#progOfficeToPdf"),
      resultEl: $("#resOfficeToPdf"),
      buttons: [$("#btnOfficeToPdf")],
    }),
  );

  /* ---------------------- PDF → Office / 图片 ---------------------- */
  const pdfToOffice = () =>
    runPickerJob({
      picker: pdfPicker,
      op: ($("#selPdfTarget").value === "png" || $("#selPdfTarget").value === "jpeg")
        ? "pdf-to-images"
        : "pdf-to-office",
      options: pdfTargetOptions(),
      progressEl: $("#progPdfToOffice"),
      resultEl: $("#resPdfToOffice"),
      galleryEl: $("#galleryPdfToOffice"),
      allowInsert: true,
      buttons: [$("#btnPdfToOffice")],
    });
  $("#btnPdfToOffice").addEventListener("click", pdfToOffice);

  return { officePicker, pdfPicker };
}

function pdfTargetOptions() {
  const target = $("#selPdfTarget").value;
  const isImage = target === "png" || target === "jpeg";
  return {
    target: isImage ? "image" : target,
    format: isImage ? target : undefined,
    mode: isImage ? "raster" : $("#selPdfMode").value,
    dpi: Number($("#inPdfDpi").value) || 150,
    pages: $("#inPdfPages").value.trim() || undefined,
  };
}

export { pdfTargetOptions };

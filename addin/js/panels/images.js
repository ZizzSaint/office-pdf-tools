/** 「图片 / PDF」面板。 */
import { runPickerJob } from "../jobflow.js";
import { createFilePicker, $ } from "../components.js";

export function initImagesPanel() {
  const imagesPicker = createFilePicker({
    dropzone: "#dzImages",
    input: "#inImages",
    list: "#listImages",
    multiple: true,
    thumbnails: true,
    sortable: true,
    accept: "image/*",
  });

  const pdfPicker = createFilePicker({
    dropzone: "#dzPdfToImages",
    input: "#inPdfToImages",
    list: "#listPdfToImages",
    multiple: false,
    accept: ".pdf",
  });

  $("#btnImagesToPdf").addEventListener("click", () =>
    runPickerJob({
      picker: imagesPicker,
      op: "images-to-pdf",
      options: {
        pageSize: $("#selPageSize").value,
        orientation: $("#selOrientation").value,
        margin: Number($("#inMargin").value) || 0,
        perPage: Number($("#selPerPage").value) || 1,
        fit: $("#selFit").value,
        separate: $("#chkImgSeparate").checked,
      },
      progressEl: $("#progImages"),
      resultEl: $("#resImages"),
      buttons: [$("#btnImagesToPdf")],
    }),
  );

  $("#btnPdfToImages").addEventListener("click", () =>
    runPickerJob({
      picker: pdfPicker,
      op: "pdf-to-images",
      options: {
        format: $("#selImgFormat").value,
        dpi: Number($("#inImgDpi").value) || 150,
        pages: $("#inImgPages").value.trim() || undefined,
      },
      progressEl: $("#progPdfToImages"),
      resultEl: $("#resPdfToImages"),
      galleryEl: $("#galleryPdfToImages"),
      buttons: [$("#btnPdfToImages")],
    }),
  );

  return { imagesPicker, pdfPicker };
}

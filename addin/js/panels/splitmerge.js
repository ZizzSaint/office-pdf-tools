/** 「拆分合并」面板。 */
import { runPickerJob } from "../jobflow.js";
import { createFilePicker, $, toast, el } from "../components.js";
import { t } from "../i18n.js";

const STRATEGIES = {
  pdf: [
    { value: "every-n-pages", label: "everyN", param: { type: "number", def: "2" } },
    { value: "ranges", label: "ranges", param: { type: "text", def: "1-3,4-6" } },
    { value: "each-page", label: "eachPage", param: null },
  ],
  docx: [
    { value: "heading", label: "heading", param: { type: "number", def: "1" } },
    { value: "section", label: "section", param: null },
  ],
  xlsx: [
    { value: "sheets", label: "sheets", param: null },
    { value: "every-n-rows", label: "everyNRows", param: { type: "number", def: "100" } },
  ],
  pptx: [
    { value: "every-n-slides", label: "everyN", param: { type: "number", def: "10" } },
    { value: "ranges", label: "ranges", param: { type: "text", def: "1-5,6-10" } },
  ],
};

export function kindOf(name = "") {
  const ext = name.toLowerCase().split(".").pop();
  if (ext === "pdf") return "pdf";
  if (["docx", "docm", "doc"].includes(ext)) return "docx";
  if (["xlsx", "xlsm", "xls"].includes(ext)) return "xlsx";
  if (["pptx", "pptm", "ppt"].includes(ext)) return "pptx";
  return null;
}

export function initSplitMergePanel() {
  const mergePicker = createFilePicker({
    dropzone: "#dzMerge",
    input: "#inMerge",
    list: "#listMerge",
    multiple: true,
    sortable: true,
    accept: ".pdf,.docx,.docm,.xlsx,.xlsm,.pptx,.pptm",
  });

  const splitPicker = createFilePicker({
    dropzone: "#dzSplit",
    input: "#inSplit",
    list: "#listSplit",
    multiple: false,
    accept: ".pdf,.docx,.docm,.xlsx,.xlsm,.pptx,.pptm",
  });

  const strategySelect = $("#selSplitStrategy");
  const paramLabel = $("#splitParamLabel");
  const paramInput = $("#inSplitParam");
  const hint = $("#splitHint");

  function refreshStrategies() {
    const file = splitPicker.files[0];
    const kind = kindOf(file?.name || "") || "pdf";
    const list = STRATEGIES[kind] || STRATEGIES.pdf;
    strategySelect.textContent = "";
    for (const s of list) {
      strategySelect.appendChild(el("option", { value: s.value, text: t(`split.strategy.${kind}.${s.label}`) }));
    }
    applyParam(kind);
  }

  function applyParam(kind) {
    const s = (STRATEGIES[kind] || STRATEGIES.pdf).find((x) => x.value === strategySelect.value)
      || (STRATEGIES[kind] || STRATEGIES.pdf)[0];
    if (s?.param) {
      paramInput.disabled = false;
      paramInput.value = s.param.def;
      paramLabel.parentElement?.classList.remove("hidden");
      paramInput.parentElement?.classList.remove("hidden");
    } else {
      paramInput.value = "";
      paramInput.disabled = true;
    }
    const label = s?.label || "everyN";
    hint.textContent = t(`split.hint.${kind}.${label}`);
  }

  splitPicker.onChange(() => refreshStrategies());
  strategySelect.addEventListener("change", () => applyParam(kindOf(splitPicker.files[0]?.name || "") || "pdf"));
  refreshStrategies();

  $("#btnMerge").addEventListener("click", () => {
    const files = mergePicker.files;
    if (files.length < 2) {
      toast(t("merge.needTwo"), "warn");
      return;
    }
    let kind = $("#selMergeKind").value;
    if (kind === "auto") {
      const kinds = new Set(files.map((f) => kindOf(f.name)).filter(Boolean));
      if (kinds.size !== 1) {
        toast(t("merge.mixed"), "warn");
        return;
      }
      kind = [...kinds][0];
      if (!kind) {
        toast(t("merge.unsupported") + files.map((f) => f.name).join(", "), "warn");
        return;
      }
    }
    return runPickerJob({
      picker: mergePicker,
      op: "merge",
      options: {
        kind,
        outputName: $("#inMergeName").value.trim() || undefined,
        pageBreak: $("#chkMergePageBreak").checked,
      },
      progressEl: $("#progMerge"),
      resultEl: $("#resMerge"),
      buttons: [$("#btnMerge")],
    });
  });

  $("#btnSplit").addEventListener("click", () => {
    const file = splitPicker.files[0];
    if (!file) {
      toast(t("msg.noFile"), "warn");
      return;
    }
    const kind = kindOf(file.name) || "auto";
    return runPickerJob({
      picker: splitPicker,
      op: "split",
      options: {
        kind,
        strategy: strategySelect.value,
        param: paramInput.value.trim() || undefined,
      },
      progressEl: $("#progSplit"),
      resultEl: $("#resSplit"),
      buttons: [$("#btnSplit")],
    });
  });

  return { mergePicker, splitPicker };
}

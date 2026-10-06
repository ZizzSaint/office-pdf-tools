/**
 * FunctionFile 入口（ribbon 命令）。
 * 目前所有 ribbon 按钮都用 ShowTaskpane 打开任务窗格，这里保留 Office.actions 注册，
 * 方便扩展 ExecuteFunction 型按钮（如“一键导出并打开”）。
 */
(function () {
  if (typeof Office === "undefined" || !Office.actions) return;

  async function showPane() {
    try {
      if (Office.addin && Office.addin.showAsTaskpane) {
        await Office.addin.showAsTaskpane();
        return;
      }
    } catch { /* fall through */ }
    try {
      if (Office.context?.ui?.displayDialogAsync) {
        Office.context.ui.displayDialogAsync(
          new URL("taskpane.html", window.location.href).href,
          { height: 70, width: 30, displayInIframe: true },
        );
      }
    } catch { /* ignore */ }
  }

  const handlers = {
    openPdfToolkit: showPane,
    openTaskpaneToPdf: showPane,
    openTaskpaneSplitMerge: showPane,
  };
  for (const [name, fn] of Object.entries(handlers)) {
    try {
      Office.actions.associate(name, fn);
    } catch { /* ignore */ }
  }
})();

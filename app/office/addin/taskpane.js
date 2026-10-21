(() => {
  let session;
  let name = "Excel workbook";
  const documentId = crypto.randomUUID();
  const status = document.getElementById("status");
  const disconnect = document.getElementById("disconnect");
  const update = (text) => { status.textContent = text; };
  disconnect.addEventListener("click", () => {
    session?.deactivate();
    session = null;
    disconnect.disabled = true;
    update("Disconnected. Reopen this add-in to connect again.");
  });
  globalThis.TeamsOfficeBridge = Object.freeze({
    describe: () => ({ documentId, name, application: "excel", ready: !!session && session.getDocumentUrl() === session.initialUrl }),
    execute: (request) => session ? session.execute(request) : Promise.resolve({ ok: false, code: "WRONG_DOCUMENT" }),
  });
  const timer = setTimeout(() => update("Excel has not finished connecting. Reopen the pane if this persists."), 15000);
  async function initialize() {
    if (!globalThis.Office) throw new Error("Office.js did not load");
    const info = await globalThis.Office.onReady();
    if (info.host !== globalThis.Office.HostType.Excel ||
      !globalThis.Office.context.requirements.isSetSupported("ExcelApi", "1.1")) throw new Error("Excel context unavailable");
    session = new globalThis.TeamsExcelOperations.ExcelSession(globalThis.Excel, {
      documentId, getDocumentUrl: () => globalThis.Office.context.document.url,
    });
    try {
      const part = new URL(globalThis.Office.context.document.url).pathname.split("/").pop();
      if (part) name = decodeURIComponent(part).slice(0, 256);
    } catch { /* Unsaved workbook or non-URL document identifier. */ }
    document.getElementById("workbook").textContent = `Workbook: ${name}`;
    clearTimeout(timer);
    disconnect.disabled = false;
    update("Ready for local agent access.");
  }
  void initialize().catch(() => {
    clearTimeout(timer);
    update("Open this add-in inside Excel in the Teams client to connect.");
  });
})();

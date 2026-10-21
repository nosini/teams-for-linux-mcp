(() => {
  let session;
  let name = "Office document";
  let application;
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
    describe: () => ({ documentId, name, application, ready: !!session && session.getDocumentUrl() === session.initialUrl }),
    execute: (request) => session ? session.execute(request) : Promise.resolve({ ok: false, code: "WRONG_DOCUMENT" }),
  });
  const timer = setTimeout(() => update("Office has not finished connecting. Reopen the pane if this persists."), 15000);
  async function initialize() {
    if (!globalThis.Office) throw new Error("Office.js did not load");
    const info = await globalThis.Office.onReady();
    const options = { isSetSupported: (name, version) => globalThis.Office.context.requirements.isSetSupported(name, version), documentId, getDocumentUrl: () => globalThis.Office.context.document.url };
    if (info.host === globalThis.Office.HostType.Excel && globalThis.Office.context.requirements.isSetSupported("ExcelApi", "1.1")) {
      application = "excel";
      name = "Excel workbook";
      session = new globalThis.TeamsExcelOperations.ExcelSession(globalThis.Excel, options);
    } else if (info.host === globalThis.Office.HostType.Word) {
      if (!globalThis.Office.context.requirements.isSetSupported("WordApi", "1.6")) {
        clearTimeout(timer);
        update("This Word host needs WordApi 1.6 to connect. Update Office or use a supported Word on the web host.");
        return;
      }
      application = "word";
      name = "Word document";
      session = new globalThis.TeamsWordOperations.WordSession(globalThis.Word, options);
    } else throw new Error("Office context unavailable");
    try {
      const part = new URL(globalThis.Office.context.document.url).pathname.split("/").pop();
      if (part) name = decodeURIComponent(part).slice(0, 256);
    } catch { /* Unsaved workbook or non-URL document identifier. */ }
    document.getElementById("workbook").textContent = `Document: ${name}`;
    clearTimeout(timer);
    disconnect.disabled = false;
    update("Ready for local agent access.");
  }
  void initialize().catch(() => {
    clearTimeout(timer);
    update("Open this add-in inside Excel or Word in the Teams client to connect.");
  });
})();

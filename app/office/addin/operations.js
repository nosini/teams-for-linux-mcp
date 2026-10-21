// Shared validation and Office.js operations. This module runs in the task pane
// and in Node tests; it never reads local files or accepts executable code.
((root, factory) => {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.TeamsExcelOperations = factory();
})(globalThis, () => {
  const messages = {
    INVALID_RANGE: "Use a single A1 rectangle of at most 1,000 cells within Excel's worksheet bounds.",
    INVALID_DATA: "Data must match the range dimensions and fit within 32 KB. Literal text cannot begin with =, + or -; use the formula tool for formulas.",
    WRONG_DOCUMENT: "The document changed or disconnected. List documents again.",
    WRONG_SHEET: "The worksheet ID changed. List worksheets again.",
    STALE_READ: "The read token expired, was used, or the cells changed. Read the range again before writing.",
    BUSY: "Another operation is running in this document. Wait for its result before continuing.",
    TIMEOUT: "The operation deadline expired. Read the cells before retrying a write; its outcome may be unknown.",
    TOO_LARGE: "The result is too large. Read a smaller range.",
    OFFICE_ERROR: "Excel could not complete the operation. Read the affected cells before retrying a write.",
  };
  function fail(code) { const error = new Error(messages[code]); error.code = code; throw error; }
  function rangeShape(address) {
    if (typeof address !== "string") fail("INVALID_RANGE");
    const match = /^([A-Z]{1,3})([1-9]\d{0,6})(?::([A-Z]{1,3})([1-9]\d{0,6}))?$/.exec(address.toUpperCase());
    if (!match) fail("INVALID_RANGE");
    const column = (s) => [...s].reduce((n, char) => n * 26 + char.charCodeAt(0) - 64, 0);
    const left = column(match[1]);
    const top = Number(match[2]);
    const right = column(match[3] || match[1]);
    const bottom = Number(match[4] || match[2]);
    const rows = bottom - top + 1;
    const columns = right - left + 1;
    if (rows < 1 || columns < 1 || right > 16384 || bottom > 1048576 || rows * columns > 1000) fail("INVALID_RANGE");
    return { rows, columns, address: address.toUpperCase() };
  }
  function bounded(value) {
    return new TextEncoder().encode(JSON.stringify(value)).length <= 32768;
  }
  function validateData(address, data, formulas = false) {
    const { rows, columns } = rangeShape(address);
    if (!Array.isArray(data) || data.length !== rows || !bounded(data)) fail("INVALID_DATA");
    for (const row of data) {
      if (!Array.isArray(row) || row.length !== columns) fail("INVALID_DATA");
      for (const value of row) {
        if (formulas) {
          if (typeof value !== "string" || !value.startsWith("=") || value.length > 8192) fail("INVALID_DATA");
        } else if (!(typeof value === "boolean" || (typeof value === "number" && Number.isFinite(value)) ||
          (typeof value === "string" && value.length <= 8192 && !/^[=+-]/.test(value)))) fail("INVALID_DATA");
      }
    }
  }

  class ExcelSession {
    constructor(excel, { documentId, getDocumentUrl, randomUUID = () => crypto.randomUUID(), now = Date.now }) {
      this.excel = excel;
      this.documentId = documentId;
      this.getDocumentUrl = getDocumentUrl;
      this.initialUrl = getDocumentUrl();
      this.randomUUID = randomUUID;
      this.now = now;
      this.snapshots = new Map();
      this.running = false;
      this.active = true;
    }

    deactivate() { this.active = false; this.snapshots.clear(); }

    check(request) {
      if (!this.active || request.documentId !== this.documentId || this.getDocumentUrl() !== this.initialUrl) fail("WRONG_DOCUMENT");
      if (!Number.isFinite(request.deadline) || this.now() >= request.deadline) fail("TIMEOUT");
    }

    snapshot(range, args) {
      const result = {
        worksheetId: args.worksheetId, address: args.address.toUpperCase(),
        values: range.values, formulas: range.formulas, text: range.text,
      };
      if (!bounded(result)) fail("TOO_LARGE");
      // Detach from Office's proxy objects and the result returned to callers.
      const saved = JSON.parse(JSON.stringify({ ...result, expires: this.now() + 120000 }));
      while (this.snapshots.size >= 64) this.snapshots.delete(this.snapshots.keys().next().value);
      const readToken = this.randomUUID();
      this.snapshots.set(readToken, saved);
      return { ...result, readToken };
    }

    async execute(request) {
      if (this.running) return { ok: false, code: "BUSY" };
      this.running = true;
      try {
        this.check(request);
        const { operation, args = {} } = request;
        if (operation === "list_worksheets") {
          const data = await this.excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/id,items/name,items/position");
            await context.sync();
            this.check(request);
            const worksheets = sheets.items.slice(0, 200).map(({ id, name, position }) => ({ id, name, position }));
            return { worksheets, truncated: sheets.items.length > 200 };
          });
          return { ok: true, data };
        }
        const writing = ["write_range", "set_formulas"].includes(operation);
        if (operation !== "read_range" && !writing) fail("OFFICE_ERROR");
        const shape = rangeShape(args.address);
        if (typeof args.worksheetId !== "string" || args.worksheetId.length > 256) fail("WRONG_SHEET");
        let expected;
        if (writing) {
          validateData(args.address, args.data, operation === "set_formulas");
          expected = this.snapshots.get(args.readToken);
          // A write attempt consumes the token even if its outcome is uncertain.
          this.snapshots.delete(args.readToken);
          if (!expected || expected.expires < this.now() || expected.worksheetId !== args.worksheetId || expected.address !== shape.address) fail("STALE_READ");
        }
        const data = await this.excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getItem(args.worksheetId);
          sheet.load("id");
          const range = sheet.getRange(shape.address);
          range.load("values,formulas,text,rowCount,columnCount");
          await context.sync();
          this.check(request);
          if (sheet.id !== args.worksheetId) fail("WRONG_SHEET");
          if (range.rowCount !== shape.rows || range.columnCount !== shape.columns) fail("INVALID_RANGE");
          if (writing) {
            if (JSON.stringify(range.values) !== JSON.stringify(expected.values) ||
              JSON.stringify(range.formulas) !== JSON.stringify(expected.formulas)) fail("STALE_READ");
            // Excel has no atomic compare-and-set. This fresh comparison narrows
            // the race window but cannot lock out coauthors during the write.
            if (operation === "set_formulas") range.formulas = args.data;
            else range.values = args.data;
            await context.sync();
            range.load("values,formulas,text");
            await context.sync();
          }
          return { ...this.snapshot(range, args), ...(writing ? { applied: true } : {}) };
        });
        return { ok: true, data };
      } catch (error) {
        return { ok: false, code: Object.hasOwn(messages, error?.code) ? error.code : "OFFICE_ERROR" };
      } finally {
        this.running = false;
      }
    }
  }
  return { ExcelSession, rangeShape, validateData, messages };
});

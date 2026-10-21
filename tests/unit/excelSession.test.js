const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { randomUUID } = require("node:crypto");
const { ExcelSession, rangeShape, validateData } = require("../../app/office/addin/operations");
const { excelFixture } = require("../fixtures/excelApi");

function fixture(initial) {
  const { excel, model } = excelFixture(initial);
  const documentId = randomUUID();
  const state = { now: 1000, url: "https://tenant.sharepoint.com/book.xlsx" };
  const session = new ExcelSession(excel, { documentId, randomUUID, now: () => state.now, getDocumentUrl: () => state.url });
  const call = (operation, args = {}, overrides = {}) => session.execute({ documentId, deadline: state.now + 10000, operation, args, ...overrides });
  return { session, model, state, call };
}

const target = { worksheetId: "sheet-id", address: "A1" };
describe("Excel live session", () => {
  it("reads live cells, writes with a snapshot, and returns a new read token", async () => {
    const { model, call } = fixture();
    const sheets = await call("list_worksheets");
    assert.equal(sheets.data.worksheets[0].id, "sheet-id");
    const before = await call("read_range", target);
    assert.equal(before.ok, true);
    assert.deepEqual(before.data.values, [[1]]);
    const result = await call("write_range", { ...target, readToken: before.data.readToken, data: [[7]] });
    assert.equal(result.ok, true);
    assert.equal(result.data.applied, true);
    assert.deepEqual(result.data.values, [[7]]);
    assert.notEqual(result.data.readToken, before.data.readToken);
    assert.equal(model.writes, 1);
    assert.equal((await call("write_range", { ...target, readToken: before.data.readToken, data: [[9]] })).code, "STALE_READ");
    assert.equal(model.writes, 1);
  });

  it("rejects coauthor edits, including changed formulas with the same calculated result", async () => {
    for (const field of ["values", "formulas"]) {
      const { model, call } = fixture();
      const { data } = await call("read_range", target);
      model[field] = field === "values" ? [[8]] : [["=1"]];
      const result = await call("write_range", { ...target, readToken: data.readToken, data: [[99]] });
      assert.equal(result.code, "STALE_READ");
      assert.equal(model.writes, 0);
    }
  });

  it("binds read tokens to the document, worksheet and exact range, and expires them", async () => {
    for (const change of ["document", "url", "sheet", "range", "expiry"]) {
      const { model, state, call } = fixture();
      const { data } = await call("read_range", target);
      const args = { ...target, readToken: data.readToken, data: [[7]] };
      const overrides = {};
      if (change === "document") overrides.documentId = randomUUID();
      if (change === "url") state.url = "https://tenant.sharepoint.com/other.xlsx";
      if (change === "sheet") args.worksheetId = "other-id";
      if (change === "range") args.address = "B1";
      if (change === "expiry") state.now += 120000; // Exactly at expiry, as for structure tokens.
      const result = await call("write_range", args, overrides);
      assert.equal(result.ok, false, change);
      if (change !== "document" && change !== "url") assert.equal(result.code, "STALE_READ", change);
      assert.equal(model.writes, 0, change);
    }
  });

  it("sets formulas explicitly and rejects formula-like text in value writes", async () => {
    const { model, call } = fixture();
    const { data } = await call("read_range", target);
    for (const value of ["=1+2", "+1+2", "-1+2"]) {
      assert.equal((await call("write_range", { ...target, readToken: data.readToken, data: [[value]] })).code, "INVALID_DATA");
    }
    const result = await call("set_formulas", { ...target, readToken: data.readToken, data: [["=1+2"]] });
    assert.equal(result.ok, true);
    assert.deepEqual(model.formulas, [["=1+2"]]);
    assert.deepEqual(result.data.values, [[3]]);
    assert.equal(model.writes, 1);
  });

  it("checks the deadline again after the fresh read and before queuing writes", async () => {
    const { model, state, call } = fixture();
    const { data } = await call("read_range", target);
    model.beforeSync = () => { state.now += 10001; };
    const result = await call("write_range", { ...target, readToken: data.readToken, data: [[7]] });
    assert.equal(result.code, "TIMEOUT");
    assert.equal(model.writes, 0);
  });

  it("honors disconnect during a pending fresh read before any write is queued", async () => {
    const { session, model, call } = fixture();
    const { data } = await call("read_range", target);
    model.beforeSync = () => session.deactivate();
    const result = await call("write_range", { ...target, readToken: data.readToken, data: [[7]] });
    assert.equal(result.code, "WRONG_DOCUMENT");
    assert.equal(model.writes, 0);
  });

  it("reports uncertain write failures without retrying and consumes their token", async () => {
    const { model, call } = fixture();
    const { data } = await call("read_range", target);
    model.beforeSync = (count) => { if (count === 4) throw new Error("private document content"); };
    const result = await call("write_range", { ...target, readToken: data.readToken, data: [[7]] });
    assert.deepEqual(result, { ok: false, code: "OFFICE_ERROR" });
    assert.equal(model.writes, 1);
    assert.equal((await call("write_range", { ...target, readToken: data.readToken, data: [[8]] })).code, "STALE_READ");
  });

  it("does not treat a worksheet name as its stable ID", async () => {
    const { model, call } = fixture();
    assert.equal((await call("read_range", { ...target, worksheetId: "Sheet One" })).code, "WRONG_SHEET");
    assert.equal(model.writes, 0);
  });

  it("bounds coordinates, data dimensions and UTF-8 payload size before accessing Excel", () => {
    for (const address of ["A0", "XFE1", "A1048577", "A1:B1000", "B2:A1", "A:A", "A1,B2", "Sheet!A1", "A1;alert(1)"]) {
      assert.throws(() => rangeShape(address), /A1 rectangle/);
    }
    assert.deepEqual(rangeShape("xfd1048576"), { rows: 1, columns: 1, address: "XFD1048576" });
    assert.throws(() => validateData("A1:B1", [[1]]));
    assert.throws(() => validateData("A1", [[null]]));
    assert.throws(() => validateData("A1", [[Infinity]]));
    assert.throws(() => validateData("A1:A2", [["漢".repeat(8000)], ["漢".repeat(8000)]]));
  });
});

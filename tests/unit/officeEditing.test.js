const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { randomUUID } = require("node:crypto");
const { ExcelSession, editing } = require("../../app/office/addin/operations");
const { WordSession } = require("../../app/office/addin/wordOperations");
const { excelFixture } = require("../fixtures/excelApi");
const { wordFixture } = require("../fixtures/wordApi");

function fixture(application = "word", data, apiSupport = () => true) {
  const api = application === "word" ? wordFixture(data) : excelFixture(data);
  let now = 1000, url = "document";
  const options = { documentId: randomUUID(), randomUUID, now: () => now, getDocumentUrl: () => url, isSetSupported: apiSupport };
  const session = application === "word" ? new WordSession(api.word, options) : new ExcelSession(api.excel, options);
  const call = (operation, args = {}) => session.execute({ documentId: options.documentId, deadline: now + 10000, operation, args });
  const paragraph = async () => (await call("read_paragraphs", { count: 1 })).data.paragraphs[0];
  return { call, session, model: api.model, paragraph, expire: () => { now += 120000; }, navigate: () => { url = "other"; } };
}
const target = { worksheetId: "sheet-id", address: "A1" };

describe("Extended Excel live editing", () => {
  it("discovers used bounds and tables, including empty sheets and table pagination", async () => {
    const { call, model } = fixture("excel");
    model.tables = [{ id: "table-id", name: "Sales", range: { address: "Sheet One!A1:B2", rowCount: 2, columnCount: 2 } }, { id: "other", name: "Costs", range: { address: "Sheet One!D1:E2", rowCount: 2, columnCount: 2 } }];
    const result = await call("inspect_worksheet", { worksheetId: target.worksheetId, count: 1 });
    assert.equal(result.ok, true);
    assert.equal(result.data.usedRange.rowCount, 1);
    assert.equal(result.data.tables[0].name, "Sales");
    assert.equal(result.data.nextStart, 1);
    assert.equal((await call("inspect_worksheet", { worksheetId: target.worksheetId, start: 1 })).data.tables[0].name, "Costs");
    model.empty = true;
    assert.equal((await call("inspect_worksheet", { worksheetId: target.worksheetId })).data.usedRange, null);
    assert.equal((await call("inspect_worksheet", { worksheetId: "Sheet One" })).code, "WRONG_SHEET");
  });

  it("searches literal text/formulas in bounded windows and continues after 50 matches", async () => {
    const { call, model } = fixture("excel", [Array(60).fill("Alpha")]);
    const args = { worksheetId: "sheet-id", address: "B3:BI3", query: "alpha" };
    const first = await call("search_range", args);
    assert.equal(first.ok, true);
    assert.equal(first.data.matches.length, 50);
    assert.equal(first.data.matches[0].address, "B3");
    assert.equal(first.data.nextOffset, 50);
    const second = await call("search_range", { ...args, offset: first.data.nextOffset });
    assert.equal(second.data.matches.length, 10);
    assert.equal(second.data.nextOffset, null);
    assert.equal((await call("search_range", { ...args, matchCase: true })).data.matches.length, 0);
    model.formulas[0][0] = "=SUM(A1:A2)";
    assert.equal((await call("search_range", { ...args, query: "SUM(", searchIn: "formulas" })).data.matches.length, 1);
    assert.equal((await call("search_range", { ...args, address: "A1:B1000" })).code, "INVALID_RANGE");
  });

  it("reads and changes only requested range formatting while preserving cell contents", async () => {
    const { call, model } = fixture("excel", [[0.25]]);
    const before = (await call("read_range_format", target)).data;
    assert.deepEqual(before.format.numberFormat, [["General"]]);
    const result = await call("format_range", { ...target, readToken: before.readToken, format: { numberFormat: "0.0%", font: { bold: true }, fillColor: "#FFFF00", wrapText: true } });
    assert.equal(result.ok, true);
    assert.equal(result.data.applied, true);
    assert.deepEqual(model.values, [[0.25]]);
    assert.deepEqual(model.numberFormat, [["0.0%"]]);
    assert.equal(model.font.bold, true);
    assert.equal(model.font.name, "Arial");
    assert.equal(model.fillColor, "#FFFF00");
    assert.equal((await call("format_range", { ...target, readToken: before.readToken, format: { font: { bold: false } } })).code, "STALE_STRUCTURE");
  });

  it("rejects stale formats, changed values/formulas, mismatched ranges, and token misuse", async () => {
    for (const change of ["format", "values", "formulas", "range", "token-kind"]) {
      const { call, model } = fixture("excel");
      const before = (await call(change === "token-kind" ? "read_range" : "read_range_format", target)).data;
      if (change === "format") model.font.bold = true;
      if (change === "values") model.values = [[9]];
      if (change === "formulas") model.formulas = [["=1"]];
      const result = await call("format_range", { ...target, ...(change === "range" ? { address: "B1" } : {}), readToken: before.readToken, format: { font: { bold: false } } });
      assert.equal(result.ok, false, change);
      assert.equal(model.writes, 0, change);
    }
  });

  it("creates sheets from a current worksheet-list token and rejects changed lists and invalid names", async () => {
    const { call, model, session } = fixture("excel");
    const before = (await call("list_worksheets")).data;
    const result = await call("create_worksheet", { name: "Calculations", readToken: before.readToken });
    assert.equal(result.ok, true);
    assert.equal(result.data.name, "Calculations");
    assert.equal(model.extraSheets.length, 1);
    assert.equal(session.snapshots.size, 0);
    const stale = (await call("list_worksheets")).data;
    model.extraSheets[0].name = "Renamed by coauthor";
    assert.equal((await call("create_worksheet", { name: "New", readToken: stale.readToken })).code, "STALE_STRUCTURE");
    const current = (await call("list_worksheets")).data;
    assert.equal((await call("create_worksheet", { name: "Sheet One", readToken: current.readToken })).code, "INVALID_EDIT");
    for (const name of ["Bad/name", "", "'Bad", "X".repeat(32)]) assert.equal((await call("create_worksheet", { name, readToken: current.readToken })).code, "INVALID_EDIT");
  });

  it("inserts whole rows/columns with fresh layout tokens and rejects layout changes", async () => {
    for (const axis of ["rows", "columns"]) {
      const { call, model, session } = fixture("excel");
      const before = (await call("inspect_worksheet", { worksheetId: "sheet-id" })).data;
      assert.equal((await call("insert_rows_columns", { worksheetId: "sheet-id", readToken: before.readToken, axis, index: 2, count: 2 })).ok, true);
      assert.equal(model.inserted[0].direction, axis === "rows" ? "Down" : "Right");
      assert.equal(session.snapshots.size, 0);
      const stale = (await call("inspect_worksheet", { worksheetId: "sheet-id" })).data;
      model.values.push([2]);
      assert.equal((await call("insert_rows_columns", { worksheetId: "sheet-id", readToken: stale.readToken, axis, index: 0, count: 1 })).code, "STALE_STRUCTURE");
    }
  });
});

describe("Extended Word live editing", () => {
  it("discovers styles and headings, paragraph formatting, tables and paged cell text", async () => {
    const { call, model } = fixture();
    model.paragraphs[0].styleBuiltIn = "Heading1";
    model.paragraphs[0].style = "Kop 1";
    model.tables = [{ rows: [["A", "B"], ["Merged"], ["C", "D"]] }];
    const structure = (await call("read_structure", { count: 1 })).data;
    assert.equal(structure.paragraphs[0].styleBuiltIn, "Heading1");
    assert.equal(structure.paragraphs[0].style, "Kop 1");
    assert.equal(structure.nextStart, 1);
    const details = (await call("read_paragraph_details", structure.paragraphs[0])).data;
    assert.equal(details.format.font.name, "Arial");
    assert.equal((await call("list_tables")).data.tables[0].rowCount, 3);
    const table = (await call("read_table", { tableIndex: 0, rowCount: 2 })).data;
    assert.deepEqual(table.rows.map((row) => row.cellCount), [2, 1]);
    assert.equal(table.nextRowStart, 2);
    assert.deepEqual((await call("read_table", { tableIndex: 0, rowStart: 2 })).data.rows[0].values, [["C", "D"]]);
  });

  it("replaces long paragraphs, inserts with stable anchoring and deletes explicit paragraphs", async () => {
    const { call, paragraph, model, session } = fixture();
    const before = await paragraph();
    const replaced = await call("replace_paragraph", { ...before, text: "Long text. ".repeat(100) });
    assert.equal(replaced.ok, true);
    assert.equal(model.paragraphs[0].text.length, 1100);
    const inserted = await call("insert_paragraph", { ...replaced.data, text: "New section", position: "After", styleBuiltIn: "Heading1" });
    assert.equal(inserted.ok, true);
    assert.equal(model.paragraphs[1].id, inserted.data.paragraphId);
    assert.equal(model.paragraphs[1].styleBuiltIn, "Heading1");
    assert.equal((await call("delete_paragraph", inserted.data)).ok, true);
    assert.equal(model.paragraphs.length, 2);
    assert.equal(session.snapshots.size, 0);
    assert.equal((await call("replace_paragraph", { ...before, text: "stale" })).code, "STALE_STRUCTURE");
  });

  it("formats only requested paragraph properties and rejects formatting-only coauthor changes", async () => {
    const { call, paragraph, model } = fixture();
    const target = await paragraph();
    const before = (await call("read_paragraph_details", target)).data;
    const result = await call("format_paragraph", { ...before, format: { styleBuiltIn: "Heading2", font: { bold: true }, alignment: "Centered" } });
    assert.equal(result.ok, true);
    assert.equal(model.paragraphs[0].font.bold, true);
    assert.equal(model.paragraphs[0].font.name, "Arial");
    assert.equal(model.paragraphs[0].text, target.text);
    const fresh = (await call("read_paragraph_details", target)).data;
    model.paragraphs[0].font.size = 20;
    assert.equal((await call("format_paragraph", { ...fresh, format: { font: { italic: true } } })).code, "STALE_PARAGRAPH");
    assert.equal(model.paragraphs[0].font.italic, false);
    assert.equal((await call("format_paragraph", { ...target, format: { font: { italic: true } } })).code, "STALE_PARAGRAPH");
  });

  it("adds, reads, replies to, resolves and reopens a stable comment thread", async () => {
    const { call, paragraph, model } = fixture();
    const target = await paragraph();
    const added = await call("add_comment", { ...target, text: "Please check this assumption." });
    assert.equal(added.ok, true);
    const read = async () => (await call("read_comments", { paragraphId: target.paragraphId })).data.comments[0];
    const comment = await read();
    const args = { paragraphId: target.paragraphId, commentId: comment.comment.id };
    assert.equal(comment.comment.content, "Please check this assumption.");
    assert.equal((await call("reply_comment", { ...args, readToken: comment.readToken, text: "Checked." })).ok, true);
    const replied = await read();
    assert.equal(replied.replies[0].content, "Checked.");
    assert.equal((await call("resolve_comment", { ...args, readToken: replied.readToken, resolved: true })).ok, true);
    const resolved = await read();
    assert.equal(resolved.comment.resolved, true);
    assert.equal((await call("resolve_comment", { ...args, readToken: resolved.readToken, resolved: false })).ok, true);
    assert.equal(model.paragraphs[0].comments[0].resolved, false);
  });

  it("rejects changed/deleted comments, new replies and truncated discussions", async () => {
    for (const change of ["content", "delete", "reply", "many-replies"]) {
      const { call, paragraph, model } = fixture();
      const target = await paragraph();
      await call("add_comment", { ...target, text: "Question" });
      const comment = (await call("read_comments", target)).data.comments[0];
      const stored = model.paragraphs[0].comments[0];
      if (change === "content") stored.content = "Changed";
      if (change === "delete") model.paragraphs[0].comments = [];
      if (change === "reply") stored.replies.push({ id: "coauthor", content: "New reply", authorName: "Other", creationDate: new Date(0) });
      if (change === "many-replies") stored.replies = Array.from({ length: 21 }, (_, i) => ({ id: String(i), content: "Reply" }));
      const before = model.writes;
      const result = await call("resolve_comment", { paragraphId: target.paragraphId, commentId: comment.comment.id, readToken: comment.readToken, resolved: true });
      assert.equal(result.code, "STALE_STRUCTURE", change);
      assert.equal(model.writes, before, change);
    }
  });

  it("changes tracking only explicitly and rejects another editor's changed mode", async () => {
    const { call, model } = fixture();
    const before = (await call("get_change_tracking")).data;
    assert.equal(before.mode, "Off");
    const changed = await call("set_change_tracking", { readToken: before.readToken, mode: "TrackAll" });
    assert.equal(changed.ok, true);
    assert.equal(model.trackingMode, "TrackAll");
    model.trackingMode = "TrackMineOnly";
    assert.equal((await call("set_change_tracking", { readToken: changed.data.readToken, mode: "Off" })).code, "STALE_STRUCTURE");
    assert.equal(model.trackingMode, "TrackMineOnly");
  });

  it("reads revisions and accepts/rejects one only when the entire inspected set is unchanged", async () => {
    for (const action of ["accept", "reject", "stale"]) {
      const { call, model, paragraph } = fixture();
      const target = await paragraph();
      model.paragraphs[0].originalText = "Original text";
      model.paragraphs[0].changes = [{ author: "Other", date: new Date(0), text: "Hello", type: "Added" }];
      const before = (await call("read_tracked_changes", target)).data;
      assert.equal(before.originalText, "Original text");
      assert.equal(before.changes.length, 1);
      if (action === "stale") model.paragraphs[0].changes.push({ author: "Other", date: new Date(1), text: "More", type: "Added" });
      const result = await call("resolve_tracked_change", { paragraphId: target.paragraphId, readToken: before.readToken, changeIndex: 0, action: action === "stale" ? "accept" : action });
      assert.equal(result.ok, action !== "stale", action);
      assert.equal(model.writes, action === "stale" ? 0 : 1);
      if (action === "reject") assert.equal(model.paragraphs[0].text, "Original text");
    }
  });

  it("bounds revision sets and text, checks capability support, consumes expired tokens and honors disconnects", async () => {
    const { call, paragraph, model, expire } = fixture();
    const target = await paragraph();
    model.paragraphs[0].changes = Array.from({ length: 51 }, () => ({ text: "x", author: "a", type: "Added", date: new Date(0) }));
    assert.equal((await call("read_tracked_changes", target)).code, "TOO_LARGE");
    assert.equal((await call("replace_paragraph", { ...target, text: "line\nbreak" })).code, "INVALID_EDIT");
    expire();
    assert.equal((await call("delete_paragraph", target)).code, "STALE_STRUCTURE");
    const unsupported = fixture("word", undefined, () => false);
    assert.equal((await unsupported.call("read_structure")).code, "UNSUPPORTED_API");
    const unsupportedExcel = fixture("excel", undefined, () => false);
    assert.equal((await unsupportedExcel.call("inspect_worksheet", { worksheetId: "sheet-id" })).code, "UNSUPPORTED_API");
    const disconnected = fixture();
    const before = await disconnected.paragraph();
    disconnected.model.beforeSync = () => disconnected.session.deactivate();
    assert.equal((await disconnected.call("delete_paragraph", before)).code, "WRONG_DOCUMENT");
    assert.equal(disconnected.model.writes, 0);
  });

  it("validates formatting allowlists instead of forwarding arbitrary proxy properties", () => {
    assert.equal(editing.validFormat({ font: { bold: true } }, "word"), true);
    for (const format of [{}, { font: {} }, { style: "Custom", styleBuiltIn: "Normal" }, { font: { color: "javascript:bad" } }, { font: { size: Infinity } }, { font: { arbitrary: true } }, { arbitrary: true }]) assert.equal(editing.validFormat(format, "word"), false);
  });
});

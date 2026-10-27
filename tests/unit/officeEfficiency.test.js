const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { randomUUID } = require("node:crypto");
const { ExcelSession, readTokenLifetime } = require("../../app/office/addin/operations");
const { WordSession } = require("../../app/office/addin/wordOperations");
const { excelFixture } = require("../fixtures/excelApi");
const { wordFixture } = require("../fixtures/wordApi");

function fixture(application, initial) {
  const api = application === "word" ? wordFixture(initial) : excelFixture(initial);
  let now = 1000;
  const documentId = randomUUID();
  const options = { documentId, now: () => now, getDocumentUrl: () => "document" };
  const session = application === "word" ? new WordSession(api.word, options) : new ExcelSession(api.excel, options);
  const call = (operation, args = {}) => session.execute({ documentId, deadline: now + 10000, operation, args });
  return { session, model: api.model, call, advance: (milliseconds) => { now += milliseconds; } };
}
const target = { worksheetId: "sheet-id", address: "A1" };
const bytes = (value) => Buffer.byteLength(JSON.stringify(value));

describe("Office context efficiency", () => {
  it("inserts a section with one token in array order and keeps unrelated tokens valid", async () => {
    for (const position of ["Before", "After"]) {
      const { call, model } = fixture("word", ["Anchor", "Unrelated"]);
      const [anchor, unrelated] = (await call("read_paragraphs")).data.paragraphs;
      const inserted = await call("insert_paragraphs", { ...anchor, position, paragraphs: [{ text: "Heading", styleBuiltIn: "Heading1" }, { text: "Body" }, { text: "End" }] });
      assert.equal(inserted.ok, true);
      assert.equal(inserted.data.paragraphs.length, 3);
      assert.equal(new Set(inserted.data.paragraphs.map(({ paragraphId }) => paragraphId)).size, 3);
      assert.deepEqual(model.paragraphs.map(({ text }) => text), position === "Before" ? ["Heading", "Body", "End", "Anchor", "Unrelated"] : ["Anchor", "Heading", "Body", "End", "Unrelated"]);
      assert.equal(model.paragraphs.find(({ text }) => text === "Heading").styleBuiltIn, "Heading1");
      assert.equal(Object.hasOwn(inserted.data.paragraphs[0], "text"), false);
      assert.equal((await call("replace_paragraph", { ...unrelated, text: "Still valid" })).ok, true);
      assert.equal((await call("replace_paragraph", { ...inserted.data.paragraphs[1], text: "Updated body" })).ok, true);
    }
  });

  it("validates an entire inserted section before writing", async () => {
    const { call, model } = fixture("word");
    const anchor = (await call("read_paragraphs")).data.paragraphs[0];
    for (const paragraphs of [[], [{ text: "Valid" }, { text: "Invalid\nbreak" }], [{ text: "x", styleBuiltIn: "Invalid" }], Array(21).fill({ text: "x" }), [{ text: "漢".repeat(8192) }, { text: "漢".repeat(8192) }]]) {
      assert.equal((await call("insert_paragraphs", { ...anchor, position: "After", paragraphs })).code, "INVALID_EDIT");
      assert.equal(model.writes, 0);
    }
    model.paragraphs[0].text = "Changed";
    assert.equal((await call("insert_paragraphs", { ...anchor, position: "After", paragraphs: [{ text: "New" }] })).code, "STALE_PARAGRAPH");
    assert.equal(model.writes, 0);
  });

  it("edits a page of Word paragraphs without rereading after each replacement", async () => {
    const { call, model } = fixture("word", Array.from({ length: 20 }, (_, i) => `Paragraph ${i}`));
    const page = (await call("read_paragraphs")).data.paragraphs;
    for (const paragraph of page) {
      const result = await call("replace_paragraph", { ...paragraph, text: `Updated ${paragraph.text}` });
      assert.equal(result.ok, true);
      assert.deepEqual(Object.keys(result.data).sort(), ["applied", "paragraphId", "readToken"]);
    }
    assert.equal(model.writes, 20);
    assert.ok(model.paragraphs.every(({ text }) => text.startsWith("Updated ")));
  });

  it("preserves other paragraph tokens after insert/delete but rejects changed or deleted targets", async () => {
    for (const operation of ["insert_paragraph", "delete_paragraph", "format_paragraph"]) {
      const { call, model } = fixture("word", ["First", "Second", "Third"]);
      const [first, second, third] = (await call("read_paragraphs")).data.paragraphs;
      const token = operation === "format_paragraph" ? (await call("read_paragraph_details", first)).data : first;
      const result = await call(operation, { ...token, text: "Inserted", position: "After", format: { font: { bold: true } } });
      assert.equal(result.ok, true, operation);
      model.paragraphs.find(({ id }) => id === second.paragraphId).text = "Coauthor edit";
      assert.equal((await call("replace_paragraph", { ...second, text: "Overwrite" })).code, "STALE_PARAGRAPH");
      assert.equal((await call("replace_paragraph", { ...third, text: "Updated third" })).ok, true, operation);
      assert.equal((await call("replace_paragraph", { ...first, text: "Reuse" })).ok, false);
    }
  });

  it("keeps tokens past two minutes, checks current content, expires at fifteen minutes and bounds retention", async () => {
    for (const application of ["word", "excel"]) {
      const { call, advance, session, model } = fixture(application);
      const read = async () => application === "word" ? (await call("read_paragraphs")).data.paragraphs[0] : (await call("read_range", target)).data;
      const write = (token) => application === "word" ? call("replace_paragraph", { ...token, text: "Updated" }) : call("write_range", { ...target, ...token, data: [[2]] });
      const first = await read();
      advance(120001);
      assert.equal((await write(first)).ok, true);
      const second = await read();
      advance(readTokenLifetime - 1);
      if (application === "word") model.paragraphs[0].text = "Coauthor edit";
      else model.formulas = [["=1+1"]];
      assert.equal((await write(second)).ok, false, "Long-lived tokens still compare live state");
      const third = await read();
      advance(readTokenLifetime);
      assert.equal((await write(third)).ok, false, "Expiry is inclusive");
      for (let i = 0; i < 520; i++) await read();
      assert.equal(session.snapshots.size, 512);
    }
  });

  it("returns sparse formula/display differences without losing the full comparison snapshot", async () => {
    const { call, model, session } = fixture("excel", [[3, "Label", null]]);
    model.formulas = [["=1+2", "Label", null]];
    const range = { ...target, address: "A1:C1" };
    const compact = (await call("read_range", range)).data;
    assert.deepEqual(compact.values, [[3, "Label", null]]);
    assert.deepEqual(compact.formulaCells, [[0, 0, "=1+2"]]);
    assert.deepEqual(compact.textCells, [[0, 2, "null"]]);
    assert.equal(Object.hasOwn(compact, "formulas"), false);
    assert.equal(Object.hasOwn(compact, "text"), false);
    assert.deepEqual(session.snapshots.get(compact.readToken).formulas, model.formulas);
    model.formulas[0][0] = "=3";
    assert.equal((await call("write_range", { ...range, readToken: compact.readToken, data: [[4, "Label", ""]] })).code, "STALE_READ");
    assert.equal(model.writes, 0);
    const full = (await call("read_range", { ...range, responseMode: "full" })).data;
    assert.deepEqual(full.formulas, model.formulas);
    assert.deepEqual(full.text, [["3", "Label", "null"]]);
    for (const responseMode of [null, false, "invalid"]) assert.equal((await call("read_range", { ...range, responseMode })).code, "INVALID_EDIT");
  });

  it("fits a compact 1,000-cell formatting read that the dense response cannot fit", async () => {
    const values = Array.from({ length: 40 }, (_, r) => Array.from({ length: 25 }, (_, c) => `Cell ${r},${c}`));
    const { call, session } = fixture("excel", values);
    const range = { ...target, address: "A1:Y40" };
    assert.equal((await call("read_range_format", { ...range, responseMode: "full" })).code, "TOO_LARGE");
    const compact = await call("read_range_format", range);
    assert.equal(compact.ok, true);
    assert.equal(compact.data.format.numberFormat, "General");
    assert.equal(Object.hasOwn(compact.data, "values"), false);
    assert.deepEqual(session.snapshots.get(compact.data.readToken).values, values);
    const changed = await call("format_range", { ...range, readToken: compact.data.readToken, format: { font: { bold: true } } });
    assert.equal(changed.ok, true);
    assert.ok(bytes(changed.data) < 200);
    assert.ok(bytes(compact.data) < 32768);
    assert.equal((await call("format_range", { ...range, readToken: changed.data.readToken, format: { wrapText: true } })).ok, true);
  });

  it("preserves mixed number formats and cuts normal read/write payloads substantially", async (t) => {
    const values = Array.from({ length: 20 }, (_, r) => Array.from({ length: 10 }, (_, c) => `Cell ${r},${c}`));
    const { call, model } = fixture("excel", values);
    const range = { ...target, address: "A1:J20" };
    const compact = (await call("read_range_format", range)).data;
    const full = (await call("read_range_format", { ...range, responseMode: "full" })).data;
    assert.ok(bytes(compact) < bytes(full) * 0.4);
    const compactWrite = (await call("format_range", { ...range, readToken: compact.readToken, format: { wrapText: true } })).data;
    const fresh = (await call("read_range_format", range)).data;
    const fullWrite = (await call("format_range", { ...range, readToken: fresh.readToken, format: { wrapText: false }, responseMode: "full" })).data;
    assert.ok(bytes(compactWrite) < bytes(fullWrite) * 0.03);
    t.diagnostic(`200-cell responses: reads ${bytes(full)} -> ${bytes(compact)} bytes; writes ${bytes(fullWrite)} -> ${bytes(compactWrite)} bytes`);
    model.numberFormat[0][0] = "0.00";
    assert.deepEqual((await call("read_range_format", range)).data.format.numberFormat, model.numberFormat);
  });
});

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { randomUUID } = require("node:crypto");
const { WordSession, validateText } = require("../../app/office/addin/wordOperations");
const { wordFixture } = require("../fixtures/wordApi");

function fixture(texts) {
  const { word, model } = wordFixture(texts);
  const documentId = randomUUID();
  const state = { now: 1000, url: "https://tenant.sharepoint.com/document.docx" };
  const session = new WordSession(word, { documentId, randomUUID, now: () => state.now, getDocumentUrl: () => state.url });
  const call = (operation, args = {}, overrides = {}) => session.execute({ documentId, deadline: state.now + 10000, operation, args, ...overrides });
  const read = async () => (await call("read_paragraphs")).data.paragraphs[0];
  return { session, model, state, call, read };
}

describe("Word live session", () => {
  it("keeps tokens from earlier paragraph and structure pages until used or expired", async () => {
    for (const operation of ["read_paragraphs", "read_structure"]) {
      const { call, session, state } = fixture(Array.from({ length: 100 }, (_, i) => "Paragraph " + i));
      const first = (await call(operation, { count: 50 })).data.paragraphs;
      await call(operation, { start: 50, count: 50 });
      const result = await call("replace_text", { ...first[10], findText: "Paragraph", replacementText: "Updated" });
      assert.equal(result.ok, true);
      assert.equal(result.data.text, "Updated 10");
      assert.equal((await call("replace_text", { ...first[10], findText: "Updated", replacementText: "Again" })).code, "STALE_PARAGRAPH");
      state.now += 120000;
      await call("read_paragraphs", { count: 1 });
      assert.equal(session.snapshots.size, 1, "Expired snapshots are pruned before issuing new tokens");
    }
  });

  it("pages main-body paragraphs and returns scoped snapshots without loading the whole document", async () => {
    const { call, model, session } = fixture(["One", "Two", "Three"]);
    const first = await call("read_paragraphs", { count: 2 });
    assert.deepEqual(first.data.paragraphs.map((p) => p.text), ["One", "Two"]);
    assert.equal(first.data.nextStart, 2);
    assert.deepEqual(model.lastPage, { select: "uniqueLocalId,text", skip: 0, top: 3 });
    const last = await call("read_paragraphs", { start: 2, count: 2 });
    assert.equal(last.data.nextStart, null);
    const paragraph = last.data.paragraphs[0];
    assert.equal(paragraph.text, "Three");
    paragraph.text = "Caller mutation";
    assert.equal(session.snapshots.get(paragraph.readToken).text, "Three");
    assert.deepEqual((await call("read_paragraphs", { start: 100 })).data, { paragraphs: [], nextStart: null });
  });

  it("replaces one literal match, keeps surrounding content, and returns a fresh snapshot", async () => {
    const { call, read, model } = fixture();
    const before = await read();
    const result = await call("replace_text", { ...before, findText: "Hello", replacementText: "Greetings" });
    assert.equal(result.ok, true);
    assert.equal(result.data.applied, true);
    assert.equal(result.data.text, "Greetings from Word.");
    assert.equal(model.paragraphs[1].text, "A second paragraph.");
    assert.deepEqual(model.paragraphs[0].formatting, { outsideMatch: "preserved" });
    assert.equal(model.searchOptions.matchWildcards, false);
    assert.equal(model.searchOptions.matchCase, true);
    assert.notEqual(result.data.readToken, before.readToken);
    assert.equal((await call("replace_text", { ...before, findText: "Greetings", replacementText: "Again" })).code, "STALE_PARAGRAPH");
    assert.equal(model.writes, 1);
    const deleted = await call("replace_text", { ...result.data, findText: "Greetings ", replacementText: "" });
    assert.equal(deleted.data.text, "from Word.");
  });

  it("keeps paragraph identity when paragraphs are inserted before the target", async () => {
    const { call, read, model } = fixture();
    const before = await read();
    model.paragraphs.unshift({ id: randomUUID(), text: "New paragraph" });
    const result = await call("replace_text", { ...before, findText: "Hello", replacementText: "Updated" });
    assert.equal(result.data.paragraphId, before.paragraphId);
    assert.equal(model.paragraphs[0].text, "New paragraph");
    assert.equal(model.paragraphs[1].text, "Updated from Word.");
  });

  it("rejects coauthor changes both before the read and during the search", async () => {
    for (const stage of ["before", "search"]) {
      const { call, read, model } = fixture();
      const before = await read();
      if (stage === "before") model.paragraphs[0].text = "Coauthor changed this";
      else model.beforeSync = (sync) => { if (sync === 3) model.paragraphs[0].text = "Hello, coauthor"; };
      const result = await call("replace_text", { ...before, findText: "Hello", replacementText: "Changed" });
      assert.equal(result.code, "STALE_PARAGRAPH", stage);
      assert.equal(model.writes, 0);
    }
  });

  it("rejects ambiguous, absent and nonliteral search results", async () => {
    for (const text of ["Hello Hello", "No matching text", "hello"]) {
      const { call, read, model } = fixture([text]);
      const before = await read();
      assert.equal((await call("replace_text", { ...before, findText: "Hello", replacementText: "Updated" })).code, "AMBIGUOUS_TEXT");
      assert.equal(model.writes, 0);
    }
    const { call, read, model } = fixture();
    const before = await read();
    model.searchOverride = [{ text: "HELLO" }];
    assert.equal((await call("replace_text", { ...before, findText: "Hello", replacementText: "Updated" })).code, "AMBIGUOUS_TEXT");
    assert.equal(model.writes, 0);
  });

  it("binds tokens to their paragraph, document and expiration", async () => {
    for (const change of ["paragraph", "document", "url", "expiry"]) {
      const { call, read, model, state } = fixture();
      const before = await read();
      const overrides = {};
      if (change === "paragraph") before.paragraphId = model.paragraphs[1].id;
      if (change === "document") overrides.documentId = randomUUID();
      if (change === "url") state.url += "?other-document";
      if (change === "expiry") state.now += 120000;
      assert.equal((await call("replace_text", { ...before, findText: "Hello", replacementText: "Updated" }, overrides)).ok, false, change);
      assert.equal(model.writes, 0);
    }
  });

  it("handles deleted paragraphs with a static error and never retargets by position", async () => {
    const { call, read, model } = fixture();
    const before = await read();
    model.paragraphs.shift();
    assert.deepEqual(await call("read_paragraph", before), { ok: false, code: "WRONG_PARAGRAPH" });
    assert.equal((await call("replace_text", { ...before, findText: "Hello", replacementText: "Updated" })).code, "WRONG_PARAGRAPH");
    assert.equal(model.writes, 0);
  });

  it("honors deadlines and disconnects after the fresh read and during search", async () => {
    for (const sync of [2, 3]) {
      for (const action of ["timeout", "disconnect"]) {
        const { call, read, model, state, session } = fixture();
        const before = await read();
        model.beforeSync = (n) => {
          if (n === sync) {
            if (action === "timeout") state.now += 10001;
            else session.deactivate();
          }
        };
        assert.equal((await call("replace_text", { ...before, findText: "Hello", replacementText: "Updated" })).code,
          action === "timeout" ? "TIMEOUT" : "WRONG_DOCUMENT");
        assert.equal(model.writes, 0);
      }
    }
  });

  it("never retries uncertain writes, consumes the token and supports explicit readback", async () => {
    const { call, read, model } = fixture();
    const before = await read();
    model.beforeSync = (sync) => { if (sync === 5) throw new Error("private content"); };
    assert.deepEqual(await call("replace_text", { ...before, findText: "Hello", replacementText: "Updated" }), { ok: false, code: "WORD_ERROR" });
    assert.equal(model.writes, 1);
    assert.equal((await call("replace_text", { ...before, findText: "Updated", replacementText: "Again" })).code, "STALE_PARAGRAPH");
    model.beforeSync = null;
    assert.equal((await call("read_paragraph", before)).data.text, "Updated from Word.");
  });

  it("rejects overlapping requests", async () => {
    const { call, model } = fixture();
    let release;
    model.beforeSync = () => new Promise((resolve) => { release = resolve; });
    const pending = call("read_paragraphs");
    assert.equal((await call("read_paragraphs")).code, "BUSY");
    release();
    assert.equal((await pending).ok, true);
  });

  it("bounds pages, UTF-8 results, replacement sizes and special search syntax", async () => {
    const { call } = fixture(["漢".repeat(12000)]);
    assert.equal((await call("read_paragraphs")).code, "TOO_LARGE");
    for (const args of [{ start: -1 }, { start: 100001 }, { count: 0 }, { count: 51 }, { count: 1.5 }]) {
      assert.equal((await call("read_paragraphs", args)).code, "INVALID_PAGE");
    }
    for (const text of ["", "^p", "line\nbreak", "\u0007", "x".repeat(256)]) assert.throws(() => validateText(text, "replacement"));
    for (const replacement of ["line\rbreak", "\u2029", "x".repeat(8193)]) assert.throws(() => validateText("find", replacement));
    assert.doesNotThrow(() => validateText("literal.*[text]", ""));
  });
});

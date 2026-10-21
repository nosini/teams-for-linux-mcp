const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const vm = require("node:vm");
const { OfficeInspector, classifyOfficeUrl, PROBE_SOURCE } = require("../../app/office/inspector");

function frame(url, executeJavaScript = async () => ({})) {
  return { url, processId: 11, routingId: 22, executeJavaScript };
}

function inspectorFor(frames, options) {
  return new OfficeInspector(() => [{ id: 7, isDestroyed: () => false, mainFrame: { framesInSubtree: frames } }], options);
}

describe("Office editor discovery", () => {
  it("recognizes Office subframes and SharePoint wrappers without exposing document URLs", () => {
    const inspector = inspectorFor([
      frame("https://teams.cloud.microsoft/"),
      frame("https://weu-excel.officeapps.live.com/x/edit?access_token=secret#private"),
      frame("https://tenant.sharepoint.com/sites/team/_layouts/15/Doc.aspx?sourcedoc=private"),
    ]);
    const result = inspector.listEditors();
    assert.equal(result.editors.length, 2);
    assert.deepEqual(result.editors[0], {
      id: "7:11:22", webContentsId: 7, application: "excel", surface: "office-frame",
      origin: "https://weu-excel.officeapps.live.com",
    });
    assert.equal(result.editors[1].surface, "document-wrapper");
    assert.doesNotMatch(JSON.stringify(result), /private|secret|access_token|sourcedoc/);
  });

  it("rejects lookalike domains, insecure pages and arbitrary SharePoint pages", () => {
    for (const url of [
      "https://excel.officeapps.live.com.evil.test/", "https://evilofficeapps.live.com/",
      "http://excel.officeapps.live.com/", "file:///excel.officeapps.live.com",
      "https://tenant.sharepoint.com/sites/team/", "https://word.cloud.microsoft.evil.test/",
      "https://user:password@word.cloud.microsoft/", "not a url",
    ]) assert.equal(classifyOfficeUrl(url), null, url);
    assert.equal(classifyOfficeUrl("https://word.cloud.microsoft/").application, "word");
  });

  it("does not execute scripts during discovery or in an unrecognized frame", async () => {
    let executions = 0;
    const inspector = inspectorFor([frame("https://evil.test/", async () => { executions++; })]);
    assert.deepEqual(inspector.listEditors(), { editors: [], truncated: false });
    await assert.rejects(inspector.inspectEditor("7:11:22"), /no longer available/);
    assert.equal(executions, 0);
  });

  it("bounds results and tolerates closed windows and detached subframes", () => {
    const detached = { get url() { throw new Error("detached"); } };
    const inspector = inspectorFor([detached, frame("https://word.cloud.microsoft/"), frame("https://excel.cloud.microsoft/")], { maxFrames: 1 });
    assert.equal(inspector.listEditors().editors.length, 1);
    assert.equal(inspector.listEditors().truncated, true);
    const closed = new OfficeInspector(() => [{ isDestroyed: () => true }]);
    assert.deepEqual(closed.listEditors().editors, []);
  });

  it("runs a fixed read-only probe and sanitizes renderer-controlled results", async () => {
    const target = frame("https://excel.cloud.microsoft/", async (source, userGesture) => {
      assert.equal(source, PROBE_SOURCE);
      assert.equal(userGesture, false);
      return { officeGlobal: true, excelRun: "yes", privateDocumentText: "secret" };
    });
    const result = await inspectorFor([target]).inspectEditor("7:11:22");
    assert.equal(result.apiPresence.officeGlobal, true);
    assert.equal(result.apiPresence.excelRun, false);
    assert.equal(result.bridgeStatus, "not-connected");
    assert.doesNotMatch(JSON.stringify(result), /secret/);
  });

  it("bounds hanging probes and does not forward renderer exceptions", async () => {
    for (const execute of [() => new Promise(() => {}), async () => { throw new Error("private URL"); }]) {
      const result = await inspectorFor([frame("https://word.cloud.microsoft/", execute)], { timeoutMs: 10 }).inspectEditor("7:11:22");
      assert.equal(result.probeStatus, "unavailable");
      assert.doesNotMatch(JSON.stringify(result), /private URL/);
    }
  });

  it("discards capability results after a navigation away from Office", async () => {
    const target = frame("https://excel.cloud.microsoft/", async () => {
      target.url = "https://evil.test/";
      return { officeGlobal: true };
    });
    assert.equal((await inspectorFor([target]).inspectEditor("7:11:22")).probeStatus, "unavailable");
  });

  it("reports missing Office.js as absence and does not try to initialize it", () => {
    const empty = vm.runInNewContext(PROBE_SOURCE, {});
    assert.ok(Object.values(empty).every((value) => value === false));
    const available = vm.runInNewContext(PROBE_SOURCE, {
      Office: {
        onReady() { throw new Error("must not initialize Office"); },
        context: { document: {}, requirements: { isSetSupported: (name) => name === "ExcelApi" } },
      },
      Excel: { run() { throw new Error("must not read the workbook"); } },
    });
    assert.equal(available.documentContext, true);
    assert.equal(available.excelRun, true);
    assert.equal(available.excelApi11, true);
    assert.equal(available.wordApi11, false);
  });
});

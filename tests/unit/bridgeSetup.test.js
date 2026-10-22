const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { runBridgeSetup, setupFrame, inspectBridgeSetup, executeBridgeSetupStep } = require("../../app/office/bridgeSetup");

function fixture() {
  const editor = { url: "https://euc-excel.officeapps.live.com/x/edit?token=private", executeJavaScript: async () => ({ actions: [] }) };
  const contents = { id: 9, isDestroyed: () => false, mainFrame: { framesInSubtree: [editor] } };
  const controller = new AbortController();
  const updates = [];
  const options = { contents, manifests: { excel: "EXCEL MANIFEST", word: "WORD MANIFEST" }, signal: controller.signal, listDocuments: async () => ({ documents: [] }), update: (state) => updates.push(state) };
  return { editor, contents, controller, updates, options };
}
const timing = { timeoutMs: 100, pollMs: 1, interaction: "pointer" };

describe("Office Bridge import job", () => {
  it("follows Add-ins, Advanced, Upload My Add-in, manifest selection and Upload in order", async () => {
    const { editor, options } = fixture();
    const actions = ["addins", "advanced", "uploadEntry", "selectManifest", "submit"];
    const performed = [];
    options.listDocuments = async () => ({ documents: performed.length === actions.length ? [{ webContentsId: 9 }] : [] });
    editor.executeJavaScript = async (source) => {
      const next = actions[performed.length];
      if (source.includes('"type":"probe"')) return { actions: next === "selectManifest" ? [] : [next], uploadInput: ["selectManifest", "submit"].includes(next), expectedFile: next === "submit" };
      if (next === "selectManifest") assert.ok(source.includes('"type":"fill"') && source.includes('"manifest":"EXCEL MANIFEST"'));
      else assert.ok(source.includes(`"action":"${next}"`));
      performed.push(next);
      return { performed: true };
    };
    const result = await runBridgeSetup(options, { timeoutMs: 200, interaction: "pointer" });
    assert.equal(result.status, "connected");
    assert.ok(result.elapsedMs < 200, "Ready controls advance without per-step delays");
    assert.deepEqual(performed, actions);
  });

  it("prepares a collapsed ribbon before sending Add-ins input", async () => {
    const { editor, options } = fixture();
    const performed = [];
    let probes = 0;
    options.listDocuments = async () => ({ documents: performed.includes("addins") ? [{ webContentsId: 9 }] : [] });
    editor.executeJavaScript = async (source) => {
      if (source.includes('"type":"probe"')) return { actions: ["home", "addins"], homeRequired: ++probes < 4 };
      performed.push(source.includes('"action":"home"') ? "home" : "addins");
      return { performed: true };
    };
    assert.equal((await runBridgeSetup(options, timing)).status, "connected");
    assert.deepEqual(performed, ["home", "addins"]);
    assert.equal(probes, 4, "Wait for Home's delayed expansion without toggling it again");
  });

  it("waits for Upload to enable without touching the ribbon behind its dialog", async () => {
    const { editor, options } = fixture();
    let probes = 0, filled = false, uploaded = false;
    options.listDocuments = async () => ({ documents: uploaded ? [{ webContentsId: 9 }] : [] });
    editor.executeJavaScript = async (source) => {
      if (source.includes('"type":"probe"')) return { uploadInput: true, expectedFile: filled,
        actions: ++probes >= 4 ? ["submit", "addins", "home"] : ["addins", "home"] };
      if (source.includes('"type":"fill"')) filled = true;
      else { assert.ok(source.includes('"action":"submit"'), "Do not act outside the pending upload dialog"); uploaded = true; }
      return { performed: true };
    };
    assert.equal((await runBridgeSetup(options, timing)).status, "connected");
    assert.equal(probes, 4);
  });

  it("grants gestures only to named UI actions and keeps single-step probes passive", async () => {
    const { editor, contents, options } = fixture();
    const gestures = [];
    editor.executeJavaScript = async (source, gesture) => {
      gestures.push(gesture);
      if (source.includes('"type":"probe"')) return { actions: ["addins"] };
      return { performed: true };
    };
    assert.deepEqual(await executeBridgeSetupStep(contents, options.manifests, "addins"), { performed: true });
    assert.deepEqual(gestures, [false, true]);
    await assert.rejects(executeBridgeSetupStep(contents, options.manifests, "deleteDocument"), /supported/);
  });

  it("waits inside the add-in flow instead of switching ribbon tabs while its menu loads", async () => {
    const { editor, options, updates } = fixture();
    editor.executeJavaScript = async (source) => source.includes('"type":"probe"') ? { actions: ["addins", "home", "insert"] } : { performed: true };
    const result = await runBridgeSetup(options, { ...timing, timeoutMs: 20 });
    assert.equal(result.stage, "opening_addins");
    assert.equal(result.code, "IMPORT_UI_NOT_FOUND");
    assert.equal(updates.some((state) => ["opening_home", "opening_insert"].includes(state.stage)), false);
  });

  it("expands Home once when a collapsed ribbon obscures Add-ins, then revalidates", async () => {
    const { editor, options } = fixture();
    const performed = [];
    editor.executeJavaScript = async (source) => {
      if (source.includes('"type":"probe"')) return { actions: ["addins", "home"] };
      const action = source.includes('"action":"home"') ? "home" : "addins";
      performed.push(action);
      return action === "addins" ? { performed: false, code: "CONTROL_OBSCURED" } : { performed: true };
    };
    const result = await runBridgeSetup(options, timing);
    assert.equal(result.code, "CONTROL_OBSCURED");
    assert.deepEqual(performed, ["addins", "home", "addins"], "Persistent overlays still stop setup after one Home attempt");
  });

  it("bounds UI diagnostics, sanitizes URLs and excludes unrelated frame contents", async () => {
    const { editor, contents } = fixture();
    editor.executeJavaScript = async () => ({ controls: Array.from({ length: 150 }, () => ({ tag: "button", label: "Add-ins https://private/?token=secret", enabled: true })), html: "private document", actions: ["home"] });
    const foreign = { url: "https://evil.test/private?token=secret", parent: editor, executeJavaScript: () => assert.fail("Untrusted frames must not be inspected") };
    const store = { url: "https://store.office.com/import?secret=token", parent: editor, executeJavaScript: editor.executeJavaScript };
    contents.mainFrame.framesInSubtree = [editor, store, foreign];
    const result = await inspectBridgeSetup(contents);
    assert.equal(result.frames.reduce((count, frame) => count + (frame.controls?.length || 0), 0), 240);
    assert.equal(result.truncated, true);
    assert.equal(result.frames[2].eligible, false);
    assert.equal(result.frames[0].controls[0].label, "Add-ins [URL]");
    assert.doesNotMatch(JSON.stringify(result), /token|secret|private|document|html/);
  });

  it("reports only bounded hit-test metadata without blocker text, identifiers or paths", async () => {
    const { editor, contents } = fixture();
    editor.executeJavaScript = async () => ({ controls: [], actions: ["addins"], actionability: Array.from({ length: 50 }, () => ({
      action: "addins", relation: "other", hitTag: "div", hitRole: "dialog", pointerEvents: "none", width: 36, height: 24,
      clipped: true, text: "private cell value", id: "secret identifier", url: "https://private/", html: "private document",
    })) });
    contents.mainFrame.framesInSubtree = [editor, { url: "https://store.office.com/import", parent: editor, executeJavaScript: editor.executeJavaScript },
      { url: "https://res.public.onecdn.static.microsoft/addins", parent: editor, executeJavaScript: editor.executeJavaScript }];
    const result = await inspectBridgeSetup(contents);
    assert.equal(result.frames.reduce((count, frame) => count + frame.actionability.length, 0), 40);
    assert.equal(result.truncated, true);
    assert.deepEqual(result.frames[0].actionability[0], { action: "addins", relation: "other", centreRelation: "unknown", hitTag: "div", hitRole: "dialog", pointerEvents: "none", width: 36, height: 24, clipped: true });
    assert.doesNotMatch(JSON.stringify(result), /private|secret|identifier|html/);
  });

  it("restricts import controls to the selected editor and trusted descendant frames", () => {
    const { editor } = fixture();
    assert.equal(setupFrame(editor, editor), true);
    assert.equal(setupFrame({ url: "https://store.office.com/import", parent: editor }, editor), true);
    assert.equal(setupFrame({ url: "https://res.public.onecdn.static.microsoft/addins", parent: editor }, editor), true);
    assert.equal(setupFrame({ url: "https://res.public.onecdn.static.microsoft/addins" }, editor), false);
    assert.equal(setupFrame({ url: "https://euc-excel.officeapps.live.com/other" }, editor), false);
    for (const url of ["https://store.office.com.evil.test/", "https://res.public.onecdn.static.microsoft.evil.test/", "https://other.onecdn.static.microsoft/", "https://res.public.onecdn.static.microsoft:444/", "https://user:password@res.public.onecdn.static.microsoft/", "https://tenant.sharepoint.com/upload", "https://evil.test/", "https://store.office.com:444/", "https://user:password@store.office.com/", "file:///tmp/upload.html"]) {
      assert.equal(setupFrame({ url, parent: editor }, editor), false);
    }
  });

  it("recognizes an existing live Bridge without clicking or uploading", async () => {
    const { editor, options } = fixture();
    editor.executeJavaScript = () => assert.fail("Existing connections need no UI automation");
    options.listDocuments = async () => ({ documents: [{ webContentsId: 9 }] });
    const result = await runBridgeSetup(options, timing);
    assert.ok(Number.isInteger(result.elapsedMs) && result.elapsedMs >= 0);
    assert.deepEqual(result, { status: "connected", stage: "connected", submitted: false, elapsedMs: result.elapsedMs });
  });

  it("reports sign-in without touching authentication controls", async () => {
    const { contents, options } = fixture();
    contents.mainFrame.framesInSubtree = [{ url: "https://login.microsoftonline.com/common/authorize" }];
    assert.equal((await runBridgeSetup(options, timing)).code, "SIGN_IN_REQUIRED");
  });

  it("does not select between multiple documents in one tab", async () => {
    const { contents, options } = fixture();
    contents.mainFrame.framesInSubtree.push({ url: "https://euc-word.officeapps.live.com/w/edit" });
    assert.equal((await runBridgeSetup(options, timing)).code, "AMBIGUOUS_EDITOR");
  });

  it("never retries an uncertain submission and sanitizes renderer errors", async () => {
    const { editor, options, updates } = fixture();
    let clicks = 0;
    editor.executeJavaScript = async (source) => {
      if (source.includes('"type":"fill"')) return { performed: true };
      if (source.includes('"type":"click"')) { clicks++; throw new Error("secret file contents https://private/?token=secret"); }
      return { actions: ["submit"], uploadInput: true, expectedFile: true };
    };
    const result = await runBridgeSetup(options, timing);
    assert.equal(result.code, "UPLOAD_OUTCOME_UNKNOWN");
    assert.equal(result.submitted, true);
    assert.equal(clicks, 1);
    assert.ok(updates.some((state) => state.submitted));
    assert.doesNotMatch(JSON.stringify([result, updates]), /secret|private|token/);
  });

  it("stops if the editor navigates during setup", async () => {
    const { editor, options } = fixture();
    editor.executeJavaScript = async () => { editor.url = "https://euc-excel.officeapps.live.com/x/different"; return { actions: [] }; };
    assert.equal((await runBridgeSetup(options, timing)).code, "DOCUMENT_CHANGED");
  });

  it("waits for an editor but can cancel immediately on tab shutdown", async () => {
    const { contents, options, controller } = fixture();
    contents.mainFrame.framesInSubtree = [];
    options.listDocuments = async () => { controller.abort(); return { documents: [] }; };
    const result = await runBridgeSetup(options, { ...timing, timeoutMs: 45000, pollMs: 400 });
    assert.equal(result.code, "CANCELLED");
    assert.equal(result.submitted, false);
  });

  it("reports unsupported UI without touching unrelated page controls", async () => {
    const { options } = fixture();
    const result = await runBridgeSetup(options, { ...timing, timeoutMs: 10 });
    assert.equal(result.code, "IMPORT_UI_NOT_FOUND");
    assert.equal(result.submitted, false);
  });

  it("stops before submitting when the upload input changes", async () => {
    const { editor, options } = fixture();
    editor.executeJavaScript = async (source) => source.includes('"type":"fill"') ? { performed: false } : { actions: [], uploadInput: true, expectedFile: false };
    const result = await runBridgeSetup(options, timing);
    assert.equal(result.code, "UPLOAD_INPUT_CHANGED");
    assert.equal(result.submitted, false);
  });
});

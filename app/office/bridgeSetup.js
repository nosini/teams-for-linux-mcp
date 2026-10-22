const { classifyOfficeUrl } = require("./inspector");
const { evaluateInFrame } = require("../mcp/evaluateInFrame");
const { bridgeSetupUiScript } = require("./bridgeSetupUiLoader");
const { nativeSetupInput } = require("./nativeSetupInput");
const { setupMousePoint } = require("./setupMousePoint");

function activateSetupControl(contents, frame, command, raw, isCurrent, interaction) {
  const mouse = interaction === "nativeMouse" || (interaction === "auto" && command.action === "addins");
  return nativeSetupInput(contents, () => raw({ ...command, type: mouse ? "point" : "focus" }), isCurrent,
    mouse ? (local) => setupMousePoint(contents, frame, local, isCurrent) : undefined, frame);
}

function setupFrame(frame, editor) {
  try {
    const url = new URL(frame.url);
    if (frame.detached || url.protocol !== "https:" || url.username || url.password || url.port ||
      !(url.hostname.endsWith(".officeapps.live.com") || ["excel.cloud.microsoft", "word.cloud.microsoft", "store.office.com", "res.public.onecdn.static.microsoft"].includes(url.hostname))) return false;
    for (let current = frame, depth = 0; current && depth < 24; current = current.parent, depth++) if (current === editor) return true;
  } catch { /* Detached frame. */ }
  return false;
}

const setupActions = ["home", "insert", "addins", "getAddins", "advanced", "moreSettings", "moreAddins", "manage", "myAddins", "uploadEntry", "bridge", "selectManifest", "submit", "trust"];

async function executeBridgeSetupStep(contents, manifests, action, interaction = "pointer") {
  if (!setupActions.includes(action) || !["click", "pointer", "native", "nativeMouse", "auto"].includes(interaction)) throw new Error("Choose a supported Bridge setup action and interaction.");
  const all = contents.mainFrame.framesInSubtree;
  const editors = all.filter((frame) => ["excel", "word"].includes(classifyOfficeUrl(frame.url)?.application));
  if (editors.length !== 1) throw new Error("Open one Excel or Word document in this tab before setting up the Bridge.");
  const editor = editors[0], editorUrl = editor.url, application = classifyOfficeUrl(editorUrl).application;
  if (!manifests?.[application]) throw new Error("Bridge setup is unavailable in this run.");
  const evaluate = (frame, command) => {
    if (contents.isDestroyed() || editor.detached || editor.url !== editorUrl || !setupFrame(frame, editor)) throw new Error("Document changed during setup.");
    const raw = (request) => evaluateInFrame(frame, bridgeSetupUiScript({ ...request, url: frame.url, deadline: Date.now() + 1500, filename: `${application}-addin.xml`, allowUploadMenu: true, interaction }), 2000, "Setup step timed out", !["probe", "focus", "point"].includes(request.type));
    if (command.type === "click" && ["native", "nativeMouse", "auto"].includes(interaction)) return activateSetupControl(contents, frame, command, raw, () => {
      try { return !contents.isDestroyed() && !editor.detached && editor.url === editorUrl && setupFrame(frame, editor); } catch { return false; }
    }, interaction);
    return raw(command);
  };
  try {
    const probes = await Promise.all(all.filter((frame) => setupFrame(frame, editor)).slice(0, 16).map(async (frame) => ({ frame, probe: await evaluate(frame, { type: "probe" }) })));
    const matches = probes.filter(({ probe }) => action === "selectManifest" ? probe.uploadInput : probe.actions?.includes(action));
    if (matches.length !== 1) return { performed: false, code: matches.length ? "AMBIGUOUS_CONTROL" : "CONTROL_NOT_FOUND" };
    const result = await evaluate(matches[0].frame, { type: action === "selectManifest" ? "fill" : "click", action, manifest: ["selectManifest", "submit"].includes(action) ? manifests[application] : undefined });
    return { performed: result?.performed === true, ...(result?.performed ? {} : { code: result?.code || "CONTROL_CHANGED" }) };
  } catch { return { performed: false, code: action === "submit" ? "UPLOAD_OUTCOME_UNKNOWN" : "SETUP_UI_UNAVAILABLE" }; }
}

async function inspectBridgeSetup(contents) {
  const all = contents.mainFrame.framesInSubtree;
  const editors = all.filter((frame) => ["excel", "word"].includes(classifyOfficeUrl(frame.url)?.application));
  if (editors.length !== 1) throw new Error("Open one Excel or Word document in this tab before inspecting Bridge setup.");
  const editor = editors[0];
  let budget = 240, actionBudget = 40;
  const frames = await Promise.all(all.slice(0, 32).map(async (frame) => {
    let origin;
    try { origin = new URL(frame.url).origin; } catch { origin = "unavailable"; }
    const eligible = setupFrame(frame, editor);
    const metadata = { origin, eligible };
    if (!eligible) return metadata;
    try {
      const probe = await evaluateInFrame(frame, bridgeSetupUiScript({ type: "inspect", url: frame.url, deadline: Date.now() + 1500 }), 2000, "Setup inspection timed out");
      if (!Array.isArray(probe?.controls)) return { ...metadata, status: "unavailable" };
      return { ...metadata, status: "completed", truncated: probe.truncated === true || probe.controls.length > 120,
        actions: Array.isArray(probe.actions) ? probe.actions.filter((action) => typeof action === "string").map((action) => action.slice(0, 40)).slice(0, 20) : [],
        actionability: Array.isArray(probe.actionability) ? probe.actionability.slice(0, 20).filter((item) => item && setupActions.includes(item.action)).map((item) => ({
          action: item.action,
          relation: ["target", "descendant", "ancestor", "other", "none", "no_rect"].includes(item.relation) ? item.relation : "unknown",
          centreRelation: ["target", "descendant", "ancestor", "other", "none", "no_rect"].includes(item.centreRelation) ? item.centreRelation : "unknown",
          hitTag: typeof item.hitTag === "string" && /^[a-z][a-z0-9-]{0,23}$/.test(item.hitTag) ? item.hitTag : null,
          hitRole: typeof item.hitRole === "string" && /^[a-z ]{1,40}$/.test(item.hitRole) ? item.hitRole : null,
          pointerEvents: item.pointerEvents === "none" ? "none" : "auto",
          width: Number.isFinite(item.width) && item.width >= 0 && item.width <= 100000 ? item.width : null,
          height: Number.isFinite(item.height) && item.height >= 0 && item.height <= 100000 ? item.height : null,
          clipped: item.clipped === true,
        })) : [],
        controls: probe.controls.slice(0, 120).filter((control) => control && typeof control.label === "string").map((control) => ({
          tag: typeof control.tag === "string" ? control.tag.slice(0, 24) : "unknown",
          role: typeof control.role === "string" ? control.role.slice(0, 40) : null,
          label: control.label.replace(/https?:\/\/\S+/g, "[URL]").slice(0, 160), enabled: control.enabled === true,
        })) };
    } catch { return { ...metadata, status: "unavailable" }; }
  }));
  // Apply a single total budget in frame order after parallel probing.
  for (const frame of frames) {
    if (!frame.controls) continue;
    if (frame.actionability.length > actionBudget) frame.truncated = true;
    frame.actionability = frame.actionability.slice(0, actionBudget);
    actionBudget -= frame.actionability.length;
    if (frame.controls.length > budget) frame.truncated = true;
    frame.controls = frame.controls.slice(0, budget);
    budget -= frame.controls.length;
  }
  return { application: classifyOfficeUrl(editor.url)?.application, frames, truncated: all.length > 32 || frames.some((frame) => frame.truncated) };
}

// Bounded background job. Progress contains only fixed stage/error names; no
// document URLs, renderer errors, UI text, file paths or manifest contents.
async function runBridgeSetup({ contents, manifests, listDocuments, update, signal }, { timeoutMs = 45000, pollMs = 400, settleMs = 800, interaction = "auto" } = {}) {
  const deadline = Date.now() + timeoutMs;
  let editor, editorUrl, application, selectedFrame, submitted = false, stage = "waiting_for_editor";
  const clicked = new Set();
  const result = (status, code) => ({ status, stage, submitted, ...(code ? { code } : {}) });
  const report = () => update(result("running"));
  const wait = (ms) => new Promise((resolve) => {
    const done = () => { clearTimeout(timer); signal.removeEventListener("abort", done); resolve(); };
    const timer = setTimeout(done, ms);
    signal.addEventListener("abort", done, { once: true });
    if (signal.aborted) done();
  });
  const evaluate = async (frame, command) => {
    if (signal.aborted || Date.now() >= deadline || !setupFrame(frame, editor) || editor.detached || editor.url !== editorUrl) throw new Error("stopped");
    const raw = (request) => evaluateInFrame(frame, bridgeSetupUiScript({ ...request, url: frame.url, deadline: Math.min(deadline, Date.now() + 1500), filename: `${application}-addin.xml`, allowUploadMenu: clicked.has("manage"), interaction }), Math.min(2000, Math.max(1, deadline - Date.now())), "Setup UI timed out", !["probe", "focus", "point"].includes(request.type));
    if (command.type === "click" && ["native", "nativeMouse", "auto"].includes(interaction)) return activateSetupControl(contents, frame, command, raw, () => {
      try { return !signal.aborted && Date.now() < deadline && !contents.isDestroyed() && !editor.detached && editor.url === editorUrl && setupFrame(frame, editor); } catch { return false; }
    }, interaction);
    return raw(command);
  };
  report();
  try {
    while (!signal.aborted && !contents.isDestroyed() && Date.now() < deadline) {
      const documents = (await listDocuments()).documents.filter((document) => document.webContentsId === contents.id);
      if (signal.aborted || contents.isDestroyed()) return result("needs_attention", "CANCELLED");
      if (documents.length) { stage = submitted ? "connected_after_upload" : "connected"; return result("connected"); }
      const all = contents.mainFrame.framesInSubtree;
      if (all.some((frame) => { try { return ["login.microsoftonline.com", "login.microsoft.com", "login.live.com"].includes(new URL(frame.url).hostname); } catch { return false; } }) && !all.some((frame) => ["excel", "word"].includes(classifyOfficeUrl(frame.url)?.application))) return result("needs_attention", "SIGN_IN_REQUIRED");
      if (!editor) {
        const editors = all.filter((frame) => ["excel", "word"].includes(classifyOfficeUrl(frame.url)?.application));
        if (editors.length > 1) return result("needs_attention", "AMBIGUOUS_EDITOR");
        if (!editors.length) { await wait(pollMs); continue; }
        editor = editors[0]; editorUrl = editor.url; application = classifyOfficeUrl(editorUrl).application;
        if (!manifests?.[application]) return result("failed", "BRIDGE_UNAVAILABLE");
      }
      if (editor.detached || editor.url !== editorUrl || !all.includes(editor)) return result("needs_attention", "DOCUMENT_CHANGED");
      const frames = all.filter((frame) => setupFrame(frame, editor)).slice(0, 16);
      const probes = await Promise.all(frames.map(async (frame) => ({ frame, probe: await evaluate(frame, { type: "probe" }) })));
      const file = probes.filter(({ probe }) => probe.uploadInput);
      if (file.length > 1) return result("needs_attention", "AMBIGUOUS_UPLOAD");
      if (file.length && !submitted) {
        const { frame, probe } = file[0];
        if (!selectedFrame) {
          stage = "selecting_manifest"; report();
          const filled = await evaluate(frame, { type: "fill", manifest: manifests[application] });
          if (!filled.performed) return result("needs_attention", "UPLOAD_INPUT_CHANGED");
          selectedFrame = frame;
          await wait(settleMs); continue;
        }
        if (selectedFrame !== frame || !probe.expectedFile) return result("needs_attention", "UPLOAD_INPUT_CHANGED");
        if (probe.actions.includes("submit")) {
          // Mark BEFORE dispatch; an uncertain click must never be retried.
          submitted = true; stage = "waiting_for_bridge"; report();
          const uploaded = await evaluate(frame, { type: "click", action: "submit", manifest: manifests[application] });
          if (!uploaded.performed) return result("needs_attention", uploaded.code || "UPLOAD_NOT_SUBMITTED");
          await wait(settleMs); continue;
        }
      }
      const priorities = submitted ? ["trust"] : ["trust", "bridge", "uploadEntry", "advanced", "manage", "moreSettings", "moreAddins", "myAddins", "getAddins", "addins", "home", "insert"];
      let acted = false;
      for (const action of priorities) {
        if (clicked.has(action)) continue;
        // Once inside the add-in flow, wait for its asynchronous UI. Switching
        // ribbon tabs after a short delay can close the menu before it loads.
        if (["home", "insert"].includes(action) && [...clicked].some((done) => !["home", "insert"].includes(done))) continue;
        const candidates = probes.filter(({ probe }) => probe.actions.includes(action));
        if (candidates.length > 1) return result("needs_attention", "AMBIGUOUS_CONTROL");
        if (!candidates.length) continue;
        clicked.add(action); stage = action === "bridge" ? "opening_existing_bridge" : `opening_${action}`; report();
        const clickedResult = await evaluate(candidates[0].frame, { type: "click", action });
        // A collapsed ribbon can leave Add-ins in the DOM underneath the
        // formula bar. Expand Home once, then hit-test Add-ins again. Do this
        // only before any add-in action has actually been dispatched.
        if (!clickedResult.performed && action === "addins" && clickedResult.code === "CONTROL_OBSCURED" &&
          [...clicked].every((done) => ["addins", "insert"].includes(done))) {
          const home = probes.filter(({ probe }) => probe.actions.includes("home"));
          if (home.length > 1) return result("needs_attention", "AMBIGUOUS_CONTROL");
          if (home.length === 1) {
            stage = "opening_home"; report();
            const expanded = await evaluate(home[0].frame, { type: "click", action: "home" });
            if (!expanded.performed) return result("needs_attention", expanded.code || "CONTROL_CHANGED");
            clicked.delete("addins"); clicked.add("home");
            acted = true; break;
          }
        }
        if (!clickedResult.performed) return result("needs_attention", clickedResult.code || "CONTROL_CHANGED");
        acted = true; break;
      }
      await wait(acted ? settleMs : pollMs);
    }
    return result("needs_attention", signal.aborted || contents.isDestroyed() ? "CANCELLED" : submitted ? "BRIDGE_NOT_CONNECTED" : "IMPORT_UI_NOT_FOUND");
  } catch {
    // Renderer failures may contain document content or URL parameters.
    return result("needs_attention", signal.aborted ? "CANCELLED" : submitted ? "UPLOAD_OUTCOME_UNKNOWN" : Date.now() >= deadline ? "IMPORT_UI_NOT_FOUND" : "SETUP_UI_UNAVAILABLE");
  }
}

module.exports = { runBridgeSetup, setupFrame, inspectBridgeSetup, executeBridgeSetupStep, setupActions };

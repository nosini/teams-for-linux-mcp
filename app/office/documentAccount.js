const { evaluateInFrame } = require("../mcp/evaluateInFrame");
const { createBridgeSetupUiLoader } = require("./bridgeSetupUiLoader");
const { nativeSetupInput } = require("./nativeSetupInput");

const accountUiScript = createBridgeSetupUiLoader({ filename: require.resolve("./documentAccountUi"), exportName: "documentAccountUi" });

// Read only this document site's remembered picker choice. This never calls
// Teams, a token provider, Graph, or cookie APIs. Only our own preference key
// on the login page is read; Microsoft authentication records are untouched.
async function getRememberedDocumentAccount(contents, scope, accountHint) {
  const frame = accountFrame(contents), url = frame?.url;
  if (!frame) return null;
  const isCurrent = () => accountFrame(contents) === frame && frame.url === url;
  try {
    const value = accountHint ?? await evaluateInFrame(frame, accountUiScript({ type: "remembered", url, scope,
      deadline: Date.now() + 1000 }), 1500, "Remembered account unavailable");
    if (!isCurrent() || typeof value !== "string" || value.length > 320 || !/^[^\s@]+@[^\s@]+$/.test(value) ||
      [...value].some((character) => character.charCodeAt(0) < 32 || character.charCodeAt(0) === 127)) return null;
    return { identifiers: [value.toLowerCase()], isCurrent };
  } catch { return null; }
}

function accountFrame(contents) {
  try {
    const frame = contents.mainFrame, url = new URL(frame.url);
    return !contents.isDestroyed() && url.protocol === "https:" && !url.username && !url.password && !url.port &&
      ["login.microsoftonline.com", "login.microsoft.com", "login.live.com"].includes(url.hostname) ? frame : null;
  } catch { return null; }
}

async function inspectDocumentAccount(contents, scope) {
  const frame = accountFrame(contents);
  if (!frame) return { picker: false, tileCount: 0, credentials: false };
  const url = frame.url;
  try {
    const result = await evaluateInFrame(frame, accountUiScript({ type: "probe", url, scope, useRememberedAccount: true, deadline: Date.now() + 1000 }), 1500, "Account picker unavailable");
    if (contents.isDestroyed() || contents.mainFrame !== frame || frame.url !== url) return { picker: false, tileCount: 0, credentials: false };
    return { picker: result?.picker === true, credentials: result?.credentials === true,
      tileCount: Number.isSafeInteger(result?.tileCount) ? Math.min(17, Math.max(0, result.tileCount)) : 0 };
  } catch { return { picker: false, tileCount: 0, credentials: false }; }
}

async function selectDocumentAccount(contents, getAccount, signal, deadline) {
  const frame = accountFrame(contents), url = frame?.url;
  if (!frame || signal.aborted || Date.now() >= deadline) return { performed: false, code: "ACCOUNT_PAGE_CHANGED" };
  const account = await getAccount();
  if (!account) return { performed: false, code: "DOCUMENT_ACCOUNT_NOT_REMEMBERED" };
  const isCurrent = () => !signal.aborted && Date.now() < deadline && frame && accountFrame(contents) === frame && frame.url === url && account.isCurrent();
  if (!isCurrent()) return { performed: false, code: "ACCOUNT_PAGE_CHANGED" };
  try {
    return await nativeSetupInput(contents, () => evaluateInFrame(frame, accountUiScript({ type: "focus", url,
      deadline: Math.min(deadline, Date.now() + 1000), identifiers: account.identifiers }), 1500, "Account picker unavailable"), isCurrent, undefined, frame);
  } catch { return { performed: false, code: "ACCOUNT_SELECTION_UNKNOWN" }; }
}

module.exports = { getRememberedDocumentAccount, inspectDocumentAccount, selectDocumentAccount };

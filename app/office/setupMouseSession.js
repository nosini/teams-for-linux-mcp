// In-process child-frame routing only. No debugging port, caller target or
// protocol command. The parent hit tests still run before any mouse press.
async function setupMouseSession(debuggerApi, frame, root, command, owned) {
  if (!frame || frame === root || frame.processId === root.processId) return { sessionId: undefined, close() {} };
  const sessions = new Map(), pending = new Set();
  let active = true, exceeded = false;
  const close = () => { active = false; debuggerApi.removeListener("message", message); };
  const message = (_event, method, params, parentSession) => {
    if (!active || !owned()) return;
    if (method === "Target.detachedFromTarget") { sessions.delete(params.sessionId); return; }
    if (method !== "Target.attachedToTarget" || params.targetInfo?.type !== "iframe") return;
    const depth = parentSession ? (sessions.get(parentSession)?.depth ?? 8) + 1 : 1;
    if (sessions.size >= 16 || depth > 8) { exceeded = true; return; }
    sessions.set(params.sessionId, { url: params.targetInfo.url, depth });
    const task = command("Target.setAutoAttach", { autoAttach: true, waitForDebuggerOnStart: false, flatten: true }, params.sessionId)
      .catch(() => {}).finally(() => pending.delete(task));
    pending.add(task);
  };
  debuggerApi.on("message", message);
  try {
    await command("Target.setAutoAttach", { autoAttach: true, waitForDebuggerOnStart: false, flatten: true });
    const deadline = Date.now() + 1500;
    do {
      if (!active || !owned() || exceeded || frame.detached) break;
      const matches = [...sessions].filter(([, info]) => info.url === frame.url);
      if (matches.length > 1) break;
      if (matches.length === 1 && !pending.size) return { sessionId: matches[0][0], close };
      await new Promise((resolve) => setTimeout(resolve, 20));
    } while (Date.now() < deadline);
    close(); return null;
  } catch (error) { close(); throw error; }
}

module.exports = { setupMouseSession };

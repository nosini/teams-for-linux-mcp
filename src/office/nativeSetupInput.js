// Native Chromium activation for an already validated, named setup control.
// No remote debugging port, arbitrary keys or caller CDP commands/coordinates.
const { setupMouseSession } = require("./setupMouseSession");

async function nativeSetupInput(contents, focus, isCurrent, mousePoint, frame) {
  const debuggerApi = contents.debugger;
  if (!debuggerApi || contents.isDevToolsOpened?.() || debuggerApi.isAttached()) return { performed: false, code: "DEBUGGER_BUSY" };
  let attached = false, focusEnabled = false, keyPressed = false, mousePressed = false, point, dispatchPoint, route;
  const detached = () => { attached = false; };
  const key = { key: "Enter", code: "Enter", windowsVirtualKeyCode: 13, nativeVirtualKeyCode: 13 };
  const command = async (method, params, sessionId) => {
    if (!attached || !debuggerApi.isAttached()) throw new Error("Native setup debugger detached");
    let timer;
    try {
      return await Promise.race([
        debuggerApi.sendCommand(method, params, sessionId),
        new Promise((_, reject) => { timer = setTimeout(() => reject(new Error("Native setup input timed out")), 2000); }),
      ]);
    } finally { clearTimeout(timer); }
  };
  try {
    debuggerApi.attach("1.3"); attached = true;
    debuggerApi.once?.("detach", detached);
    if (mousePoint) {
      route = await setupMouseSession(debuggerApi, frame, contents.mainFrame, command, () => attached && debuggerApi.isAttached());
      if (!route) return { performed: false, code: "FRAME_INPUT_UNAVAILABLE" };
    }
    focusEnabled = true;
    await command("Emulation.setFocusEmulationEnabled", { enabled: true });
    if (!isCurrent()) return { performed: false, code: "DOCUMENT_CHANGED" };
    const focused = await focus();
    if (!focused?.performed) return { performed: false, code: focused?.code || "CONTROL_NOT_FOCUSED" };
    if (!isCurrent()) return { performed: false, code: "DOCUMENT_CHANGED" };
    if (mousePoint) {
      point = await mousePoint(focused);
      if (!point || ![point.x, point.y].every(Number.isFinite)) return { performed: false, code: "CONTROL_NOT_POSITIONED" };
      dispatchPoint = route.sessionId ? focused.point : point;
      await command("Input.dispatchMouseEvent", { type: "mouseMoved", ...dispatchPoint, button: "none", buttons: 0 }, route.sessionId);
      // Hover can rearrange a menu. Revalidate its named target before press.
      const checked = await focus(), checkedPoint = checked?.performed && await mousePoint(checked);
      if (!checkedPoint || Math.abs(point.x - checkedPoint.x) > 0.5 || Math.abs(point.y - checkedPoint.y) > 0.5) return { performed: false, code: "CONTROL_MOVED" };
      if (!isCurrent()) return { performed: false, code: "DOCUMENT_CHANGED" };
      dispatchPoint = route.sessionId ? checked.point : checkedPoint;
      mousePressed = true;
      await command("Input.dispatchMouseEvent", { type: "mousePressed", ...dispatchPoint, button: "left", buttons: 1, clickCount: 1 }, route.sessionId);
      await command("Input.dispatchMouseEvent", { type: "mouseReleased", ...dispatchPoint, button: "left", buttons: 0, clickCount: 1 }, route.sessionId);
      mousePressed = false;
      return { performed: true };
    }
    keyPressed = true;
    await command("Input.dispatchKeyEvent", { type: "keyDown", ...key, text: "\r", unmodifiedText: "\r" });
    await command("Input.dispatchKeyEvent", { type: "keyUp", ...key });
    keyPressed = false;
    return { performed: true };
  } finally {
    try {
      if (attached) {
        try {
          try {
            if (keyPressed) await command("Input.dispatchKeyEvent", { type: "keyUp", ...key });
            if (mousePressed) await command("Input.dispatchMouseEvent", { type: "mouseReleased", ...dispatchPoint, button: "left", buttons: 0, clickCount: 1 }, route?.sessionId);
          }
          finally { if (focusEnabled) await command("Emulation.setFocusEmulationEnabled", { enabled: false }); }
        } finally { if (attached && debuggerApi.isAttached()) debuggerApi.detach(); }
      }
    } finally { route?.close(); debuggerApi.removeListener?.("detach", detached); }
  }
}

module.exports = { nativeSetupInput };

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { EventEmitter } = require("node:events");
const { nativeSetupInput } = require("../../app/office/nativeSetupInput");

function fixture() {
  const calls = [];
  let attached = false;
  const debuggerApi = Object.assign(new EventEmitter(), {
    isAttached: () => attached,
    attach: () => { attached = true; calls.push(["attach"]); },
    detach: () => { attached = false; calls.push(["detach"]); debuggerApi.emit("detach", {}, "fixture"); },
    sendCommand: async (method, args) => { calls.push([method, args]); return {}; },
  });
  return { contents: { debugger: debuggerApi }, calls };
}

describe("native setup activation", () => {
  it("sends native mouse input only after stable positioning and never sends Enter", async () => {
    const { contents, calls } = fixture();
    let checked = 0;
    const result = await nativeSetupInput(contents, async () => ({ performed: true }), () => true, async () => { checked++; return { x: 45, y: 67 }; });
    assert.deepEqual(result, { performed: true });
    assert.equal(checked, 2, "Hover must be followed by a new named-control check");
    assert.deepEqual(calls.filter(([method]) => method === "Input.dispatchMouseEvent").map(([, args]) => [args.type, args.button, args.x, args.y]), [["mouseMoved", "none", 45, 67], ["mousePressed", "left", 45, 67], ["mouseReleased", "left", 45, 67]]);
    assert.equal(calls.some(([method]) => method === "Input.dispatchKeyEvent"), false);
    assert.equal(contents.debugger.isAttached(), false);
  });

  it("refuses missing or moved mouse targets before pressing", async () => {
    for (const points of [[null], [{ x: 1, y: 2 }, { x: 9, y: 2 }]]) {
      const { contents, calls } = fixture();
      const result = await nativeSetupInput(contents, async () => ({ performed: true }), () => true, async () => points.shift());
      assert.equal(result.performed, false);
      assert.equal(calls.some(([method, args]) => method === "Input.dispatchMouseEvent" && args.type === "mousePressed"), false);
      assert.equal(contents.debugger.isAttached(), false);
    }
  });

  it("releases the mouse and detaches if native press dispatch fails", async () => {
    const { contents, calls } = fixture();
    const send = contents.debugger.sendCommand;
    contents.debugger.sendCommand = async (method, args) => {
      await send(method, args);
      if (args.type === "mousePressed") throw new Error("Interrupted mouse press");
      return {};
    };
    await assert.rejects(nativeSetupInput(contents, async () => ({ performed: true }), () => true, async () => ({ x: 1, y: 2 })), /Interrupted/);
    assert.ok(calls.some(([method, args]) => method === "Input.dispatchMouseEvent" && args.type === "mouseReleased"));
    assert.equal(contents.debugger.isAttached(), false);
  });

  it("presses only Enter after validated focus, then restores focus emulation and detaches", async () => {
    const { contents, calls } = fixture();
    const result = await nativeSetupInput(contents, async () => ({ performed: true }), () => true);
    assert.deepEqual(result, { performed: true });
    assert.deepEqual(calls.filter(([method]) => method === "Input.dispatchKeyEvent").map(([, args]) => [args.type, args.key]), [["keyDown", "Enter"], ["keyUp", "Enter"]]);
    assert.deepEqual(calls.at(-2), ["Emulation.setFocusEmulationEnabled", { enabled: false }]);
    assert.deepEqual(calls.at(-1), ["detach"]);
  });

  it("does not interfere with an existing debugger", async () => {
    const { contents, calls } = fixture();
    contents.debugger.attach(); calls.length = 0;
    assert.equal((await nativeSetupInput(contents, () => assert.fail("No focus allowed"), () => true)).code, "DEBUGGER_BUSY");
    assert.equal(contents.debugger.isAttached(), true);
    assert.equal(calls.length, 0);
    contents.debugger.detach(); calls.length = 0;
    contents.isDevToolsOpened = () => true;
    assert.equal((await nativeSetupInput(contents, () => assert.fail("No focus allowed"), () => true)).code, "DEBUGGER_BUSY");
    assert.equal(calls.length, 0);
  });

  it("does not send keys or detach a replacement debugger after losing ownership", async () => {
    const { contents, calls } = fixture();
    await assert.rejects(nativeSetupInput(contents, async () => {
      contents.debugger.detach();
      contents.debugger.attach();
      return { performed: true };
    }, () => true), /detached/);
    assert.equal(contents.debugger.isAttached(), true);
    assert.equal(calls.some(([method]) => method === "Input.dispatchKeyEvent"), false);
    assert.equal(contents.debugger.listenerCount("detach"), 0);
  });

  it("sends no key if the named control cannot be focused or the document changes", async () => {
    for (const changed of [false, true]) {
      const { contents, calls } = fixture();
      let checks = 0;
      const result = await nativeSetupInput(contents, async () => ({ performed: changed }), () => ++checks === 1 || !changed);
      assert.equal(result.code, changed ? "DOCUMENT_CHANGED" : "CONTROL_NOT_FOCUSED");
      assert.equal(calls.some(([method]) => method === "Input.dispatchKeyEvent"), false);
      assert.equal(contents.debugger.isAttached(), false);
    }
  });

  it("releases Enter and detaches even if dispatch fails after key down", async () => {
    const { contents, calls } = fixture();
    const send = contents.debugger.sendCommand;
    contents.debugger.sendCommand = async (method, args) => {
      await send(method, args);
      if (args.type === "keyDown") throw new Error("Interrupted dispatch");
      return {};
    };
    await assert.rejects(nativeSetupInput(contents, async () => ({ performed: true }), () => true), /Interrupted/);
    assert.ok(calls.some(([method, args]) => method === "Input.dispatchKeyEvent" && args.type === "keyUp"));
    assert.equal(contents.debugger.isAttached(), false);
    assert.deepEqual(calls.at(-2), ["Emulation.setFocusEmulationEnabled", { enabled: false }]);
  });
});

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { EventEmitter } = require("node:events");
const { setupMouseSession } = require("../../app/office/setupMouseSession");

describe("native setup child-frame routing", () => {
  const root = { processId: 1 };
  const frame = { processId: 2, url: "https://euc-excel.officeapps.live.com/x/edit" };
  it("routes only to one matching descendant session and releases its listener", async () => {
    const api = new EventEmitter(), calls = [];
    const command = async (method, params, sessionId) => {
      calls.push({ method, params, sessionId });
      if (!sessionId) api.emit("message", {}, "Target.attachedToTarget", { sessionId: "child", targetInfo: { type: "iframe", url: frame.url } });
    };
    const route = await setupMouseSession(api, frame, root, command, () => true);
    assert.equal(route.sessionId, "child");
    assert.equal(calls.length, 2);
    assert.ok(calls.every((call) => call.method === "Target.setAutoAttach" && call.params.waitForDebuggerOnStart === false));
    route.close();
    assert.equal(api.listenerCount("message"), 0);
    api.emit("message", {}, "Target.attachedToTarget", { sessionId: "late", targetInfo: { type: "iframe", url: frame.url } });
    assert.equal(calls.length, 2);
  });

  it("refuses ambiguous sessions and preserves debugger ownership boundaries", async () => {
    for (const owned of [true, false]) {
      const api = new EventEmitter();
      const command = async (_method, _params, sessionId) => {
        if (!sessionId) for (const id of ["one", "two"]) api.emit("message", {}, "Target.attachedToTarget", { sessionId: id, targetInfo: { type: "iframe", url: frame.url } });
      };
      assert.equal(await setupMouseSession(api, frame, root, command, () => owned), null);
      assert.equal(api.listenerCount("message"), 0);
    }
  });

  it("bounds attached descendants and removes listeners after protocol failures", async () => {
    const api = new EventEmitter();
    const command = async (_method, _params, sessionId) => {
      if (!sessionId) for (let id = 0; id < 17; id++) api.emit("message", {}, "Target.attachedToTarget", { sessionId: String(id), targetInfo: { type: "iframe", url: "https://other.officeapps.live.com/" } });
    };
    assert.equal(await setupMouseSession(api, frame, root, command, () => true), null);
    assert.equal(api.listenerCount("message"), 0);
    await assert.rejects(setupMouseSession(api, frame, root, async () => { throw new Error("Failed"); }, () => true), /Failed/);
    assert.equal(api.listenerCount("message"), 0);
    assert.equal((await setupMouseSession(api, root, root, () => assert.fail(), () => true)).sessionId, undefined);
  });
});

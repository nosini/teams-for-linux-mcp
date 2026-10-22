const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { z } = require("zod");
const { registerJsonTool } = require("../../app/mcp/tools");
const { evaluateInFrame } = require("../../app/mcp/evaluateInFrame");

function fakeServer() {
  const tools = new Map();
  return { tools, registerTool: (name, options, handler) => tools.set(name, { ...options, handler }) };
}

describe("MCP tool helpers", () => {
  it("returns JSON results, sanitized or fixed errors and explicit annotations", async () => {
    const server = fakeServer();
    const inputSchema = z.object({}).strict();
    registerJsonTool(server, "local_read", { description: "Read.", inputSchema, openWorld: false }, async () => ({ value: 1 }));
    registerJsonTool(server, "remote_write", { description: "Write.", inputSchema, writing: true, openWorld: true }, () => { throw new Error("Sanitized failure."); });
    registerJsonTool(server, "fixed_error", { description: "Probe.", inputSchema, openWorld: false, errorMessage: () => "Select again." }, async () => { throw new Error("private renderer detail"); });
    const read = server.tools.get("local_read");
    assert.deepEqual(read.annotations, { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false });
    assert.deepEqual(await read.handler({}), { content: [{ type: "text", text: "{\"value\":1}" }], structuredContent: { value: 1 } });
    const write = server.tools.get("remote_write");
    assert.deepEqual(write.annotations, { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true });
    assert.deepEqual(await write.handler({}), { isError: true, content: [{ type: "text", text: "Sanitized failure." }] });
    assert.deepEqual(await server.tools.get("fixed_error").handler({}), { isError: true, content: [{ type: "text", text: "Select again." }] });
    assert.throws(() => registerJsonTool(server, "unstated", { description: "?", inputSchema }, () => null), /open-world/);
  });

  it("evaluates without a user gesture and rejects with the caller's timeout message", async () => {
    const gestures = [];
    const frame = (result) => ({ executeJavaScript: (source, gesture) => { gestures.push(gesture); return result; } });
    assert.equal(await evaluateInFrame(frame(Promise.resolve(3)), "1 + 2", 1000, "Timed out"), 3);
    await assert.rejects(evaluateInFrame(frame(new Promise(() => {})), "stall()", 10, "Timed out"), { message: "Timed out" });
    await assert.rejects(evaluateInFrame(frame(Promise.reject(new Error("renderer"))), "fail()", 1000, "Timed out"), { message: "renderer" });
    assert.deepEqual(gestures, [false, false, false]);
    assert.equal(await evaluateInFrame(frame(Promise.resolve(4)), "authorized UI action", 1000, "Timed out", true), 4);
    assert.equal(gestures.at(-1), true);
  });
});

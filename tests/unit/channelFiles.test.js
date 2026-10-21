const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { ChannelFiles, registerChannelFileTools, LIMIT } = require("../../app/mcp/channelFiles");

const base = "https://graph.microsoft.com/v1.0";
const item = { id: "file", name: "test.txt", size: 5, file: {}, eTag: '"one"' };
function setup(responses) {
  const calls = [];
  const files = new ChannelFiles({ acquireToken: async () => ({ success: true, token: "secret-token" }) }, async (url, options) => {
    calls.push({ url, ...options });
    assert.ok(responses.length, `Unexpected request: ${url}`);
    const value = responses.shift();
    if (value instanceof Error) throw value;
    return value instanceof Response ? value : Response.json(value);
  });
  return { files, calls };
}

describe("Teams channel files", () => {
  it("downloads exact bytes without forwarding Graph credentials to the download host", async () => {
    const { files, calls } = setup([item, new Response(null, { status: 302, headers: { location: "https://tenant.sharepoint.com/download?secret=link" } }), new Response("hello"), item]);
    const result = await files.read({ driveId: "drive", itemId: "file" });
    assert.equal(Buffer.from(result.contentBase64, "base64").toString(), "hello");
    assert.equal(result.item.eTag, item.eTag);
    assert.equal(calls[0].headers.Authorization, "Bearer secret-token");
    assert.equal(calls[2].headers.Authorization, undefined);
    assert.equal(calls[2].redirect, "manual");
    assert.doesNotMatch(JSON.stringify(result), /secret/);
    await assert.rejects(files.request("https://evil.test/v1.0/me"), /Invalid Graph/);
  });

  it("rejects changed files, folders and oversized transfers", async () => {
    const changed = setup([item, new Response("hello"), { ...item, eTag: '"two"' }]);
    await assert.rejects(changed.files.read({ driveId: "d", itemId: "f" }), /changed/);
    const folder = setup([{ folder: {} }]);
    await assert.rejects(folder.files.read({ driveId: "d", itemId: "f" }), /folder/);
    const large = setup([{ ...item, size: LIMIT + 1 }]);
    await assert.rejects(large.files.read({ driveId: "d", itemId: "f" }), /1 MiB/);
    const streamed = setup([item, new Response(Buffer.alloc(2 * LIMIT + 1))]);
    await assert.rejects(streamed.files.read({ driveId: "d", itemId: "f" }), /transfer limit/);
  });

  it("creates with conflict failure and replaces using the exact eTag and raw bytes", async () => {
    const { files, calls } = setup([item, item, item]);
    const contentBase64 = Buffer.from("hello").toString("base64");
    await files.write({ driveId: "d", parentId: "p", name: "test #.txt", contentBase64 });
    assert.match(calls[0].url, /test%20%23.txt:\/content\?@microsoft.graph.conflictBehavior=fail$/);
    assert.equal(calls[0].headers["If-None-Match"], "*");
    await files.write({ driveId: "d", itemId: "file", expectedETag: item.eTag, contentBase64 });
    assert.equal(calls[2].headers["If-Match"], item.eTag);
    assert.equal(calls[2].method, "PUT");
    assert.deepEqual(calls[2].body, Buffer.from("hello"));
    const stale = setup([{ ...item, eTag: '"two"' }]);
    await assert.rejects(stale.files.write({ driveId: "d", itemId: "f", expectedETag: item.eTag, contentBase64 }), /changed/);
    assert.equal(stale.calls.length, 1);
    await assert.rejects(files.write({ contentBase64: "a===" }), /canonical/);
  });

  it("sanitizes HTTP and network errors and never retries an uncertain write", async () => {
    const denied = setup([new Response('secret-server-body', { status: 403 })]);
    await assert.rejects(denied.files.json(`${base}/me/joinedTeams`), /HTTP 403.*permission/);
    const failed = setup([new Error("https://host/upload?secret=value")]);
    await assert.rejects(failed.files.write({ driveId: "d", parentId: "p", name: "test.txt", contentBase64: "" }), /inspect the destination/);
    assert.equal(failed.calls.length, 1);
    const invalid = setup([new Response("secret-content")]);
    await assert.rejects(invalid.files.json(`${base}/me`), { message: "Invalid Graph JSON response" });
  });

  it("reports only fixed diagnostic labels from bounded Microsoft error bodies", async () => {
    const { files } = setup([Response.json({ error: { code: "invalidRequest", message: "deferCommit not supported: https://host/?secret=token-value" } }, { status: 400 })]);
    await assert.rejects(files.json(`${base}/me`), { message: "File request failed (HTTP 400). Microsoft diagnostic: invalidRequest, defercommit, token, not supported." });
    const malicious = setup([Response.json({ error: { code: "secret-token", message: "https://secret.test/?credential=value" } }, { status: 400 })]);
    await assert.rejects(malicious.files.json(`${base}/me`), { message: "File request failed (HTTP 400)." });
  });

  it("paginates with an opaque cursor and retains the selected folder", async () => {
    const { files, calls } = setup([{ value: [{ ...item, "@microsoft.graph.downloadUrl": "secret-download" }], "@odata.nextLink": `${base}/drives/d/items/p/children?$skiptoken=abc%2Bdef` }, { value: [] }]);
    const endpoint = "/drives/d/items/p/children?$top=50";
    const result = await files.list(endpoint);
    assert.equal(result.nextCursor, "abc+def");
    assert.doesNotMatch(JSON.stringify(result), /secret-download/);
    await files.list(endpoint, result.nextCursor);
    assert.equal(calls[1].url, `${base}${endpoint}&$skiptoken=abc%2Bdef`);
  });

  it("validates file names, base64, identifiers and write preconditions at the MCP boundary", async () => {
    const tools = new Map();
    registerChannelFileTools({ registerTool: (name, options, handler) => tools.set(name, { ...options, handler }) }, setup([]).files);
    assert.equal(tools.size, 14);
    const create = tools.get("teams_create_file");
    const args = { driveId: "d", parentId: "p", name: "test.txt", contentBase64: "" };
    assert.ok(create.inputSchema.safeParse(args).success);
    for (const name of ["../test.txt", "test/other.txt", "..", "name."]) assert.equal(create.inputSchema.safeParse({ ...args, name }).success, false);
    assert.equal(create.inputSchema.safeParse({ ...args, driveId: "../me" }).success, false);
    const replace = tools.get("teams_replace_file");
    assert.equal(replace.inputSchema.safeParse({ driveId: "d", itemId: "f", contentBase64: "", expectedETag: "*" }).success, false);
    assert.equal(create.annotations.readOnlyHint, false);
    assert.equal(tools.get("teams_read_file").annotations.readOnlyHint, true);
  });
});

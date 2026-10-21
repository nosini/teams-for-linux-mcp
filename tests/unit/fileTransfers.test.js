const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { createHash } = require("node:crypto");
const { ChannelFiles, registerChannelFileTools } = require("../../app/mcp/channelFiles");
const { CHUNK_SIZE } = require("../../app/mcp/fileTransfers");
const item = { id: "file", name: "test.bin", file: {}, eTag: '"one"', size: 7 };
const session = () => ({ uploadUrl: "https://tenant.sharepoint.com/upload?secret=session", expirationDateTime: new Date(Date.now() + 3600000).toISOString() });
const progress = (offset) => Response.json({ ...session(), nextExpectedRanges: [`${offset}-`] }, { status: 202 });
const digest = (bytes) => createHash("sha256").update(bytes).digest("hex");
function setup(responses) {
  const calls = [];
  const files = new ChannelFiles({ acquireToken: async () => ({ success: true, token: "secret-token" }) }, async (url, options) => {
    calls.push({ url, ...options });
    assert.ok(responses.length, `Unexpected request: ${url}`);
    const value = responses.shift();
    if (value instanceof Error) throw value;
    return value instanceof Response ? value : Response.json(value);
  });
  return { transfers: files.transfers, files, calls };
}
const start = (transfers, size, extra = {}) => transfers.beginUpload({ driveId: "d", parentId: "p", name: "test.bin", size, ...extra });
const send = (transfers, transferId, offset, bytes) => transfers.uploadChunk({ transferId, offset, contentBase64: bytes.toString("base64") });

describe("chunked channel file transfers", () => {
  it("downloads more than 1 MiB with exact ranges, version checks, retries and no credential leakage", async () => {
    const bytes = Buffer.alloc(CHUNK_SIZE * 2 + 7, 91);
    const metadata = { ...item, size: bytes.length, "@microsoft.graph.downloadUrl": "https://tenant.sharepoint.com/download?secret=link" };
    const responses = [metadata];
    for (const offset of [0, CHUNK_SIZE, CHUNK_SIZE * 2, CHUNK_SIZE * 2]) {
      const end = Math.min(offset + CHUNK_SIZE, bytes.length);
      responses.push(metadata, new Response(bytes.subarray(offset, end), { status: 206, headers: { "Content-Range": `bytes ${offset}-${end - 1}/${bytes.length}` } }), metadata);
    }
    const { transfers, calls } = setup(responses);
    const transfer = await transfers.beginDownload({ driveId: "d", itemId: "file" });
    const chunks = [];
    for (let offset = 0; offset < bytes.length;) {
      const result = await transfers.readChunk({ transferId: transfer.transferId, offset });
      const chunk = Buffer.from(result.contentBase64, "base64");
      assert.equal(result.chunkSha256, digest(chunk));
      chunks.push(chunk);
      offset = result.nextOffset;
      assert.doesNotMatch(JSON.stringify(result), /secret/);
    }
    assert.deepEqual(Buffer.concat(chunks), bytes);
    const repeated = await transfers.readChunk({ transferId: transfer.transferId, offset: CHUNK_SIZE * 2 });
    assert.equal(repeated.status, "completed");
    assert.equal(repeated.done, true);
    for (const call of calls.filter((call) => call.url.includes("sharepoint"))) assert.equal(call.headers.Authorization, undefined);
    await assert.rejects(transfers.readChunk({ transferId: transfer.transferId, offset: 0 }), /nextOffset/);
  });

  it("resolves a content redirect when Graph omits the download annotation", async () => {
    const { transfers, calls } = setup([item, item, new Response(null, { status: 302, headers: { location: "https://tenant.sharepoint.com/download?secret=link" } }), new Response(Buffer.alloc(7), { status: 206, headers: { "content-range": "bytes 0-6/7" } }), item]);
    const { transferId } = await transfers.beginDownload({ driveId: "d", itemId: "file" });
    const chunk = await transfers.readChunk({ transferId, offset: 0 });
    assert.equal(chunk.done, true);
    assert.equal(calls[2].headers.Authorization, "Bearer secret-token");
    assert.equal(calls[3].headers.Authorization, undefined);
    assert.equal(calls[3].headers.Range, "bytes=0-6");
    assert.doesNotMatch(JSON.stringify(chunk), /secret/);
  });

  it("rejects ignored, incorrect and oversized range responses", async () => {
    const metadata = { ...item, size: CHUNK_SIZE + 1, "@microsoft.graph.downloadUrl": "https://tenant.sharepoint.com/download" };
    for (const response of [new Response(Buffer.alloc(CHUNK_SIZE)), new Response(Buffer.alloc(CHUNK_SIZE), { status: 206, headers: { "content-range": "bytes 1-2/3" } }), new Response(Buffer.alloc(CHUNK_SIZE + 1))]) {
      const { transfers } = setup([metadata, metadata, response]);
      const { transferId } = await transfers.beginDownload({ driveId: "d", itemId: "file" });
      await assert.rejects(transfers.readChunk({ transferId, offset: 0 }), /exact requested|transfer limit/);
    }
  });

  it("discards bytes when the version changes during download and handles empty files", async () => {
    const metadata = { ...item, "@microsoft.graph.downloadUrl": "https://tenant.sharepoint.com/download" };
    const { transfers } = setup([metadata, metadata, new Response(Buffer.alloc(7)), { ...metadata, eTag: '"two"' }]);
    const { transferId } = await transfers.beginDownload({ driveId: "d", itemId: "file" });
    await assert.rejects(transfers.readChunk({ transferId, offset: 0 }), /changed/);
    assert.equal((await transfers.getTransfer({ transferId })).status, "stale");
    const empty = setup(Array(3).fill({ ...item, size: 0 }));
    const zero = await empty.transfers.beginDownload({ driveId: "d", itemId: "file" });
    const result = await empty.transfers.readChunk({ transferId: zero.transferId, offset: 0 });
    assert.equal(result.contentBase64, "");
    assert.equal(result.done, true);
    assert.equal(empty.calls.length, 3);
  });

  it("stages a multi-chunk upload, hashes source bytes and publishes only at finish", async () => {
    const bytes = Buffer.alloc(CHUNK_SIZE * 2 + 7, 42);
    const { transfers, calls } = setup([session(), progress(CHUNK_SIZE), progress(CHUNK_SIZE * 2), progress(bytes.length), item]);
    const { transferId } = await start(transfers, bytes.length, { expectedSha256: digest(bytes) });
    assert.equal(JSON.parse(calls[0].body).deferCommit, true);
    assert.equal(JSON.parse(calls[0].body).item["@microsoft.graph.conflictBehavior"], "fail");
    for (let offset = 0; offset < bytes.length; offset += CHUNK_SIZE) {
      await send(transfers, transferId, offset, bytes.subarray(offset, offset + CHUNK_SIZE));
    }
    const ready = await send(transfers, transferId, CHUNK_SIZE * 2, bytes.subarray(CHUNK_SIZE * 2));
    assert.equal(ready.status, "ready");
    assert.equal(ready.sourceSha256, digest(bytes));
    assert.equal(calls.length, 4); // Duplicate last chunk was not sent.
    for (const [index, call] of calls.slice(1).entries()) {
      assert.equal(call.method, "PUT");
      assert.equal(call.headers.Authorization, undefined);
      assert.equal(Number(call.headers["Content-Length"]), call.body.length);
      assert.equal(call.headers["Content-Range"], `bytes ${index * CHUNK_SIZE}-${Math.min((index + 1) * CHUNK_SIZE, bytes.length) - 1}/${bytes.length}`);
    }
    const done = await transfers.finishUpload({ transferId });
    assert.equal(done.status, "completed");
    assert.equal(calls[4].method, "POST");
    assert.equal(calls[4].headers["Content-Length"], "0");
    assert.doesNotMatch(JSON.stringify(done), /secret/);
    await transfers.finishUpload({ transferId });
    assert.equal(calls.length, 5);
    assert.match((await transfers.cancelTransfer({ transferId })).instruction, /already saved/);
  });

  it("serializes OData conflict annotations before name in create, replace and fallback requests", async () => {
    for (const replacing of [false, true]) {
      const { transfers, calls } = setup([...(replacing ? [item] : []), new Response("invalid request", { status: 400 }), session()]);
      const args = replacing ? { itemId: "file", expectedETag: item.eTag } : { parentId: "p", name: "test.bin" };
      await transfers.beginUpload({ driveId: "d", size: 7, ...args });
      const behavior = replacing ? "replace" : "fail";
      for (const request of calls.filter((call) => call.method === "POST")) {
        assert.match(request.body, new RegExp('^\\{"item":\\{"@microsoft\\.graph\\.conflictBehavior":"' + behavior + '","name":"test\\.bin"\\}'));
      }
    }
  });

  it("falls back only after definite HTTP 400 session rejection and reports automatic publishing", async () => {
    const { transfers, calls } = setup([new Response("unsupported", { status: 400 }), session(), item]);
    const upload = await start(transfers, 7);
    assert.equal(upload.commitMode, "automatic");
    assert.equal(JSON.parse(calls[0].body).deferCommit, true);
    assert.equal("deferCommit" in JSON.parse(calls[1].body), false);
    assert.equal(JSON.parse(calls[1].body).item["@microsoft.graph.conflictBehavior"], "fail");
    assert.equal(calls[1].headers.Accept, "application/json");
    assert.equal((await send(transfers, upload.transferId, 0, Buffer.alloc(7))).status, "completed");
    assert.equal(calls[2].headers.Accept, "application/json");
    await transfers.finishUpload(upload);
    assert.equal(calls.length, 3);
    for (const response of [new Error("uncertain"), new Response(null, { status: 403 }), new Response(null, { status: 409 }), new Response(null, { status: 412 })]) {
      const failed = setup([response]);
      await assert.rejects(start(failed.transfers, 7));
      assert.equal(failed.calls.length, 1);
    }
  });

  it("requires sequential aligned chunks and verifies source hash before sending the final chunk", async () => {
    const { transfers, calls } = setup([session()]);
    const { transferId } = await start(transfers, CHUNK_SIZE + 1, { expectedSha256: "0".repeat(64) });
    for (const [offset, bytes] of [[1, Buffer.alloc(CHUNK_SIZE)], [0, Buffer.alloc(7)], [0, Buffer.alloc(CHUNK_SIZE + 1)]]) {
      await assert.rejects(send(transfers, transferId, offset, bytes), /sequential|canonical/);
    }
    await assert.rejects(transfers.uploadChunk({ transferId, offset: 0, contentBase64: "AB==" }), /canonical/);
    assert.equal(calls.length, 1);
    const small = setup([session()]);
    const upload = await start(small.transfers, 7, { expectedSha256: "0".repeat(64) });
    await assert.rejects(send(small.transfers, upload.transferId, 0, Buffer.alloc(7)), /SHA-256/);
    assert.equal(small.calls.length, 1);
    await assert.rejects(small.transfers.finishUpload(upload), /all chunks/);
  });

  it("checks replacement versions before chunks and finish", async () => {
    for (const changedBeforeFinish of [false, true]) {
      const responses = [item, session(), ...(changedBeforeFinish ? [item, progress(7)] : []), { ...item, eTag: '"two"' }];
      const { transfers, calls } = setup(responses);
      const { transferId } = await transfers.beginUpload({ driveId: "d", itemId: "file", expectedETag: item.eTag, size: 7 });
      assert.equal(calls[1].headers["If-Match"], item.eTag);
      if (changedBeforeFinish) {
        await send(transfers, transferId, 0, Buffer.alloc(7));
        await assert.rejects(transfers.finishUpload({ transferId }), /changed/);
      } else await assert.rejects(send(transfers, transferId, 0, Buffer.alloc(7)), /changed/);
      assert.equal((await transfers.getTransfer({ transferId })).status, "stale");
      assert.equal(calls.filter((call) => call.url.includes("sharepoint") && call.method === "POST").length, 0);
    }
  });

  it("commits a replacement with If-Match and accepts an empty missing-range list", async () => {
    const { transfers, calls } = setup([item, session(), item, Response.json({ nextExpectedRanges: [] }, { status: 202 }), item, { ...item, eTag: '"two"' }]);
    const { transferId } = await transfers.beginUpload({ driveId: "d", itemId: "file", expectedETag: item.eTag, size: 7 });
    await send(transfers, transferId, 0, Buffer.alloc(7));
    const done = await transfers.finishUpload({ transferId });
    assert.equal(done.status, "completed");
    assert.equal(done.item.eTag, '"two"');
    assert.equal(calls.at(-1).headers["If-Match"], item.eTag);
    assert.equal(calls.at(-1).headers.Authorization, undefined);
  });

  it("accepts bounded missing ranges and validates both boundaries without exposing URLs", async () => {
    const { transfers } = setup([]);
    assert.equal(transfers.expectedOffset({ nextExpectedRanges: ["655360-1310736"] }, 1310737), 655360);
    assert.equal(transfers.expectedOffset({ nextExpectedRanges: ["655360-"] }, 1310737), 655360);
    assert.equal(transfers.expectedOffset({ nextExpectedRanges: ["1310737-"] }, 1310737), 1310737);
    assert.equal(transfers.expectedOffset({ nextExpectedRanges: [] }, 1310737), 1310737);
    for (const range of ["7-6", "0-7", "8-", "0-9007199254740992", "9007199254740992-", "-1-6"]) assert.throws(() => transfers.expectedOffset({ nextExpectedRanges: [range] }, 7), /invalid|unsupported/);
    assert.throws(() => transfers.expectedOffset({ nextExpectedRanges: ["0-2", "4-6"] }, 7), /unsupported/);
    assert.throws(() => transfers.expectedOffset({ nextExpectedRanges: ["https://host/?secret=token"] }, 7), { message: "Microsoft returned unsupported upload ranges ([unrecognized]). Inspect or cancel this transfer." });
    const live = setup([session(), Response.json({ nextExpectedRanges: [`${CHUNK_SIZE}-${CHUNK_SIZE + 6}`] }, { status: 202 }), Response.json({ nextExpectedRanges: [] }, { status: 202 }), item]);
    const upload = await start(live.transfers, CHUNK_SIZE + 7);
    assert.equal((await send(live.transfers, upload.transferId, 0, Buffer.alloc(CHUNK_SIZE))).nextOffset, CHUNK_SIZE);
    assert.equal((await send(live.transfers, upload.transferId, CHUNK_SIZE, Buffer.alloc(7))).status, "ready");
    assert.equal((await live.transfers.finishUpload(upload)).status, "completed");
  });

  it("reconciles lost chunk responses without blindly resending or double-hashing", async () => {
    for (const accepted of [true, false]) {
      const { transfers, calls } = setup([session(), new Error("secret-network-url"), progress(accepted ? 7 : 0), ...(!accepted ? [progress(7)] : []), item]);
      const { transferId } = await start(transfers, 7);
      await assert.rejects(send(transfers, transferId, 0, Buffer.alloc(7)), /inspect the destination/);
      await assert.rejects(send(transfers, transferId, 0, Buffer.alloc(7)), /not active/);
      const status = await transfers.getTransfer({ transferId });
      assert.equal(status.nextOffset, accepted ? 7 : 0);
      if (!accepted) await send(transfers, transferId, 0, Buffer.alloc(7));
      const done = await transfers.finishUpload({ transferId });
      assert.equal(done.sourceSha256, digest(Buffer.alloc(7)));
      assert.equal(calls.filter((call) => call.method === "PUT").length, accepted ? 1 : 2);
    }
  });

  it("does not retry uncertain commits or infer success from a vanished session", async () => {
    const { transfers, calls } = setup([session(), progress(7), new Error("secret-network-url"), progress(7), new Response("not JSON", { status: 404 }), new Response(null, { status: 404 })]);
    const { transferId } = await start(transfers, 7);
    await send(transfers, transferId, 0, Buffer.alloc(7));
    await assert.rejects(transfers.finishUpload({ transferId }), /inspect the destination/);
    const uncertain = await transfers.getTransfer({ transferId });
    assert.equal(uncertain.uncertainOperation, "commit");
    await assert.rejects(transfers.finishUpload({ transferId }), /all chunks/);
    const gone = await transfers.getTransfer({ transferId });
    assert.equal(gone.status, "unknown");
    assert.match(gone.instruction, /may have completed or expired/);
    const cancelled = await transfers.cancelTransfer({ transferId });
    assert.match(cancelled.instruction, /may already have been saved/);
    assert.equal(calls.filter((call) => call.method === "POST").length, 2); // Create + one commit.
  });

  it("distinguishes definite commit rejections from uncertain commits", async () => {
    const cases = [[412, "stale", /changed before the commit/], [409, "rejected", /rejected the commit/], [403, "rejected", /rejected the commit/], [410, "unknown", /may have completed/], [500, "uncertain", /HTTP 500/], [408, "uncertain", /HTTP 408/]];
    for (const [code, status, message] of cases) {
      const { transfers, calls } = setup([session(), progress(7), new Response("secret response", { status: code }), new Response(null, { status: 204 })]);
      const { transferId } = await start(transfers, 7);
      await send(transfers, transferId, 0, Buffer.alloc(7));
      await assert.rejects(transfers.finishUpload({ transferId }), (error) => message.test(error.message) && !error.message.includes("secret"));
      if (status === "uncertain") {
        assert.equal(transfers.entries.get(transferId).uncertainOperation, "commit");
        continue;
      }
      const inspected = await transfers.getTransfer({ transferId });
      assert.equal(inspected.status, status, String(code));
      assert.equal(inspected.uncertainOperation, undefined);
      const cancelled = await transfers.cancelTransfer({ transferId });
      assert.equal(cancelled.status, "cancelled");
      assert.equal(calls.at(-1).method, "DELETE");
      assert.equal("instruction" in cancelled, status === "unknown", String(code));
    }
    for (const code of [423, 429]) {
      const { transfers, calls } = setup([session(), progress(7), new Response(null, { status: code }), progress(7), item]);
      const { transferId } = await start(transfers, 7);
      await send(transfers, transferId, 0, Buffer.alloc(7));
      await assert.rejects(transfers.finishUpload({ transferId }), /did not save this upload yet/);
      assert.equal((await transfers.getTransfer({ transferId })).status, "ready");
      assert.equal((await transfers.finishUpload({ transferId })).status, "completed");
      assert.equal(calls.filter((call) => call.method === "POST").length, 3); // Create + two commits.
    }
  });

  it("reports provider auto-completion, sanitizes conflicts and cancels staged uploads", async () => {
    const auto = setup([session(), item]);
    const transfer = await start(auto.transfers, 7);
    assert.equal((await send(auto.transfers, transfer.transferId, 0, Buffer.alloc(7))).status, "completed");
    await auto.transfers.finishUpload(transfer);
    assert.equal(auto.calls.length, 2);
    const conflict = setup([new Response("secret response", { status: 409 })]);
    await assert.rejects(start(conflict.transfers, 7), { message: "File request failed (HTTP 409)." });
    const cancelled = setup([session(), new Response(null, { status: 204 })]);
    const staging = await start(cancelled.transfers, 7);
    assert.equal((await cancelled.transfers.cancelTransfer(staging)).status, "cancelled");
    assert.equal(cancelled.calls[1].method, "DELETE");
    assert.equal(cancelled.calls[1].headers.Authorization, undefined);
    await assert.rejects(send(cancelled.transfers, staging.transferId, 0, Buffer.alloc(7)), /not active/);
  });

  it("bounds active sessions, expires idle transfers and serializes operations", async () => {
    const { transfers } = setup(Array(10).fill(item));
    let now = Date.now(); transfers.now = () => now;
    const opened = [];
    for (let i = 0; i < 8; i++) opened.push(await transfers.beginDownload({ driveId: "d", itemId: "file" }));
    await assert.rejects(transfers.beginDownload({ driveId: "d", itemId: "file" }), /Eight/);
    let release;
    const locked = transfers.withTransfer(opened[0].transferId, () => new Promise((resolve) => { release = resolve; }));
    await assert.rejects(transfers.getTransfer(opened[0]), /Another operation/);
    release(); await locked;
    now += 30 * 60 * 1000 + 1;
    assert.equal((await transfers.getTransfer(opened[0])).status, "expired");
    await transfers.beginDownload({ driveId: "d", itemId: "file" });
    now += 10 * 60 * 1000;
    await assert.rejects(transfers.getTransfer(opened[0]), /unknown/);
  });

  it("validates upload modes and refuses caller-provided transfer URLs at the MCP boundary", async () => {
    const tools = new Map();
    registerChannelFileTools({ registerTool: (name, options) => tools.set(name, options) }, setup([]).files);
    const schema = tools.get("teams_begin_upload").inputSchema;
    const args = { driveId: "d", parentId: "p", name: "test.bin", size: 7 };
    assert.equal(schema.safeParse(args).success, true);
    for (const extra of [{ itemId: "file" }, { expectedETag: '"one"' }, { size: 0 }, { size: 1.5 }, { expectedSha256: "invalid" }, { uploadUrl: "https://evil.test" }]) assert.equal(schema.safeParse({ ...args, ...extra }).success, false);
    assert.equal(schema.safeParse({ driveId: "d", itemId: "file", expectedETag: '"one"', size: 7 }).success, true);
    assert.equal(tools.get("teams_upload_chunk").inputSchema.safeParse({ transferId: "a2f1bb41-0f1d-4dbc-9c65-16c0aa334841", offset: 0, contentBase64: Buffer.alloc(CHUNK_SIZE).toString("base64") }).success, true);
  });
});

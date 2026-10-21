const { randomUUID, createHash } = require("node:crypto");
const { z } = require("zod");
const { registerJsonTool } = require("./tools");
const { itemFields, pickItem, id, filename, eTag, base64 } = require("./channelFileTypes");
const CHUNK_SIZE = 640 * 1024; // Multiple of Graph's 320 KiB fragment quantum.
const IDLE_MS = 30 * 60 * 1000;
const RETAIN_MS = 10 * 60 * 1000;
const sha256 = (bytes) => createHash("sha256").update(bytes).digest("hex");
const active = (transfer) => ["active", "ready", "uncertain"].includes(transfer.status);
function transferUrl(value) {
  try {
    const url = new URL(value);
    if (url.protocol === "https:" && !url.username && !url.password) return url.href;
  } catch { /* Invalid service response. */ }
  throw new Error("Invalid transfer location returned by Microsoft.");
}
function parseJson(bytes) {
  try { return JSON.parse(bytes.toString("utf8")); } catch { throw new Error("Invalid transfer response from Microsoft."); }
}

class FileTransfers {
  #files;
  #now;
  #entries = new Map();
  #starts = 0;

  constructor(files, { now = Date.now } = {}) {
    this.#files = files;
    this.#now = now;
  }
  prune() {
    for (const [key, transfer] of this.#entries) {
      if (transfer.busy) continue;
      if (active(transfer) && transfer.expires <= this.#now()) {
        transfer.status = "expired";
        transfer.pending = null;
        transfer.uploadUrl = null;
        transfer.finished = this.#now();
      }
      if (!active(transfer) && this.#now() - (transfer.finished ?? transfer.expires) >= RETAIN_MS) this.#entries.delete(key);
    }
  }
  reserve() {
    this.prune();
    if (this.#starts + [...this.#entries.values()].filter(active).length >= 8) throw new Error("Eight transfers are already open. Finish or cancel one first.");
    this.#starts++;
    while (this.#entries.size >= 32) {
      const old = [...this.#entries].find(([, transfer]) => !active(transfer) && !transfer.busy);
      if (!old) break;
      this.#entries.delete(old[0]);
    }
  }
  remember(transfer) {
    transfer.id = randomUUID();
    transfer.expires = Math.min(this.#now() + IDLE_MS, transfer.serverExpires ?? Infinity);
    this.#entries.set(transfer.id, transfer);
    return this.describe(transfer);
  }
  describe(transfer) {
    return {
      transferId: transfer.id, direction: transfer.direction, status: transfer.status,
      size: transfer.size, nextOffset: transfer.offset, chunkSize: CHUNK_SIZE,
      expiresAt: new Date(transfer.expires).toISOString(),
      ...(transfer.item ? { item: pickItem(transfer.item) } : {}),
      ...(transfer.sourceSha256 ? { sourceSha256: transfer.sourceSha256 } : {}),
      ...(transfer.commitMode ? { commitMode: transfer.commitMode } : {}),
      ...(transfer.uncertainOperation ? { uncertainOperation: transfer.uncertainOperation } : {}),
    };
  }
  async withTransfer(transferId, operation) {
    this.prune();
    const transfer = this.#entries.get(transferId);
    if (!transfer) throw new Error("Transfer is unknown or no longer retained. Start a new transfer.");
    if (transfer.busy) throw new Error("Another operation is running on this transfer.");
    transfer.busy = true;
    try { return await operation(transfer); }
    finally { transfer.busy = false; }
  }
  touch(transfer, response) {
    if (response?.expirationDateTime) {
      const expiry = Date.parse(response.expirationDateTime);
      if (!Number.isFinite(expiry) || expiry <= this.#now()) throw new Error("Microsoft returned an expired transfer session.");
      transfer.serverExpires = expiry;
    }
    transfer.expires = Math.min(this.#now() + IDLE_MS, transfer.serverExpires ?? Infinity);
  }
  finishState(transfer, status) {
    transfer.status = status;
    transfer.finished = this.#now();
    transfer.pending = null;
    if (status === "completed" || status === "cancelled") transfer.uploadUrl = null;
  }
  async metadata(driveId, itemId, includeUrl = false) {
    return this.#files.json(`${this.#files.itemPath(driveId, itemId)}?$select=${itemFields}${includeUrl ? ",@microsoft.graph.downloadUrl" : ""}`);
  }
  async unchanged(transfer, includeUrl = false) {
    const item = await this.metadata(transfer.driveId, transfer.itemId, includeUrl);
    if (!item.file || item.eTag !== transfer.eTag || (transfer.direction === "download" && item.size !== transfer.size)) {
      this.finishState(transfer, "stale");
      throw new Error("File changed during transfer. Cancel and start again from a fresh read.");
    }
    return item;
  }
  async beginDownload({ driveId, itemId }) {
    this.reserve();
    try {
      const item = await this.metadata(driveId, itemId);
      if (!item.file || !item.eTag || !Number.isSafeInteger(item.size) || item.size < 0) throw new Error("Select a file with a valid size and version.");
      return this.remember({ direction: "download", status: "active", driveId, itemId, item: pickItem(item), eTag: item.eTag, size: item.size, offset: 0 });
    } finally { this.#starts--; }
  }
  async readChunk({ transferId, offset }) {
    return this.withTransfer(transferId, async (transfer) => {
      if (transfer.direction !== "download" || !["active", "completed"].includes(transfer.status)) throw new Error("This download is not active.");
      if (!Number.isSafeInteger(offset) || offset < 0 || (offset !== transfer.offset && offset !== transfer.lastOffset) || (offset === transfer.size && transfer.size !== 0)) throw new Error("Read the nextOffset or repeat the most recent download chunk.");
      const before = await this.unchanged(transfer, true);
      const end = Math.min(offset + CHUNK_SIZE, transfer.size), length = end - offset;
      let bytes = Buffer.alloc(0);
      if (length) {
        let response;
        const options = { headers: { Range: `bytes=${offset}-${end - 1}` } };
        const bounds = { details: true, maxBytes: length };
        try {
          if (before["@microsoft.graph.downloadUrl"]) {
            response = await this.#files.request(transferUrl(before["@microsoft.graph.downloadUrl"]), options, false, bounds);
          } else {
            // Some tenants omit the download annotation even when selected.
            // /content either serves the range itself or redirects to a
            // private URL, where Range is applied again.
            response = await this.#files.request(`${this.#files.itemPath(transfer.driveId, transfer.itemId)}/content`, options, true, bounds);
            if (response.redirect) response = await this.#files.request(transferUrl(response.redirect), options, false, bounds);
          }
        } catch (error) {
          // A server that ignores Range sends more than the requested chunk.
          if (error.code !== "RESPONSE_TOO_LARGE") throw error;
          response = {};
        }
        if (!response.body || response.body.length !== length || !((response.status === 206 && response.headers.get("content-range") === `bytes ${offset}-${end - 1}/${transfer.size}`) || (response.status === 200 && offset === 0 && end === transfer.size))) throw new Error("Microsoft did not return the exact requested byte range.");
        bytes = response.body;
      }
      await this.unchanged(transfer);
      transfer.offset = Math.max(transfer.offset, end);
      transfer.lastOffset = offset;
      this.touch(transfer);
      if (end === transfer.size) this.finishState(transfer, "completed");
      return { ...this.describe(transfer), offset, nextOffset: end, done: end === transfer.size, contentBase64: bytes.toString("base64"), chunkSha256: sha256(bytes) };
    });
  }
  async beginUpload(args) {
    const { driveId, parentId, itemId, name, expectedETag, size, expectedSha256 } = args;
    if (!Number.isSafeInteger(size) || size < 1) throw new Error("Provide a positive total byte size. Use the single-call tools for empty files.");
    this.reserve();
    try {
      let path, uploadName;
      if (itemId) {
        const item = await this.metadata(driveId, itemId);
        if (!item.file || !expectedETag || item.eTag !== expectedETag) throw new Error("File changed. Read it again before replacing it.");
        path = this.#files.itemPath(driveId, itemId);
        uploadName = item.name;
      } else {
        filename.parse(name); id.parse(parentId);
        path = `${this.#files.itemPath(driveId, parentId)}:/${encodeURIComponent(name)}:`;
        uploadName = name;
      }
      // SharePoint's OData parser requires annotations before ordinary fields.
      // JSON property order is otherwise semantically irrelevant, but putting
      // name first yields HTTP 400 invalidRequest on affected hosts.
      // https://github.com/microsoftgraph/msgraph-sdk-java/issues/1351
      const item = { "@microsoft.graph.conflictBehavior": itemId ? "replace" : "fail", name: uploadName };
      const create = (deferred) => this.#files.json(`${path}/createUploadSession`, {
        method: "POST", headers: { Accept: "application/json", "Content-Type": "application/json", ...(itemId ? { "If-Match": expectedETag } : {}) },
        body: JSON.stringify({ item, ...(deferred ? { deferCommit: true } : {}) }),
      });
      let session, commitMode = "deferred";
      try { session = await create(true); }
      catch (error) {
        // A definite HTTP 400 rejects session creation before bytes are sent.
        // Some SharePoint hosts reject the optional deferred-commit request.
        // Never fall back after a timeout, conflict or fragment/commit failure.
        if (error.status !== 400) throw error;
        session = await create(false);
        commitMode = "automatic";
      }
      const expires = Date.parse(session.expirationDateTime);
      if (!Number.isFinite(expires) || expires <= this.#now()) throw new Error("Microsoft returned an expired transfer session.");
      return this.remember({ direction: "upload", status: "active", driveId, itemId, parentId, name: uploadName, eTag: expectedETag, size, offset: 0, commitMode, expectedSha256: expectedSha256?.toLowerCase(), hash: createHash("sha256"), uploadUrl: transferUrl(session.uploadUrl), serverExpires: expires });
    } finally { this.#starts--; }
  }
  async sessionRequest(transfer, options = {}, allowStatus = []) {
    const response = await this.#files.request(transfer.uploadUrl, { ...options, headers: { Accept: "application/json", ...options.headers } }, false, { details: true, maxBytes: 65536, allowStatus });
    if (!response.body) throw new Error("Unexpected upload-session redirect.");
    return { status: response.status, data: response.body.length && ![404, 410].includes(response.status) ? parseJson(response.body) : null };
  }
  expectedOffset(data, size) {
    const ranges = data?.nextExpectedRanges;
    if (Array.isArray(ranges) && ranges.length === 0) return size;
    // Graph can return both open ("655360-") and bounded ("655360-1310736")
    // missing ranges. We upload sequentially and cannot reconcile disjoint holes.
    const diagnostic = Array.isArray(ranges) && ranges.length <= 4
      ? ranges.map((range) => typeof range === "string" && /^[0-9-]{1,40}$/.test(range) ? range : "[unrecognized]").join(", ")
      : "[unrecognized]";
    const match = Array.isArray(ranges) && ranges.length === 1 && typeof ranges[0] === "string" && /^(\d+)-(\d*)$/.exec(ranges[0]);
    if (!match) throw new Error(`Microsoft returned unsupported upload ranges (${diagnostic}). Inspect or cancel this transfer.`);
    const offset = Number(match[1]);
    const end = match[2] ? Number(match[2]) : null;
    if (!Number.isSafeInteger(offset) || offset < 0 || offset > size || (end !== null && (!Number.isSafeInteger(end) || end < offset || end >= size))) throw new Error(`Microsoft returned invalid upload progress (${diagnostic}).`);
    return offset;
  }
  acknowledge(transfer, end) {
    const pending = transfer.pending;
    transfer.hash.update(pending.bytes);
    transfer.offset = end;
    transfer.last = { offset: pending.offset, length: pending.bytes.length, hash: pending.hash };
    transfer.pending = null;
    transfer.status = end === transfer.size ? "ready" : "active";
    transfer.uncertainOperation = null;
    if (end === transfer.size) transfer.sourceSha256 = transfer.hash.copy().digest("hex");
  }
  async uploadChunk({ transferId, offset, contentBase64 }) {
    const bytes = Buffer.from(contentBase64, "base64");
    if (!bytes.length || bytes.length > CHUNK_SIZE || bytes.toString("base64") !== contentBase64) throw new Error("Provide canonical base64 of one non-empty upload chunk.");
    return this.withTransfer(transferId, async (transfer) => {
      if (transfer.direction !== "upload") throw new Error("Select an upload transfer.");
      const digest = sha256(bytes);
      if (["active", "ready", "completed"].includes(transfer.status) && transfer.last?.offset === offset && transfer.last.length === bytes.length && transfer.last.hash === digest) return this.describe(transfer);
      if (transfer.status !== "active") throw new Error("Upload is not active. Inspect its status before continuing.");
      const end = offset + bytes.length;
      if (!Number.isSafeInteger(offset) || offset !== transfer.offset || end > transfer.size || (end !== transfer.size && bytes.length % (320 * 1024) !== 0)) throw new Error("Upload sequentially at nextOffset; nonfinal fragments must be multiples of 320 KiB, at most 640 KiB.");
      if (end === transfer.size && transfer.expectedSha256 && transfer.hash.copy().update(bytes).digest("hex") !== transfer.expectedSha256) throw new Error("Source SHA-256 did not match. Cancel the staged upload.");
      if (transfer.itemId) await this.unchanged(transfer);
      transfer.pending = { offset, bytes, hash: digest };
      try {
        const response = await this.sessionRequest(transfer, { method: "PUT", headers: { "Content-Type": "application/octet-stream", "Content-Length": String(bytes.length), "Content-Range": `bytes ${offset}-${end - 1}/${transfer.size}` }, body: bytes });
        if (response.status === 202) {
          const next = this.expectedOffset(response.data, transfer.size);
          if (next !== end) throw new Error("Upload progress did not match the sent chunk. Inspect the transfer status.");
          this.touch(transfer, response.data);
          this.acknowledge(transfer, end);
        } else if ([200, 201].includes(response.status) && end === transfer.size && response.data?.id) {
          this.acknowledge(transfer, end);
          transfer.item = pickItem(response.data);
          this.finishState(transfer, "completed");
        } else throw new Error("Unexpected upload response. Inspect the transfer status.");
        return this.describe(transfer);
      } catch (error) {
        transfer.status = "uncertain";
        transfer.uncertainOperation = "chunk";
        throw error;
      }
    });
  }
  async getTransfer({ transferId }) {
    return this.withTransfer(transferId, async (transfer) => {
      if (transfer.direction !== "upload" || !active(transfer)) return this.describe(transfer);
      const response = await this.sessionRequest(transfer, {}, [404, 410]);
      if ([404, 410].includes(response.status)) {
        this.finishState(transfer, "unknown");
        return { ...this.describe(transfer), instruction: "The session is gone. Inspect the destination; it may have completed or expired. Do not blindly upload again." };
      }
      const offset = this.expectedOffset(response.data, transfer.size);
      this.touch(transfer, response.data);
      if (transfer.uncertainOperation === "commit") return { ...this.describe(transfer), instruction: "Commit outcome remains uncertain. Inspect the destination before any further write." };
      if (transfer.pending) {
        const end = transfer.pending.offset + transfer.pending.bytes.length;
        if (offset === end) this.acknowledge(transfer, end);
        else if (offset === transfer.pending.offset) { transfer.pending = null; transfer.status = "active"; transfer.uncertainOperation = null; }
        else throw new Error("Remote progress cannot be reconciled. Cancel and inspect the destination.");
      } else if (offset !== transfer.offset) throw new Error("Remote upload progress changed unexpectedly. Cancel and inspect the destination.");
      return this.describe(transfer);
    });
  }
  async finishUpload({ transferId }) {
    return this.withTransfer(transferId, async (transfer) => {
      if (transfer.direction !== "upload") throw new Error("Select an upload transfer.");
      if (transfer.status === "completed") return this.describe(transfer);
      if (transfer.status !== "ready" || transfer.offset !== transfer.size) throw new Error("Send all chunks and reconcile any uncertain requests before finishing.");
      if (transfer.expectedSha256 && transfer.sourceSha256 !== transfer.expectedSha256) throw new Error("Source SHA-256 did not match. Cancel the staged upload; it has not been explicitly committed.");
      if (transfer.itemId) await this.unchanged(transfer);
      try {
        const response = await this.sessionRequest(transfer, { method: "POST", headers: { "Content-Length": "0", ...(transfer.itemId ? { "If-Match": transfer.eTag } : {}) } });
        if (![200, 201].includes(response.status) || !response.data?.id) throw new Error("Microsoft did not confirm upload completion. Inspect the destination.");
        transfer.item = pickItem(response.data);
        this.finishState(transfer, "completed");
        return this.describe(transfer);
      } catch (error) { throw this.commitFailure(transfer, error); }
    });
  }
  commitFailure(transfer, error) {
    const status = error.status;
    // An HTTP client error is Microsoft's definite answer that it did not
    // commit. Network failures, timeouts and server errors leave it unknown.
    if (status === 412) {
      this.finishState(transfer, "stale");
      return new Error(`${error.message} The file changed before the commit; Microsoft did not save this upload. Cancel and start again from a fresh read.`, { cause: error });
    }
    if (status === 404 || status === 410) {
      this.finishState(transfer, "unknown");
      return new Error(`${error.message} The session is gone. Inspect the destination; it may have completed or expired. Do not blindly upload again.`, { cause: error });
    }
    if (status === 423 || status === 429) return new Error(`${error.message} Microsoft did not save this upload yet. Wait before finishing again.`, { cause: error });
    if (Number.isInteger(status) && status >= 400 && status < 500 && status !== 408) {
      this.finishState(transfer, "rejected");
      return new Error(`${error.message} Microsoft rejected the commit and did not save this upload. Cancel this transfer.`, { cause: error });
    }
    transfer.status = "uncertain";
    transfer.uncertainOperation = "commit";
    return error;
  }
  async cancelTransfer({ transferId }) {
    return this.withTransfer(transferId, async (transfer) => {
      if (transfer.status === "completed") return { ...this.describe(transfer), instruction: "The file is already saved; cancelling a transfer cannot delete or revert it." };
      if (transfer.uploadUrl) {
        const response = await this.#files.request(transfer.uploadUrl, { method: "DELETE" }, false, { details: true, maxBytes: 65536, allowStatus: [404, 410] });
        if (![204, 404, 410].includes(response.status)) throw new Error("Microsoft did not confirm cancellation. Inspect the transfer before continuing.");
      }
      const uncertain = transfer.status === "uncertain" || transfer.status === "unknown";
      this.finishState(transfer, "cancelled");
      return { ...this.describe(transfer), ...(uncertain ? { instruction: "Cancellation cannot undo a file that may already have been saved. Inspect the destination." } : {}) };
    });
  }
}

function registerTransferTools(server, transfers) {
  const token = z.string().uuid(), integer = z.number().int().min(0).max(Number.MAX_SAFE_INTEGER);
  const register = (name, description, inputSchema, action, writing = false) =>
    registerJsonTool(server, name, { description, inputSchema, writing, openWorld: true }, action);
  register("teams_begin_download", "Start a version-bound chunked file download of any supported provider size. Returns a private transfer ID and 640 KiB chunk size. Does not return temporary download URLs. Sessions expire after 30 idle minutes and are lost on Teams restart.", z.object({ driveId: id, itemId: id }).strict(), (args) => transfers.beginDownload(args));
  register("teams_read_file_chunk", "Read the next 640 KiB byte range as base64 using transferId and nextOffset. The most recent chunk can be reread safely. Checks original eTag/size before and after every range; discard all assembled bytes if the file changes. Returns chunkSha256 and done. This does not load the whole file into the server.", z.object({ transferId: token, offset: integer }).strict(), (args) => transfers.readChunk(args));
  const upload = z.object({ driveId: id, size: integer.min(1), parentId: id.optional(), name: filename.optional(), itemId: id.optional(), expectedETag: eTag.optional(), expectedSha256: z.string().regex(/^[a-f\d]{64}$/i).optional() }).strict().refine((args) => args.itemId ? !!args.expectedETag && !args.parentId && !args.name : !!args.parentId && !!args.name && !args.expectedETag, "Choose either itemId + expectedETag for replacement or parentId + name for creation.");
  register("teams_begin_upload", "Start a resumable upload with a fixed positive total byte size. Create: parentId+name, conflict behavior fail. Replace: itemId+expectedETag from a recent read. Optional expectedSha256 checks assembled source bytes before explicit finish. Requests deferred commit; on a definite HTTP 400 retries session creation without deferred commit. Returns commitMode: deferred requires teams_finish_upload; automatic publishes on the final chunk. eTag checks are not an atomic lock across the upload. Existing Office files should be edited live for coauthoring. Empty files use the single-call tools.", upload, (args) => transfers.beginUpload(args), true);
  register("teams_upload_chunk", "Send one sequential base64 chunk at nextOffset, at most 640 KiB; nonfinal chunks must be multiples of 320 KiB. Total size is fixed. Identical repeat of the last acknowledged chunk returns its receipt without resending. After an error call teams_get_transfer; do not blindly retry. Check commitMode: automatic publishes on the final chunk; deferred normally stages the file for teams_finish_upload. Unexpected provider auto-completion is reported as completed.", z.object({ transferId: token, offset: integer, contentBase64: base64(CHUNK_SIZE).refine((value) => value.length > 0) }).strict(), (args) => transfers.uploadChunk(args), true);
  register("teams_get_transfer", "Inspect transfer progress. Uploads query Microsoft and reconcile uncertain chunk outcomes before resuming. A vanished session or uncertain commit does not prove whether a file was saved; inspect the destination before any new write. Transfer IDs survive MCP reconnects while Teams remains running.", z.object({ transferId: token }).strict(), (args) => transfers.getTransfer(args));
  register("teams_finish_upload", "Explicitly publish a staged upload after all chunks and optional source SHA-256 validation. Replacement checks the original eTag again immediately before committing; no atomic coauthor lock is claimed. Never automatically retry an uncertain commit (network failure, timeout or server error). A definite rejection saves nothing: stale after a changed eTag, rejected otherwise; locked or throttled uploads stay ready to finish later. Microsoft may rewrite Office package metadata, so saved bytes/size can differ from source. Returns projected saved item metadata when confirmed.", z.object({ transferId: token }).strict(), (args) => transfers.finishUpload(args), true);
  register("teams_cancel_transfer", "Cancel a local download or ask Microsoft to delete an unfinished upload session. This does not delete or undo an already saved file, including one whose commit outcome was uncertain. Temporary provider storage otherwise expires automatically.", z.object({ transferId: token }).strict(), (args) => transfers.cancelTransfer(args), true);
}
module.exports = { FileTransfers, registerTransferTools, CHUNK_SIZE };

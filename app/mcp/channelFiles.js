const { z } = require("zod");

const LIMIT = 1024 * 1024;
const BASE = "https://graph.microsoft.com/v1.0";
const itemFields = "id,name,size,eTag,file,folder,parentReference,lastModifiedDateTime";
const pickItem = (data) => Object.fromEntries(itemFields.split(",").filter((key) => key in data).map((key) => [key, data[key]]));
const segment = (value) => encodeURIComponent(value);
const hasControl = (value) => [...value].some((character) => character.charCodeAt(0) < 32);
const id = z.string().min(1).max(256).refine((value) => !hasControl(value) && !/[ /\\?#]/.test(value) && value !== "." && value !== "..");

// Neither Graph tokens nor preauthenticated download URLs cross the MCP boundary.
class ChannelFiles {
  constructor(graph, fetcher = globalThis.fetch) {
    this.graph = graph;
    this.fetch = fetcher;
  }

  async request(endpoint, options = {}, authenticated = true) {
    const url = new URL(endpoint, `${BASE}/`);
    if (authenticated && (url.origin !== "https://graph.microsoft.com" || !url.pathname.startsWith("/v1.0/"))) {
      throw new Error("Invalid Graph endpoint");
    }
    if (url.protocol !== "https:" || url.username || url.password) throw new Error("Invalid transfer URL");
    let token;
    if (authenticated) {
      const result = await this.graph.acquireToken();
      if (!result.success || !result.token) throw new Error("Sign in to Teams before accessing channel files.");
      token = result.token;
    }
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 30000);
    try {
      const response = await this.fetch(url.href, {
        ...options, redirect: "manual", signal: controller.signal,
        headers: { ...options.headers, ...(token ? { Authorization: `Bearer ${token}` } : {}) },
      });
      if (!response.ok && response.status !== 302) {
        await response.body?.cancel();
        const hint = response.status === 403 ? " Teams' session lacks permission, or tenant policy blocks this operation." : "";
        throw new Error(`File request failed (HTTP ${response.status}).${hint}`);
      }
      if (response.status === 302) {
        await response.body?.cancel();
        return { redirect: response.headers.get("location") };
      }
      const chunks = [];
      let size = 0;
      for await (const chunk of response.body || []) {
        size += chunk.length;
        if (size > 2 * LIMIT) throw new Error("Response exceeds the transfer limit.");
        chunks.push(chunk);
      }
      return Buffer.concat(chunks);
    } catch (error) {
      if (/^(File request failed|Response exceeds)/.test(error.message)) throw error;
      // Drop fetch causes because they can contain preauthenticated URLs.
      // eslint-disable-next-line preserve-caught-error
      throw new Error("File transfer failed or timed out. For writes, inspect the destination before retrying.");
    } finally {
      clearTimeout(timeout);
    }
  }

  async json(endpoint, options) {
    const response = await this.request(endpoint, options);
    if (!Buffer.isBuffer(response)) throw new Error("Unexpected Graph redirect");
    try { return JSON.parse(response.toString("utf8")); } catch { throw new Error("Invalid Graph JSON response"); }
  }

  async list(endpoint, cursor) {
    const path = cursor ? `${BASE}${endpoint}${endpoint.includes("?") ? "&" : "?"}$skiptoken=${segment(cursor)}` : `${BASE}${endpoint}`;
    const data = await this.json(path);
    let nextCursor = null;
    if (data["@odata.nextLink"]) {
      const next = new URL(data["@odata.nextLink"]);
      // Opaque Graph skip tokens only; never accept arbitrary URLs from agents.
      nextCursor = next.searchParams.get("$skiptoken");
      if (!nextCursor) throw new Error("Unsupported Graph pagination. Narrow the requested location.");
    }
    const fields = [...itemFields.split(","), "displayName", "description", "membershipType", "isArchived", "tenantId"];
    return { items: data.value.map((item) => Object.fromEntries(fields.filter((key) => key in item).map((key) => [key, item[key]]))), nextCursor };
  }

  itemPath(driveId, itemId) {
    return `${BASE}/drives/${segment(driveId)}/items/${segment(itemId)}`;
  }

  async read({ driveId, itemId }) {
    const path = this.itemPath(driveId, itemId);
    const before = await this.json(`${path}?$select=${itemFields}`);
    if (!before.file) throw new Error("Select a file, not a folder.");
    if (before.size > LIMIT) throw new Error("Files are limited to 1 MiB per transfer.");
    let bytes = await this.request(`${path}/content`);
    if (!Buffer.isBuffer(bytes)) {
      if (!bytes.redirect) throw new Error("Missing download location");
      bytes = await this.request(bytes.redirect, {}, false);
    }
    if (!Buffer.isBuffer(bytes) || bytes.length > LIMIT) throw new Error("File exceeds 1 MiB or returned an unsupported redirect.");
    const after = await this.json(`${path}?$select=${itemFields}`);
    if (!before.eTag || before.eTag !== after.eTag) throw new Error("File changed during download. Read again.");
    return { item: pickItem(after), contentBase64: bytes.toString("base64") };
  }

  async write({ driveId, itemId, parentId, name, expectedETag, contentBase64 }) {
    const bytes = Buffer.from(contentBase64, "base64");
    if (bytes.length > LIMIT || bytes.toString("base64") !== contentBase64) throw new Error("Use canonical base64 of at most 1 MiB.");
    const path = itemId ? this.itemPath(driveId, itemId) : `${this.itemPath(driveId, parentId)}:/${segment(name)}:`;
    if (itemId) {
      const current = await this.json(`${path}?$select=${itemFields}`);
      if (!current.file || current.eTag !== expectedETag) throw new Error("File changed. Read it again before replacing it.");
    }
    const data = await this.json(`${path}/content${itemId ? "" : "?@microsoft.graph.conflictBehavior=fail"}`, {
      method: "PUT", headers: { "Content-Type": "application/octet-stream", ...(itemId ? { "If-Match": expectedETag } : { "If-None-Match": "*" }) }, body: bytes,
    });
    return { item: pickItem(data) };
  }
}

function registerChannelFileTools(server, files) {
  const read = { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: true };
  function register(name, description, shape, operation, write = false) {
    server.registerTool(name, { description, inputSchema: z.object(shape).strict(), annotations: write ? { ...read, readOnlyHint: false, destructiveHint: true, idempotentHint: false } : read }, async (args) => {
      try {
        const value = await operation(args);
        return { content: [{ type: "text", text: JSON.stringify(value) }], structuredContent: value };
      } catch (error) {
        return { isError: true, content: [{ type: "text", text: error.message }] };
      }
    });
  }
  const cursor = z.string().max(8192).optional();
  register("teams_list_teams", "List teams joined by the signed-in Teams user. A 403 may mean this session lacks team discovery permission. Pass nextCursor as cursor to continue.", { cursor }, ({ cursor }) => files.list("/me/joinedTeams", cursor));
  register("teams_list_channels", "List channels of a selected team. Pass nextCursor as cursor to continue.", { teamId: id, cursor }, ({ teamId, cursor }) => files.list(`/teams/${segment(teamId)}/channels?$select=id,displayName,membershipType`, cursor));
  register("teams_get_channel_folder", "Resolve the Files folder of an explicitly selected Teams channel. Returns driveId in parentReference and folder id. No Office add-in is needed.", { teamId: id, channelId: id }, async ({ teamId, channelId }) => {
    const data = await files.json(`${BASE}/teams/${segment(teamId)}/channels/${segment(channelId)}/filesFolder`);
    return { item: pickItem(data) };
  });
  register("teams_list_files", "List children of a selected channel folder or subfolder. Pass nextCursor as cursor to continue.", { driveId: id, folderId: id, cursor }, ({ driveId, folderId, cursor }) => files.list(`/drives/${segment(driveId)}/items/${segment(folderId)}/children?$select=${itemFields}&$top=50`, cursor));
  register("teams_read_file", "Download an explicitly selected file, up to 1 MiB, as base64 with its eTag. Decode as appropriate for its format. Word/Excel editing in an open document is available through the separate Office tools.", { driveId: id, itemId: id }, (args) => files.read(args));
  const contentBase64 = z.string().max(4 * Math.ceil(LIMIT / 3)).regex(/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/);
  register("teams_create_file", "Upload a new file of up to 1 MiB to a selected folder using base64 bytes. Requests conflict failure rather than overwriting an existing name. Do not automatically retry uncertain writes; list and read back first.", { driveId: id, parentId: id, name: z.string().min(1).max(255).refine((value) => !hasControl(value) && !/["*:<>?/\\|]/.test(value) && !/[. ]$/.test(value) && value !== "." && value !== ".."), contentBase64 }, (args) => files.write(args), true);
  register("teams_replace_file", "Replace the entire contents of a selected file with up to 1 MiB of base64 bytes. Requires the exact eTag from a recent read, checks it and sends If-Match. Concurrent-save protection needs live validation; coordinate edits. This replaces the complete file, including Office packages. Do not automatically retry uncertain writes; read back first.", { driveId: id, itemId: id, expectedETag: z.string().min(1).max(256).refine((value) => value !== "*" && !/[\r\n]/.test(value)), contentBase64 }, (args) => files.write(args), true);
}

module.exports = { ChannelFiles, registerChannelFileTools, LIMIT };

const { randomUUID } = require("node:crypto");
const { evaluateInFrame } = require("../mcp/evaluateInFrame");
const { filename } = require("../mcp/channelFileTypes");
const { createBridgeSetupUiLoader } = require("./bridgeSetupUiLoader");

const locationUiScript = createBridgeSetupUiLoader({ filename: require.resolve("./documentLocationUi"), exportName: "documentLocationUi" });
const requestUiScript = createBridgeSetupUiLoader({ filename: require.resolve("./documentRequestUi"), exportName: "documentRequestUi" });

const LIFETIME = 5 * 60 * 1000;
const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const literal = (value) => "'" + encodeURIComponent(value.replace(/'/g, "''")).replace(/'/g, "%27") + "'";
const fields = "Name,ServerRelativeUrl,UniqueId";

function wordDocumentName(value) {
  if (!filename.safeParse(value).success || /\u007f/.test(value)) throw new Error("Use a valid Word document name.");
  const name = /\.docx$/i.test(value) ? value : value + ".docx";
  if (!filename.safeParse(name).success || name.length <= 5 || name.startsWith("~$")) throw new Error("Use a valid Word document name of at most 255 characters.");
  return name;
}

function serverPath(value) {
  if (typeof value !== "string" || value.length > 4096 || !value.startsWith("/") || value.startsWith("//") ||
    value.includes("\\") || [...value].some((character) => character.charCodeAt(0) < 32 || character.charCodeAt(0) === 127) ||
    value.split("/").some((part) => part === "." || part === "..")) return null;
  return value.replace(/\/$/, "") || "/";
}

function sharepointPage(value) {
  try {
    const url = new URL(value);
    if (url.protocol !== "https:" || url.username || url.password || url.port || !url.hostname.endsWith(".sharepoint.com")) return null;
    const path = decodeURIComponent(url.pathname.replace(/^\/:[wx]:\/r(?=\/)/i, ""));
    const webPath = path.match(/^(?:\/(?:sites|teams|personal)\/[^/]+)/i)?.[0] || "";
    const layout = path.toLowerCase().indexOf("/_layouts/");
    const webUrl = url.origin + (layout >= 0 ? path.slice(0, layout) : webPath).split("/").map(encodeURIComponent).join("/");
    return { url, path, origin: url.origin, webUrl, source: "sharepoint_view" };
  } catch { return null; }
}

function locationSource(value) {
  try {
    const page = sharepointPage(value);
    if (!page) return null;
    const { url, path, webUrl } = page;
    if (/\/(?:allitems|onedrive)\.aspx$/i.test(path)) {
      const folder = url.searchParams.get("id") || url.searchParams.get("RootFolder") || url.searchParams.get("rootFolder");
      const folderPath = serverPath(folder || (/\/Forms\/AllItems\.aspx$/i.test(path) ? path.replace(/\/Forms\/AllItems\.aspx$/i, "") : ""));
      return folderPath ? { origin: url.origin, webUrl, folderPath, source: "folder" } : null;
    }
    if (/\/_layouts\/15\/(?:doc|wopiframe)\.aspx$/i.test(path)) {
      const fileId = (url.searchParams.get("sourcedoc") || "").replace(/^\{|\}$/g, "");
      return guid.test(fileId) ? { origin: url.origin, webUrl, fileId, source: "document_parent" } : null;
    }
    if (/\.(?:docx|docm|doc|xlsx|xlsm|xlsb|xls)$/i.test(path) && serverPath(path)) {
      return { origin: url.origin, webUrl, filePath: path, source: "document_parent" };
    }
  } catch { /* Unsupported or detached page: do not guess its destination. */ }
  return null;
}

function fileMetadata(data, origin, folderPath) {
  const path = serverPath(data?.ServerRelativeUrl);
  if (!path || (path.slice(0, path.lastIndexOf("/")) || "/") !== folderPath || typeof data.Name !== "string" ||
    data.Name.length > 255 || path.slice(path.lastIndexOf("/") + 1) !== data.Name) throw new Error("INVALID_SHAREPOINT_RESPONSE");
  const url = new URL(origin);
  url.pathname = path.split("/").map(encodeURIComponent).join("/");
  url.searchParams.set("web", "1");
  return { name: data.Name, url: url.href, ...(guid.test(data.UniqueId) ? { fileId: data.UniqueId } : {}) };
}

class DocumentLocations {
  #getWebContents;
  #locations = new Map();
  #closed = false;

  constructor(getWebContents) { this.#getWebContents = getWebContents; }

  #current(location) {
    try { return !this.#closed && Date.now() < location.expires && !location.contents.isDestroyed() &&
      location.frame.url === location.frameUrl && location.contents.mainFrame.framesInSubtree.includes(location.frame); }
    catch { return false; }
  }

  async #json(location, endpoint, options = {}, writing = false) {
    if (!this.#current(location)) throw new Error("LOCATION_CHANGED. Call office_list_locations again.");
    try {
      const url = new URL(location.webUrl + "/_api/" + endpoint);
      if (url.origin !== location.origin) throw new Error("INVALID_LOCATION");
      const response = await evaluateInFrame(location.frame, requestUiScript({ pageUrl: location.frameUrl, webUrl: location.webUrl,
        url: url.href, deadline: Date.now() + 15000, method: options.method || "GET",
        headers: { Accept: "application/json;odata=nometadata", ...options.headers },
        ...(options.body ? { bodyBase64: options.body.toString("base64") } : {}) }), 16000, "SharePoint request timed out");
      if (response?.code) throw new Error(writing ? "CREATION_OUTCOME_UNKNOWN" : response.code);
      if (!Number.isInteger(response?.status)) throw new Error("SHAREPOINT_REQUEST_UNAVAILABLE");
      if (response.status < 200 || response.status >= 300) {
        if (response.status >= 300 && response.status < 400 || response.status === 401) throw new Error("SIGN_IN_REQUIRED");
        if (response.status === 403) throw new Error("SHAREPOINT_ACCESS_DENIED");
        if (response.status === 409) throw new Error("DOCUMENT_NAME_EXISTS. List the folder before choosing another name.");
        throw new Error(writing && response.status >= 500 ? "CREATION_OUTCOME_UNKNOWN" : "SHAREPOINT_REQUEST_FAILED");
      }
      let data;
      try { data = JSON.parse(response.text); }
      catch { throw new Error(writing ? "CREATION_OUTCOME_UNKNOWN" : "SIGN_IN_REQUIRED"); }
      return data.d || data;
    } catch (error) {
      const allowed = ["SIGN_IN_REQUIRED", "SHAREPOINT_ACCESS_DENIED", "DOCUMENT_NAME_EXISTS", "CREATION_OUTCOME_UNKNOWN", "SHAREPOINT_REQUEST_FAILED", "SHAREPOINT_RESPONSE_TOO_LARGE"];
      const code = allowed.find((value) => error.message?.startsWith(value));
      // eslint-disable-next-line preserve-caught-error -- Causes may contain credentials or server text.
      if (code) throw new Error(code + (writing ? ". Do not retry creation automatically; use office_list_location_files to inspect the destination." : ""));
      // Never forward fetch causes, server text, cookie values or request digests.
      // eslint-disable-next-line preserve-caught-error
      throw new Error(writing ? "CREATION_OUTCOME_UNKNOWN. Use office_list_location_files before retrying." : "SHAREPOINT_REQUEST_UNAVAILABLE");
    }
  }

  #folderEndpoint(path) { return "web/GetFolderByServerRelativePath(decodedurl=" + literal(path) + ")"; }

  async list() {
    for (const [key, value] of this.#locations) if (!this.#current(value)) this.#locations.delete(key);
    const result = [], unavailable = []; let visited = 0, truncated = false;
    const seen = [];
    for (const contents of this.#getWebContents()) {
      if (this.#closed) break;
      try {
        if (contents.isDestroyed()) continue;
        for (const frame of contents.mainFrame.framesInSubtree) {
          const source = locationSource(frame.url) || sharepointPage(frame.url);
          if (!source) continue;
          if (++visited > 16) { truncated = true; break; }
          const frameUrl = frame.url;
          const candidate = { ...source, contents, frame, frameUrl, expires: Date.now() + LIFETIME };
          let page = {};
          try {
            try { page = await evaluateInFrame(frame, locationUiScript({ url: frameUrl, deadline: Date.now() + 1000 }), 1500, "Page context unavailable") || {}; }
            catch { /* Fall back to recognized URL formats, never a guessed folder. */ }
            const web = page.webUrl;
            if (web) {
              const parsed = new URL(web);
              if (parsed.origin !== candidate.origin || parsed.username || parsed.password || parsed.port || parsed.search || parsed.hash || !serverPath(decodeURIComponent(parsed.pathname))) throw new Error("INVALID_LOCATION");
              candidate.webUrl = parsed.href.replace(/\/$/, "");
            }
            if (!this.#current(candidate)) continue;
            let path = source.folderPath || serverPath(page.folderPath);
            if (!path) {
              if (!source.fileId && !source.filePath) throw new Error("FOLDER_NOT_IDENTIFIED");
              const endpoint = source.fileId ? "web/GetFileById(" + literal(source.fileId) + ")" : "web/GetFileByServerRelativePath(decodedurl=" + literal(source.filePath) + ")";
              const file = await this.#json(candidate, endpoint + "?$select=" + fields);
              const filePath = serverPath(file.ServerRelativeUrl);
              if (!filePath) throw new Error("INVALID_LOCATION");
              path = filePath.slice(0, filePath.lastIndexOf("/"));
            }
            const webPath = decodeURIComponent(new URL(candidate.webUrl).pathname).replace(/\/$/, "");
            if (!serverPath(path) || webPath && !path.startsWith(webPath + "/")) throw new Error("INVALID_LOCATION");
            const folder = await this.#json(candidate, this.#folderEndpoint(path) + "?$select=" + fields);
            if (serverPath(folder.ServerRelativeUrl) !== path || !this.#current(candidate)) throw new Error("INVALID_LOCATION");
            const existing = [...this.#locations.values()].find((location) => location.contents === contents && location.frame === frame && location.frameUrl === frameUrl && location.folderPath === path);
            const locationId = existing?.locationId || randomUUID();
            const sourceKind = source.source === "sharepoint_view" ? "folder" : source.source;
            const location = { ...candidate, source: sourceKind, folderPath: path, locationId, name: String(folder.Name || path.split("/").at(-1)).slice(0, 255) };
            this.#locations.set(locationId, location);
            if (seen.some((item) => item.session === contents.session && item.origin === candidate.origin && item.path === path)) continue;
            seen.push({ session: contents.session, origin: candidate.origin, path });
            result.push({ locationId, name: location.name, origin: location.origin, folderPath: path, source: location.source, webContentsId: contents.id });
          } catch (error) {
            const code = ["SIGN_IN_REQUIRED", "SHAREPOINT_ACCESS_DENIED", "SHAREPOINT_REQUEST_UNAVAILABLE", "FOLDER_NOT_IDENTIFIED"].includes(error.message) ? error.message : "LOCATION_UNAVAILABLE";
            const kinds = ["allitems", "listallitems", "onedrive", "doclibembed", "filebrowser", "teamslogon", "doc", "wopiframe", "other"];
            const keys = ["id", "RootFolder", "rootFolder", "folder", "folderPath", "folderUrl", "listUrl", "dest", "listId", "siteId"];
            unavailable.push({ source: source.source, origin: source.origin, webContentsId: contents.id, code,
              pageKind: kinds.includes(page.pageKind) ? page.pageKind : "unavailable",
              pageName: typeof page.pageName === "string" && /^[a-z][a-z0-9_-]{0,63}\.aspx$/i.test(page.pageName) ? page.pageName : null,
              queryKeys: keys.filter((key) => Array.isArray(page.queryKeys) && page.queryKeys.includes(key)), hasPageContext: page.hasPageContext === true });
          }
        }
      } catch { /* Closed contents: keep other available locations. */ }
      if (truncated) break;
    }
    return { locations: result, unavailable, truncated,
      ...(!result.length ? { next: "Open a SharePoint folder or an Office document in Teams, then call office_list_locations again." } : {}) };
  }

  async #resolve(locationId) {
    if (!locationId) {
      const report = await this.list();
      if (report.locations.length !== 1 || report.truncated || report.unavailable.length) throw new Error("Select a locationId from office_list_locations; the current folder is missing or ambiguous.");
      locationId = report.locations[0].locationId;
    }
    const location = this.#locations.get(locationId);
    if (!location || !this.#current(location)) throw new Error("LOCATION_CHANGED. Call office_list_locations again.");
    return location;
  }

  async listFiles(locationId, name) {
    if (name !== undefined && !filename.safeParse(name).success) throw new Error("Use an exact file name.");
    const location = await this.#resolve(locationId);
    const data = await this.#json(location, this.#folderEndpoint(location.folderPath) + "/Files?$select=" + fields + "&$top=50" +
      (name === undefined ? "" : "&$filter=Name%20eq%20" + literal(name)));
    const values = data.value || data.results;
    if (!Array.isArray(values) || !this.#current(location)) throw new Error("LOCATION_CHANGED. Call office_list_locations again.");
    return { locationId: location.locationId, files: values.slice(0, 50).map((file) => fileMetadata(file, location.origin, location.folderPath)),
      truncated: values.length > 50 || !!data["@odata.nextLink"] || !!data["odata.nextLink"] || !!data.__next };
  }

  async create(locationId, name) {
    name = wordDocumentName(name);
    const location = await this.#resolve(locationId);
    const context = await this.#json(location, "contextinfo", { method: "POST" });
    const digest = (context.GetContextWebInformation || context).FormDigestValue;
    if (typeof digest !== "string" || !digest || digest.length > 16384) throw new Error("SHAREPOINT_REQUEST_UNAVAILABLE");
    if (!this.#current(location)) throw new Error("LOCATION_CHANGED. Call office_list_locations again.");
    const { readFile } = require("node:fs/promises");
    const body = await readFile(require("node:path").join(__dirname, "templates", "blank.docx"));
    const data = await this.#json(location, this.#folderEndpoint(location.folderPath) + "/Files/AddUsingPath(decodedurl=" + literal(name) + ",overwrite=false)?$select=" + fields,
      { method: "POST", headers: { "X-RequestDigest": digest, "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }, body }, true);
    let item;
    try { item = fileMetadata(data, location.origin, location.folderPath); if (item.name !== name) throw new Error("Unexpected file name"); }
    catch { throw new Error("CREATION_OUTCOME_UNKNOWN. Use office_list_location_files before retrying."); }
    return { item, contents: location.contents, locationId: location.locationId };
  }

  close() { this.#closed = true; this.#locations.clear(); }
}

module.exports = { DocumentLocations, locationSource, wordDocumentName };

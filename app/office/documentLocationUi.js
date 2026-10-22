// Fixed read-only SharePoint location probe. Only public page context and
// known folder-address parameters are read; no content, storage or credentials.
function documentLocationUi(request) {
  if (location.href !== request.url || Date.now() >= request.deadline || location.protocol !== "https:" ||
    !location.hostname.endsWith(".sharepoint.com")) return null;
  let url = new URL(location.href);
  const kinds = ["allitems", "listallitems", "onedrive", "doclibembed", "filebrowser", "teamslogon", "doc", "wopiframe"];
  const keys = ["id", "RootFolder", "rootFolder", "folder", "folderPath", "folderUrl", "listUrl", "dest", "listId", "siteId"];
  const kind = (target) => {
    const name = target.pathname.split("/").at(-1).replace(/\.aspx$/i, "").toLowerCase();
    return kinds.includes(name) ? name : "other";
  };
  const pageKind = kind(url), queryKeys = keys.filter((key) => url.searchParams.has(key));
  const leaf = url.pathname.split("/").at(-1);
  const pageName = /^[a-z][a-z0-9_-]{0,63}\.aspx$/i.test(leaf) ? leaf : null;
  const context = globalThis._spPageContextInfo;
  let webUrl = typeof context?.webAbsoluteUrl === "string" && context.webAbsoluteUrl.length <= 4096 ? context.webAbsoluteUrl : null;
  for (let depth = 0; depth < 3 && kind(url) === "teamslogon"; depth++) {
    try {
      const dest = new URL(url.searchParams.get("dest"), url);
      if (dest.origin !== location.origin || dest.username || dest.password || dest.port) break;
      url = dest;
      if (!webUrl) {
        const layout = url.pathname.toLowerCase().indexOf("/_layouts/");
        if (layout >= 0) webUrl = url.origin + url.pathname.slice(0, layout);
      }
    } catch { break; }
  }
  let folderPath = null;
  if (["allitems", "listallitems", "onedrive", "doclibembed", "filebrowser"].includes(kind(url))) {
    // listallitems.aspx is the library view embedded by Teams. listUrl is
    // its root only; explicit subfolder parameters take precedence.
    const folder = ["id", "RootFolder", "rootFolder", "folder", "folderPath", "folderUrl"].map((key) => url.searchParams.get(key)).find((value) => value?.startsWith("/"));
    folderPath = folder || (kind(url) === "listallitems" ? url.searchParams.get("listUrl") : null);
  }
  if (kind(url) === "filebrowser" && folderPath) {
    // Teams may host filebrowser.aspx on the tenant root while its id points
    // into another site. The hosting page's web context is not that folder's
    // REST web. Keep a matching subweb; otherwise use the addressed site.
    const site = folderPath.match(/^\/(?:sites|teams|personal)\/[^/]+/i)?.[0];
    if (site) {
      let matches = false;
      try {
        const contextUrl = new URL(webUrl);
        const contextPath = decodeURIComponent(contextUrl.pathname).replace(/\/$/, "");
        if (contextUrl.origin !== location.origin) return { webUrl, folderPath, pageKind, pageName, queryKeys, hasPageContext: !!context };
        matches = (contextPath === site || contextPath.startsWith(site + "/")) && folderPath.startsWith(contextPath + "/");
      } catch { /* No public web context: derive only the explicitly addressed site. */ }
      if (!matches) webUrl = location.origin + site.split("/").map(encodeURIComponent).join("/");
    }
  }
  return { webUrl, folderPath, pageKind, pageName, queryKeys, hasPageContext: !!context };
}

module.exports = { documentLocationUi };

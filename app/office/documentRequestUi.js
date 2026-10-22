// App-owned SharePoint REST transport. Requests execute in the originating
// browser frame so its normal request context and credentials apply. No MCP
// caller supplies URLs, headers, JavaScript or document bytes to this function.
async function documentRequestUi(request) {
  if (location.href !== request.pageUrl || Date.now() >= request.deadline) return { code: "LOCATION_CHANGED" };
  let url, web;
  try { url = new URL(request.url); web = new URL(request.webUrl); } catch { return { code: "INVALID_LOCATION" }; }
  if (location.protocol !== "https:" || !location.hostname.endsWith(".sharepoint.com") ||
    url.origin !== location.origin || web.origin !== location.origin || web.username || web.password || web.port ||
    url.username || url.password || url.port || web.search || web.hash ||
    !url.pathname.startsWith(web.pathname.replace(/\/$/, "") + "/_api/") || !["GET", "POST"].includes(request.method)) return { code: "INVALID_LOCATION" };
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), Math.max(1, request.deadline - Date.now()));
  try {
    const body = request.bodyBase64 === undefined ? undefined : Uint8Array.from(atob(request.bodyBase64), (character) => character.charCodeAt(0));
    const response = await fetch(url.href, { method: request.method, headers: request.headers, body,
      credentials: "include", mode: "same-origin", redirect: "error", cache: "no-store", signal: controller.signal });
    const chunks = []; let size = 0;
    const reader = response.body?.getReader();
    if (reader) {
      while (true) {
        const chunk = await reader.read();
        if (chunk.done) break;
        size += chunk.value.byteLength;
        if (size > 256 * 1024) { await reader.cancel(); return { code: "SHAREPOINT_RESPONSE_TOO_LARGE" }; }
        chunks.push(chunk.value);
      }
    }
    const bytes = new Uint8Array(size); let offset = 0;
    for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.byteLength; }
    const text = new TextDecoder().decode(bytes);
    // Error bodies can echo tokens, URLs or filenames. Return only recognized
    // categories, never their raw text or response headers.
    if (!response.ok) {
      let errorCode;
      try {
        const data = JSON.parse(text), code = data.error?.code || data["odata.error"]?.code;
        if (typeof code === "string") {
          if (code.endsWith("System.UnauthorizedAccessException")) errorCode = "access_denied";
          else if (code.endsWith("Microsoft.SharePoint.SPException")) errorCode = "sharepoint_error";
        }
      } catch { /* Non-JSON failures carry only the HTTP status. */ }
      return { status: response.status, ...(errorCode ? { errorCode } : {}) };
    }
    return { status: response.status, text };
  } catch { return { code: "SHAREPOINT_REQUEST_UNAVAILABLE" }; }
  finally { clearTimeout(timer); }
}

module.exports = { documentRequestUi };

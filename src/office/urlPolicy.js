// Shared Microsoft URL checks for the main process. documentAccountUi keeps
// a copy for hot-reloaded renderer probes; urlPolicy.test.js compares them.
function officeUrl(value, kind = "https") {
  try {
    const url = new URL(value);
    if (url.protocol !== "https:" || url.username || url.password || url.port) return null;
    const login = ["login.microsoftonline.com", "login.microsoft.com", "login.live.com"].includes(url.hostname);
    const editor = url.hostname.endsWith(".officeapps.live.com") || ["excel.cloud.microsoft", "word.cloud.microsoft"].includes(url.hostname);
    if (kind === "https" || kind === "login" && login ||
      kind === "sharepoint" && url.hostname.endsWith(".sharepoint.com") ||
      kind === "editor" && editor ||
      kind === "setup" && (editor || ["store.office.com", "res.public.onecdn.static.microsoft"].includes(url.hostname)) ||
      kind === "navigation" && (login || editor || url.hostname === "m365.cloud.microsoft")) return url;
  } catch { /* Malformed URLs fail closed in Node and renderer probes. */ }
  return null;
}

module.exports = { officeUrl };

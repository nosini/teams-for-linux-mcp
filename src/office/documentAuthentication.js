const { officeUrl } = require("./urlPolicy");
// Read-only diagnostics for the selected document's session. Never return
// cookie values, account identifiers, document URLs or profile paths.
function cookieAttributes(cookie, origin) {
  const host = new URL(origin).hostname;
  const domain = typeof cookie.domain === "string" ? cookie.domain.replace(/^\./, "").toLowerCase() : "";
  const flag = (value) => typeof value === "boolean" ? value : null;
  return {
    persistence: cookie.session ? "session" : "persistent",
    scope: !domain ? "unknown" : domain === host ? "opened_host" : host.endsWith(`.${domain}`) ? "parent_domain" : "other",
    secure: flag(cookie.secure),
    httpOnly: flag(cookie.httpOnly),
    rootPath: typeof cookie.path === "string" ? cookie.path === "/" : null,
    partitioned: cookie.partitionKey !== undefined,
  };
}

async function inspectDocumentAuthentication(session, documentUrl) {
  let persistentSession = null;
  try { persistentSession = session.isPersistent(); } catch { /* Test doubles or unavailable session. */ }
  let sharepointOrigin;
  try {
    sharepointOrigin = officeUrl(documentUrl, "sharepoint")?.origin;
  } catch { /* Cloud document links do not identify a SharePoint tenant. */ }
  const status = async (url, name, attributes = false) => {
    if (!url) return { state: "unavailable", attributes: null };
    let timer;
    try {
      const cookies = await Promise.race([
        session.cookies.get({ url, name }),
        new Promise((_, reject) => { timer = setTimeout(() => reject(new Error("Cookie inspection timed out")), 1500); }),
      ]);
      const sessionOnly = cookies.some((cookie) => cookie.session);
      const persistent = cookies.some((cookie) => !cookie.session);
      return {
        state: !cookies.length ? "missing" : sessionOnly && persistent ? "mixed" : sessionOnly ? "session" : "persistent",
        attributes: attributes ? { entries: cookies.slice(0, 8).map((cookie) => cookieAttributes(cookie, url)), truncated: cookies.length > 8 } : null,
      };
    } catch { return { state: "unavailable", attributes: null }; }
    finally { clearTimeout(timer); }
  };
  const [microsoftSession, microsoftPersistent, sharepointSite, sharepointRoot] = await Promise.all([
    status("https://login.microsoftonline.com/", "ESTSAUTH"),
    status("https://login.microsoftonline.com/", "ESTSAUTHPERSISTENT"),
    status(sharepointOrigin, "FedAuth", true), status(sharepointOrigin, "rtFa", true),
  ]);
  return { persistentSession, cookies: { microsoftSession: microsoftSession.state, microsoftPersistent: microsoftPersistent.state,
    sharepointSite: sharepointSite.state, sharepointRoot: sharepointRoot.state },
  cookieAttributes: { sharepointSite: sharepointSite.attributes, sharepointRoot: sharepointRoot.attributes } };
}

module.exports = { inspectDocumentAuthentication };

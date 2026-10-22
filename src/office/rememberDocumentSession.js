const DAY = 24 * 60 * 60;

// Opt-in browser persistence for opened SharePoint document sites and their
// shared rtFa root cookie. Values stay in Electron's cookie store, never in app
// settings, MCP output or logs. Microsoft still validates the signed session.
class RememberDocumentSession {
  #sessions = new Map();
  #days;
  #closed = false;

  constructor({ days = 7 } = {}) {
    this.#days = Number.isFinite(days) && days >= 1 && days <= 30 ? days : 7;
  }

  #key(cookie) { return `${cookie.domain}|${cookie.name}`; }

  #matches(state, cookie) {
    const domain = (cookie.domain || "").replace(/^\./, "").toLowerCase();
    // rtFa is also issued on the shared SharePoint parent domain. Do not
    // generalize this exception to FedAuth or arbitrary parent domains.
    const matchingHost = state.hosts.has(domain) || (cookie.name === "rtFa" && domain === "sharepoint.com" &&
      cookie.hostOnly === false && state.hosts.size > 0);
    return cookie.secure === true && cookie.httpOnly === true && cookie.path === "/" &&
      cookie.partitionKey === undefined && matchingHost &&
      (/^FedAuth(?:[1-9]\d?)?$/.test(cookie.name) || cookie.name === "rtFa");
  }

  watch(session, documentUrl) {
    if (this.#closed || !session.isPersistent()) return;
    let host;
    try {
      const url = new URL(documentUrl);
      if (url.protocol !== "https:" || url.username || url.password || url.port || !url.hostname.endsWith(".sharepoint.com")) return;
      host = url.hostname;
    } catch { return; }
    let state = this.#sessions.get(session);
    if (!state) {
      state = { hosts: new Set(), versions: new Map(), pending: new Map(), removed: new Map(), worker: null, closed: false, scans: new Set() };
      state.listener = (_event, cookie, cause, removed) => {
        if (state.closed || !this.#matches(state, cookie)) return;
        const key = this.#key(cookie);
        state.versions.set(key, (state.versions.get(key) || 0) + 1);
        if (removed) {
          state.pending.delete(key);
          if (cause !== "overwrite") state.removed.set(key, state.versions.get(key));
          return;
        }
        if (cookie.session) this.#enqueue(session, state, cookie);
      };
      session.cookies.on("changed", state.listener);
      this.#sessions.set(session, state);
    }
    if (state.hosts.has(host)) return;
    state.hosts.add(host);
    const scan = session.cookies.get({ url: `https://${host}/` }).then((cookies) => {
      if (state.closed) return;
      for (const cookie of cookies) if (cookie.session && this.#matches(state, cookie)) this.#enqueue(session, state, cookie);
    }).catch(() => { /* Leave normal login untouched if cookie access fails. */ }).finally(() => state.scans.delete(scan));
    state.scans.add(scan);
  }

  #enqueue(session, state, cookie) {
    const key = this.#key(cookie);
    state.pending.set(key, { cookie, version: state.versions.get(key) || 0 });
    this.#start(session, state);
  }

  #start(session, state) {
    if (!state.worker) {
      state.worker = this.#drain(session, state).finally(() => {
        state.worker = null;
        if (!state.closed && state.pending.size) this.#start(session, state);
      });
    }
  }

  async #drain(session, state) {
    while (!state.closed && state.pending.size) {
      const [key, { cookie, version }] = state.pending.entries().next().value;
      state.pending.delete(key);
      const host = cookie.domain.replace(/^\./, "");
      const url = `https://${host}/`;
      try {
        // Never restore a cookie that disappeared or rotated during a scan.
        const current = (await session.cookies.get({ url, name: cookie.name })).find((item) =>
          item.domain === cookie.domain && item.path === cookie.path);
        if (state.closed || !current?.session || current.value !== cookie.value ||
          (state.versions.get(key) || 0) !== version) continue;
        const expiry = Math.floor(Date.now() / 1000) + this.#days * DAY;
        const removedAt = state.removed.get(key) || 0;
        const details = { url, name: current.name, value: current.value, path: current.path,
          secure: current.secure, httpOnly: current.httpOnly, sameSite: current.sameSite, expirationDate: expiry };
        if (current.hostOnly === false || current.domain.startsWith(".")) details.domain = current.domain;
        await session.cookies.set(details);
        // An explicit sign-out while set() was pending must win. Remove only
        // our exact promotion, never a newer login or a server-issued cookie.
        if ((state.removed.get(key) || 0) > removedAt) {
          const after = (await session.cookies.get({ url, name: cookie.name })).find((item) => item.domain === cookie.domain && item.path === cookie.path);
          if (after?.value === current.value && after.expirationDate === expiry) await session.cookies.remove(url, cookie.name);
        }
        await session.cookies.flushStore();
      } catch { /* No retries or raw cookie-store errors containing secrets. */ }
    }
  }

  async settle() {
    for (const state of this.#sessions.values()) {
      await Promise.allSettled([...state.scans]);
      while (state.worker) await state.worker;
    }
  }

  close() {
    this.#closed = true;
    for (const [session, state] of this.#sessions) {
      state.closed = true;
      state.pending.clear();
      session.cookies.removeListener("changed", state.listener);
    }
    this.#sessions.clear();
  }
}

module.exports = { RememberDocumentSession };

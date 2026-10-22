// Fixed Microsoft account-picker command. The chosen username stays private
// to this login page and the local app; no caller-provided selectors or code.
function documentAccountUi(request) {
  if (location.href !== request.url || Date.now() >= request.deadline ||
    location.protocol !== "https:" || !["login.microsoftonline.com", "login.microsoft.com", "login.live.com"].includes(location.hostname)) return { performed: false, code: "ACCOUNT_PAGE_CHANGED" };
  const visible = (element) => {
    if (!element || getComputedStyle(element).visibility !== "visible") return false;
    for (let node = element; node; node = node.parentElement) {
      if (node.hidden || node.getAttribute("aria-hidden") === "true" || getComputedStyle(node).display === "none") return false;
    }
    return true;
  };
  // Only presence is checked. Never read, fill or submit credential fields.
  const credentials = [...document.querySelectorAll('input[type=password],input[name=loginfmt],input[type=email],input[name=otc],input[autocomplete=one-time-code]')].some(visible);
  const holder = document.getElementById("tilesHolder");
  const tiles = holder && visible(holder) ? [...holder.querySelectorAll('button,[role=button],.tile[tabindex],a')].filter((element) =>
    visible(element) && !element.disabled && element.getAttribute("aria-disabled") !== "true") : [];
  const picker = !credentials && !!holder && visible(holder);
  const normalize = (value) => typeof value === "string" ? value.trim().toLowerCase() : "";
  const validIdentifier = (value) => typeof value === "string" && value.length <= 320 && /^[^\s@]+@[^\s@]+$/.test(value) &&
    [...value].every((character) => character.charCodeAt(0) >= 32 && character.charCodeAt(0) !== 127);
  let scope;
  try {
    // Older running builds can install the passive recorder through their
    // hot-reloaded probe before restarting, using Microsoft's redirect target.
    const target = new URL(request.scope || new URL(location.href).searchParams.get("redirect_uri"));
    if (target.protocol === "https:" && !target.username && !target.password && !target.port &&
      (target.hostname.endsWith(".sharepoint.com") || ["word.cloud.microsoft", "excel.cloud.microsoft", "m365.cloud.microsoft"].includes(target.hostname))) scope = target.origin;
  } catch { /* No recognized document site: do not guess a storage scope. */ }
  const key = scope ? `teams-office.document-account.v1:${scope}` : null;
  if (request.type === "remembered") {
    if (!key) return null;
    try { const value = localStorage.getItem(key); return validIdentifier(value) ? normalize(value) : null; } catch { return null; }
  }
  if (picker && key && !globalThis.__teamsOfficeAccountRecorder) {
    const recorder = { holder, key };
    globalThis.__teamsOfficeAccountRecorder = recorder;
    const remember = (event) => {
      if (!event.isTrusted || !(event.target instanceof Element) || (event.type === "keydown" && event.key !== "Enter" && event.key !== " ")) return;
      if ([...document.querySelectorAll('input[type=password],input[name=loginfmt],input[type=email],input[name=otc],input[autocomplete=one-time-code]')].some(visible)) return;
      const target = event.target.closest('button,[role=button],.tile[tabindex],a');
      if (!target || !recorder.holder.contains(target) || !visible(target) || target.disabled || target.getAttribute("aria-disabled") === "true") return;
      const values = new Set([target, ...target.querySelectorAll("*")].slice(0, 128).flatMap((node) =>
        [node.getAttribute("data-test-id"), node.getAttribute("data-username"), node.getAttribute("title"), node.textContent]).map(normalize).filter(validIdentifier));
      if (values.size !== 1) return;
      try { localStorage.setItem(recorder.key, [...values][0]); } catch { /* Storage blocked: ordinary manual login still works. */ }
    };
    document.addEventListener("click", remember, true);
    document.addEventListener("keydown", remember, true);
  }
  if (picker && key && globalThis.__teamsOfficeAccountRecorder) Object.assign(globalThis.__teamsOfficeAccountRecorder, { holder, key });
  // Old main-process builds queried Teams tokens on picker:true. Suppress
  // that removed path until they restart with the remembered-choice handler.
  if (request.type === "probe") return { picker: picker && request.useRememberedAccount === true, tileCount: Math.min(tiles.length, 17), credentials };
  if (request.type !== "focus" || !picker || tiles.length > 16) return { performed: false, code: "ACCOUNT_PICKER_UNAVAILABLE" };
  const identifiers = new Set(request.identifiers.map(normalize));
  // Match complete identifiers in known attributes or a bounded text node,
  // never substrings (a@example.test must not match ba@example.test).
  const matches = tiles.filter((tile) => [tile, ...tile.querySelectorAll("*")].slice(0, 128).some((node) =>
    [node.getAttribute("data-test-id"), node.getAttribute("data-username"), node.getAttribute("title"), node.getAttribute("aria-label"), node.textContent].some((value) => identifiers.has(normalize(value)))));
  if (matches.length !== 1) return { performed: false, code: matches.length ? "ACCOUNT_AMBIGUOUS" : "ACCOUNT_NOT_FOUND" };
  const target = matches[0];
  if (target.closest('[contenteditable=true]')) return { performed: false, code: "ACCOUNT_PICKER_UNAVAILABLE" };
  target.focus({ preventScroll: true });
  return document.activeElement === target ? { performed: true } : { performed: false, code: "ACCOUNT_NOT_FOCUSED" };
}

module.exports = { documentAccountUi };

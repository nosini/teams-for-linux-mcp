// Serialized into a selected Office frame. This operates only on named add-in
// controls, never document text, private Office APIs or caller-supplied scripts.
async function bridgeSetupUi(command) {
  if (location.href !== command.url || Date.now() >= command.deadline) return { expired: true };
  const normalize = (value) => (value || "").toLowerCase().replace(/[\u2010-\u2015]/g, "-").replace(/\s+/g, " ").replace(/\s*\([^)]*\)$/, "").replace(/[….]+$/, "").trim();
  const parent = (element) => element.parentElement || element.getRootNode()?.host;
  const within = (element, ancestor) => {
    for (let current = element; current; current = parent(current)) if (current === ancestor) return true;
    return false;
  };
  const excluded = (element, selector) => {
    for (let current = element; current; current = parent(current)) if (current.matches(selector)) return true;
    return false;
  };
  const label = (element) => normalize(element.getAttribute("aria-label") || element.getAttribute("title") ||
    (element.getAttribute("aria-labelledby") || "").split(/\s+/).map((id) => (element.getRootNode().getElementById?.(id) || document.getElementById(id))?.textContent || "").join(" ").trim() || (element.matches('input[type="button"],input[type="submit"]') ? element.value : element.textContent));
  // Chromium can skip layout in a hidden cross-origin child frame, yielding
  // no client rects even for its active controls. Check CSS/ARIA state instead
  // of requiring on-screen geometry, which would force us to show the tab.
  const visible = (element) => {
    if (excluded(element, '[hidden],[aria-hidden="true"],[contenteditable="true"],[role="grid"],[role="document"],[role="textbox"],[role="combobox"]')) return false;
    // Visibility is inherited but descendants may override it. Office mounts
    // its visible Add-ins menu inside visibility:hidden portal containers.
    if (["hidden", "collapse"].includes(getComputedStyle(element).visibility)) return false;
    for (let current = element; current; current = parent(current)) {
      const style = getComputedStyle(current);
      if (style.display === "none") return false;
    }
    return true;
  };
  // Include open web-component menus while keeping document surfaces excluded.
  // Bound the walk; unsupported/closed components must not appear to succeed.
  const roots = [document];
  let visited = 0, rootLimit = false;
  for (let index = 0; index < roots.length; index++) {
    for (const element of roots[index].querySelectorAll("*")) {
      if (++visited > 20000) { rootLimit = true; break; }
      if (element.shadowRoot && !excluded(element, '[contenteditable="true"],[role="grid"],[role="document"],[role="textbox"],[role="combobox"]')) {
        if (roots.length >= 32) { rootLimit = true; continue; }
        roots.push(element.shadowRoot);
      }
    }
    if (visited > 20000) break;
  }
  if (rootLimit && command.type !== "inspect") return { actions: [], uploadInput: false, performed: false };
  const queryAll = (selector, scope) => roots.flatMap((root) => [...root.querySelectorAll(selector)]).filter((element) => !scope || within(element, scope));
  const enabled = (element) => !element.disabled && element.getAttribute("aria-disabled") !== "true";
  const locate = (target) => {
    const rect = target.getBoundingClientRect();
    const left = Math.max(0, rect.left), right = Math.min(innerWidth, rect.right);
    const top = Math.max(0, rect.top), bottom = Math.min(innerHeight, rect.bottom);
    const geometry = { width: rect.width, height: rect.height, clipped: left !== rect.left || right !== rect.right || top !== rect.top || bottom !== rect.bottom };
    if (right <= left || bottom <= top) return { performed: false, code: "CONTROL_NOT_POSITIONED", relation: "no_rect", geometry };
    const hitAt = (x, y) => {
      const point = { x: left + (right - left) * x, y: top + (bottom - top) * y };
      let hit = document.elementFromPoint(point.x, point.y);
      while (hit?.shadowRoot) {
        const next = hit.shadowRoot.elementFromPoint(point.x, point.y);
        if (!next || next === hit) break;
        hit = next;
      }
      const relation = !hit ? "none" : hit === target ? "target" : within(hit, target) ? "descendant" : within(target, hit) ? "ancestor" : "other";
      return { point, hit, relation };
    };
    const centre = hitAt(0.5, 0.5);
    let chosen = centre;
    if (!["target", "descendant"].includes(centre.relation)) {
      for (const [x, y] of [[0.25, 0.25], [0.75, 0.25], [0.25, 0.75], [0.75, 0.75]]) {
        const candidate = hitAt(x, y);
        if (["target", "descendant"].includes(candidate.relation)) { chosen = candidate; break; }
      }
    }
    const { point, hit, relation } = chosen;
    const detail = { relation, geometry, hitTag: hit?.tagName.toLowerCase() || null, hitRole: hit?.getAttribute("role") || null,
      centreRelation: centre.relation, pointerEvents: getComputedStyle(target).pointerEvents === "none" ? "none" : "auto" };
    if (!["target", "descendant"].includes(relation)) return { performed: false, code: "CONTROL_OBSCURED", ...detail };
    return { performed: true, point, viewport: { width: innerWidth, height: innerHeight }, ...detail };
  };
  const controls = queryAll('button,a,input[type="button"],input[type="submit"],[role="button"],[role="tab"],[role="menuitem"],[role="option"]').filter(visible);
  const names = {
    home: /^(home|start)$/,
    insert: /^(insert|invoegen)$/,
    addins: /^(add-ins|office add-ins|invoegtoepassingen|office-invoegtoepassingen)$/,
    getAddins: /^(get add-ins|invoegtoepassingen downloaden)$/,
    advanced: /^(advanced|geavanceerd)$/,
    moreSettings: /^(more settings|meer instellingen)$/,
    moreAddins: /^(more add-ins|meer invoegtoepassingen)$/,
    manage: /^(manage my add-ins?|manage add-ins|mijn invoegtoepassingen beheren|invoegtoepassingen beheren)$/,
    myAddins: /^(my add-ins|mijn invoegtoepassingen)$/,
    uploadEntry: /^(upload my add-in|upload add-in|mijn invoegtoepassing uploaden|invoegtoepassing uploaden)$/,
    bridge: /^teams office bridge$/,
  };
  const actions = {};
  for (const [name, pattern] of Object.entries(names)) {
    let matches = controls.filter((element) => pattern.test(label(element)) && enabled(element));
    if (["home", "insert"].includes(name)) {
      const tabs = matches.filter((element) => element.getAttribute("role") === "tab");
      if (tabs.length) matches = tabs;
    }
    if (matches.length === 1) actions[name] = matches[0];
  }
  // Word's compact Home ribbon can put Add-ins under More options. Limit
  // this intermediate step to the named ribbon toolbar, never a document
  // or comment menu, and prefer an already exposed Add-ins command.
  if (!actions.addins) {
    const ribbons = queryAll('[role="toolbar"]').filter((element) => /^(ribbon|lint)$/.test(label(element)) && visible(element));
    const overflow = controls.filter((element) => /^(more options|meer opties)$/.test(label(element)) && enabled(element) && ribbons.some((ribbon) => within(element, ribbon)));
    if (overflow.length === 1) actions.ribbonOverflow = overflow[0];
  }
  // Some Office versions put a plain Upload item under Manage My Add-ins.
  if (command.allowUploadMenu) {
    const matches = controls.filter((element) => /^(upload|uploaden)$/.test(label(element)) && enabled(element) && element.getAttribute("role") === "menuitem");
    if (matches.length === 1) actions.uploadEntry = matches[0];
  }
  const dialogs = queryAll('[role="dialog"],[role="alertdialog"]').filter(visible);
  const uploadDialogs = dialogs.filter((dialog) => /(?:upload|uploaden)/.test(label(dialog)) && /(?:add-in|invoegtoepassing)/.test(label(dialog)));
  // A file input may be visually hidden behind Browse. Never click Browse or
  // open a native picker, and never fill a generic document-upload input.
  const inputs = [...new Set(uploadDialogs.flatMap((dialog) => queryAll('input[type="file"]', dialog)))];
  const input = inputs.length === 1 ? inputs[0] : null;
  const dialog = input && uploadDialogs.find((item) => within(input, item));
  const uploads = dialog ? controls.filter((element) => within(element, dialog) && /^(upload|uploaden)$/.test(label(element)) && enabled(element)) : [];
  if (input && !input.disabled && uploads.length === 1) actions.submit = uploads[0];
  const trustDialogs = dialogs.filter((dialog) => /teams office bridge/.test(normalize(dialog.textContent)));
  const trust = controls.filter((element) => trustDialogs.some((dialog) => within(element, dialog)) && /^(trust it|trust this add-in|add|continue|vertrouwen|toevoegen|doorgaan)$/.test(label(element)) && enabled(element));
  if (trust.length === 1) actions.trust = trust[0];
  const expectedFile = input?.files?.length === 1 && input.files[0].name === command.filename;
  if (command.type === "inspect") {
    // Blocker metadata contains tags/roles only, never blocker text, IDs,
    // HTML, input values, document content or URLs.
    const candidates = queryAll('button,a,input[type="button"],input[type="submit"],[role="button"],[role="tab"],[role^="menuitem"],[role="option"],[aria-haspopup],[tabindex]')
      .filter((element) => visible(element) && !excluded(element, '[role="document"],[role="textbox"],[role="combobox"]') && !element.matches('input:not([type="button"]):not([type="submit"]),textarea,select'));
    return {
      actions: Object.keys(actions), truncated: rootLimit || candidates.length > 120,
      actionability: Object.entries(actions).slice(0, 20).map(([action, target]) => {
        const result = locate(target);
        return { action, relation: result.relation, centreRelation: result.centreRelation || result.relation, hitTag: result.hitTag || null, hitRole: result.hitRole || null,
          pointerEvents: result.pointerEvents || "auto", ...result.geometry };
      }),
      controls: candidates.slice(0, 120).map((element) => ({
        tag: element.tagName.toLowerCase(), role: element.getAttribute("role") || null,
        label: label(element).replace(/https?:\/\/\S+/g, "[URL]").slice(0, 160), enabled: enabled(element),
      })),
    };
  }
  if (command.type === "probe") return { actions: Object.keys(actions), uploadInput: !!input, expectedFile,
    homeRequired: !!actions.home && !!actions.addins && !excluded(actions.addins, '[role="menu"],[role="dialog"],[role="alertdialog"]') && locate(actions.addins).code === "CONTROL_OBSCURED" };
  if (command.type === "fill") {
    if (!input || input.disabled || !/^(excel|word)-addin\.xml$/.test(command.filename)) return { performed: false };
    const transfer = new DataTransfer();
    transfer.items.add(new File([command.manifest], command.filename, { type: "text/xml" }));
    input.files = transfer.files;
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
    return { performed: true };
  }
  const target = actions[command.action];
  if (!["click", "focus", "point"].includes(command.type) || !target || (command.action === "submit" && !expectedFile)) return { performed: false };
  if (command.action === "submit" && await input.files[0].text() !== command.manifest) return { performed: false };
  if (location.href !== command.url || Date.now() >= command.deadline) return { expired: true, performed: false };
  if (["focus", "point"].includes(command.type)) {
    // A collapsed ribbon leaves the Add-ins command behind the formula bar.
    // Do not activate it through that obstruction, even with the keyboard.
    // Missing layout in hidden frames is not itself evidence of an overlay.
    if (command.type === "focus" && command.action === "addins" && !excluded(target, '[role="menu"],[role="dialog"],[role="alertdialog"]')) {
      const position = locate(target);
      if (position.code === "CONTROL_OBSCURED") return position;
    }
    target.focus({ preventScroll: true });
    let active = document.activeElement;
    while (active?.shadowRoot?.activeElement) active = active.shadowRoot.activeElement;
    if (active !== target || !target.isConnected) return { performed: false };
    if (command.type === "focus") return { performed: true };
    return locate(target);
  }
  if (command.interaction === "pointer") {
    target.focus({ preventScroll: true });
    for (const [type, pointer, buttons] of [["pointerdown", true, 1], ["mousedown", false, 1], ["pointerup", true, 0], ["mouseup", false, 0]]) {
      const options = { bubbles: true, cancelable: true, composed: true, button: 0, buttons, pointerId: 1, pointerType: "mouse", isPrimary: true };
      target.dispatchEvent(pointer ? new PointerEvent(type, options) : new MouseEvent(type, options));
      if (!target.isConnected) return { performed: true };
    }
  }
  target.click();
  return { performed: true };
}

module.exports = { bridgeSetupUi };

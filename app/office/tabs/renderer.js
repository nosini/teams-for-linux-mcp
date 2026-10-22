const api = globalThis.documentTabs;
const status = document.getElementById("status");
async function request(value) {
  try {
    const state = await api.request(value);
    if (state.error) { status.textContent = state.error; return; }
    if (state.tabs) render(state);
  } catch { status.textContent = "Could not update document tabs. Reopen Documents from the Teams menu."; }
}
function render(state) {
  const tabs = document.getElementById("tabs");
  tabs.replaceChildren();
  for (const tab of state.tabs) {
    const item = document.createElement("div"); item.className = "tab";
    item.setAttribute("aria-current", String(tab.id === state.selectedTabId));
    const select = document.createElement("button"); select.className = "select";
    select.textContent = tab.title; select.title = tab.title;
    select.onclick = () => request({ action: "select", tabId: tab.id });
    const close = document.createElement("button"); close.textContent = "×"; close.setAttribute("aria-label", `Close ${tab.title}`);
    close.onclick = () => request({ action: "close", tabId: tab.id });
    item.append(select, close); tabs.append(item);
  }
  const sources = document.getElementById("sources"); const previous = sources.value;
  sources.replaceChildren();
  for (const source of state.sources) {
    const option = document.createElement("option"); option.value = source.editorId;
    option.textContent = `Teams file (${source.webContentsId})`; sources.append(option);
  }
  if ([...sources.options].some((option) => option.value === previous)) sources.value = previous;
  document.getElementById("capture").disabled = !state.sources.length;
  const tab = state.tabs.find((tab) => tab.id === state.selectedTabId);
  const setup = document.getElementById("setup");
  setup.disabled = !tab || !state.bridgeAvailable || tab.connected || tab.bridgeSetup?.status === "running";
  setup.onclick = () => request({ action: "setup", tabId: tab.id, retry: !!tab.bridgeSetup });
  status.textContent = !state.bridgeAvailable ? "Live Office access is unavailable in this run. Enable the Office Bridge and restart Teams to let agents edit." :
    !tab ? "Open a file here. Closing this window keeps your tabs running." :
    tab.connected ? "Bridge connected. Agents can edit this document while this window is hidden." :
    tab.bridgeSetup?.status === "running" ? (["waiting_for_sign_in", "selecting_account"].includes(tab.bridgeSetup.stage) ? "Signing in with your remembered document account…" : "Connecting the Bridge through Office’s add-in menu…") :
    tab.bridgeSetup?.code ? `Bridge setup stopped (${tab.bridgeSetup.code}). Check Office’s add-in menu or sign-in before trying again.` :
    tab.blockedNavigation ? "This document requested another page or sign-in window. Open it visibly to finish setup; external navigation is blocked." :
    tab.status === "loading" ? "Opening document…" : tab.status === "failed" || tab.status === "crashed" ? "The document could not load. Close this tab and reopen the file." :
    "Waiting for the Bridge. Open the Teams Office Bridge add-in in this document; finish sign-in or trust prompts here if needed.";
}
document.getElementById("open").onsubmit = (event) => { event.preventDefault(); void request({ action: "open", url: document.getElementById("url").value }); };
document.getElementById("capture").onclick = () => request({ action: "open", editorId: document.getElementById("sources").value });
document.getElementById("hide").onclick = () => request({ action: "hide" });
api.onState(render);
void request({ action: "list" });
setInterval(() => { if (!document.hidden) void request({ action: "list" }); }, 3000);

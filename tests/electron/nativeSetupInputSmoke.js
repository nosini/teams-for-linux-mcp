// Real Chromium native input and cross-origin focus; no real Office/network.
const { app, BrowserWindow, session, webContents } = require("electron");
const path = require("node:path");
const assert = require("node:assert/strict");
const { executeBridgeSetupStep, inspectBridgeSetup, runBridgeSetup } = require("../../app/office/bridgeSetup");

app.setPath("userData", path.resolve(__dirname, "../../test-results/native-setup-profile"));
app.disableHardwareAcceleration();
const timer = setTimeout(() => app.exit(1), 25000);

app.whenReady().then(async () => {
  let window, userWindow;
  try {
    session.defaultSession.protocol.handle("https", (request) => {
      const url = new URL(request.url);
      if (url.hostname === "tenant.sharepoint.com") return new Response('<iframe src="https://euc-excel.officeapps.live.com/x/edit"></iframe>', { headers: { "Content-Type": "text/html" } });
      if (url.hostname === "euc-excel.officeapps.live.com") return new Response(`
        <button id="addins">Add-ins</button><script>
        globalThis.audit = { events: [], keys: [], opened: 0 };
        for (const type of ['keydown', 'keypress', 'keyup']) document.addEventListener(type, event => audit.keys.push({type, key:event.key, target:event.target.id, trusted:event.isTrusted}));
        document.getElementById('addins').onclick = event => {
          audit.events.push({ trusted: event.isTrusted, activation: navigator.userActivation.isActive });
          if (event.isTrusted) audit.opened++;
        };
        </script>`, { headers: { "Content-Type": "text/html" } });
      return new Response("No external network", { status: 403 });
    });
    window = new BrowserWindow({ show: false, webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false, backgroundThrottling: false } });
    await window.loadURL("https://tenant.sharepoint.com/book.xlsx");
    userWindow = new BrowserWindow({ show: true });
    await userWindow.loadURL("data:text/html,<p>User activity stays focused</p>");
    userWindow.focus();
    await new Promise((resolve) => setTimeout(resolve, 100));
    assert.equal(BrowserWindow.getFocusedWindow(), userWindow);
    const focusedContents = webContents.getFocusedWebContents();
    const editor = window.webContents.mainFrame.framesInSubtree.find((frame) => frame.url.startsWith("https://euc-excel.officeapps.live.com/"));
    assert.ok(editor);
    const manifests = { excel: "fixture manifest" };
    for (const interaction of ["click", "pointer"]) {
      assert.equal((await executeBridgeSetupStep(window.webContents, manifests, "addins", interaction)).performed, true);
    }
    assert.equal((await editor.executeJavaScript("audit")).opened, 0, "DOM events must not be mistaken for native input");
    const focusedWindow = BrowserWindow.getFocusedWindow();
    const native = await executeBridgeSetupStep(window.webContents, manifests, "addins", "auto");
    assert.equal(native.performed, true, JSON.stringify(native));
    await new Promise((resolve) => setTimeout(resolve, 100));
    const audit = await editor.executeJavaScript("audit");
    assert.equal(audit.opened, 1, JSON.stringify(audit));
    assert.deepEqual(audit.events.map((event) => event.trusted), [false, false, true]);
    assert.equal(audit.events.at(-1).activation, true);
    assert.equal(window.isVisible(), false);
    assert.equal(BrowserWindow.getFocusedWindow(), focusedWindow);
    assert.equal(webContents.getFocusedWebContents(), focusedContents);
    assert.equal(window.webContents.debugger.isAttached(), false);
    window.webContents.debugger.attach("1.3");
    assert.equal((await executeBridgeSetupStep(window.webContents, manifests, "addins", "native")).code, "DEBUGGER_BUSY");
    assert.equal(window.webContents.debugger.isAttached(), true);
    window.webContents.debugger.detach();
    await editor.executeJavaScript(`
      globalThis.mouseAudit = { presses: 0, opened: 0 };
      const button = document.getElementById('addins');
      button.onpointerdown = event => { if (event.isTrusted) mouseAudit.presses++; };
      button.onclick = event => { if (event.isTrusted && mouseAudit.presses) mouseAudit.opened++; };
    `);
    assert.equal((await executeBridgeSetupStep(window.webContents, manifests, "addins", "native")).performed, true);
    assert.equal((await editor.executeJavaScript("mouseAudit")).opened, 0, "Enter does not satisfy a pointer-only menu");
    const mouse = await executeBridgeSetupStep(window.webContents, manifests, "addins", "nativeMouse");
    assert.equal(mouse.performed, true, JSON.stringify(mouse));
    assert.deepEqual(await editor.executeJavaScript("mouseAudit"), { presses: 1, opened: 1 });
    await editor.executeJavaScript("const blocker = document.createElement('div'); blocker.id = 'blocker'; blocker.setAttribute('role', 'dialog'); blocker.textContent = 'private blocker contents'; blocker.style = 'position:fixed;inset:0;z-index:1000'; document.body.append(blocker)");
    const blockedInspection = await inspectBridgeSetup(window.webContents);
    const blocked = blockedInspection.frames.flatMap((frame) => frame.actionability || []).find((item) => item.action === "addins");
    assert.equal(blocked.relation, "other");
    assert.equal(blocked.hitTag, "div");
    assert.equal(blocked.hitRole, "dialog");
    assert.doesNotMatch(JSON.stringify(blockedInspection), /private blocker|blocker contents/);
    assert.equal((await executeBridgeSetupStep(window.webContents, manifests, "addins", "auto")).code, "CONTROL_OBSCURED");
    assert.equal((await editor.executeJavaScript("mouseAudit")).presses, 1);
    await editor.executeJavaScript("document.getElementById('blocker').remove()");
    await editor.executeJavaScript(`
      const rect = document.getElementById('addins').getBoundingClientRect();
      const part = document.createElement('div'); part.id = 'partial';
      part.style.cssText = 'position:fixed;z-index:1000;left:' + (rect.left + rect.width * .375) + 'px;top:' + (rect.top + rect.height / 3) + 'px;width:' + (rect.width / 4) + 'px;height:' + (rect.height / 3) + 'px';
      document.body.append(part);
    `);
    const partialInspection = await inspectBridgeSetup(window.webContents);
    const partial = partialInspection.frames.flatMap((frame) => frame.actionability || []).find((item) => item.action === "addins");
    assert.equal(partial.centreRelation, "other");
    assert.ok(["target", "descendant"].includes(partial.relation));
    assert.equal((await executeBridgeSetupStep(window.webContents, manifests, "addins", "nativeMouse")).performed, true);
    assert.equal((await editor.executeJavaScript("mouseAudit")).opened, 2, "Click only an uncovered point that hits the same named button");
    await editor.executeJavaScript("document.getElementById('partial').remove()");
    for (const zoom of [0.75, 1.25, 2]) {
      await editor.executeJavaScript(`document.body.style.zoom = ${zoom}`);
      const zoomInspection = await inspectBridgeSetup(window.webContents);
      const state = zoomInspection.frames.flatMap((frame) => frame.actionability || []).find((item) => item.action === "addins");
      assert.ok(["target", "descendant"].includes(state.relation), JSON.stringify(state));
    }
    await editor.executeJavaScript("document.body.style.zoom = 1");
    await window.webContents.executeJavaScript("const frame = document.querySelector('iframe'); frame.style.margin = '25px 40px'; frame.style.transformOrigin = '0 0'; frame.style.transform = 'scale(.75)'");
    assert.equal((await executeBridgeSetupStep(window.webContents, manifests, "addins", "nativeMouse")).performed, true, "Mouse coordinates must account for frame offsets, borders and scale");
    assert.equal((await editor.executeJavaScript("mouseAudit")).opened, 3);
    await window.webContents.executeJavaScript("const cover = document.createElement('div'); cover.id = 'cover'; cover.style = 'position:fixed;inset:0;z-index:1000'; document.body.append(cover)");
    const obscured = await executeBridgeSetupStep(window.webContents, manifests, "addins", "nativeMouse");
    assert.equal(obscured.performed, false, "Never click an overlay instead of the validated child frame");
    assert.equal((await editor.executeJavaScript("mouseAudit")).presses, 3);
    await window.webContents.executeJavaScript("document.getElementById('cover').remove(); document.querySelector('iframe').style.transform = 'rotate(5deg)'");
    assert.equal((await executeBridgeSetupStep(window.webContents, manifests, "addins", "nativeMouse")).performed, false, "Unsupported geometry must stop before mouse press");
    await window.webContents.executeJavaScript("document.querySelector('iframe').style.transform = 'none'");
    await editor.executeJavaScript(`
      document.body.replaceChildren();
      const host = document.createElement('div'); host.id = 'menu'; document.body.append(host);
      const root = host.attachShadow({mode:'open'});
      const advanced = document.createElement('button'); advanced.textContent = 'Advanced';
      globalThis.advancedClicks = 0;
      advanced.onclick = event => { if (event.isTrusted) advancedClicks++; };
      root.append(advanced);
      const privateGrid = document.createElement('section'); privateGrid.setAttribute('role','grid'); root.append(privateGrid);
      const privateHost = document.createElement('div'); privateGrid.append(privateHost);
      privateHost.attachShadow({mode:'open'}).innerHTML = '<button>private shadow cell value</button><button>Advanced</button>';
    `);
    const inspection = await inspectBridgeSetup(window.webContents);
    assert.ok(inspection.frames.some((frame) => frame.actions?.includes("advanced")));
    assert.doesNotMatch(JSON.stringify(inspection), /private shadow cell/);
    assert.equal((await executeBridgeSetupStep(window.webContents, manifests, "advanced", "nativeMouse")).performed, true);
    assert.equal(await editor.executeJavaScript("advancedClicks"), 1, "Native input must reach the validated control inside a shadow root");
    await editor.executeJavaScript(`document.getElementById('menu').style.visibility = 'hidden'; document.getElementById('menu').shadowRoot.querySelector('button').style.visibility = 'visible'`);
    assert.equal((await executeBridgeSetupStep(window.webContents, manifests, "advanced", "native")).performed, true, "A visible menu can override its hidden portal's inherited visibility");
    assert.equal(await editor.executeJavaScript("advancedClicks"), 2);
    await editor.executeJavaScript(`document.getElementById('menu').shadowRoot.querySelector('button').style.visibility = ''`);
    assert.equal((await executeBridgeSetupStep(window.webContents, manifests, "advanced", "native")).code, "CONTROL_NOT_FOUND", "Inherited hidden visibility still prevents activation");
    await editor.executeJavaScript(`document.getElementById('menu').shadowRoot.querySelector('button').style.visibility = 'visible'; document.getElementById('menu').style.display = 'none'`);
    assert.equal((await executeBridgeSetupStep(window.webContents, manifests, "advanced", "native")).code, "CONTROL_NOT_FOUND", "Visible descendants cannot override display:none");
    await editor.executeJavaScript(`document.getElementById('menu').style.display = ''; document.getElementById('menu').style.visibility = ''`);
    await editor.executeJavaScript("document.getElementById('menu').hidden = true");
    assert.equal((await executeBridgeSetupStep(window.webContents, manifests, "advanced", "native")).code, "CONTROL_NOT_FOUND");
    await editor.executeJavaScript(`
      document.body.innerHTML = '<button id="home" role="tab">Home</button><button id="addins">Add-ins</button><div id="cover"></div>';
      const ribbonRect = document.getElementById('addins').getBoundingClientRect();
      document.getElementById('cover').style.cssText = 'position:absolute;z-index:1000;left:' + ribbonRect.left + 'px;top:' + ribbonRect.top + 'px;width:' + ribbonRect.width + 'px;height:' + ribbonRect.height + 'px';
      globalThis.prepared = [];
      document.getElementById('home').onclick = () => { prepared.push('home'); setTimeout(() => document.getElementById('cover').remove(), 75); };
      document.getElementById('addins').onclick = event => { if (event.isTrusted) prepared.push('addins'); };
    `);
    assert.equal((await executeBridgeSetupStep(window.webContents, manifests, "addins", "auto")).code, "CONTROL_OBSCURED", "Keyboard default must not activate an obscured ribbon command");
    const prepared = await runBridgeSetup({ contents: window.webContents, manifests, signal: new AbortController().signal,
      update() {}, listDocuments: async () => ({ documents: await editor.executeJavaScript("prepared.includes('addins')") ? [{ webContentsId: window.webContents.id }] : [] }) });
    assert.equal(prepared.status, "connected");
    assert.deepEqual(await editor.executeJavaScript("prepared"), ["home", "addins"], "Automatic setup expands a collapsed ribbon once and waits until Add-ins is reachable");
    await editor.executeJavaScript(`
      document.body.innerHTML = '<button>More options</button><div role="toolbar" aria-label="Ribbon"><button id="overflow">More options</button></div>';
      globalThis.overflowClicks = 0;
      document.getElementById('overflow').onclick = () => {
        overflowClicks++;
        const menu = document.createElement('div'); menu.setAttribute('role', 'menu'); menu.style.pointerEvents = 'none';
        const addins = document.createElement('button'); addins.textContent = 'Add-ins'; addins.setAttribute('role', 'menuitem');
        addins.onclick = event => { if (event.isTrusted) overflowClicks++; }; menu.append(addins); document.body.append(menu);
      };
    `);
    assert.equal((await executeBridgeSetupStep(window.webContents, manifests, "ribbonOverflow", "auto")).performed, true, "Ignore matching More options outside the named ribbon");
    assert.equal((await executeBridgeSetupStep(window.webContents, manifests, "addins", "auto")).performed, true, "Keyboard activation reaches an Add-ins menu item whose portal has pointer events disabled");
    assert.equal(await editor.executeJavaScript("overflowClicks"), 2);
    await editor.executeJavaScript("document.querySelector('[role=toolbar]').remove(); document.querySelector('[role=menu]').remove()");
    assert.equal((await executeBridgeSetupStep(window.webContents, manifests, "ribbonOverflow", "auto")).code, "CONTROL_NOT_FOUND", "Never use an unscoped More options menu");
    assert.equal(window.isVisible(), false);
    assert.equal(BrowserWindow.getFocusedWindow(), focusedWindow);
    assert.equal(webContents.getFocusedWebContents(), focusedContents);
    console.log("Native setup input smoke passed: DOM click/pointer rejected, trusted native activation accepted in hidden cross-origin frame, unchanged native focus, debugger released and busy debugger preserved. No real Office UI used.");
  } finally { window?.destroy(); userWindow?.destroy(); clearTimeout(timer); }
}).then(() => app.exit(0), (error) => { console.error(error); app.exit(1); });

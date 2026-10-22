// A representative import UI, not a copy of Microsoft's live implementation.
function officeImportFixture({ origin, application, file, store = false, panel = false, native = false }) {
  const dutch = application === "word";
  return `<!doctype html><meta charset="utf-8"><title>Office import fixture</title>
    <main id="ui"></main>
    <section role="grid"><button>private cell value</button></section>
    <section contenteditable="true"><button>private document text</button></section>
    <section role="document"><button>private document body</button></section>
    <input value="private input value">
    <aside><button>More options</button></aside>
    <div role="textbox" tabindex="0" aria-label="private cell accessibility text"></div>
    <script>
    const application = ${JSON.stringify(application)};
    const file = ${JSON.stringify(file)};
    const origin = ${JSON.stringify(origin)};
    const dutch = ${dutch};
    const requireTrusted = ${native};
    const ui = document.getElementById('ui');
    // Office owns the task pane; it is not embedded in the store/CDN dialog.
    window.addEventListener('message', event => {
      const expectedOrigin = ${panel} ? 'https://store.office.com' : 'https://res.public.onecdn.static.microsoft';
      if (event.origin !== expectedOrigin || event.data?.fixtureBridge !== file) return;
      if (${panel}) parent.postMessage(event.data, 'https://euc-' + application + '.officeapps.live.com');
      else if (!${store}) {
        const pane = document.createElement('iframe');
        pane.src = origin + '/taskpane.html?application=' + application + '&file=' + encodeURIComponent(file);
        ui.append(pane);
      }
    });
    globalThis.importAudit = { advanced: 0, overflow: 0, uploads: 0, browse: 0, file: null, type: null, blocked: 0, gesture: false, trusted: false };
    importAudit.mouse = [];
    for (const type of ['pointerdown', 'click']) document.addEventListener(type, event => importAudit.mouse.push({ type, trusted: event.isTrusted, target: event.target.tagName, label: event.target.textContent.slice(0, 40), x: event.clientX, y: event.clientY }), true);
    function button(text, fn, container = ui) {
      const b = document.createElement('button'); b.textContent = text;
      let pointer = false;
      b.onpointerdown = event => { pointer = !requireTrusted || event.isTrusted; };
      b.onclick = (event) => {
        if (requireTrusted && !event.isTrusted) { importAudit.blocked++; return; }
        if (/^(Add-ins|Invoegtoepassingen)$/.test(text)) {
          importAudit.gesture = navigator.userActivation.isActive;
          importAudit.trusted = event.isTrusted;
          if ((!pointer && !event.isTrusted) || !importAudit.gesture) { importAudit.blocked++; return; }
        }
        pointer = false; fn(event);
      };
      container.append(b); return b;
    }
    function uploadDialog() {
      ui.replaceChildren();
      const dialog = document.createElement('section'); dialog.setAttribute('role', 'dialog');
      dialog.setAttribute('aria-label', dutch ? 'Invoegtoepassing uploaden' : 'Upload Add-in');
      const input = document.createElement('input'); input.type = 'file'; input.accept = '.xml'; input.style.display = 'none';
      const browse = document.createElement(dutch ? 'button' : 'input');
      if (dutch) browse.textContent = 'Bladeren'; else { browse.type = 'button'; browse.value = 'Browse...'; }
      browse.onclick = () => { importAudit.browse++; input.click(); };
      const upload = document.createElement(dutch ? 'button' : 'input');
      if (dutch) upload.textContent = 'Uploaden'; else { upload.type = 'button'; upload.value = 'Upload'; }
      upload.disabled = true;
      // A matching filename alone must not make the installer submit a file
      // somebody selected earlier. It must replace this with its own manifest.
      const stale = new DataTransfer(); stale.items.add(new File(['wrong manifest'], application + '-addin.xml'));
      input.files = stale.files; upload.disabled = false;
      input.onchange = () => { upload.disabled = input.files.length !== 1; };
      upload.onclick = async () => {
        importAudit.uploads++;
        const manifest = input.files[0]; importAudit.file = manifest.name;
        const xml = new DOMParser().parseFromString(await manifest.text(), 'application/xml');
        const host = xml.getElementsByTagName('Host')[0]?.getAttribute('Name');
        importAudit.type = host;
        if (host !== (application === 'word' ? 'Document' : 'Workbook') ||
          xml.getElementsByTagName('SourceLocation')[0]?.getAttribute('DefaultValue') !== origin + '/taskpane.html' ||
          manifest.name !== application + '-addin.xml') throw new Error('Wrong manifest uploaded');
        ui.replaceChildren();
        const consent = document.createElement('section'); consent.setAttribute('role', 'dialog');
        consent.innerHTML = '<h2>Teams Office Bridge</h2><p>Read and write this document</p>';
        const trust = document.createElement('button'); trust.textContent = dutch ? 'Toevoegen' : 'Trust it';
        trust.onclick = () => {
          ui.replaceChildren();
          parent.postMessage({ fixtureBridge: file }, ${panel} ? 'https://euc-' + application + '.officeapps.live.com' : 'https://res.public.onecdn.static.microsoft');
        };
        consent.append(trust); ui.append(consent);
      };
      const host = document.createElement('div'); dialog.append(host);
      host.attachShadow({ mode: 'open' }).append(input, browse, upload); ui.append(dialog);
    }
    function uploadEntry() {
      ui.replaceChildren();
      button(dutch ? 'Mijn invoegtoepassingen beheren' : 'Upload My Add-in', dutch ? () => {
        ui.replaceChildren(); const entry = button('Uploaden', uploadDialog); entry.setAttribute('role', 'menuitem');
      } : uploadDialog);
    }
    function panelMenu() {
      button(dutch ? 'Meer instellingen' : 'More Settings', () => {
        ui.replaceChildren();
        const frame = document.createElement('iframe');
        frame.src = 'https://store.office.com/import?application=' + application + '&file=' + encodeURIComponent(file);
        ui.append(frame);
      });
      button(dutch ? 'Meer invoegtoepassingen' : 'More Add-ins', uploadEntry);
    }
    function addins(container = ui) {
      return button(dutch ? 'Invoegtoepassingen' : 'Add-ins', () => {
        ui.replaceChildren();
        // Exercise a menu that takes longer than the job's old 800ms delay.
        setTimeout(() => {
        const portal = document.createElement('div'); portal.style.visibility = 'hidden'; ui.append(portal);
        const host = document.createElement('div'); host.style.visibility = 'visible'; portal.append(host);
        const menu = host.attachShadow({ mode: 'open' });
        // Exercise inspection and native focus through a web-component menu.
        button(dutch ? 'Geavanceerd' : 'Advanced', () => {
          importAudit.advanced++;
          ui.replaceChildren();
          const frame = document.createElement('iframe');
          frame.src = 'https://res.public.onecdn.static.microsoft/addins?application=' + application + '&file=' + encodeURIComponent(file);
          ui.append(frame);
        }, menu);
        const privateGrid = document.createElement('section'); privateGrid.setAttribute('role', 'grid');
        menu.append(privateGrid);
        button('private shadow cell value', () => {}, privateGrid);
        }, 1200);
      }, container);
    }
    if (${store}) uploadEntry();
    else if (${panel}) { if (dutch) panelMenu(); else uploadEntry(); }
    else button(dutch ? 'Start' : 'Home', () => {
      ui.replaceChildren();
      if (!dutch) { addins(); return; }
      const ribbon = document.createElement('div'); ribbon.setAttribute('role', 'toolbar'); ribbon.setAttribute('aria-label', 'Lint'); ui.append(ribbon);
      button('Meer opties', () => {
        importAudit.overflow++;
        const overflow = document.createElement('div'); overflow.setAttribute('role', 'menu'); overflow.style.pointerEvents = 'none'; ui.append(overflow);
        addins(overflow).setAttribute('role', 'menuitem');
      }, ribbon);
    });
    </script>`;
}

module.exports = { officeImportFixture };

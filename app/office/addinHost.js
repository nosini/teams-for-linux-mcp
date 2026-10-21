const fs = require("node:fs/promises");
const path = require("node:path");
const https = require("node:https");
const { X509Certificate, createPrivateKey } = require("node:crypto");
const { execFile } = require("node:child_process");
const { promisify } = require("node:util");
const runFile = promisify(execFile);

async function credentials(directory) {
  await fs.mkdir(directory, { recursive: true, mode: 0o700 });
  const stat = await fs.lstat(directory);
  if (!stat.isDirectory() || stat.uid !== process.getuid() || (stat.mode & 0o077)) throw new Error("Add-in certificate directory must be private and owned by this user");
  const certPath = path.join(directory, "localhost.pem");
  const keyPath = path.join(directory, "localhost-key.pem");
  const read = async () => {
    const cert = await fs.readFile(certPath);
    const key = await fs.readFile(keyPath);
    const keyStat = await fs.lstat(keyPath);
    const certificate = new X509Certificate(cert);
    if (!keyStat.isFile() || keyStat.uid !== process.getuid() || (keyStat.mode & 0o077) ||
      !certificate.checkIP("127.0.0.1") || !certificate.checkPrivateKey(createPrivateKey(key)) ||
      Date.parse(certificate.validFrom) > Date.now() || Date.parse(certificate.validTo) < Date.now() + 86400000) {
      throw new Error("Invalid local add-in certificate or private key");
    }
    return { cert, key, certificate };
  };
  try { return await read(); } catch (error) { if (error.code !== "ENOENT") throw error; }
  const temporary = await fs.mkdtemp(path.join(directory, "generate-"));
  try {
    await runFile("openssl", ["req", "-x509", "-newkey", "rsa:2048", "-nodes", "-days", "365", "-subj", "/CN=localhost",
      "-addext", "subjectAltName=IP:127.0.0.1,DNS:localhost", "-keyout", path.join(temporary, "key.pem"), "-out", path.join(temporary, "cert.pem")]);
    await fs.chmod(path.join(temporary, "key.pem"), 0o600);
    await fs.rename(path.join(temporary, "key.pem"), keyPath);
    await fs.rename(path.join(temporary, "cert.pem"), certPath);
    return await read();
  } finally {
    await fs.rm(temporary, { recursive: true, force: true });
  }
}

function createManifest(origin) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>2b7fbc27-ea2d-4db4-94df-2b770f6cda41</Id>
  <Version>0.1.0.0</Version>
  <ProviderName>Teams for Linux</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Teams Office Bridge"/>
  <Description DefaultValue="Connect this workbook to local agents through the Teams client."/>
  <Hosts><Host Name="Workbook"/></Hosts>
  <Requirements><Sets DefaultMinVersion="1.1"><Set Name="ExcelApi"/></Sets></Requirements>
  <DefaultSettings><SourceLocation DefaultValue="${origin}/taskpane.html"/></DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
`;
}

async function startAddinHost({ directory, port = 39847 }) {
  if (!Number.isInteger(port) || (port !== 0 && (port < 1024 || port > 65535))) throw new Error("Invalid add-in port");
  const { cert, key, certificate } = await credentials(path.join(directory, "addin"));
  const assets = new Map();
  for (const [name, type] of [["taskpane.html", "text/html"], ["taskpane.js", "text/javascript"], ["operations.js", "text/javascript"], ["taskpane.css", "text/css"]]) {
    assets.set(`/${name}`, { content: await fs.readFile(path.join(__dirname, "addin", name)), type });
  }
  let origin;
  const sockets = new Set();
  const server = https.createServer({ cert, key, minVersion: "TLSv1.2" }, (request, response) => {
    if (request.headers.host !== new URL(origin).host || !["GET", "HEAD"].includes(request.method)) {
      response.writeHead(403).end();
      return;
    }
    let asset;
    try { asset = assets.get(new URL(request.url, origin).pathname); } catch { /* Invalid request URL. */ }
    if (!asset) { response.writeHead(404).end(); return; }
    response.writeHead(200, {
      "Content-Type": `${asset.type}; charset=utf-8`, "Cache-Control": "no-store", "X-Content-Type-Options": "nosniff",
      "Referrer-Policy": "no-referrer",
      // Office.js can load Microsoft's legacy Ajax runtime, which uses eval
      // for type resolution. These script origins are Microsoft's Office CDNs.
      "Content-Security-Policy": "default-src 'none'; script-src 'self' 'unsafe-eval' https://appsforoffice.microsoft.com https://officeapis.public.onecdn.static.microsoft https://ajax.aspnetcdn.com; style-src 'self'; frame-ancestors https://*.officeapps.live.com https://*.office.com https://*.sharepoint.com https://*.cloud.microsoft https://teams.microsoft.com",
    });
    response.end(request.method === "HEAD" ? undefined : asset.content);
  });
  server.on("connection", (socket) => { sockets.add(socket); socket.once("close", () => sockets.delete(socket)); });
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, "127.0.0.1", resolve);
  });
  origin = `https://127.0.0.1:${server.address().port}`;
  const manifestPath = path.join(directory, "excel-addin.xml");
  let closing;
  const close = () => {
    if (!closing) {
      for (const socket of sockets) socket.destroy();
      closing = new Promise((resolve) => server.close(resolve));
    }
    return closing;
  };
  try {
    await fs.writeFile(manifestPath, createManifest(origin), { mode: 0o600 });
  } catch (error) { await close(); throw error; }
  server.on("error", () => {});
  return {
    origin, manifestPath, close,
    acceptsCertificate(url, error, peer) {
      try {
        const candidate = new X509Certificate(peer.data);
        return new URL(url).origin === origin && error === "net::ERR_CERT_AUTHORITY_INVALID" &&
          candidate.fingerprint256 === certificate.fingerprint256 && candidate.checkIP("127.0.0.1") === "127.0.0.1" &&
          Date.parse(candidate.validFrom) <= Date.now() && Date.parse(candidate.validTo) > Date.now();
      } catch { return false; }
    },
  };
}

module.exports = { startAddinHost, createManifest };

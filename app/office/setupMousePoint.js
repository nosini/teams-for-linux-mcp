const { evaluateInFrame } = require("../mcp/evaluateInFrame");

// Fixed parent-frame hit test, never a caller selector or coordinate. A child
// frame's validated control point is projected through its actual ancestors.
function projectSetupPoint(request) {
  if (location.href !== request.url || Date.now() >= request.deadline) return null;
  const roots = [document];
  let scanned = 0;
  for (let index = 0; index < roots.length; index++) {
    for (const element of roots[index].querySelectorAll("*")) {
      if (++scanned > 20000) return null;
      if (element.shadowRoot) {
        if (roots.length >= 32) return null;
        roots.push(element.shadowRoot);
      }
    }
  }
  const frames = roots.flatMap((root) => [...root.querySelectorAll("iframe,frame")]);
  const named = request.name ? frames.filter((element) => element.name === request.name) : [];
  const matches = named.length ? named : frames.filter((element) => element.src === request.childUrl);
  if (matches.length !== 1) return null;
  const frame = matches[0];
  for (let current = frame; current; current = current.parentElement || current.getRootNode()?.host) {
    const style = getComputedStyle(current);
    if (current.hidden || current.getAttribute("aria-hidden") === "true" || style.display === "none" || ["hidden", "collapse"].includes(style.visibility)) return null;
    if (style.transform !== "none") {
      const matrix = new DOMMatrix(style.transform);
      if (!matrix.is2D || matrix.b || matrix.c || matrix.a <= 0 || matrix.d <= 0) return null;
    }
  }
  const rect = frame.getBoundingClientRect();
  if (!frame.offsetWidth || !frame.offsetHeight || !frame.clientWidth || !frame.clientHeight || rect.width <= 0 || rect.height <= 0) return null;
  const point = {
    x: rect.left + (frame.clientLeft + request.point.x) * rect.width / frame.offsetWidth,
    y: rect.top + (frame.clientTop + request.point.y) * rect.height / frame.offsetHeight,
  };
  if (point.x < 0 || point.y < 0 || point.x >= innerWidth || point.y >= innerHeight) return null;
  let hit = document.elementFromPoint(point.x, point.y);
  while (hit?.shadowRoot) {
    const next = hit.shadowRoot.elementFromPoint(point.x, point.y);
    if (!next || next === hit) break;
    hit = next;
  }
  if (hit !== frame) return null;
  return { point, viewport: { width: innerWidth, height: innerHeight } };
}

function validPoint(value) {
  return value && [value.point?.x, value.point?.y, value.viewport?.width, value.viewport?.height].every(Number.isFinite) &&
    value.point.x >= 0 && value.point.y >= 0 && value.viewport.width > value.point.x && value.viewport.height > value.point.y;
}

async function setupMousePoint(contents, frame, local, isCurrent) {
  if (!validPoint(local)) return null;
  const chain = [];
  for (let current = frame; current && chain.length < 8; current = current.parent) chain.push({ frame: current, url: current.url, name: current.name });
  if (chain.at(-1)?.frame !== contents.mainFrame) return null;
  const current = () => isCurrent() && chain.every((item) => !item.frame.detached && item.frame.url === item.url && item.frame.name === item.name);
  let position = local;
  for (let index = 1; index < chain.length; index++) {
    if (!current()) return null;
    const child = chain[index - 1], parent = chain[index];
    const request = { url: parent.url, childUrl: child.url, name: child.name, ...position, deadline: Date.now() + 1500 };
    position = await evaluateInFrame(parent.frame, `(${projectSetupPoint.toString()})(${JSON.stringify(request)})`, 2000, "Setup frame positioning timed out");
    if (!validPoint(position)) return null;
  }
  return current() ? position.point : null;
}

module.exports = { setupMousePoint };

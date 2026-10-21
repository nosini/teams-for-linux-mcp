const vm = require("node:vm");
const { randomUUID } = require("node:crypto");

function codeMirrorFixture(text = "# Markdown\n\nHello, coauthors.\n", link = "cmTile") {
  const model = { dispatches: [], updates: [], filtered: false, deferred: false, error: null, url: "https://tenant.sharepoint.com/editor?file=test", now: Date.now() };
  const doc = (value) => Object.freeze({ length: value.length, sliceString: (from, to) => value.slice(from, to) });
  const content = { getAttribute: (name) => name === "contenteditable" ? (model.editable === false ? "false" : "true") : null };
  const root = { isConnected: true, querySelector: (selector) => selector === ".cm-content" ? content : null, contains: (node) => node === content };
  const view = {
    dom: root, contentDOM: content, state: { doc: doc(text), config: {}, readOnly: false }, composing: false,
    dispatch(spec) {
      model.dispatches.push(spec);
      if (model.error) throw model.error;
      if (model.filtered || model.deferred) return;
      const old = this.state.doc.sliceString(0, this.state.doc.length);
      const changes = spec.changes;
      this.state = { ...this.state, doc: doc(old.slice(0, changes.from) + changes.insert + old.slice(changes.to)) };
      model.updates.push(this.state.doc.sliceString(0, this.state.doc.length));
    },
  };
  content[link] = { parent: { view } };
  model.roots = [root];
  const scope = vm.createContext({
    document: { querySelectorAll: () => model.roots },
    location: { get href() { return model.url; } },
    crypto: { randomUUID }, TextEncoder,
    Date: { now: () => model.now },
  });
  return { model, root, content, view, scope, doc, evaluate: (source) => vm.runInContext(source, scope) };
}
module.exports = { codeMirrorFixture };

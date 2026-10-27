// Word proxy fixture: loads, searches and writes execute at context.sync().
function wordFixture(texts = ["Hello from Word.", "A second paragraph."]) {
  const model = {
    paragraphs: texts.map((text, index) => ({
      id: `00000000-0000-4000-8000-${String(index + 1).padStart(12, "0")}`, text,
      style: "Normal", styleBuiltIn: "Normal", alignment: "Left",
      font: { bold: false, italic: false, name: "Arial", size: 11, color: "#000000" },
      comments: [], changes: [],
      formatting: { outsideMatch: "preserved" },
    })),
    tables: [], trackingMode: "Off", syncs: 0, writes: 0, beforeSync: null, lastPage: null, searchOverride: null,
  };
  let nextParagraphId = model.paragraphs.length + 1;
  const word = {
    async run(callback) {
      let queued = [];
      const enqueue = (work) => queued.push(work);
      const find = (id) => {
        const item = model.paragraphs.find((entry) => entry.id === id);
        if (!item) { const error = new Error("private document information"); error.code = "ItemNotFound"; throw error; }
        return item;
      };
      const collection = (getItems, selectProxy = (item) => item) => ({
        items: [], load(options = {}) { enqueue(() => { this.items = getItems().slice(options.skip ?? 0, (options.skip ?? 0) + (options.top ?? 100000)).map(selectProxy); }); },
      });
      const commentProxy = (item) => ({
        ...item,
        load() { enqueue(() => { for (const key of ["id", "content", "authorName", "creationDate"]) this[key] = item[key]; }); },
        replies: collection(() => item.replies || []),
        reply(content) { enqueue(() => { (item.replies ||= []).push({ id: `reply-${item.replies.length}`, content, authorName: "Agent", creationDate: new Date(0) }); model.writes++; }); },
        get resolved() { return item.resolved; },
        set resolved(value) { enqueue(() => { item.resolved = value; model.writes++; }); },
      });
      const proxy = (id, seed = {}) => {
        const state = { ...seed, font: { ...seed.font } };
        const paragraph = {
          font: Object.fromEntries([]),
          getRange(location) {
            if (location !== "Content") throw new Error("Expected content range");
            return {
              insertText(text, position) { if (position !== "Replace") throw new Error("Expected Replace"); enqueue(() => { find(id).text = text; model.writes++; }); },
              insertComment(content) {
                const comment = { id: `comment-${model.writes}`, content, authorName: "Agent", creationDate: new Date(0), resolved: false, replies: [] };
                enqueue(() => { find(id).comments.push(comment); model.writes++; });
                return commentProxy(comment);
              },
            };
          },
          insertParagraph(text, position) {
            const newId = `10000000-0000-4000-8000-${String(nextParagraphId++).padStart(12, "0")}`;
            enqueue(() => {
              const item = find(id), index = model.paragraphs.indexOf(item);
              model.paragraphs.splice(index + (position === "After" ? 1 : 0), 0, { ...item, id: newId, text, font: { ...item.font }, comments: [], changes: [] });
              model.writes++;
            });
            return proxy(newId);
          },
          delete() { enqueue(() => { const item = find(id); model.paragraphs.splice(model.paragraphs.indexOf(item), 1); model.writes++; }); },
          getComments() { return collection(() => find(id).comments || [], commentProxy); },
          getTrackedChanges() { return collection(() => find(id).changes || [], (change) => ({
            ...change,
            accept() { enqueue(() => { const item = find(id); item.changes.splice(item.changes.indexOf(change), 1); model.writes++; }); },
            reject() { enqueue(() => { const item = find(id); item.text = item.originalText ?? item.text; item.changes.splice(item.changes.indexOf(change), 1); model.writes++; }); },
          })); },
          getReviewedText(version) { const result = {}; enqueue(() => { const item = find(id); result.value = version === "Original" ? (item.originalText ?? item.text) : item.text; }); return result; },
          load() {
            enqueue(() => { const item = find(id); paragraph.uniqueLocalId = item.id; paragraph.text = item.text;
              for (const key of ["style", "styleBuiltIn", "alignment"]) state[key] = item[key];
              state.font = { ...item.font };
            });
          },
          search(text, options) {
            const matches = {
              items: [],
              load(query) {
                enqueue(() => {
                  const item = find(id);
                  const ranges = [];
                  for (let start = item.text.indexOf(text); start !== -1; start = item.text.indexOf(text, start + text.length)) {
                    ranges.push({
                      text,
                      insertText(replacement, location) {
                        if (location !== "Replace") throw new Error("Expected range replacement");
                        enqueue(() => {
                          const current = find(id);
                          current.text = current.text.slice(0, start) + replacement + current.text.slice(start + text.length);
                          model.writes++;
                        });
                      },
                    });
                    if (ranges.length >= (query?.top ?? Infinity)) break;
                  }
                  matches.items = model.searchOverride || ranges;
                });
              },
            };
            model.searchOptions = options;
            return matches;
          },
        };
        for (const key of ["style", "styleBuiltIn", "alignment"]) Object.defineProperty(paragraph, key, { enumerable: true, get: () => state[key], set: (value) => { enqueue(() => { find(id)[key] = value; state[key] = value; model.writes++; }); } });
        for (const key of ["bold", "italic", "name", "size", "color"]) Object.defineProperty(paragraph.font, key, { enumerable: true, get: () => state.font?.[key], set: (value) => { enqueue(() => { find(id).font[key] = value; model.writes++; }); } });
        return paragraph;
      };
      const paragraphs = {
        items: [],
        load(options) {
          model.lastPage = options;
          enqueue(() => {
            paragraphs.items = model.paragraphs.slice(options.skip, options.skip + options.top).map((item) => {
              const paragraph = proxy(item.id, item);
              paragraph.uniqueLocalId = item.id;
              paragraph.text = item.text;
              return paragraph;
            });
          });
        },
      };
      const context = {
        document: {
          body: { paragraphs, tables: collection(() => model.tables, (table) => ({ rowCount: table.rows.length, rows: collection(() => table.rows.map((values) => ({ values: [values], cellCount: values.length }))) })) },
          getParagraphByUniqueLocalId: proxy,
          load() {},
          get changeTrackingMode() { return model.trackingMode; },
          set changeTrackingMode(value) { enqueue(() => { model.trackingMode = value; model.writes++; }); },
        },
        async sync() {
          model.syncs++;
          await model.beforeSync?.(model.syncs);
          const batch = queued;
          queued = [];
          for (const work of batch) work();
        },
      };
      return callback(context);
    },
  };
  return { word, model };
}

module.exports = { wordFixture };

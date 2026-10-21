// Word proxy fixture: loads, searches and writes execute at context.sync().
function wordFixture(texts = ["Hello from Word.", "A second paragraph."]) {
  const model = {
    paragraphs: texts.map((text, index) => ({
      id: `00000000-0000-4000-8000-${String(index + 1).padStart(12, "0")}`, text,
      formatting: { outsideMatch: "preserved" },
    })),
    syncs: 0, writes: 0, beforeSync: null, lastPage: null, searchOverride: null,
  };
  const word = {
    async run(callback) {
      let queued = [];
      const enqueue = (work) => queued.push(work);
      const find = (id) => {
        const item = model.paragraphs.find((entry) => entry.id === id);
        if (!item) { const error = new Error("private document information"); error.code = "ItemNotFound"; throw error; }
        return item;
      };
      const proxy = (id) => {
        const paragraph = {
          load() {
            enqueue(() => { const item = find(id); paragraph.uniqueLocalId = item.id; paragraph.text = item.text; });
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
        return paragraph;
      };
      const paragraphs = {
        items: [],
        load(options) {
          model.lastPage = options;
          enqueue(() => {
            paragraphs.items = model.paragraphs.slice(options.skip, options.skip + options.top).map((item) => {
              const paragraph = proxy(item.id);
              paragraph.uniqueLocalId = item.id;
              paragraph.text = item.text;
              return paragraph;
            });
          });
        },
      };
      const context = {
        document: { body: { paragraphs }, getParagraphByUniqueLocalId: proxy },
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

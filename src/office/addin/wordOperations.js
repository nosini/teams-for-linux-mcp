// Word commands run only in the add-in's Office.js context.
((root, factory) => {
  if (typeof module === "object" && module.exports) module.exports = factory(require("./operations"));
  else root.TeamsWordOperations = factory(root.TeamsExcelOperations);
})(globalThis, ({ OfficeSession, bounded, fail, messages, editing }) => {
  const paragraphIdPattern = /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i;
  const hasBreaks = (text) => [...text].some((char) => char.charCodeAt(0) < 32 || ["\u007f", "\u2028", "\u2029"].includes(char));

  function validateText(findText, replacementText) {
    // Word interprets caret codes even with wildcards disabled. Reject codes
    // and paragraph/control characters rather than expanding a search scope.
    if (typeof findText !== "string" || !findText.length || findText.length > 255 ||
      findText.includes("^") || hasBreaks(findText) || typeof replacementText !== "string" ||
      replacementText.length > 8192 || hasBreaks(replacementText) ||
      !bounded({ findText, replacementText })) fail("INVALID_TEXT");
  }

  class WordSession extends OfficeSession {
    constructor(word, options) { super(options); this.word = word; }

    snapshot(paragraph) {
      if (!paragraphIdPattern.test(paragraph.uniqueLocalId)) fail("WRONG_PARAGRAPH");
      const result = { paragraphId: paragraph.uniqueLocalId, text: paragraph.text };
      if (!bounded({ ...result, readToken: "0".repeat(36) })) fail("TOO_LARGE");
      return this.remember(result);
    }

    async execute(request) {
      if (this.running) return { ok: false, code: "BUSY" };
      this.running = true;
      try {
        this.check(request);
        const { operation, args = {} } = request;
        if (editing.wordOperations.includes(operation)) {
          return { ok: true, data: await editing.executeWord(this, request, { fail, bounded }) };
        }
        if (operation === "read_paragraphs") {
          const { start = 0, count = 20 } = args;
          if (!Number.isInteger(start) || start < 0 || start > 100000 ||
            !Number.isInteger(count) || count < 1 || count > 50) fail("INVALID_PAGE");
          const data = await this.word.run(async (context) => {
            const collection = context.document.body.paragraphs;
            collection.load({ select: "uniqueLocalId,text", skip: start, top: count + 1 });
            await context.sync();
            this.check(request);
            const items = collection.items.slice(0, count);
            const paragraphs = items.map((paragraph) => this.snapshot(paragraph));
            const result = { paragraphs, nextStart: collection.items.length > count ? start + count : null };
            if (!bounded(result)) fail("TOO_LARGE");
            return result;
          });
          return { ok: true, data };
        }
        const writing = operation === "replace_text";
        if (operation !== "read_paragraph" && !writing) fail("WORD_ERROR");
        if (!paragraphIdPattern.test(args.paragraphId)) fail("WRONG_PARAGRAPH");
        let expected;
        if (writing) {
          validateText(args.findText, args.replacementText);
          expected = this.consume(args.readToken, "STALE_PARAGRAPH");
          if (expected.paragraphId !== args.paragraphId) fail("STALE_PARAGRAPH");
        }
        const data = await this.word.run(async (context) => {
          const paragraph = context.document.getParagraphByUniqueLocalId(args.paragraphId);
          paragraph.load("uniqueLocalId,text");
          await context.sync();
          this.check(request);
          if (paragraph.uniqueLocalId !== args.paragraphId) fail("WRONG_PARAGRAPH");
          if (writing) {
            if (paragraph.text !== expected.text) fail("STALE_PARAGRAPH");
            const matches = paragraph.search(args.findText, {
              matchCase: true, matchWholeWord: false, matchWildcards: false,
              matchPrefix: false, matchSuffix: false, ignorePunct: false, ignoreSpace: false,
            });
            matches.load({ select: "text", top: 2 });
            // Refresh paragraph text in the same batch as the search, then
            // check again immediately before queuing the replacement.
            paragraph.load("uniqueLocalId,text");
            await context.sync();
            this.check(request);
            if (paragraph.uniqueLocalId !== args.paragraphId || paragraph.text !== expected.text) fail("STALE_PARAGRAPH");
            if (matches.items.length !== 1 || matches.items[0].text !== args.findText) fail("AMBIGUOUS_TEXT");
            matches.items[0].insertText(args.replacementText, "Replace");
            await context.sync();
            paragraph.load("uniqueLocalId,text");
            await context.sync();
          }
          return { ...this.snapshot(paragraph), ...(writing ? { applied: true } : {}) };
        });
        return { ok: true, data };
      } catch (error) {
        const code = error?.code === "ItemNotFound" ? "WRONG_PARAGRAPH" : error?.code;
        return { ok: false, code: Object.hasOwn(messages, code) ? code : "WORD_ERROR" };
      } finally { this.running = false; }
    }
  }
  return { WordSession, validateText, paragraphIdPattern };
});

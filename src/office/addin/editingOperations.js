// Bounded Office.js editing operations, shared by the pane and Node tests.
((root, factory) => {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.TeamsOfficeEditing = factory();
})(globalThis, () => {
  const excelOperations = ["inspect_worksheet", "search_range", "read_range_format", "format_range", "create_worksheet", "insert_rows_columns"];
  const wordOperations = ["read_structure", "read_paragraph_details", "list_tables", "read_table", "replace_paragraph", "insert_paragraph", "delete_paragraph", "format_paragraph", "read_comments", "add_comment", "reply_comment", "resolve_comment", "get_change_tracking", "set_change_tracking", "read_tracked_changes", "resolve_tracked_change"];
  const styles = ["Normal", "Title", "Subtitle", ...Array.from({ length: 9 }, (_, i) => `Heading${i + 1}`), "Quote", "Caption", "ListParagraph"];
  const trackingModes = ["Off", "TrackAll", "TrackMineOnly"];
  const wordAlignments = ["Left", "Centered", "Right", "Justified"];
  const excelAlignments = ["General", "Left", "Center", "Right", "Fill", "Justify", "CenterAcrossSelection", "Distributed"];
  const plain = (value) => value !== null && typeof value === "object" && !Array.isArray(value);
  const short = (value, max = 256) => typeof value === "string" && value.length > 0 && value.length <= max && ![...value].some((c) => c.charCodeAt(0) < 32);
  const paragraphText = (value) => typeof value === "string" && value.length <= 8192 && ![...value].some((c) => c.charCodeAt(0) < 32 || [127, 0x2028, 0x2029].includes(c.charCodeAt(0)));
  const integer = (value, min, max) => Number.isInteger(value) && value >= min && value <= max;
  const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
  const color = (value) => typeof value === "string" && /^#[\da-f]{6}$/i.test(value);
  function validFont(font) {
    return plain(font) && Object.keys(font).length > 0 && Object.entries(font).every(([key, value]) => {
      if (["bold", "italic"].includes(key)) return typeof value === "boolean";
      if (key === "name") return short(value, 100);
      if (key === "size") return Number.isFinite(value) && value >= 1 && value <= 200;
      return key === "color" && color(value);
    });
  }
  function validFormat(format, application) {
    return plain(format) && Object.keys(format).length > 0 && Object.entries(format).every(([key, value]) => {
      if (key === "font") return validFont(value);
      if (application === "word") {
        if (key === "style") return short(value, 128) && !Object.hasOwn(format, "styleBuiltIn");
        if (key === "styleBuiltIn") return styles.includes(value) && !Object.hasOwn(format, "style");
        return key === "alignment" && wordAlignments.includes(value);
      }
      if (key === "numberFormat") return short(value, 256);
      if (key === "fillColor") return color(value);
      if (key === "wrapText") return typeof value === "boolean";
      if (key === "columnWidth" || key === "rowHeight") return Number.isFinite(value) && value >= 0 && (key !== "rowHeight" || value <= 409) && !Object.hasOwn(format, key === "columnWidth" ? "autoFitColumns" : "autoFitRows");
      if (key === "autoFitColumns" || key === "autoFitRows") return value === true && !Object.hasOwn(format, key === "autoFitColumns" ? "columnWidth" : "rowHeight");
      if (key === "horizontalAlignment") return excelAlignments.includes(value);
      return key === "verticalAlignment" && ["Top", "Center", "Bottom", "Justify", "Distributed"].includes(value);
    });
  }
  function validName(name) { return short(name, 31) && !/[\]:*?/\\[]/.test(name) && name.trim().length > 0 && !/^'|'$/.test(name); }
  function validate(application, operation, args) {
    if (!plain(args)) return false;
    const page = () => integer(args.start ?? 0, 0, 100000) && integer(args.count ?? 20, 1, 50);
    if (["read_structure", "list_tables", "read_comments", "inspect_worksheet"].includes(operation) && !page()) return false;
    if (operation === "read_table" && !(integer(args.tableIndex, 0, 100000) && integer(args.rowStart ?? 0, 0, 100000) && integer(args.rowCount ?? 20, 1, 50))) return false;
    if (operation === "search_range" && !(short(args.query, 256) && ["text", "formulas"].includes(args.searchIn ?? "text") && (args.matchCase === undefined || typeof args.matchCase === "boolean") && integer(args.offset ?? 0, 0, 1000))) return false;
    if (["format_range", "format_paragraph"].includes(operation) && !validFormat(args.format, application)) return false;
    if (["replace_paragraph", "insert_paragraph"].includes(operation) && !paragraphText(args.text)) return false;
    if (operation === "insert_paragraph" && (!["Before", "After"].includes(args.position) || (args.styleBuiltIn !== undefined && !styles.includes(args.styleBuiltIn)))) return false;
    if (["add_comment", "reply_comment"].includes(operation) && !(short(args.text, 4096))) return false;
    if (operation === "resolve_comment" && typeof args.resolved !== "boolean") return false;
    if (operation === "set_change_tracking" && !trackingModes.includes(args.mode)) return false;
    if (operation === "resolve_tracked_change" && (!integer(args.changeIndex, 0, 49) || !["accept", "reject"].includes(args.action))) return false;
    if (operation === "create_worksheet" && !validName(args.name)) return false;
    if (operation === "insert_rows_columns" && (!["rows", "columns"].includes(args.axis) || !integer(args.index, 0, args.axis === "rows" ? 1048575 : 16383) || !integer(args.count, 1, args.axis === "rows" ? 100 : 20) || args.index + args.count > (args.axis === "rows" ? 1048576 : 16384))) return false;
    return true;
  }
  const fontData = (font) => Object.fromEntries(["bold", "italic", "name", "size", "color"].map((key) => [key, font[key]]));
  const wordFormat = (paragraph) => ({ style: paragraph.style, styleBuiltIn: paragraph.styleBuiltIn, alignment: paragraph.alignment, font: fontData(paragraph.font) });
  const rangeFormat = (range, dimensions) => ({ ...dimensions, numberFormat: range.numberFormat, font: fontData(range.format.font), fillColor: range.format.fill.color, horizontalAlignment: range.format.horizontalAlignment, verticalAlignment: range.format.verticalAlignment, wrapText: range.format.wrapText });
  const wordLoads = "uniqueLocalId,text,style,styleBuiltIn,alignment,font/bold,font/italic,font/name,font/size,font/color";
  const rangeLoads = "values,formulas,text,rowCount,columnCount,numberFormat,format/font/bold,format/font/italic,format/font/name,format/font/size,format/font/color,format/fill/color,format/horizontalAlignment,format/verticalAlignment,format/wrapText";
  const usedData = (used) => used.isNullObject ? null : Object.fromEntries(["address", "rowIndex", "columnIndex", "rowCount", "columnCount"].map((key) => [key, used[key]]));
  const columnName = (index) => { let name = ""; for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + (n - 1) % 26) + name; return name; };

  async function executeExcel(session, request, { fail, rangeShape, bounded }) {
    const { operation, args = {} } = request;
    if (!validate("excel", operation, args) || !bounded(args)) fail("INVALID_EDIT");
    session.requireApi("ExcelApi", operation === "inspect_worksheet" || operation === "insert_rows_columns" ? "1.4" : "1.1");
    if (operation !== "create_worksheet" && !short(args.worksheetId)) fail("WRONG_SHEET");
    const sizingSupported = session.isSetSupported("ExcelApi", "1.2");
    if (operation === "format_range" && ["columnWidth", "rowHeight", "autoFitColumns", "autoFitRows"].some((key) => Object.hasOwn(args.format, key))) session.requireApi("ExcelApi", "1.2");
    const expected = ["format_range", "create_worksheet", "insert_rows_columns"].includes(operation) ? session.consume(args.readToken) : null;
    if (operation === "create_worksheet") {
      if (expected.kind !== "worksheets") fail("STALE_STRUCTURE");
      return session.excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/id,items/name,items/position");
        await context.sync(); session.check(request);
        const current = sheets.items.map(({ id, name, position }) => ({ id, name, position }));
        if (!same(current, expected.worksheets)) fail("STALE_STRUCTURE");
        if (current.some(({ name }) => name.toLowerCase() === args.name.toLowerCase())) fail("INVALID_EDIT");
        const added = sheets.add(args.name);
        added.load("id,name,position");
        await context.sync(); session.check(request);
        session.snapshots.clear();
        return { applied: true, worksheetId: added.id, name: added.name, position: added.position };
      });
    }
    return session.excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(args.worksheetId);
      sheet.load("id,name");
      if (["inspect_worksheet", "insert_rows_columns"].includes(operation)) {
        const used = sheet.getUsedRangeOrNullObject();
        used.load("address,rowIndex,columnIndex,rowCount,columnCount");
        await context.sync(); session.check(request);
        if (sheet.id !== args.worksheetId) fail("WRONG_SHEET");
        const usedRange = usedData(used);
        if (operation === "insert_rows_columns") {
          if (expected.kind !== "worksheetStructure" || expected.worksheetId !== sheet.id || !same(expected.usedRange, usedRange)) fail("STALE_STRUCTURE");
          const anchor = args.axis === "rows" ? sheet.getRangeByIndexes(args.index, 0, args.count, 1).getEntireRow() : sheet.getRangeByIndexes(0, args.index, 1, args.count).getEntireColumn();
          anchor.insert(args.axis === "rows" ? "Down" : "Right");
          await context.sync(); session.check(request);
          session.snapshots.clear();
          return { applied: true, worksheetId: sheet.id, axis: args.axis, index: args.index, count: args.count };
        }
        const start = args.start ?? 0, count = args.count ?? 20;
        const tables = sheet.tables;
        tables.load({ select: "id,name", skip: start, top: count + 1 });
        await context.sync(); session.check(request);
        const tableRanges = tables.items.slice(0, count).map((table) => { const range = table.getRange(); range.load("address,rowCount,columnCount"); return { table, range }; });
        await context.sync(); session.check(request);
        return session.rememberBounded({ kind: "worksheetStructure", worksheetId: sheet.id, name: sheet.name, usedRange, tables: tableRanges.map(({ table, range }) => ({ id: table.id, name: table.name, address: range.address, rowCount: range.rowCount, columnCount: range.columnCount })), nextStart: tables.items.length > count ? start + count : null });
      }
      const shape = rangeShape(args.address);
      const range = sheet.getRange(shape.address);
      // Inspect each dimension separately: an aggregate null for mixed sizes
      // would otherwise miss coauthor changes within a mixed-size range.
      const columns = [], rows = [];
      if (operation !== "search_range" && sizingSupported) {
        for (let i = 0; i < shape.columns; i++) columns.push(range.getColumn(i));
        for (let i = 0; i < shape.rows; i++) rows.push(range.getRow(i));
      }
      const loadDimensions = () => {
        for (const column of columns) column.load("format/columnWidth");
        for (const row of rows) row.load("format/rowHeight");
      };
      const dimensions = () => sizingSupported ? { columnWidths: columns.map((column) => column.format.columnWidth), rowHeights: rows.map((row) => row.format.rowHeight) } : {};
      loadDimensions();
      range.load(operation === "search_range" ? "text,formulas,rowCount,columnCount,rowIndex,columnIndex" : rangeLoads);
      await context.sync(); session.check(request);
      if (sheet.id !== args.worksheetId) fail("WRONG_SHEET");
      if (range.rowCount !== shape.rows || range.columnCount !== shape.columns) fail("INVALID_RANGE");
      if (operation === "search_range") {
        const cells = range[args.searchIn ?? "text"], matches = [];
        const query = args.matchCase ? args.query : args.query.toLowerCase();
        let offset = args.offset ?? 0;
        for (; offset < shape.rows * shape.columns; offset++) {
          const row = Math.floor(offset / shape.columns), column = offset % shape.columns;
          const value = String(cells[row][column] ?? "");
          if ((args.matchCase ? value : value.toLowerCase()).includes(query)) {
            matches.push({ address: `${columnName(range.columnIndex + column)}${range.rowIndex + row + 1}`, value });
            if (matches.length === 50) { offset++; break; }
          }
        }
        const result = { worksheetId: sheet.id, address: shape.address, matches, nextOffset: offset < shape.rows * shape.columns ? offset : null };
        if (!bounded(result)) fail("TOO_LARGE");
        return result;
      }
      let format = rangeFormat(range, dimensions());
      if (operation === "format_range") {
        if (expected.kind !== "rangeFormat" || expected.worksheetId !== sheet.id || expected.address !== shape.address || !same(expected.format, format) || !same(expected.values, range.values) || !same(expected.formulas, range.formulas)) fail("STALE_READ");
        const input = args.format;
        if (input.font) Object.assign(range.format.font, input.font);
        if (input.fillColor) range.format.fill.color = input.fillColor;
        for (const key of ["horizontalAlignment", "verticalAlignment", "wrapText", "columnWidth", "rowHeight"]) if (Object.hasOwn(input, key)) range.format[key] = input[key];
        if (input.numberFormat) range.numberFormat = Array.from({ length: shape.rows }, () => Array(shape.columns).fill(input.numberFormat));
        if (input.autoFitColumns) range.format.autofitColumns();
        if (input.autoFitRows) range.format.autofitRows();
        await context.sync();
        range.load(rangeLoads);
        loadDimensions();
        await context.sync(); session.check(request);
        format = rangeFormat(range, dimensions());
      }
      return session.rememberBounded({ kind: "rangeFormat", worksheetId: sheet.id, address: shape.address, values: range.values, formulas: range.formulas, text: range.text, format, ...(operation === "format_range" ? { applied: true } : {}) });
    });
  }

  const commentFields = "id,content,authorName,creationDate,resolved";
  const commentData = (comment) => ({ id: comment.id, content: comment.content, authorName: comment.authorName, creationDate: comment.creationDate, resolved: comment.resolved });
  const changeData = (change) => ({ author: change.author, date: change.date, text: change.text, type: change.type });
  async function executeWord(session, request, { fail, bounded }) {
    const { operation, args = {} } = request;
    if (!validate("word", operation, args) || !bounded(args)) fail("INVALID_EDIT");
    session.requireApi("WordApi", "1.6");
    const reads = ["read_structure", "read_paragraph_details", "list_tables", "read_table", "read_comments", "get_change_tracking", "read_tracked_changes"];
    const writing = !reads.includes(operation);
    const expected = writing ? session.consume(args.readToken) : null;
    return session.word.run(async (context) => {
      if (["get_change_tracking", "set_change_tracking"].includes(operation)) {
        context.document.load("changeTrackingMode");
        await context.sync(); session.check(request);
        if (writing) {
          if (expected.kind !== "trackingMode" || expected.mode !== context.document.changeTrackingMode) fail("STALE_STRUCTURE");
          context.document.changeTrackingMode = args.mode;
          await context.sync(); session.check(request);
          session.snapshots.clear();
        }
        return session.rememberBounded({ kind: "trackingMode", mode: context.document.changeTrackingMode, ...(writing ? { applied: true } : {}) });
      }
      if (operation === "read_structure") {
        const start = args.start ?? 0, count = args.count ?? 20;
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load({ select: "uniqueLocalId,text,style,styleBuiltIn,alignment", skip: start, top: count + 1 });
        await context.sync(); session.check(request);
        const result = { paragraphs: paragraphs.items.slice(0, count).map((p) => ({ ...session.snapshot(p), style: p.style, styleBuiltIn: p.styleBuiltIn, alignment: p.alignment })), nextStart: paragraphs.items.length > count ? start + count : null };
        if (!bounded(result)) fail("TOO_LARGE");
        return result;
      }
      if (["list_tables", "read_table"].includes(operation)) {
        const start = operation === "read_table" ? args.tableIndex : (args.start ?? 0), count = operation === "read_table" ? 1 : (args.count ?? 20);
        const tables = context.document.body.tables;
        tables.load({ select: "rowCount", skip: start, top: count + 1 });
        await context.sync(); session.check(request);
        if (operation === "list_tables") return session.boundedResult({ tables: tables.items.slice(0, count).map((table, i) => ({ tableIndex: start + i, rowCount: table.rowCount })), nextStart: tables.items.length > count ? start + count : null });
        if (!tables.items.length) fail("STALE_STRUCTURE");
        const table = tables.items[0], rows = table.rows, rowStart = args.rowStart ?? 0, rowCount = args.rowCount ?? 20;
        rows.load({ select: "values,cellCount", skip: rowStart, top: rowCount + 1 });
        await context.sync(); session.check(request);
        return session.boundedResult({ tableIndex: args.tableIndex, rowCount: table.rowCount, rows: rows.items.slice(0, rowCount).map((row, i) => ({ rowIndex: rowStart + i, values: row.values, cellCount: row.cellCount })), nextRowStart: rows.items.length > rowCount ? rowStart + rowCount : null });
      }
      if (!/^[a-f\d]{8}(?:-[a-f\d]{4}){3}-[a-f\d]{12}$/i.test(args.paragraphId)) fail("WRONG_PARAGRAPH");
      const paragraph = context.document.getParagraphByUniqueLocalId(args.paragraphId);
      paragraph.load(operation === "format_paragraph" || operation === "read_paragraph_details" ? wordLoads : "uniqueLocalId,text");
      await context.sync(); session.check(request);
      if (paragraph.uniqueLocalId !== args.paragraphId) fail("WRONG_PARAGRAPH");
      if (writing && (expected.paragraphId !== args.paragraphId || expected.text !== paragraph.text)) fail("STALE_PARAGRAPH");
      if (operation === "read_paragraph_details") return session.rememberBounded({ kind: "paragraphFormat", paragraphId: paragraph.uniqueLocalId, text: paragraph.text, format: wordFormat(paragraph) });
      if (["read_comments", "reply_comment", "resolve_comment"].includes(operation)) {
        const comments = paragraph.getComments();
        // Resolve by stable comment ID, never by the current collection index.
        const start = operation === "read_comments" ? (args.start ?? 0) : 0;
        const count = operation === "read_comments" ? (args.count ?? 20) : 50;
        comments.load({ select: commentFields, skip: start, top: count + 1 });
        paragraph.load("uniqueLocalId,text");
        await context.sync(); session.check(request);
        if (writing && expected.text !== paragraph.text) fail("STALE_PARAGRAPH");
        const selected = writing ? comments.items.find((comment) => comment.id === args.commentId) : null;
        if (writing && (!selected || comments.items.length > 50 || expected.kind !== "comment" || expected.comment.id !== args.commentId || !same(expected.comment, commentData(selected)))) fail("STALE_STRUCTURE");
        if (writing) {
          if (expected.repliesTruncated) fail("STALE_STRUCTURE");
          selected.load(commentFields);
          selected.replies.load({ select: "id,content,authorName,creationDate", top: 21 });
          paragraph.load("uniqueLocalId,text");
          await context.sync(); session.check(request);
          const replies = selected.replies.items.map(({ id, content, authorName, creationDate }) => ({ id, content, authorName, creationDate }));
          if (expected.text !== paragraph.text || !same(expected.comment, commentData(selected)) || !same(expected.replies, replies)) fail("STALE_STRUCTURE");
          if (operation === "reply_comment") selected.reply(args.text);
          else selected.resolved = args.resolved;
          await context.sync(); session.check(request);
          return { applied: true, commentId: selected.id };
        }
        const selectedComments = comments.items.slice(0, count);
        const replies = selectedComments.map((comment) => { comment.replies.load({ select: "id,content,authorName,creationDate", top: 21 }); return comment.replies; });
        await context.sync(); session.check(request);
        return session.boundedResult({ comments: selectedComments.map((comment, i) => session.rememberBounded({ kind: "comment", paragraphId: args.paragraphId, text: paragraph.text, comment: commentData(comment), replies: replies[i].items.slice(0, 20).map(({ id, content, authorName, creationDate }) => ({ id, content, authorName, creationDate })), repliesTruncated: replies[i].items.length > 20 })), nextStart: comments.items.length > count ? start + count : null });
      }
      if (["read_tracked_changes", "resolve_tracked_change"].includes(operation)) {
        const changes = paragraph.getTrackedChanges();
        changes.load({ select: "author,date,text,type", top: 51 });
        const original = paragraph.getReviewedText("Original"), current = paragraph.getReviewedText("Current");
        paragraph.load("uniqueLocalId,text");
        await context.sync(); session.check(request);
        if (changes.items.length > 50) fail("TOO_LARGE");
        const data = changes.items.map(changeData);
        if (writing) {
          if (expected.kind !== "trackedChanges" || expected.text !== paragraph.text || expected.originalText !== original.value || expected.currentText !== current.value || !same(expected.changes, data) || !changes.items[args.changeIndex]) fail("STALE_STRUCTURE");
          changes.items[args.changeIndex][args.action]();
          await context.sync(); session.check(request);
          session.snapshots.clear();
          return { applied: true, paragraphId: args.paragraphId, changeIndex: args.changeIndex, action: args.action };
        }
        return session.rememberBounded({ kind: "trackedChanges", paragraphId: args.paragraphId, text: paragraph.text, originalText: original.value, currentText: current.value, changes: data });
      }
      if (operation === "add_comment") {
        if (![undefined, "paragraphFormat"].includes(expected.kind)) fail("STALE_PARAGRAPH");
        const comment = paragraph.getRange("Content").insertComment(args.text);
        comment.load(commentFields);
        await context.sync(); session.check(request);
        return { applied: true, comment: commentData(comment) };
      }
      if (operation === "format_paragraph") {
        if (expected.kind !== "paragraphFormat" || !same(expected.format, wordFormat(paragraph))) fail("STALE_PARAGRAPH");
        if (args.format.font) Object.assign(paragraph.font, args.format.font);
        for (const key of ["style", "styleBuiltIn", "alignment"]) if (Object.hasOwn(args.format, key)) paragraph[key] = args.format[key];
        await context.sync();
        paragraph.load(wordLoads);
        await context.sync(); session.check(request);
        return session.rememberBounded({ kind: "paragraphFormat", paragraphId: args.paragraphId, text: paragraph.text, format: wordFormat(paragraph), applied: true });
      }
      if (![undefined, "paragraphFormat"].includes(expected.kind)) fail("STALE_PARAGRAPH");
      if (operation === "replace_paragraph") paragraph.getRange("Content").insertText(args.text, "Replace");
      else if (operation === "delete_paragraph") paragraph.delete();
      else if (operation === "insert_paragraph") {
        const inserted = paragraph.insertParagraph(args.text, args.position);
        if (args.styleBuiltIn) inserted.styleBuiltIn = args.styleBuiltIn;
        inserted.load("uniqueLocalId,text");
        await context.sync(); session.check(request);
        session.snapshots.clear();
        return { ...session.snapshot(inserted), applied: true };
      } else fail("WORD_ERROR");
      await context.sync(); session.check(request);
      session.snapshots.clear();
      if (operation === "delete_paragraph") return { applied: true, paragraphId: args.paragraphId };
      paragraph.load("uniqueLocalId,text");
      await context.sync(); session.check(request);
      return { ...session.snapshot(paragraph), applied: true };
    });
  }
  return { excelOperations, wordOperations, executeExcel, executeWord, validate, validFormat, paragraphText, validName, styles, trackingModes, wordAlignments, excelAlignments };
});

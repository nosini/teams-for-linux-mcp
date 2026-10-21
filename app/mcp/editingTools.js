const { z } = require("zod");
const { rangeShape, editing } = require("../office/addin/operations");
const { paragraphIdPattern } = require("../office/addin/wordOperations");

function registerEditingTools(server, bridge) {
  const documentId = z.string().uuid(), readToken = z.string().uuid();
  const worksheetId = z.string().min(1).max(256);
  const paragraphId = z.string().regex(paragraphIdPattern);
  const address = z.string().max(24).refine((value) => { try { rangeShape(value); return true; } catch { return false; } });
  const paragraph = { documentId, paragraphId }, range = { documentId, worksheetId, address };
  const page = { start: z.number().int().min(0).max(100000).default(0), count: z.number().int().min(1).max(50).default(20) };
  const font = z.object({ bold: z.boolean().optional(), italic: z.boolean().optional(), name: z.string().min(1).max(100).optional(), size: z.number().min(1).max(200).optional(), color: z.string().regex(/^#[\da-f]{6}$/i).optional() }).strict();
  const color = z.string().regex(/^#[\da-f]{6}$/i);
  const wordFormat = z.object({ font: font.optional(), style: z.string().min(1).max(128).optional(), styleBuiltIn: z.enum(editing.styles).optional(), alignment: z.enum(editing.wordAlignments).optional() }).strict();
  const excelFormat = z.object({ font: font.optional(), numberFormat: z.string().min(1).max(256).optional(), fillColor: color.optional(), horizontalAlignment: z.enum(editing.excelAlignments).optional(), verticalAlignment: z.enum(["Top", "Center", "Bottom", "Justify", "Distributed"]).optional(), wrapText: z.boolean().optional() }).strict();
  function register(application, operation, description, shape, writing = false) {
    const warning = writing ? " Requires a recent single-use read token. Comparison and writing are not atomic against coauthors. Do not automatically retry an uncertain write; inspect the result first." : "";
    server.registerTool(`${application}_${operation}`, {
      description: description + warning,
      inputSchema: z.object(shape).strict().refine((args) => editing.validate(application, operation, args), "Invalid edit arguments, text or formatting."),
      annotations: { readOnlyHint: !writing, destructiveHint: writing, idempotentHint: !writing, openWorldHint: false },
    }, async ({ documentId: id, ...args }) => {
      try {
        const value = await bridge.execute(id, operation, args);
        return { content: [{ type: "text", text: JSON.stringify(value) }], structuredContent: value };
      } catch (error) { return { isError: true, content: [{ type: "text", text: error.message }] }; }
    });
  }
  register("excel", "inspect_worksheet", "Inspect used-range bounds, name, and a page of structured tables without reading the entire sheet. Empty sheets return usedRange: null. Returns a token for row/column insertion that checks used-range bounds only, not all cell contents or table structure. ExcelApi 1.4.", { documentId, worksheetId, ...page });
  register("excel", "search_range", "Find literal text or formula substrings in an A1 window of up to 1,000 cells. Returns up to 50 matches with cell addresses. Continue with nextOffset using the same window/query. Use inspect_worksheet to plan windows across a large sheet; search does not read beyond the supplied window.", { ...range, query: z.string().min(1).max(256), searchIn: z.enum(["text", "formulas"]).default("text"), matchCase: z.boolean().default(false), offset: z.number().int().min(0).max(1000).default(0) });
  register("excel", "read_range_format", "Read cell values, formulas, text, number formats and basic formatting for up to 1,000 cells. Font/alignment/fill may be null or mixed across a range. Returns the token required by format_range.", range);
  register("excel", "format_range", "Apply a uniform number format, font, fill color, alignment and/or wrapping to a selected rectangle. Requires a token from read_range_format for unchanged values, formulas and inspected formatting. Only supplied properties change.", { ...range, readToken, format: excelFormat }, true);
  register("excel", "create_worksheet", "Create a worksheet with an explicit unique name. Requires a token from a non-truncated excel_list_worksheets result; checks that worksheet IDs, names and positions have not changed. Invalidates earlier document tokens.", { documentId, readToken, name: z.string().min(1).max(31) }, true);
  register("excel", "insert_rows_columns", "Insert whole rows (up to 100) or columns (up to 20) BEFORE a zero-based index. Excel shifts existing content and adjusts references. Requires inspect_worksheet token; checks worksheet ID and used-range bounds, not all cell contents. Coordinate structural edits with coauthors. Invalidates all earlier document tokens. ExcelApi 1.4.", { documentId, worksheetId, readToken, axis: z.enum(["rows", "columns"]), index: z.number().int().min(0).max(1048575), count: z.number().int().min(1).max(100) }, true);

  register("word", "read_structure", "Read a page of main-body paragraph IDs, text, localized styles, built-in styles and alignment. Identify headings using styleBuiltIn (Heading1–Heading9). Returns ordinary paragraph read tokens; page positions may shift during coauthoring. Tables have separate inspection tools.", { documentId, ...page });
  register("word", "read_paragraph_details", "Inspect one stable paragraph ID, text, style, alignment and basic font formatting. Mixed fonts may return null or empty properties. Returns the token required by format_paragraph.", paragraph);
  register("word", "list_tables", "List a page of main-body tables with their current positional indexes and row counts. Indexes can shift during coauthoring; these indexes are for inspection only.", { documentId, ...page });
  register("word", "read_table", "Read a page of table rows and their cell text values, preserving row boundaries and cell counts. Uses a current positional tableIndex from list_tables. Read-only; merged cells may give unequal cell counts. Results are capped at 32 KiB.", { documentId, tableIndex: z.number().int().min(0).max(100000), rowStart: z.number().int().min(0).max(100000).default(0), rowCount: z.number().int().min(1).max(50).default(20) });
  const text = z.string().max(8192).refine(editing.paragraphText);
  register("word", "replace_paragraph", "Replace the complete text of one paragraph (up to 8,192 characters) using its stable ID and an ordinary paragraph read token. Supports long replacements without Word's 255-character search limit. Paragraph marks are excluded. Whole-content replacement can remove inline formatting, fields, links and other rich content; prefer replace_text for small corrections. Existing change-tracking settings apply.", { ...paragraph, readToken, text }, true);
  register("word", "insert_paragraph", "Insert one plain-text paragraph immediately before or after an existing paragraph identified by stable ID and ordinary paragraph read token. Optionally apply a built-in paragraph style. Returns the inserted paragraph ID and a fresh token; invalidates earlier document tokens. Existing change-tracking settings apply.", { ...paragraph, readToken, text, position: z.enum(["Before", "After"]), styleBuiltIn: z.enum(editing.styles).optional() }, true);
  register("word", "delete_paragraph", "Delete an explicitly selected paragraph and its paragraph mark using its stable ID and ordinary read token. May remove rich content and affect neighboring structure. Existing change-tracking settings apply; deletion may remain as a tracked revision. Invalidates earlier document tokens.", { ...paragraph, readToken }, true);
  register("word", "format_paragraph", "Apply a localized/custom style name OR a built-in style, alignment and/or font properties to one paragraph. Font properties affect all paragraph content. Requires a read_paragraph_details token for unchanged text and inspected formatting.", { ...paragraph, readToken, format: wordFormat }, true);
  register("word", "read_comments", "Read a page of comment threads attached to one paragraph, with IDs, author, date, resolution state and up to 20 replies per thread. Returns a token per thread for reply/resolve operations. Read other paragraphs to inspect their comments.", { ...paragraph, ...page });
  register("word", "add_comment", "Attach a plain-text comment to one paragraph's content. Requires an ordinary paragraph read token. This creates an actual Office comment visible to collaborators; no @mentions or notifications are explicitly requested.", { ...paragraph, readToken, text: z.string().min(1).max(4096) }, true);
  const commentId = z.string().min(1).max(256);
  register("word", "reply_comment", "Reply to an explicitly selected comment thread on a paragraph. Requires that thread's read_comments token. Threads with more than 20 replies or paragraphs with more than 50 comments cannot be edited through this tool.", { ...paragraph, readToken, commentId, text: z.string().min(1).max(4096) }, true);
  register("word", "resolve_comment", "Resolve or reopen an explicitly selected paragraph comment thread. Requires that thread's read_comments token. Threads with more than 20 replies or paragraphs with more than 50 comments cannot be edited through this tool.", { ...paragraph, readToken, commentId, resolved: z.boolean() }, true);
  register("word", "get_change_tracking", "Read the document-wide change-tracking mode and obtain a token for explicitly changing it. This does not accept or reject any changes.", { documentId });
  register("word", "set_change_tracking", "Explicitly set document-wide tracking to Off, TrackAll or TrackMineOnly using a get_change_tracking token. This changes a shared document setting; it is not scoped to this agent. Does not accept or reject existing changes.", { documentId, readToken, mode: z.enum(editing.trackingModes) }, true);
  register("word", "read_tracked_changes", "Inspect up to 50 tracked revisions in one stable paragraph, plus Original and Current reviewed text. Returns a token bound to the ordered revisions and reviewed text. WordApi 1.6; large revision sets are rejected rather than silently truncated.", paragraph);
  register("word", "resolve_tracked_change", "Accept or reject exactly one revision by its zero-based index in a read_tracked_changes result. Requires that result's token and an unchanged revision list and reviewed text. Invalidates earlier document tokens. WordApi 1.6.", { ...paragraph, readToken, changeIndex: z.number().int().min(0).max(49), action: z.enum(["accept", "reject"]) }, true);
}
module.exports = { registerEditingTools };

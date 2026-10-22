const { McpServer } = require("@modelcontextprotocol/server");
const { z } = require("zod");
const { rangeShape, validateData } = require("../office/addin/operations");
const { validateText, paragraphIdPattern } = require("../office/addin/wordOperations");
const { registerJsonTool } = require("./tools");

function createOfficeMcpServer(inspector, liveBridge = null, channelFiles = null, surfaceInspector = null, codeMirrorBridge = null, documentTabs = null) {
  const server = new McpServer({ name: "teams-for-linux-office", version: "0.1.0" }, {
    instructions: liveBridge ? "List connected documents first and choose tools matching the application. Excel writes require a read token from the same worksheet and range. Word text edits require a read token from the same paragraph; literal replacements require exactly one matching occurrence. Formatting and structure tools specify their own inspection tokens. Inspect comments/revisions before explicitly editing them; do not automatically change tracking mode or accept/reject changes. Do not automatically retry a timed-out write. API presence diagnostics alone do not indicate a live bridge. To create a Word file without a pasted link, discover its folder with office_list_locations and call office_create_document; Electron creates it, opens it hidden and starts Bridge setup. For another existing folder, use office_resolve_location with a source location, decoded folderPath and optional owning webUrl, or browse office_list_location_folders. Inspect an uncertain creation with office_list_location_files before retrying." : "This server discovers Office frames in Teams. API presence is diagnostic only. Channel file tools, when enabled, use the signed-in Teams session. List teams, channels and files before choosing a file. Do not automatically retry timed-out writes.",
  });
  // Every tool here acts on the local Teams client, not Microsoft services.
  const tool = (name, options, operation) => registerJsonTool(server, name, { ...options, openWorld: false }, operation);

  tool("office_list_editors", {
    description: "List candidate Word/Excel frames and SharePoint document wrappers in the running Teams client. Returns temporary IDs and origins, without document content or full URLs.",
    inputSchema: z.object({}).strict(),
  }, () => inspector.listEditors());

  tool("office_inspect_editor", {
    description: "Probe a listed frame for Office.js API presence without reading or changing the document. Presence alone does not confirm that live editing is available.",
    inputSchema: z.object({ editorId: z.string().regex(/^\d+:\d+:\d+$/).max(80) }).strict(),
    errorMessage: () => "Editor is no longer available. Call office_list_editors again.",
  }, ({ editorId }) => inspector.inspectEditor(editorId));
  if (liveBridge) {
    const documentId = z.string().uuid();
    const worksheetId = z.string().min(1).max(256);
    const address = z.string().max(24).refine((value) => {
      try { rangeShape(value); return true; } catch { return false; }
    }, "Use an A1 rectangle of at most 1,000 cells.");
    tool("office_list_documents", {
      description: "List Excel and Word documents connected through the installed Teams Office Bridge task pane. Open the corresponding add-in in each document to connect it. Returns temporary document IDs and application names; call this before document operations.",
      inputSchema: z.object({}).strict(),
    }, () => liveBridge.listDocuments());
    const paragraphId = z.string().regex(paragraphIdPattern);
    tool("word_read_paragraphs", {
      description: "Read a page of main-body Word paragraphs with session-local paragraph IDs, text and single-use two-minute read tokens. start is a zero-based discovery offset; use paragraph IDs for edits. Pages can shift during coauthoring. Returns nextStart, or null at the end. Headers, footers and text boxes are outside this scope. Requires WordApi 1.6.",
      inputSchema: z.object({ documentId, start: z.number().int().min(0).max(100000).default(0), count: z.number().int().min(1).max(50).default(20) }).strict(),
    }, ({ documentId: id, ...args }) => liveBridge.execute(id, "read_paragraphs", args));
    tool("word_read_paragraph", {
      description: "Read one Word paragraph by its session-local ID and obtain a fresh single-use read token. Use this to refresh a paragraph after edits or an uncertain write.",
      inputSchema: z.object({ documentId, paragraphId }).strict(),
    }, ({ documentId: id, ...args }) => liveBridge.execute(id, "read_paragraph", args));
    tool("word_replace_text", {
      description: "Replace exactly one case-sensitive literal text match inside an explicitly selected Word paragraph. Requires a recent read token for unchanged paragraph text. findText is 1–255 characters without caret codes; replacementText is at most 8,192 characters, with an empty string deleting the match. Control characters and paragraph breaks are rejected. Edits only the matched range; replacement formatting follows Word. Coauthor comparison and replacement are not atomic. Read again after uncertain writes.",
      inputSchema: z.object({ documentId, paragraphId, readToken: z.string().uuid(), findText: z.string().min(1).max(255), replacementText: z.string().max(8192) }).strict()
        .refine((args) => { try { validateText(args.findText, args.replacementText); return true; } catch { return false; } }, "Use bounded literal text without caret search codes or paragraph/control characters."),
      writing: true,
    }, ({ documentId: id, ...args }) => liveBridge.execute(id, "replace_text", args));
    tool("excel_list_worksheets", {
      description: "List worksheet IDs and names in an explicitly selected connected Excel document.",
      inputSchema: z.object({ documentId }).strict(),
    }, ({ documentId: id }) => liveBridge.execute(id, "list_worksheets", {}));
    tool("excel_read_range", {
      description: "Read values, formulas and displayed text from an A1 rectangle of up to 1,000 cells, addressed by worksheet ID. Returns a single-use read token valid for two minutes, required for writes to this exact range.",
      inputSchema: z.object({ documentId, worksheetId, address }).strict(),
    }, ({ documentId: id, ...args }) => liveBridge.execute(id, "read_range", args));
    for (const [name, operation, description] of [
      ["excel_write_range", "write_range", "Write literal values to a range. Use numbers, booleans or strings; an empty string clears a cell. Text beginning with =, + or - is rejected. Requires a recent read token; rejects changed cells. The comparison and write are not atomic against coauthors."],
      ["excel_set_formulas", "set_formulas", "Set formulas starting with = in every cell of a range. Requires a recent read token; rejects changed cells. The comparison and write are not atomic against coauthors."],
    ]) {
      tool(name, {
        description,
        inputSchema: z.object({
          documentId, worksheetId, address, readToken: z.string().uuid(),
          data: z.array(z.array(z.union([z.string().max(8192), z.number().finite(), z.boolean()])).min(1).max(1000)).min(1).max(1000),
        }).strict().refine((args) => {
          try { validateData(args.address, args.data, operation === "set_formulas"); return true; } catch { return false; }
        }, "Data must match the range and fit within 32 KB. Use the formula tool for formulas."),
        writing: true,
      }, ({ documentId: id, ...args }) => liveBridge.execute(id, operation, args));
    }
  }
  if (surfaceInspector) {
    tool("editor_list_frames", {
      description: "List HTTPS frame IDs and origins in the running Teams client, including non-Office file editors. Returns no document URLs, names or content. inspectable identifies Microsoft-hosted frames eligible for the fixed read-only editor probe.",
      inputSchema: z.object({}).strict(),
    }, () => surfaceInspector.listFrames());
    tool("editor_inspect_frame", {
      description: "Run a fixed read-only probe on a listed inspectable frame to identify editor types (Monaco, CodeMirror, ProseMirror, Lexical, Slate, Quill), editable-element counts and adapter-handle presence. Does not read text or invoke editor/save/collaboration methods. Presence does not prove model or coauthor access. No caller JavaScript is accepted.",
      inputSchema: z.object({ frameId: z.string().regex(/^\d+:\d+:\d+$/).max(80) }).strict(),
      errorMessage: () => "Select an inspectable frame from editor_list_frames again.",
    }, ({ frameId }) => surfaceInspector.inspectFrame(frameId));
  }
  if (liveBridge) require("./editingTools").registerEditingTools(server, liveBridge);
  if (channelFiles) require("./channelFiles").registerChannelFileTools(server, channelFiles);
  if (codeMirrorBridge) require("./codeMirrorTools").registerCodeMirrorTools(server, codeMirrorBridge);
  if (documentTabs) require("./documentTabTools").registerDocumentTabTools(server, documentTabs);
  return server;
}

module.exports = { createOfficeMcpServer };

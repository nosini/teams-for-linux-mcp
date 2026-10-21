const { McpServer } = require("@modelcontextprotocol/server");
const { z } = require("zod");
const { rangeShape, validateData } = require("../office/addin/operations");

function createOfficeMcpServer(inspector, liveBridge = null) {
  const server = new McpServer({ name: "teams-for-linux-office", version: "0.1.0" }, {
    instructions: liveBridge ? "List connected documents first. Excel writes require a read token from a recent read of the same worksheet and range. Do not automatically retry a timed-out write. API presence diagnostics alone do not indicate a live bridge." : "This server discovers Office frames in Teams. It cannot yet read or edit document content. API presence is diagnostic only.",
  });
  const annotations = { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false };
  const result = (value) => ({ content: [{ type: "text", text: JSON.stringify(value) }], structuredContent: value });

  server.registerTool("office_list_editors", {
    description: "List candidate Word/Excel frames and SharePoint document wrappers in the running Teams client. Returns temporary IDs and origins, without document content or full URLs.",
    inputSchema: z.object({}).strict(),
    annotations,
  }, () => result(inspector.listEditors()));

  server.registerTool("office_inspect_editor", {
    description: "Probe a listed frame for Office.js API presence without reading or changing the document. Presence alone does not confirm that live editing is available.",
    inputSchema: z.object({ editorId: z.string().regex(/^\d+:\d+:\d+$/).max(80) }).strict(),
    annotations,
  }, async ({ editorId }) => {
    try {
      return result(await inspector.inspectEditor(editorId));
    } catch {
      return { isError: true, content: [{ type: "text", text: "Editor is no longer available. Call office_list_editors again." }] };
    }
  });
  if (liveBridge) {
    const documentId = z.string().uuid();
    const worksheetId = z.string().min(1).max(256);
    const address = z.string().max(24).refine((value) => {
      try { rangeShape(value); return true; } catch { return false; }
    }, "Use an A1 rectangle of at most 1,000 cells.");
    const respond = async (operation) => {
      try { return result(await operation()); } catch (error) {
        return { isError: true, content: [{ type: "text", text: error.message }] };
      }
    };
    server.registerTool("office_list_documents", {
      description: "List Excel documents connected through the installed Teams Office Bridge task pane. Open the add-in in each workbook to connect it. Returns temporary document IDs; call this before document operations.",
      inputSchema: z.object({}).strict(), annotations,
    }, () => respond(() => liveBridge.listDocuments()));
    server.registerTool("excel_list_worksheets", {
      description: "List worksheet IDs and names in an explicitly selected connected Excel document.",
      inputSchema: z.object({ documentId }).strict(), annotations,
    }, ({ documentId: id }) => respond(() => liveBridge.execute(id, "list_worksheets", {})));
    server.registerTool("excel_read_range", {
      description: "Read values, formulas and displayed text from an A1 rectangle of up to 1,000 cells, addressed by worksheet ID. Returns a single-use read token valid for two minutes, required for writes to this exact range.",
      inputSchema: z.object({ documentId, worksheetId, address }).strict(), annotations,
    }, ({ documentId: id, ...args }) => respond(() => liveBridge.execute(id, "read_range", args)));
    for (const [name, operation, description] of [
      ["excel_write_range", "write_range", "Write literal values to a range. Use numbers, booleans or strings; an empty string clears a cell. Text beginning with =, + or - is rejected. Requires a recent read token; rejects changed cells. The comparison and write are not atomic against coauthors."],
      ["excel_set_formulas", "set_formulas", "Set formulas starting with = in every cell of a range. Requires a recent read token; rejects changed cells. The comparison and write are not atomic against coauthors."],
    ]) {
      server.registerTool(name, {
        description,
        inputSchema: z.object({
          documentId, worksheetId, address, readToken: z.string().uuid(),
          data: z.array(z.array(z.union([z.string().max(8192), z.number().finite(), z.boolean()])).min(1).max(1000)).min(1).max(1000),
        }).strict().refine((args) => {
          try { validateData(args.address, args.data, operation === "set_formulas"); return true; } catch { return false; }
        }, "Data must match the range and fit within 32 KB. Use the formula tool for formulas."),
        annotations: { ...annotations, readOnlyHint: false, destructiveHint: true, idempotentHint: false },
      }, ({ documentId: id, ...args }) => respond(() => liveBridge.execute(id, operation, args)));
    }
  }
  return server;
}

module.exports = { createOfficeMcpServer };

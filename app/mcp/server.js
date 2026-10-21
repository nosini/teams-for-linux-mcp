const { McpServer } = require("@modelcontextprotocol/server");
const { z } = require("zod");

function createOfficeMcpServer(inspector) {
  const server = new McpServer({ name: "teams-for-linux-office", version: "0.1.0" }, {
    instructions: "This server discovers Office frames in Teams. It cannot yet read or edit document content. API presence is diagnostic only.",
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
  return server;
}

module.exports = { createOfficeMcpServer };

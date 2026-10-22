const { z } = require("zod");
const { registerJsonTool } = require("./tools");
function registerCodeMirrorTools(server, bridge) {
  const register = (name, description, shape, operation, writing = false) =>
    registerJsonTool(server, name, { description, inputSchema: z.object(shape).strict(), writing, openWorld: false }, operation);
  register("codemirror_list_editors", "Discover live CodeMirror 6 editors in SharePoint frames, including the native Markdown editor. Uses version-specific DOM links; unsupported instances count as unavailable. IDs bind to an editor instance, not a file name. Identify the intended editor with the user before editing. Requires mcp.codemirror.enabled. Native saving and coauthor compatibility are not yet verified.", {}, () => bridge.listEditors());
  const editorId = z.string().uuid();
  register("codemirror_read_text", "Read a window of the live editor model, including unsaved text, without using the DOM's visible text. start/count are UTF-16 code units, count at most 8,192 (default 4,096). Newlines are LF. Returns a two-minute single-use token bound to this editor instance and the entire document revision; use nextStart for another window.", { editorId, start: z.number().int().min(0).default(0), count: z.number().int().min(1).max(8192).default(4096) }, (args) => bridge.execute("read", args));
  register("codemirror_replace_text", "Replace live text in a half-open UTF-16 [start,end) range fully contained in the read window. Equal offsets insert; empty text deletes. Requires that window's recent read token and unchanged entire document revision. Dispatches through the existing CodeMirror transaction path; never replaces the SharePoint file or rewrites DOM text. Read-only/IME composition states reject writes. applied confirms local model readback only; saveStatus stays unverified. Native save/coauthor behavior needs live testing. Never automatically retry uncertain writes; reread first.", {
    editorId, readToken: z.string().uuid(), start: z.number().int().min(0), end: z.number().int().min(0),
    text: z.string().max(8192).refine((text) => text.isWellFormed() && !text.includes("\r") && !text.includes("\0") && !text.includes("\u2028") && !text.includes("\u2029")),
  }, (args) => bridge.execute("replace", args), true);
}
module.exports = { registerCodeMirrorTools };

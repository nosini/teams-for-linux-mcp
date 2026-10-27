const { z } = require("zod");
const { bounded } = require("../office/addin/operations");
const { documentId } = require("./officeSchemas");
const { registerJsonTool } = require("./tools");

// Reuse the individual tools' schemas and dispatchers so batching cannot bypass
// their validation, read-token checks, application checks or per-document lock.
function registerBatchTools(server, definitions, { now = Date.now, maxDurationMs = 30000 } = {}) {
  for (const [application, operations] of [
    ["word", ["replace_text", "replace_paragraph", "insert_paragraph", "delete_paragraph", "format_paragraph"]],
    ["excel", ["write_range", "set_formulas", "format_range"]],
  ]) {
    const available = definitions.filter(({ name }) => operations.some((operation) => name === `${application}_${operation}`));
    if (!available.length) continue;
    const byOperation = new Map(available.map((definition) => [definition.name.slice(application.length + 1), definition]));
    const editSchemas = available.map(({ name, inputSchema }) => {
      const shape = Object.fromEntries(Object.entries(inputSchema.shape).filter(([key]) => !["documentId", "responseMode"].includes(key)));
      return z.object({ operation: z.literal(name.slice(application.length + 1)), ...shape }).strict();
    });
    const inputSchema = z.object({ documentId, edits: z.array(z.discriminatedUnion("operation", editSchemas)).min(1).max(20) }).strict()
      .superRefine((args, context) => {
        const tokens = new Set();
        for (const [index, { operation, ...edit }] of args.edits.entries()) {
          const parsed = byOperation.get(operation).inputSchema.safeParse({ ...edit, documentId: args.documentId });
          if (!parsed.success || tokens.has(edit.readToken)) context.addIssue({ code: "custom", path: ["edits", index], message: "Use valid single-edit arguments and a distinct read token per edit." });
          tokens.add(edit.readToken);
        }
        if (!bounded(args)) context.addIssue({ code: "custom", message: "Batch arguments must fit within 32 KiB." });
      });
    registerJsonTool(server, `${application}_batch_edit`, {
      description: `Apply up to 20 ${application} edits to one document in one call (32 KiB total). Prefer this for independently inspected targets. Each edit uses the arguments of the corresponding ${application}_ tool, without documentId/responseMode. Validates all arguments first, then executes sequentially with ordinary single-use token and coauthor checks. Returns compact confirmations and fresh tokens. Stops at the first failure or time budget; completed edits are not rolled back. Inspect a failed edit before retrying; never replay the whole batch.`,
      inputSchema, writing: true, openWorld: false,
    }, async ({ documentId: id, edits }) => {
      const results = [], started = now();
      for (const [index, { operation, ...edit }] of edits.entries()) {
        if (now() - started >= maxDurationMs) return { completed: results.length, results, nextIndex: index, stopped: "time_budget", notAttempted: edits.length - index };
        try {
          const result = await byOperation.get(operation).operation({ ...edit, documentId: id, responseMode: "compact" });
          results.push({ index, ...result });
        } catch (error) {
          return { completed: results.length, results, failedIndex: index, error: error.message, uncertain: true, notAttempted: edits.length - index - 1 };
        }
      }
      return { completed: results.length, results };
    });
  }
}

module.exports = { registerBatchTools };

const read = { readOnlyHint: true, destructiveHint: false, idempotentHint: true };
const write = { readOnlyHint: false, destructiveHint: true, idempotentHint: false };

const jsonResult = (value) => ({ content: [{ type: "text", text: JSON.stringify(value) }], structuredContent: value });
const toolError = (message) => ({ isError: true, content: [{ type: "text", text: message }] });

// Registers a tool that returns its value as JSON text and structured content.
// Bridges throw only sanitized messages; tools with a fixed reply pass
// errorMessage. openWorld states whether the tool reaches Microsoft services
// rather than the local Teams client.
function registerJsonTool(server, name, { description, inputSchema, writing = false, openWorld, errorMessage = (error) => error.message }, operation) {
  if (typeof openWorld !== "boolean") throw new TypeError(`State whether ${name} is open-world.`);
  server.registerTool(name, {
    description, inputSchema, annotations: { ...(writing ? write : read), openWorldHint: openWorld },
  }, async (args) => {
    try { return jsonResult(await operation(args)); }
    catch (error) { return toolError(errorMessage(error)); }
  });
  return { name, inputSchema, operation };
}

module.exports = { registerJsonTool };

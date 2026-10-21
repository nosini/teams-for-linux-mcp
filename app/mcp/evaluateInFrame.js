// Runs a fixed script in a renderer frame without a user gesture. Rejects with
// `message` after `timeoutMs`; the renderer may still finish the script.
async function evaluateInFrame(frame, source, timeoutMs, message) {
  let timer;
  try {
    return await Promise.race([
      frame.executeJavaScript(source, false),
      new Promise((_, reject) => { timer = setTimeout(() => reject(new Error(message)), timeoutMs); }),
    ]);
  } finally { clearTimeout(timer); }
}

module.exports = { evaluateInFrame };

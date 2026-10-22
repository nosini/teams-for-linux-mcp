// Runs a fixed script, without a user gesture unless explicitly requested. Rejects with
// `message` after `timeoutMs`; the renderer may still finish the script.
async function evaluateInFrame(frame, source, timeoutMs, message, userGesture = false) {
  let timer;
  try {
    return await Promise.race([
      frame.executeJavaScript(source, userGesture),
      new Promise((_, reject) => { timer = setTimeout(() => reject(new Error(message)), timeoutMs); }),
    ]);
  } finally { clearTimeout(timer); }
}

module.exports = { evaluateInFrame };

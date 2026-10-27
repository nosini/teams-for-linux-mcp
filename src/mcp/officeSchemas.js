const { z } = require("zod");
const { rangeShape } = require("../office/addin/operations");
const { paragraphIdPattern } = require("../office/addin/wordOperations");

const documentId = z.string().uuid();
const worksheetId = z.string().min(1).max(256);
const paragraphId = z.string().regex(paragraphIdPattern);
const responseMode = z.enum(["compact", "full"]).default("compact").describe("Compact reads avoid duplicate cell matrices; successful writes return IDs and fresh tokens. Full returns complete inspected content. Prefer compact.");
const address = z.string().max(24).refine((value) => {
  try { rangeShape(value); return true; } catch { return false; }
}, "Use an A1 rectangle of at most 1,000 cells.");

module.exports = { documentId, worksheetId, paragraphId, address, responseMode };

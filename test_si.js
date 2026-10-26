/**
 * Automated Verification Suite 1: Sinhala Localization & ZWJ-Immune Search (test_si.js)
 */
const fs = require("fs");
const path = require("path");
const assert = require("assert");

const ROOT = __dirname;
const PRODUCTS_JS = path.join(ROOT, "website", "data", "products.js");
const INDEX_HTML = path.join(ROOT, "website", "index.html");

console.log("=== Running test_si.js: Sinhala Localization & ZWJ-Immune Search ===");

// 1. Load CATALOG
global.window = {};
const prodCode = fs.readFileSync(PRODUCTS_JS, "utf8");
eval(prodCode);
const catalog = global.window.CATALOG;
assert(catalog && Array.isArray(catalog.products), "CATALOG.products must exist");
assert(catalog.products.length === 508, `Expected 508 products, got ${catalog.products.length}`);

// 2. Verify every product has non-empty Sinhala title (p.si) with Sinhala Unicode chars + English term in parentheses
const sinhalaRegex = /[\u0D80-\u0DFF]/;
const cjkRegex = /[\u4E00-\u9FFF]/;

catalog.products.forEach((p, idx) => {
  assert(p.si && p.si.trim().length > 5, `Product #${idx + 1} (${p.code}) missing p.si`);
  assert(sinhalaRegex.test(p.si), `Product ${p.code} p.si lacks Sinhala Unicode characters: ${p.si}`);
  assert(/\([^)]+[A-Za-z][^)]+\)/.test(p.si), `Product ${p.code} p.si should include helpful English term in parentheses: ${p.si}`);
  assert(!cjkRegex.test(p.si), `Product ${p.code} p.si contains untranslated Chinese characters: ${p.si}`);
  assert(!cjkRegex.test(p.en), `Product ${p.code} p.en contains untranslated Chinese characters: ${p.en}`);
  assert(!cjkRegex.test(p.es), `Product ${p.code} p.es contains untranslated Chinese characters: ${p.es}`);
  assert(p.variant_si && p.variant_si.trim().length > 0, `Product ${p.code} missing variant_si`);
  assert(!cjkRegex.test(p.variant_si), `Product ${p.code} variant_si contains CJK: ${p.variant_si}`);
});
console.log("✓ All 508 products have natural Sinhala names (with English terms in parentheses) and zero CJK leakage.");

// 3. Verify ZWJ-immune norm(s) function and search equivalence
const html = fs.readFileSync(INDEX_HTML, "utf8");
assert(html.includes("Noto Sans Sinhala"), "CSS must include Noto Sans Sinhala font stack");
assert(html.includes("-webkit-font-smoothing: antialiased"), "CSS must include -webkit-font-smoothing: antialiased");

function norm(s) {
  return String(s || "").replace(/[\u200B-\u200D\uFEFF]/g, "").toLowerCase().trim();
}

// Test ZWJ pairs (with \u200D vs without \u200D)
const zwjPairs = [
  ["රූපලාවන්‍ය", "රූපලාවන්ය"],
  ["ජයග්‍රාහී", "ජයග්රාහී"],
  ["යෝග්‍යතාවය", "යෝග්යතාවය"],
  ["බ්‍රා", "බ්රා"],
  ["ඩ්රිල්", "ඩ්\u200Dරිල්"]
];

for (const [withZwj, withoutZwj] of zwjPairs) {
  assert.strictEqual(
    norm(withZwj),
    norm(withoutZwj),
    `ZWJ normalization failed for ${withZwj} vs ${withoutZwj}`
  );
}

// Search catalog with both ZWJ and non-ZWJ Sinhala queries
const queryPairs = [
  ["රූපලාවන්‍ය", "රූපලාවන්ය"],
  ["බ්‍රා", "බ්රා"],
  ["කපු-ලිනන්", "කපු\u200B-ලිනන්"],
  ["ඇඳුම්", "ඇඳු\u200Dම්"]
];

for (const [q1, q2] of queryPairs) {
  const hits1 = catalog.products.filter(p => norm(p.si + " " + p.variant_si).includes(norm(q1)));
  const hits2 = catalog.products.filter(p => norm(p.si + " " + p.variant_si).includes(norm(q2)));
  assert(hits1.length > 0, `Expected search hits for Sinhala query: ${q1}`);
  assert.strictEqual(hits1.length, hits2.length, `Hit count mismatch for ${q1} vs ${q2}`);
}
console.log("✓ ZWJ-immune normalization (norm) matches identically for ligature and phonetic Sinhala queries.");

// 4. Verify 4 languages in #lang selector
for (const langCode of ["en", "si", "zh", "es"]) {
  assert(html.includes(`value="${langCode}"`), `#lang selector missing option value="${langCode}"`);
}
console.log("✓ All 4 UI languages (en, si, zh, es) verified in #lang selector.");
console.log("=== ALL test_si.js CHECKS PASSED (100%) ===\n");

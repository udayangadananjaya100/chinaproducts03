/**
 * Automated Verification Suite 4: Data Integrity, SL-WPI Math & Edge Cases (test_edge_cases.js)
 */
const fs = require("fs");
const path = require("path");
const assert = require("assert");

const ROOT = __dirname;
const WEBSITE_DIR = path.join(ROOT, "website");
const PRODUCTS_JS = path.join(WEBSITE_DIR, "data", "products.js");

console.log("=== Running test_edge_cases.js: Data Integrity, SL-WPI Math & Edge Cases ===");

global.window = {};
eval(fs.readFileSync(PRODUCTS_JS, "utf8"));
const catalog = global.window.CATALOG;
const products = catalog.products;

assert.strictEqual(products.length, 508, `Expected 508 products across all 3 spreadsheets, got ${products.length}`);

const seenCodes = new Set();
const tierCounts = { mega: 0, strong: 0, niche: 0, low: 0 };

products.forEach((p, idx) => {
  // 1. Unique ID & Code
  assert.strictEqual(p.id, idx + 1, `Product index ${idx} id mismatch`);
  assert(p.code && p.code.trim().length > 0, `Product #${p.id} missing code`);
  assert(!seenCodes.has(p.code), `Duplicate product code: ${p.code}`);
  seenCodes.add(p.code);

  // 2. Packing numeric validation
  assert(Number.isInteger(p.pcsPerCarton) && p.pcsPerCarton >= 1, `Invalid pcsPerCarton on ${p.code}: ${p.pcsPerCarton}`);
  assert(typeof p.price === "number" && p.price > 0, `Invalid price on ${p.code}: ${p.price}`);
  assert(typeof p.cbm === "number" && p.cbm > 0, `Invalid cbm on ${p.code}: ${p.cbm}`);
  assert(p.kg === null || (typeof p.kg === "number" && p.kg > 0), `Invalid kg on ${p.code}: ${p.kg}`);

  // 3. Image existence in website/images/
  assert(p.img && p.img.endsWith("_t.webp"), `Product ${p.code} missing _t.webp thumbnail`);
  assert(p.imgLarge && p.imgLarge.endsWith("_l.webp"), `Product ${p.code} missing _l.webp large image`);
  const thumbPath = path.join(WEBSITE_DIR, p.img);
  const largePath = path.join(WEBSITE_DIR, p.imgLarge);
  assert(fs.existsSync(thumbPath), `Missing thumbnail file on disk: ${thumbPath}`);
  assert(fs.existsSync(largePath), `Missing large image file on disk: ${largePath}`);

  // 4. SL-WPI 100-Point Sea Cargo Model Mathematical Invariant
  const bd = p.winBreakdown;
  assert(bd && bd.freight && bd.margin && bd.demand && bd.risk, `Missing winBreakdown pillars on ${p.code}`);
  for (const k of ["freight", "margin", "demand", "risk"]) {
    const sc = bd[k].score;
    assert(Number.isInteger(sc) && sc >= 0 && sc <= 25, `Invalid ${k} score on ${p.code}: ${sc}`);
    for (const lang of ["en", "si", "zh", "es"]) {
      assert(bd[k].desc && bd[k].desc[lang] && bd[k].desc[lang].length > 10, `Missing ${k}.desc.${lang} on ${p.code}`);
    }
  }

  const exactSum = bd.freight.score + bd.margin.score + bd.demand.score + bd.risk.score;
  assert.strictEqual(
    exactSum,
    p.winScore,
    `SL-WPI math invariant failed on ${p.code}: ${exactSum} !== ${p.winScore}`
  );
  assert.strictEqual(p.winRate, `${p.winScore}%`, `winRate mismatch on ${p.code}`);

  // 5. Tier classification boundaries
  if (p.winScore >= 85) assert.strictEqual(p.winTier, "mega");
  else if (p.winScore >= 70) assert.strictEqual(p.winTier, "strong");
  else if (p.winScore >= 50) assert.strictEqual(p.winTier, "niche");
  else assert.strictEqual(p.winTier, "low");

  tierCounts[p.winTier]++;

  // 6. Verify low-scoring items do NOT have hardcoded positive text
  if (p.winTier === "low") {
    assert(
      !p.winReason.en.includes("Mega Winner") && !p.winReason.en.includes("Strong Winner"),
      `Low-tier item ${p.code} has positive tier text in winReason`
    );
  }
});

console.log("✓ All 508 products verified: unique codes, valid packing numbers, existing WebP pairs, and 100% exact SL-WPI math:", tierCounts);

// 7. Verify deployment & launcher files (.gitignore, vercel.json, Open_Website.bat, root index.html)
const gitignore = fs.readFileSync(path.join(ROOT, ".gitignore"), "utf8");
assert(gitignore.includes("*.xlsx") && gitignore.includes("*.xls"), ".gitignore must exclude *.xlsx and *.xls");

const vercelJson = JSON.parse(fs.readFileSync(path.join(ROOT, "vercel.json"), "utf8"));
assert(
  Array.isArray(vercelJson.redirects) &&
  vercelJson.redirects.some(r => r.source === "/" && r.destination === "/website/"),
  "vercel.json must configure root redirect to /website/"
);

const bat = fs.readFileSync(path.join(ROOT, "Open_Website.bat"), "utf8");
assert(bat.includes('start "" "%~dp0website\\index.html"'), "Open_Website.bat must launch website\\index.html");

assert(fs.existsSync(path.join(ROOT, "index.html")), "Root index.html redirect must exist");
console.log("✓ Deployment & launcher files (.gitignore, vercel.json, Open_Website.bat, index.html) verified.");
console.log("=== ALL test_edge_cases.js CHECKS PASSED (100%) ===\n");

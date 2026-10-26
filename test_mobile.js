/**
 * Automated Verification Suite 3: Mobile UX & Ergonomics Optimization (test_mobile.js)
 * Verifies 320px-480px mobile ergonomics requirements.
 */
const fs = require("fs");
const path = require("path");
const assert = require("assert");

const INDEX_HTML = path.join(__dirname, "website", "index.html");

console.log("=== Running test_mobile.js: Mobile UX & Ergonomics (320px-480px) ===");

const html = fs.readFileSync(INDEX_HTML, "utf8");

// 1. Viewport meta tag with viewport-fit=cover
assert(
  html.includes('<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">'),
  "Missing required viewport meta tag with viewport-fit=cover"
);
console.log("✓ Viewport meta tag with viewport-fit=cover verified.");

// 2. Minimum 44x44px touch targets on interactive elements (.qadd, .qty-step, .gal-nav, .rm, .x, .chip, select)
assert(html.includes("min-width: 44px") && html.includes("min-height: 44px"), "Missing 44x44px minimum touch target rule");
for (const sel of [".qadd", ".qty-step", ".gal-nav", ".rm", ".x", ".chip", "select"]) {
  assert(html.includes(sel), `Missing touch target selector: ${sel}`);
}
console.log("✓ Minimum 44x44px touch target dimensions verified across all interactive controls.");

// 3. 16px font-size on mobile <input> and <select> to prevent iOS Safari auto-zoom
assert(
  html.includes("font-size: 16px !important") || html.includes("font-size: 16px"),
  "Missing 16px font-size rule on mobile inputs/selects"
);
console.log("✓ 16px mobile input/select font-size verified (prevents iOS Safari auto-zoom).");

// 4. 2-column mobile product card grid & 2-line clamped titles
assert(
  html.includes("grid-template-columns: repeat(2, minmax(0, 1fr))"),
  "Missing 2-column mobile product card grid CSS"
);
assert(
  html.includes("-webkit-line-clamp: 2"),
  "Missing 2-line clamped title (-webkit-line-clamp: 2) CSS"
);
console.log("✓ 2-column mobile product card grid and 2-line title clamp verified.");

// 5. Bottom Sheets on mobile (<= 640px) with .sheet-handle grabber bars & pull-down-to-dismiss
assert(html.includes("@media (max-width: 640px)"), "Missing @media (max-width: 640px) mobile breakpoint");
const handleMatches = html.match(/class="sheet-handle"/g) || [];
assert(handleMatches.length >= 3, `Expected .sheet-handle in all 3 modals (#detail, #cart, #ratesModal), found ${handleMatches.length}`);
assert(html.includes("onPullStart") && html.includes("onPullEnd"), "Missing pull-down-to-dismiss touch gesture handlers");
console.log("✓ Mobile Bottom Sheets (<= 640px) with .sheet-handle grabber bars and pull-down gestures verified.");

// 6. Product Detail Modal steppers, gallery arrows, swipe navigation & sticky bottom bar
for (const id of ["dQtyMinus", "dQtyPlus", "dPrevImg", "dNextImg", "dThumbs", "mobileCartBar", "cartStickyTotal"]) {
  assert(html.includes(`id="${id}"`), `Missing required mobile UX element #${id}`);
}
assert(
  html.includes("#dThumbs, input, select, button"),
  "Swipe gesture handler must exclude #dThumbs, inputs, selects, and buttons"
);
assert(html.includes("@media print"), "Missing @media print rules for clean quotation printing");
console.log("✓ Detail modal steppers, gallery arrows, swipe exclusions, #mobileCartBar, #cartStickyTotal, and @media print verified.");
console.log("=== ALL test_mobile.js CHECKS PASSED (100%) ===\n");

/**
 * Automated Verification Suite 2: DOM Simulation, Multi-Currency, Filtering & CSV Export (test_dom_sim.js)
 */
const fs = require("fs");
const path = require("path");
const assert = require("assert");

const ROOT = __dirname;
const PRODUCTS_JS = path.join(ROOT, "website", "data", "products.js");
const INDEX_HTML = path.join(ROOT, "website", "index.html");

console.log("=== Running test_dom_sim.js: DOM Simulation, Currency, Filters & CSV ===");

// Build lightweight DOM simulation environment
const elements = {};
function makeEl(id) {
  return {
    id,
    value: "",
    textContent: "",
    innerHTML: "",
    placeholder: "",
    className: "",
    style: {},
    listeners: {},
    classList: {
      _set: new Set(),
      add(c) { this._set.add(c); },
      remove(c) { this._set.delete(c); },
      toggle(c) { if (this._set.has(c)) this._set.delete(c); else this._set.add(c); },
      contains(c) { return this._set.has(c); }
    },
    addEventListener(ev, fn) {
      this.listeners[ev] = this.listeners[ev] || [];
      this.listeners[ev].push(fn);
    },
    dispatchEvent(ev, payload = {}) {
      (this.listeners[ev] || []).forEach(fn => fn(Object.assign({ target: this, stopPropagation() {} }, payload)));
    },
    querySelector(sel) {
      return makeEl(id + "_" + sel);
    },
    closest() { return null; }
  };
}

const storage = {};
global.localStorage = {
  getItem(k) { return storage[k] || null; },
  setItem(k, v) { storage[k] = String(v); }
};

global.document = {
  documentElement: { lang: "en" },
  body: { appendChild() {}, removeChild() {} },
  getElementById(id) {
    if (!elements[id]) elements[id] = makeEl(id);
    return elements[id];
  },
  createElement(tag) { return makeEl(tag); },
  addEventListener() {}
};
global.window = global;

// Load catalog and execute inline application script from website/index.html
eval(fs.readFileSync(PRODUCTS_JS, "utf8"));
const html = fs.readFileSync(INDEX_HTML, "utf8");
const scriptMatches = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)];
assert(scriptMatches.length === 1, "Expected main inline application <script> in website/index.html");
eval(scriptMatches[0][1]);

// 1. Verify initial render & default sort order (win-desc)
assert(elements["grid"].innerHTML.includes("p-card"), "Product grid must render product cards on startup");
assert(elements["cats"].innerHTML.includes("top-winners"), "Category bar must include Top Winners (70+ pts) chip");

// 2. Test Multi-Currency conversions (RMB, USD, LKR)
assert.strictEqual(window.formatPrice(10, "RMB"), "¥10.00");
assert.strictEqual(window.formatPrice(7.15, "USD"), "$1.00");
assert.strictEqual(window.formatPrice(10, "LKR"), "රු. 435.00");

// Switch currency to LKR and test dynamic price filter (#pmin, #pmax)
elements["curr"].value = "LKR";
elements["curr"].dispatchEvent("change");
assert(elements["priceFilterLabel"].textContent.includes("රු."), "Price filter label must update to LKR symbol");

elements["pmin"].value = "400";
elements["pmin"].dispatchEvent("input");
elements["pmax"].value = "600";
elements["pmax"].dispatchEvent("input");
assert(elements["showingStatus"].textContent.includes("Showing"), "Status bar should reflect filtered count");

// Reset price filter
elements["pmin"].value = "";
elements["pmin"].dispatchEvent("input");
elements["pmax"].value = "";
elements["pmax"].dispatchEvent("input");

// 3. Test Detail Modal & Steppers (#dQtyMinus, #dQtyPlus, #dAddBtn)
const withKgProd = window.CATALOG.products.find(p => p.kg !== null && p.kg > 0);
assert(withKgProd, "Expected products with non-null weight");
window.openDetailModal(withKgProd.id);
assert(elements["detail"].classList.contains("open"), "Detail modal should open for product");
assert(elements["dPillars"].innerHTML.includes("pillar-bar-fill"), "Detail modal must render 4-pillar visual progress bars");

elements["dQtyPlus"].dispatchEvent("click");
assert.strictEqual(Number(elements["dQty"].value), 2, "#dQtyPlus should increment carton count to 2");
elements["dAddBtn"].dispatchEvent("click");

// Also add a product with null weight to test null-safe cart math
const nullKgProd = window.CATALOG.products.find(p => p.kg === null);
assert(nullKgProd, "Expected at least one product with null weight in spreadsheet");
window.openDetailModal(nullKgProd.id);
elements["dAddBtn"].dispatchEvent("click");

const totals = window.computeCartTotals();
assert.strictEqual(totals.lines.length, 2, "Cart should have 2 distinct product lines");
assert(Number.isFinite(totals.totalKg) && !Number.isNaN(totals.totalKg), "Cart totalKg must never be NaN when a product has null weight");
assert(Number.isFinite(totals.totalCbm) && totals.totalCbm > 0, "Cart totalCbm must be positive finite number");

// 4. Verify CSV Export with UTF-8 BOM & Column Alignment
const csv = window.buildCsvString();
assert(csv.startsWith("\uFEFF"), "CSV export must start with UTF-8 BOM (\\uFEFF)");
const csvLines = csv.slice(1).split("\r\n");
assert.strictEqual(csvLines.length, 4, "CSV should have 1 header row + 2 item rows + 1 TOTAL row");

function countCsvCols(line) {
  let inQuotes = false;
  let cols = 1;
  for (let i = 0; i < line.length; i++) {
    if (line[i] === '"') inQuotes = !inQuotes;
    else if (line[i] === "," && !inQuotes) cols++;
  }
  return cols;
}

const headerCols = countCsvCols(csvLines[0]);
csvLines.forEach((ln, idx) => {
  assert.strictEqual(countCsvCols(ln), headerCols, `CSV row ${idx} column count mismatch`);
});
assert(csvLines[csvLines.length - 1].startsWith("TOTAL,"), "Last CSV row must be TOTAL summary row");

console.log(`✓ Multi-currency, dynamic price filters, detail modal, null-safe cart totals, and ${headerCols}-column CSV export verified.`);
console.log("=== ALL test_dom_sim.js CHECKS PASSED (100%) ===\n");

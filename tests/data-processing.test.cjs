const test = require("node:test");
const assert = require("node:assert/strict");
const { createScriptContext } = require("./test-context.cjs");

function installDataset(context, evaluate, rows, fields = Object.keys(rows[0] || {})) {
  context.__testRows = rows;
  context.__testFields = fields;
  evaluate(`(() => {
    state.rows = __testRows;
    state.filteredRows = state.rows.slice();
    state.filterSourceRows = state.rows;
    state.fields = __testFields;
    state.profiles = state.fields.map((field) => buildColumnProfile(field, state.rows));
  })()`);
}

function filteredIds(evaluate, filters) {
  return JSON.parse(evaluate(
    `JSON.stringify(filterRowsByConfig(state.rows, ${JSON.stringify(filters)}).map((row) => row.id))`
  ));
}

function createFilterDataset() {
  return [
    { id: 1, category: "A", amount: "10", date: "2026-01-01" },
    { id: 2, category: "B", amount: "20", date: "2026-01-10" },
    { id: 3, category: "C", amount: "30", date: "2026-01-20" },
    { id: 4, category: "A", amount: "40", date: "2026-02-01" },
    { id: 5, category: "B", amount: "50", date: "2026-02-10" },
    { id: 6, category: "A", amount: "", date: "invalid-date" }
  ];
}

test("numeric parsing handles common tabular formats safely", () => {
  const { evaluate } = createScriptContext();

  assert.equal(evaluate('toNumber("1,234.50")'), 1234.5);
  assert.equal(evaluate('toNumber("$2,000")'), 2000);
  assert.equal(evaluate('toNumber("(1,250)")'), -1250);
  assert.equal(evaluate('toNumber("25%")'), 0.25);
  assert.equal(evaluate('toNumber("-3.5")'), -3.5);
  assert.equal(evaluate('toNumber(42)'), 42);
  assert.equal(evaluate('toNumber("")'), null);
  assert.equal(evaluate('toNumber("not-a-number")'), null);
  assert.equal(evaluate('toNumber("2026-07-18")'), null);
});

test("field inference distinguishes numeric, categorical, date and ID fields", () => {
  const { context, evaluate } = createScriptContext();
  const rows = Array.from({ length: 12 }, (_, index) => ({
    order_id: `ORD-${String(index + 1).padStart(3, "0")}`,
    amount: String(100 + index * 10),
    order_date: `2026-07-${String(index + 1).padStart(2, "0")}`,
    region: index % 2 === 0 ? "East" : "West"
  }));
  context.__testRows = rows;

  assert.equal(evaluate('buildColumnProfile("order_id", __testRows).typeKey'), "id");
  assert.equal(evaluate('buildColumnProfile("amount", __testRows).typeKey'), "numeric");
  assert.equal(evaluate('buildColumnProfile("order_date", __testRows).typeKey'), "date");
  assert.equal(evaluate('buildColumnProfile("region", __testRows).typeKey'), "category");
});

test("missing value statistics count empty, whitespace, null and undefined", () => {
  const { context, evaluate } = createScriptContext();
  context.__testRows = [
    { value: "10" },
    { value: "" },
    { value: "   " },
    { value: null },
    { value: undefined },
    { value: "20" }
  ];

  const profile = JSON.parse(evaluate('JSON.stringify(buildColumnProfile("value", __testRows))'));
  assert.equal(profile.nonMissingCount, 2);
  assert.equal(profile.missingCount, 4);
  assert.equal(profile.missingRate, 4 / 6);
});

test("duplicate row detection compares every field", () => {
  const { context, evaluate } = createScriptContext();
  const rows = [
    { name: "Alice", score: "90" },
    { name: "Alice", score: "90" },
    { name: " Alice ", score: "90" },
    { name: "Alice", score: "91" },
    { name: "Bob", score: "90" }
  ];
  installDataset(context, evaluate, rows, ["name", "score"]);

  assert.equal(evaluate('countDuplicateRows(state.rows, state.fields)'), 2);
});

test("mean, median and population standard deviation are correct", () => {
  const { context, evaluate } = createScriptContext();
  context.__values = [1, 2, 3, 4, 5];

  assert.equal(evaluate("mean(__values)"), 3);
  assert.equal(evaluate("median(__values)"), 3);
  assert.ok(Math.abs(evaluate("standardDeviation(__values)") - Math.sqrt(2)) < 1e-12);
  assert.equal(evaluate("median([1, 2, 3, 4])"), 2.5);
});

test("IQR outlier detection identifies values outside the fences", () => {
  const { context, evaluate } = createScriptContext();
  const rows = [1, 2, 3, 4, 5, 6, 7, 100].map((value) => ({ value: String(value) }));
  installDataset(context, evaluate, rows, ["value"]);

  const stats = JSON.parse(evaluate('JSON.stringify(buildNumericStats(state.rows)[0])'));
  assert.equal(stats.outlierCount, 1);
  assert.equal(stats.max, 100);
  assert.ok(stats.upperBound < 100);
});

test("categorical filtering supports one or multiple selected values", () => {
  const { context, evaluate } = createScriptContext();
  installDataset(context, evaluate, createFilterDataset());

  assert.deepEqual(filteredIds(evaluate, {
    category: { field: "category", values: ["A", "C"] },
    numeric: null,
    date: null
  }), [1, 3, 4, 6]);
});

test("numeric range filtering is inclusive and excludes missing numbers", () => {
  const { context, evaluate } = createScriptContext();
  installDataset(context, evaluate, createFilterDataset());

  assert.deepEqual(filteredIds(evaluate, {
    category: null,
    numeric: { field: "amount", min: 20, max: 40 },
    date: null
  }), [2, 3, 4]);
});

test("date range filtering is inclusive and excludes invalid dates", () => {
  const { context, evaluate } = createScriptContext();
  installDataset(context, evaluate, createFilterDataset());

  assert.deepEqual(filteredIds(evaluate, {
    category: null,
    numeric: null,
    date: { field: "date", start: "2026-01-10", end: "2026-02-01" }
  }), [2, 3, 4]);
});

test("categorical, numeric and date filters combine with AND semantics", () => {
  const { context, evaluate } = createScriptContext();
  installDataset(context, evaluate, createFilterDataset());

  assert.deepEqual(filteredIds(evaluate, {
    category: { field: "category", values: ["A", "B"] },
    numeric: { field: "amount", min: 15, max: 50 },
    date: { field: "date", start: "2026-01-05", end: "2026-02-05" }
  }), [2, 4]);
});

test("clearing filters restores all rows without mutating the source", () => {
  const { context, evaluate } = createScriptContext();
  const rows = createFilterDataset();
  const originalSnapshot = JSON.stringify(rows);
  installDataset(context, evaluate, rows);
  context.__filters = {
    category: { field: "category", values: ["A"] },
    numeric: { field: "amount", min: 10, max: 40 },
    date: null
  };

  evaluate(`(() => {
    state.filters = __filters;
    state.filteredRows = filterRowsByConfig(state.rows, state.filters);
  })()`);
  assert.equal(evaluate("getAnalysisRows().length"), 2);

  evaluate("resetFilterState()");
  assert.equal(evaluate("getAnalysisRows().length"), rows.length);
  assert.equal(evaluate("state.filteredRows === state.rows"), false);
  assert.equal(JSON.stringify(rows), originalSnapshot);
});

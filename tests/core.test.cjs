const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

class FakeElement {
  constructor() {
    this.value = "";
    this.innerHTML = "";
    this.textContent = "";
    this.dataset = {};
    this.className = "";
    this.options = [];
    this.selectedIndex = 0;
    this.classList = { add() {}, remove() {}, toggle() {} };
  }

  addEventListener() {}
  removeAttribute() {}
  setAttribute() {}
  querySelectorAll() { return []; }
  getContext() {
    return { clearRect() {}, save() {}, restore() {}, fillText() {} };
  }
}

const elements = new Map();
const document = {
  getElementById(id) {
    if (!elements.has(id)) elements.set(id, new FakeElement());
    return elements.get(id);
  },
  querySelectorAll() { return []; },
  createElement() { return new FakeElement(); },
  body: new FakeElement()
};

const context = vm.createContext({
  console,
  document,
  window: { Chart: null, print() {} },
  TextDecoder,
  Uint8Array,
  Intl,
  Blob,
  URL,
  setTimeout,
  clearTimeout
});

const source = fs.readFileSync(path.join(__dirname, "..", "script.js"), "utf8");
vm.runInContext(source, context, { filename: "script.js" });

function evaluate(expression) {
  return vm.runInContext(expression, context);
}

assert.equal(evaluate('toNumber("10%")'), 0.1);
assert.equal(evaluate('toNumber("(10%)")'), -0.1);
assert.equal(evaluate('toNumber("￥1,234.50")'), 1234.5);
assert.equal(evaluate('toNumber("(1,250)")'), -1250);
assert.equal(evaluate('toNumber("2026-07-17")'), null);

assert.equal(evaluate('toDate("2026-07-17").toISOString()'), "2026-07-17T00:00:00.000Z");
assert.equal(evaluate('toDate("2026-02-30")'), null);
assert.equal(evaluate('toDate("07/08/2026")'), null);
assert.equal(evaluate('toDate("17/07/2026").toISOString()'), "2026-07-17T00:00:00.000Z");

const excelData = evaluate(`JSON.stringify(buildExcelTabularData(
  [["", "", ""], ["订单编号", "地区", "销售额"], ["A-001", "华东", "1,200"], ["A-002", "华南", "980"]],
  [["", "", ""], ["订单编号", "地区", "销售额"], ["A-001", "华东", 1200], ["A-002", "华南", 980]],
  "销售数据"
))`);
assert.deepEqual(JSON.parse(excelData), {
  fields: ["订单编号", "地区", "销售额"],
  rows: [
    { 订单编号: "A-001", 地区: "华东", 销售额: "1,200" },
    { 订单编号: "A-002", 地区: "华南", 销售额: "980" }
  ]
});

function excelError(formattedRows, rawRows, sheetName = "测试表") {
  return evaluate(`(() => {
    try {
      buildExcelTabularData(${JSON.stringify(formattedRows)}, ${JSON.stringify(rawRows)}, ${JSON.stringify(sheetName)});
      return "";
    } catch (error) {
      return error.message;
    }
  })()`);
}

assert.match(excelError([], []), /为空/);
assert.match(excelError([["字段A", "字段B"]], [["字段A", "字段B"]]), /只有表头/);
assert.match(excelError([["张三", "20"]], [["张三", 20]]), /未识别到有效文本表头/);
assert.match(excelError([["字段A", ""], ["值A", "值B"]], [["字段A", ""], ["值A", "值B"]]), /缺少表头/);
assert.match(excelError([["字段A", "字段A"], ["值A", "值B"]], [["字段A", "字段A"], ["值A", "值B"]]), /重复表头/);

assert.doesNotThrow(() => evaluate('validateExcelFileSignature(new Uint8Array([0x50,0x4b,0x03,0x04]), "valid.xlsx")'));
assert.doesNotThrow(() => evaluate('validateExcelFileSignature(new Uint8Array([0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1]), "valid.xls")'));
assert.match(
  evaluate(`(() => {
    try {
      validateExcelFileSignature(new Uint8Array([0x61,0x62,0x63]), "fake.xlsx");
      return "";
    } catch (error) {
      return error.message;
    }
  })()`),
  /格式与扩展名不匹配/
);

assert.equal(
  evaluate('buildColumnProfile("用户ID", [{"用户ID":"u1"},{"用户ID":"u1"},{"用户ID":"u2"}]).typeKey'),
  "id"
);

const firstImport = evaluate("beginImport()");
const secondImport = evaluate("beginImport()");
assert.equal(evaluate(`isCurrentImport(${firstImport})`), false);
assert.equal(evaluate(`isCurrentImport(${secondImport})`), true);

evaluate(`handleParsedData({
  errors: [{type: "Delimiter", code: "UndetectableDelimiter", message: "bad delimiter"}],
  meta: {fields: ["字段"]},
  data: [{"字段": "1"}]
}, "broken.csv", ${secondImport})`);
assert.equal(evaluate("state.rows.length"), 0);
assert.equal(elements.get("statusPanel").className, "status-panel error");

const warningImport = evaluate("beginImport()");
evaluate(`handleParsedData({
  errors: [{type: "FieldMismatch", code: "TooFewFields", message: "row mismatch", row: 0}],
  meta: {fields: ["用户ID", "数值"]},
  data: [{"用户ID": "u1", "数值": "10"}, {"用户ID": "u2", "数值": "20"}]
}, "warning.csv", ${warningImport})`);
assert.equal(evaluate("state.rows.length"), 2);
assert.equal(evaluate("state.parseWarnings.length"), 1);
assert.equal(elements.get("statusPanel").className, "status-panel warning");

evaluate('applyFieldTypeOverride("数值", "category")');
assert.equal(evaluate('state.profiles.find((item) => item.field === "数值").typeKey'), "category");
assert.equal(evaluate("state.numericStats.length"), 0);

console.log("core tests passed");

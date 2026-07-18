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
    this.disabled = false;
    this.classList = { add() {}, remove() {}, toggle() {} };
  }

  addEventListener() {}
  focus() {}
  removeAttribute() {}
  scrollIntoView() {}
  setAttribute() {}
  querySelectorAll() { return []; }
  getContext() {
    return { clearRect() {}, save() {}, restore() {}, fillText() {} };
  }
}

function createScriptContext() {
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
    window: {
      Chart: null,
      addEventListener() {},
      innerHeight: 900,
      print() {},
      requestAnimationFrame(callback) { callback(); return 1; },
      scrollY: 0
    },
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

  return {
    context,
    elements,
    source,
    evaluate(expression) {
      return vm.runInContext(expression, context);
    }
  };
}

module.exports = { createScriptContext };

const state = {
  rows: [],
  fields: [],
  profiles: [],
  numericStats: [],
  categoryStats: [],
  duplicateRows: 0,
  totalMissing: 0,
  parseWarnings: [],
  typeOverrides: {},
  fieldTypeDrafts: {},
  pendingExcel: null,
  activeImportId: 0,
  charts: {}
};

const dom = {
  fileInput: document.getElementById("csvFile"),
  fileEncoding: document.getElementById("fileEncoding"),
  loadSample: document.getElementById("loadSample"),
  dropZone: document.getElementById("dropZone"),
  statusPanel: document.getElementById("statusPanel"),
  sheetSelectionPanel: document.getElementById("sheetSelectionPanel"),
  excelWorkbookName: document.getElementById("excelWorkbookName"),
  excelSheetSelect: document.getElementById("excelSheetSelect"),
  confirmSheetSelection: document.getElementById("confirmSheetSelection"),
  sheetSelectionMessage: document.getElementById("sheetSelectionMessage"),
  results: document.getElementById("results"),
  sourceName: document.getElementById("sourceName"),
  overviewCards: document.getElementById("overviewCards"),
  insights: document.getElementById("insights"),
  previewTable: document.getElementById("previewTable"),
  schemaTable: document.getElementById("schemaTable"),
  fieldConfigMessage: document.getElementById("fieldConfigMessage"),
  applyFieldConfig: document.getElementById("applyFieldConfig"),
  restoreAutoFieldTypes: document.getElementById("restoreAutoFieldTypes"),
  qualitySummary: document.getElementById("qualitySummary"),
  qualityTable: document.getElementById("qualityTable"),
  outlierTable: document.getElementById("outlierTable"),
  numericStatsTable: document.getElementById("numericStatsTable"),
  categoryStats: document.getElementById("categoryStats"),
  numericChartField: document.getElementById("numericChartField"),
  categoryChartField: document.getElementById("categoryChartField"),
  dateChartField: document.getElementById("dateChartField"),
  metricField: document.getElementById("metricField"),
  groupField: document.getElementById("groupField"),
  aggregateMethod: document.getElementById("aggregateMethod"),
  v2MetricField: document.getElementById("v2MetricField"),
  v2GroupField: document.getElementById("v2GroupField"),
  v2DateField: document.getElementById("v2DateField"),
  v2AnalyzeButton: document.getElementById("v2AnalyzeButton"),
  v2FieldPrompt: document.getElementById("v2FieldPrompt"),
  v2AnalysisResults: document.getElementById("v2AnalysisResults"),
  v2SummaryCards: document.getElementById("v2SummaryCards"),
  v2SumTable: document.getElementById("v2SumTable"),
  v2AvgTable: document.getElementById("v2AvgTable"),
  v2TrendCard: document.getElementById("v2TrendCard"),
  scenarioTemplate: document.getElementById("scenarioTemplate"),
  templateMappingForm: document.getElementById("templateMappingForm"),
  runTemplateAnalysis: document.getElementById("runTemplateAnalysis"),
  templatePrompt: document.getElementById("templatePrompt"),
  templateResults: document.getElementById("templateResults"),
  templateSummaryCards: document.getElementById("templateSummaryCards"),
  templateInsights: document.getElementById("templateInsights"),
  templatePrimaryTable: document.getElementById("templatePrimaryTable"),
  templateMainChartTitle: document.getElementById("templateMainChartTitle"),
  templateSecondaryChartTitle: document.getElementById("templateSecondaryChartTitle"),
  templateTrendCard: document.getElementById("templateTrendCard"),
  templateTrendChartTitle: document.getElementById("templateTrendChartTitle"),
  exportHtmlReport: document.getElementById("exportHtmlReport"),
  printReport: document.getElementById("printReport")
};

const ENCODING_LABELS = {
  auto: "自动识别",
  "utf-8": "UTF-8",
  gbk: "GBK",
  gb18030: "GB18030"
};

const FIELD_TYPE_LABELS = {
  numeric: "数值",
  category: "分类",
  date: "日期",
  id: "ID",
  ignore: "忽略"
};

const DATE_TIME_SUFFIX_PATTERN = "(?:[ T](?:[01]?\\d|2[0-3]):[0-5]\\d(?::[0-5]\\d(?:\\.\\d{1,3})?)?(?:Z|[+-]\\d{2}:?\\d{2})?)?";

const CHART_THEME = {
  primary: "#625bde",
  primaryFill: "rgba(98, 91, 222, 0.12)",
  secondary: "#0f9f8f",
  secondaryFill: "rgba(15, 159, 143, 0.12)",
  warning: "#e09a32",
  text: "#697084",
  grid: "rgba(211, 215, 227, 0.72)"
};

if (window.Chart) {
  Chart.defaults.color = CHART_THEME.text;
  Chart.defaults.font.family = 'Inter, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif';
  Chart.defaults.font.size = 11;
  Chart.defaults.elements.bar.borderRadius = 6;
  Chart.defaults.elements.bar.borderSkipped = false;
  Chart.defaults.elements.line.borderWidth = 2.25;
  Chart.defaults.elements.point.radius = 0;
  Chart.defaults.elements.point.hoverRadius = 4;
}

const TEMPLATE_DEFINITIONS = {
  generic: {
    name: "通用数据分析",
    fields: [
      { key: "metric", label: "分析指标字段", required: true, expected: "numeric" },
      { key: "group", label: "分组维度字段", required: true, expected: "category" },
      { key: "date", label: "时间字段（可选）", required: false, expected: "date" }
    ]
  },
  sales: {
    name: "销售数据分析",
    fields: [
      { key: "salesAmount", label: "销售额字段（可选）", required: false, expected: "numeric" },
      { key: "unitPrice", label: "单价字段（可选）", required: false, expected: "numeric" },
      { key: "quantity", label: "数量字段（可选）", required: false, expected: "numeric" },
      { key: "category", label: "品类字段", required: true, expected: "category" },
      { key: "region", label: "地区字段", required: true, expected: "category" },
      { key: "date", label: "日期字段（可选）", required: false, expected: "date" }
    ]
  },
  score: {
    name: "成绩数据分析",
    fields: [
      { key: "score", label: "成绩字段", required: true, expected: "numeric" },
      { key: "class", label: "班级字段（可选）", required: false, expected: "category" },
      { key: "subject", label: "科目字段（可选）", required: false, expected: "category" },
      { key: "examDate", label: "考试日期字段（可选）", required: false, expected: "date" }
    ]
  },
  usedGoods: {
    name: "二手商品价格分析",
    fields: [
      { key: "price", label: "价格字段", required: true, expected: "numeric" },
      { key: "item", label: "商品字段（可选）", required: false, expected: "category" },
      { key: "platform", label: "平台字段（可选）", required: false, expected: "category" },
      { key: "city", label: "城市字段（可选）", required: false, expected: "category" },
      { key: "condition", label: "成色字段（可选）", required: false, expected: "category" },
      { key: "date", label: "日期字段（可选）", required: false, expected: "date" }
    ]
  },
  survey: {
    name: "问卷调查分析",
    fields: [
      { key: "satisfaction", label: "满意度字段（可选）", required: false, expected: "numeric" },
      { key: "age", label: "年龄字段（可选）", required: false, expected: "numeric" },
      { key: "gender", label: "性别字段（可选）", required: false, expected: "category" },
      { key: "city", label: "城市字段（可选）", required: false, expected: "category" },
      { key: "choice", label: "选择题字段（可选）", required: false, expected: "category" }
    ]
  },
  behavior: {
    name: "用户行为分析",
    fields: [
      { key: "userId", label: "用户 ID 字段", required: true, expected: "any" },
      { key: "event", label: "事件字段", required: true, expected: "category" },
      { key: "channel", label: "渠道字段（可选）", required: false, expected: "category" },
      { key: "device", label: "设备字段（可选）", required: false, expected: "category" },
      { key: "date", label: "日期字段（可选）", required: false, expected: "date" }
    ]
  }
};

setupResultNavigation();

dom.fileInput.addEventListener("change", (event) => {
  const file = event.target.files[0];
  if (file) parseDataFile(file);
  event.target.value = "";
});

dom.loadSample.addEventListener("click", async () => {
  const importId = beginImport();
  try {
    setStatus("正在加载 sample-data.csv ...");
    const response = await fetch("sample-data.csv");
    if (!response.ok) throw new Error("示例数据加载失败");
    const text = await response.text();
    if (!isCurrentImport(importId)) return;
    parseCsvText(text, "sample-data.csv", importId);
  } catch (error) {
    showError(error.message, importId);
  }
});

["dragenter", "dragover"].forEach((eventName) => {
  dom.dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dom.dropZone.classList.add("dragging");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  dom.dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dom.dropZone.classList.remove("dragging");
  });
});

dom.dropZone.addEventListener("drop", (event) => {
  const file = event.dataTransfer.files[0];
  if (!file) return;
  parseDataFile(file);
});

[
  dom.numericChartField,
  dom.categoryChartField,
  dom.dateChartField,
  dom.metricField,
  dom.groupField,
  dom.aggregateMethod
].forEach((control) => control.addEventListener("change", renderCharts));

dom.v2AnalyzeButton.addEventListener("click", renderV2Analysis);

dom.scenarioTemplate.addEventListener("change", () => {
  renderTemplateMappingForm();
  resetTemplateResults(`已切换到「${getCurrentTemplate().name}」，请选择字段并生成分析。`);
});

dom.runTemplateAnalysis.addEventListener("click", runTemplateAnalysis);
dom.exportHtmlReport.addEventListener("click", exportHtmlReport);
dom.printReport.addEventListener("click", () => {
  if (ensureFieldConfigurationApplied("打印或导出 PDF")) window.print();
});
dom.confirmSheetSelection.addEventListener("click", analyzeSelectedExcelSheet);
dom.excelSheetSelect.addEventListener("change", () => {
  dom.sheetSelectionMessage.className = "sheet-selection-message";
  dom.sheetSelectionMessage.textContent = "";
});

dom.schemaTable.addEventListener("change", (event) => {
  const select = event.target.closest("select[data-field-type]");
  if (!select) return;
  stageFieldTypeChange(select.dataset.fieldType, select.value);
});
dom.applyFieldConfig.addEventListener("click", applyFieldConfiguration);
dom.restoreAutoFieldTypes.addEventListener("click", restoreAutomaticFieldTypes);

[dom.v2MetricField, dom.v2GroupField, dom.v2DateField].forEach((control) => {
  control.addEventListener("change", () => {
    if (state.rows.length) {
      dom.v2FieldPrompt.className = "analysis-prompt";
      dom.v2FieldPrompt.textContent = "字段已更新，点击“重新选择字段并分析”生成新的结果。";
    }
  });
});

function setupResultNavigation() {
  const links = Array.from(document.querySelectorAll(".results-nav nav a"));
  const sections = links
    .map((link) => document.querySelector(link.getAttribute("href")))
    .filter(Boolean);

  const setActiveLink = (sectionId) => {
    links.forEach((link) => {
      link.classList.toggle("active", link.getAttribute("href") === `#${sectionId}`);
    });
  };

  links.forEach((link) => {
    link.addEventListener("click", () => {
      const sectionId = link.getAttribute("href").slice(1);
      setActiveLink(sectionId);
    });
  });

  if ("IntersectionObserver" in window) {
    const observer = new IntersectionObserver((entries) => {
      const visibleSection = entries
        .filter((entry) => entry.isIntersecting)
        .sort((a, b) => b.intersectionRatio - a.intersectionRatio)[0];
      if (visibleSection) setActiveLink(visibleSection.target.id);
    }, {
      rootMargin: "-15% 0px -68% 0px",
      threshold: [0, 0.1, 0.35]
    });
    sections.forEach((section) => observer.observe(section));
  }
}

async function parseCsvFile(file) {
  const importId = beginImport();
  try {
    assertRuntimeDependency("Papa", "CSV 解析组件加载失败，请检查网络连接后刷新页面重试。");
    const selectedEncoding = dom.fileEncoding.value || "auto";
    setStatus(`正在读取 ${file.name} ...`);
    const buffer = await readFileAsArrayBuffer(file);
    if (!isCurrentImport(importId)) return;
    const decoded = decodeCsvBuffer(buffer, selectedEncoding);
    setStatus(`正在解析 ${file.name}（${decoded.label}）...`);
    parseCsvText(decoded.text, file.name, importId);
  } catch (error) {
    showError(error.message, importId);
  }
}

function parseDataFile(file) {
  if (/\.csv$/i.test(file.name)) return parseCsvFile(file);
  if (/\.(xlsx|xls)$/i.test(file.name)) return parseExcelFile(file);
  const importId = beginImport();
  showError("不支持的文件格式。请上传 CSV、XLSX 或 XLS 文件。", importId);
}

async function parseExcelFile(file) {
  const importId = beginImport();
  try {
    assertRuntimeDependency("XLSX", "Excel 解析组件加载失败，请检查网络连接后刷新页面重试。");
    setStatus(`正在读取 ${file.name} ...`);
    const buffer = await readFileAsArrayBuffer(file);
    if (!isCurrentImport(importId)) return;
    setStatus(`正在解析 ${file.name} ...`);
    validateExcelFileSignature(buffer, file.name);
    let workbook;
    try {
      workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    } catch {
      throw new Error("Excel 文件格式异常或文件已损坏，无法解析。请确认文件是有效的 XLSX 或 XLS 文件。");
    }

    const sheetNames = Array.isArray(workbook.SheetNames) ? workbook.SheetNames.filter(Boolean) : [];
    if (!sheetNames.length) throw new Error("Excel 文件中没有可读取的工作表。");

    if (sheetNames.length > 1) {
      showExcelSheetSelection(workbook, file.name, sheetNames, importId);
      return;
    }

    importExcelSheet(workbook, file.name, sheetNames[0], importId);
  } catch (error) {
    showError(error.message, importId);
  }
}

function showExcelSheetSelection(workbook, fileName, sheetNames, importId) {
  if (!isCurrentImport(importId)) return;
  state.pendingExcel = { workbook, fileName, importId };
  dom.excelWorkbookName.textContent = `${fileName} · 共 ${sheetNames.length} 个工作表`;
  dom.excelSheetSelect.innerHTML = "";
  sheetNames.forEach((sheetName) => {
    const option = document.createElement("option");
    option.value = sheetName;
    option.textContent = sheetName;
    dom.excelSheetSelect.appendChild(option);
  });
  dom.sheetSelectionMessage.className = "sheet-selection-message";
  dom.sheetSelectionMessage.textContent = "请选择一个工作表，确认后再生成分析结果。";
  dom.sheetSelectionPanel.classList.remove("hidden");
  dom.statusPanel.className = "status-panel hidden";
  dom.statusPanel.textContent = "";
}

function analyzeSelectedExcelSheet() {
  const pending = state.pendingExcel;
  if (!pending || !isCurrentImport(pending.importId)) {
    showError("Excel 工作簿状态已失效，请重新上传文件。");
    return;
  }

  const sheetName = dom.excelSheetSelect.value;
  if (!sheetName) {
    showSheetSelectionError("请选择一个工作表后再继续。");
    return;
  }

  try {
    importExcelSheet(pending.workbook, pending.fileName, sheetName, pending.importId);
  } catch (error) {
    showSheetSelectionError(error.message);
  }
}

function importExcelSheet(workbook, fileName, sheetName, importId) {
  if (!isCurrentImport(importId)) return;
  const worksheet = workbook.Sheets?.[sheetName];
  if (!worksheet || !worksheet["!ref"]) {
    throw new Error(`工作表「${sheetName}」为空，请选择包含表头和数据的工作表。`);
  }

  const formattedRows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw: false,
    blankrows: true,
    dateNF: "yyyy-mm-dd"
  });
  const rawRows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw: true,
    blankrows: true
  });
  const { fields, rows } = buildExcelTabularData(formattedRows, rawRows, sheetName);

  state.pendingExcel = null;
  dom.sheetSelectionPanel.classList.add("hidden");
  setStatus(`正在分析 ${fileName} · ${sheetName} ...`);
  commitTabularData(fields, rows, `${fileName} · ${sheetName}`, importId);
}

function buildExcelTabularData(formattedRows, rawRows, sheetName = "当前工作表") {
  const rowCount = Math.max(formattedRows?.length || 0, rawRows?.length || 0);
  const rowPairs = Array.from({ length: rowCount }, (_, index) => ({
    formatted: Array.isArray(formattedRows?.[index]) ? formattedRows[index] : [],
    raw: Array.isArray(rawRows?.[index]) ? rawRows[index] : []
  }));

  while (rowPairs.length && !excelRowHasValue(rowPairs[0].formatted) && !excelRowHasValue(rowPairs[0].raw)) {
    rowPairs.shift();
  }
  while (rowPairs.length && !excelRowHasValue(rowPairs[rowPairs.length - 1].formatted) && !excelRowHasValue(rowPairs[rowPairs.length - 1].raw)) {
    rowPairs.pop();
  }

  if (!rowPairs.length) {
    throw new Error(`工作表「${sheetName}」为空，请选择包含表头和数据的工作表。`);
  }

  const headerPair = rowPairs[0];
  const dataPairs = rowPairs.slice(1);
  const maxColumns = rowPairs.reduce((max, pair) => Math.max(max, pair.formatted.length, pair.raw.length), 0);
  const usedColumnIndexes = Array.from({ length: maxColumns }, (_, index) => index).filter((columnIndex) =>
    !isMissing(headerPair.formatted[columnIndex])
      || dataPairs.some((pair) => !isMissing(pair.formatted[columnIndex]) || !isMissing(pair.raw[columnIndex]))
  );

  if (!usedColumnIndexes.length) {
    throw new Error(`工作表「${sheetName}」为空，请选择包含表头和数据的工作表。`);
  }

  const fields = usedColumnIndexes.map((columnIndex) => {
    const formattedHeader = headerPair.formatted[columnIndex];
    const rawHeader = headerPair.raw[columnIndex];
    if (isMissing(formattedHeader)) {
      throw new Error(`工作表「${sheetName}」的第 ${columnIndex + 1} 列缺少表头，请补充字段名后重新上传。`);
    }
    if (typeof rawHeader !== "string") {
      throw new Error(`工作表「${sheetName}」未识别到有效文本表头；第 ${columnIndex + 1} 列首行不是字段名。`);
    }
    return normalizeFieldName(formattedHeader);
  });

  const duplicateFields = fields.filter((field, index) => fields.indexOf(field) !== index);
  if (duplicateFields.length) {
    throw new Error(`工作表「${sheetName}」包含重复表头「${duplicateFields[0]}」，请修改后重新上传。`);
  }

  const rows = dataPairs
    .filter((pair) => usedColumnIndexes.some((columnIndex) => !isMissing(pair.formatted[columnIndex]) || !isMissing(pair.raw[columnIndex])))
    .map((pair) => fields.reduce((row, field, fieldIndex) => {
      const value = pair.formatted[usedColumnIndexes[fieldIndex]];
      row[field] = value === null || value === undefined ? "" : String(value);
      return row;
    }, {}));

  if (!rows.length) {
    throw new Error(`工作表「${sheetName}」只有表头，没有可分析的数据行。`);
  }

  return { fields, rows };
}

function excelRowHasValue(row) {
  return Array.isArray(row) && row.some((value) => !isMissing(value));
}

function validateExcelFileSignature(buffer, fileName) {
  const bytes = new Uint8Array(buffer);
  const isXlsx = /\.xlsx$/i.test(fileName);
  const isXls = /\.xls$/i.test(fileName);
  const hasZipSignature = bytes.length >= 4
    && bytes[0] === 0x50
    && bytes[1] === 0x4b
    && [[0x03, 0x04], [0x05, 0x06], [0x07, 0x08]]
      .some(([third, fourth]) => bytes[2] === third && bytes[3] === fourth);
  const oleSignature = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  const hasOleSignature = bytes.length >= oleSignature.length
    && oleSignature.every((value, index) => bytes[index] === value);

  if ((isXlsx && !hasZipSignature) || (isXls && !hasOleSignature)) {
    throw new Error("Excel 文件格式与扩展名不匹配或文件已损坏。请重新导出为有效的 XLSX 或 XLS 文件后再上传。");
  }
}

function showSheetSelectionError(message) {
  dom.sheetSelectionMessage.className = "sheet-selection-message error";
  dom.sheetSelectionMessage.textContent = message;
}

function parseCsvText(text, sourceName, importId = beginImport()) {
  if (!isCurrentImport(importId)) return;
  assertRuntimeDependency("Papa", "CSV 解析组件加载失败，请检查网络连接后刷新页面重试。");
  Papa.parse(text, {
    header: true,
    skipEmptyLines: true,
    complete: (results) => handleParsedData(results, sourceName, importId),
    error: (error) => showError(error.message, importId)
  });
}

function beginImport() {
  state.activeImportId += 1;
  resetAnalysisState();
  return state.activeImportId;
}

function isCurrentImport(importId) {
  return importId === state.activeImportId;
}

function assertRuntimeDependency(globalName, message) {
  if (!window[globalName]) throw new Error(message);
}

function readFileAsArrayBuffer(file) {
  if (typeof file.arrayBuffer === "function") {
    return file.arrayBuffer();
  }

  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error || new Error("文件读取失败"));
    reader.readAsArrayBuffer(file);
  });
}

function decodeCsvBuffer(buffer, selectedEncoding) {
  if (selectedEncoding === "auto") {
    return autoDecodeCsvBuffer(buffer);
  }

  if (!Object.prototype.hasOwnProperty.call(ENCODING_LABELS, selectedEncoding)) {
    throw new Error("不支持的文件编码。");
  }

  return {
    text: decodeText(buffer, selectedEncoding),
    label: ENCODING_LABELS[selectedEncoding]
  };
}

function autoDecodeCsvBuffer(buffer) {
  if (hasUtf8Bom(buffer)) {
    return {
      text: decodeText(buffer, "utf-8"),
      label: "UTF-8"
    };
  }

  try {
    return {
      text: decodeText(buffer, "utf-8", true),
      label: "UTF-8"
    };
  } catch {
    return {
      text: decodeText(buffer, "gb18030"),
      label: "GB18030 自动识别"
    };
  }
}

function hasUtf8Bom(buffer) {
  const bytes = new Uint8Array(buffer);
  return bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
}

function decodeText(buffer, encoding, fatal = false) {
  try {
    return new TextDecoder(encoding, { fatal }).decode(buffer);
  } catch (error) {
    if (fatal) throw error;
    throw new Error(`当前浏览器不支持 ${ENCODING_LABELS[encoding] || encoding} 解码，请尝试选择其他编码或将文件另存为 CSV UTF-8。`);
  }
}

function handleParsedData(results, sourceName, importId) {
  if (!isCurrentImport(importId)) return;
  const parseErrors = Array.isArray(results.errors) ? results.errors : [];
  const fatalError = parseErrors.find((error) => error.type === "Delimiter" || error.code === "MissingQuotes");
  if (fatalError) {
    showError(`CSV 解析失败：${fatalError.message}`, importId);
    return;
  }

  const parseWarnings = parseErrors.slice(0, 20).map((error) => ({
    row: Number.isInteger(error.row) ? error.row + 2 : null,
    code: error.code || "ParseWarning",
    message: error.message || "CSV 行结构异常"
  }));
  const fieldPairs = (results.meta.fields || [])
    .map((rawField) => ({
      raw: rawField,
      clean: normalizeFieldName(rawField)
    }))
    .filter((pair) => pair.clean.trim() !== "");
  const fields = fieldPairs.map((pair) => pair.clean);
  const normalizedRows = (results.data || []).map((row) =>
    normalizeParsedRow(row, fieldPairs)
  );
  const rows = normalizedRows.filter((row) =>
    fields.some((field) => !isMissing(row[field]))
  );

  if (!fields.length || !rows.length) {
    showError("CSV 中没有可分析的数据。请确认文件包含表头和至少一行数据。", importId);
    return;
  }

  commitTabularData(fields, rows, sourceName, importId, parseWarnings);
}

function commitTabularData(fields, rows, sourceName, importId, parseWarnings = []) {
  if (!isCurrentImport(importId)) return;
  if (!fields.length || !rows.length) {
    showError("文件中没有可分析的数据。请确认数据包含表头和至少一行数据。", importId);
    return;
  }

  state.rows = rows;
  state.fields = fields;
  state.parseWarnings = parseWarnings;
  state.pendingExcel = null;
  state.profiles = fields.map((field) => buildColumnProfile(field, rows));
  state.typeOverrides = {};
  state.fieldTypeDrafts = Object.fromEntries(
    state.profiles.map((profile) => [profile.field, profile.inferredTypeKey])
  );
  state.duplicateRows = countDuplicateRows(rows, fields);
  state.totalMissing = state.profiles.reduce((sum, profile) => sum + profile.missingCount, 0);
  state.numericStats = buildNumericStats();
  state.categoryStats = buildCategoryStats();

  dom.sourceName.textContent = `已载入 · ${sourceName}`;
  dom.sourceName.title = sourceName;
  dom.sheetSelectionPanel.classList.add("hidden");
  dom.results.classList.remove("hidden");
  if (state.parseWarnings.length) {
    const affectedRows = new Set(state.parseWarnings.map((warning) => warning.row).filter(Boolean)).size;
    setStatus(`CSV 已载入，但发现 ${state.parseWarnings.length} 个解析警告${affectedRows ? `，涉及 ${affectedRows} 行` : ""}。请核对数据预览和字段数量。`, "warning");
  } else {
    dom.statusPanel.className = "status-panel hidden";
    dom.statusPanel.textContent = "";
  }

  renderOverview();
  renderInsights();
  renderPreview();
  renderSchema();
  renderQuality();
  renderNumericStats();
  renderCategoryStats();
  populateControls();
  renderCharts();
  resetV2AnalysisState();
  renderTemplateMappingForm();
  resetTemplateResults("字段识别已完成。请选择分析模板并映射字段。");
}

function buildColumnProfile(field, rows) {
  const values = rows.map((row) => row[field]);
  const nonMissingValues = values.filter((value) => !isMissing(value));
  const total = rows.length;
  const nonMissingCount = nonMissingValues.length;
  const uniqueValues = new Set(nonMissingValues.map((value) => normalizeValue(value)));
  const uniqueCount = uniqueValues.size;
  const uniqueRatio = nonMissingCount ? uniqueCount / nonMissingCount : 0;
  const numericValues = nonMissingValues.map(toNumber).filter((value) => value !== null);
  const dateStrategy = buildDateStrategy(field, nonMissingValues);
  const dateValues = nonMissingValues.map((value) => toDate(value, dateStrategy)).filter(Boolean);
  const numericRatio = nonMissingCount ? numericValues.length / nonMissingCount : 0;
  const dateRatio = nonMissingCount ? dateValues.length / nonMissingCount : 0;
  const idName = isIdFieldName(field);

  let typeKey = "category";

  if (idName || (uniqueRatio >= 0.98 && nonMissingCount >= 12 && numericRatio < 0.7 && dateRatio < 0.7)) {
    typeKey = "id";
  } else if (dateRatio >= 0.8) {
    typeKey = "date";
  } else if (numericRatio >= 0.8) {
    typeKey = "numeric";
  }

  const type = FIELD_TYPE_LABELS[typeKey];
  const conversionFailuresByType = {
    numeric: nonMissingCount - numericValues.length,
    date: nonMissingCount - dateValues.length
  };
  const numericFormat = detectNumericFormat(nonMissingValues);
  const sampleValues = [];
  const sampleKeys = new Set();
  nonMissingValues.forEach((value) => {
    const displayed = displayValue(value);
    const key = normalizeValue(value);
    if (sampleValues.length >= 3 || sampleKeys.has(key)) return;
    sampleKeys.add(key);
    sampleValues.push(displayed);
  });

  return {
    field,
    type,
    typeKey,
    inferredType: type,
    inferredTypeKey: typeKey,
    total,
    nonMissingCount,
    missingCount: total - nonMissingCount,
    missingRate: total ? (total - nonMissingCount) / total : 0,
    uniqueCount,
    uniqueRatio,
    numericRatio,
    dateRatio,
    dateStrategy,
    numericFormat,
    sampleValues,
    conversionFailuresByType,
    conversionFailureCount: conversionFailuresByType[typeKey] || 0
  };
}

function countConversionFailuresForValues(values, typeKey) {
  if (typeKey !== "numeric" && typeKey !== "date") return 0;
  const convert = typeKey === "numeric" ? toNumber : toDate;
  return values.reduce((count, value) => (
    count + (!isMissing(value) && convert(value) === null ? 1 : 0)
  ), 0);
}

function countFieldConversionFailures(field, typeKey) {
  if (typeKey !== "numeric" && typeKey !== "date") return 0;
  const profile = getFieldProfile(field);
  if (profile?.conversionFailuresByType) {
    return profile.conversionFailuresByType[typeKey] || 0;
  }
  return countConversionFailuresForValues(
    state.rows.map((row) => row[field]).filter((value) => !isMissing(value)),
    typeKey
  );
}

function buildNumericStats() {
  return getProfilesByType("numeric").flatMap((profile) => {
    const values = getNumericValues(profile.field);
    if (!values.length) return [];
    const sorted = [...values].sort((a, b) => a - b);
    const avg = mean(values);
    const min = sorted[0];
    const max = sorted[sorted.length - 1];
    const q1 = quantile(sorted, 0.25);
    const q3 = quantile(sorted, 0.75);
    const iqr = q3 - q1;
    const lowerBound = q1 - 1.5 * iqr;
    const upperBound = q3 + 1.5 * iqr;
    const outliers = values.filter((value) => value < lowerBound || value > upperBound);

    return [{
      field: profile.field,
      count: values.length,
      mean: avg,
      median: median(values),
      min,
      max,
      std: standardDeviation(values, avg),
      q1,
      q3,
      lowerBound,
      upperBound,
      outlierCount: outliers.length
    }];
  });
}

function buildCategoryStats() {
  return getProfilesByType("category").map((profile) => {
    const counts = countValues(profile.field);
    const top = Array.from(counts.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([name, count]) => ({
        name,
        count,
        ratio: profile.nonMissingCount ? count / profile.nonMissingCount : 0
      }));

    return {
      field: profile.field,
      total: profile.nonMissingCount,
      top
    };
  });
}

function renderOverview() {
  const activeProfiles = getActiveProfiles();
  const numericCount = getProfilesByType("numeric").length;
  const categoryCount = getProfilesByType("category").length;
  const dateCount = getProfilesByType("date").length;
  const ignoredCount = getProfilesByType("ignore").length;
  const cards = [
    ["总行数", state.rows.length],
    ["总字段数", state.fields.length],
    ["参与分析字段数", activeProfiles.length],
    ["数值字段数量", numericCount],
    ["分类字段数量", categoryCount],
    ["日期字段数量", dateCount],
    ["忽略字段数量", ignoredCount],
    ["缺失值总数", state.totalMissing],
    ["重复行数量", state.duplicateRows]
  ];

  dom.overviewCards.innerHTML = cards.map(([label, value]) => `
    <div class="metric-card">
      <div class="metric-label">${escapeHtml(label)}</div>
      <div class="metric-value">${formatInteger(value)}</div>
    </div>
  `).join("");
}

function renderInsights() {
  const insights = [];
  const activeProfiles = getActiveProfiles();
  const activeFields = activeProfiles.map((profile) => profile.field);
  const ignoredCount = state.profiles.length - activeProfiles.length;
  const activeMissing = activeProfiles.reduce((sum, profile) => sum + profile.missingCount, 0);
  const analysisDuplicates = activeFields.length ? countDuplicateRows(state.rows, activeFields) : 0;
  const missingRate = state.rows.length && activeFields.length
    ? activeMissing / (state.rows.length * activeFields.length)
    : 0;

  if (!activeFields.length) {
    insights.push(`数据集包含 ${formatInteger(state.rows.length)} 行、${formatInteger(state.fields.length)} 个原始字段。`);
    insights.push("所有字段均已设为忽略，因此未生成统计、图表或字段相关的自动结论。");
    dom.insights.innerHTML = insights.map((text) => `
      <div class="insight-item">${escapeHtml(text)}</div>
    `).join("");
    return;
  }

  const topNumeric = state.numericStats
    .slice()
    .sort((a, b) => Math.abs(b.max - b.min) - Math.abs(a.max - a.min))[0];
  const firstCategory = state.categoryStats.find((item) => item.top.length);
  const dateProfile = getProfilesByType("date")
    .find((profile) => fieldHasDateData(profile.field));

  insights.push(`数据集包含 ${formatInteger(state.rows.length)} 行，当前有 ${formatInteger(activeFields.length)} 个字段参与分析${ignoredCount ? `，另有 ${formatInteger(ignoredCount)} 个字段已忽略` : ""}；其中数值字段 ${getProfilesByType("numeric").length} 个，分类字段 ${getProfilesByType("category").length} 个，日期字段 ${getProfilesByType("date").length} 个。`);
  insights.push(`参与分析的字段共发现 ${formatInteger(activeMissing)} 个缺失值，缺失率为 ${formatPercent(missingRate)}；按参与分析字段计算的重复行数量为 ${formatInteger(analysisDuplicates)}。`);

  if (topNumeric) {
    insights.push(`数值字段「${topNumeric.field}」的平均值为 ${formatFieldNumber(topNumeric.field, topNumeric.mean)}，最低值为 ${formatFieldNumber(topNumeric.field, topNumeric.min)}，最高值为 ${formatFieldNumber(topNumeric.field, topNumeric.max)}。`);
  } else {
    insights.push("未识别到可用于统计的数值字段。");
  }

  if (firstCategory) {
    const top = firstCategory.top[0];
    insights.push(`分类字段「${firstCategory.field}」中占比最高的类别是「${top.name}」，出现 ${formatInteger(top.count)} 次，占比 ${formatPercent(top.ratio)}。`);
  } else {
    insights.push("未识别到适合生成频数统计的分类字段。");
  }

  if (dateProfile) {
    insights.push(buildTrendInsight(dateProfile.field));
  } else {
    insights.push("未识别到日期字段，因此未生成时间趋势结论。");
  }

  dom.insights.innerHTML = insights.map((text) => `
    <div class="insight-item">${escapeHtml(text)}</div>
  `).join("");
}

function renderPreview() {
  const previewRows = state.rows.slice(0, 10);
  dom.previewTable.innerHTML = `
    <thead>
      <tr>${state.fields.map((field) => `<th>${escapeHtml(field)}</th>`).join("")}</tr>
    </thead>
    <tbody>
      ${previewRows.map((row) => `
        <tr>${state.fields.map((field) => `<td>${escapeHtml(displayValue(row[field]))}</td>`).join("")}</tr>
      `).join("")}
    </tbody>
  `;
}

function renderSchema() {
  dom.schemaTable.innerHTML = `
    <thead>
      <tr>
        <th>字段名</th>
        <th>自动识别类型</th>
        <th>确认类型</th>
        <th>非空值数量</th>
        <th>唯一值数量</th>
        <th>示例值</th>
        <th>转换失败数量</th>
      </tr>
    </thead>
    <tbody>
      ${state.profiles.map((profile) => {
        const draftTypeKey = getDraftTypeKey(profile);
        const failureCount = countFieldConversionFailures(profile.field, draftTypeKey);
        const requiresConversion = draftTypeKey === "numeric" || draftTypeKey === "date";
        const exampleText = profile.sampleValues.join("、") || "—";
        const failureLabel = requiresConversion
          ? `${profile.field} 转换失败 ${formatInteger(failureCount)} 个`
          : `${profile.field} 当前类型无需转换`;
        return `
          <tr>
            <td>${escapeHtml(profile.field)}</td>
            <td><span class="type-badge ${profile.inferredTypeKey}">${escapeHtml(profile.inferredType)}</span></td>
            <td>
              <select class="schema-type-select" data-field-type="${escapeHtml(profile.field)}" aria-label="调整 ${escapeHtml(profile.field)} 的字段类型">
                ${Object.entries(FIELD_TYPE_LABELS).map(([typeKey, label]) => `
                  <option value="${typeKey}"${draftTypeKey === typeKey ? " selected" : ""}>${escapeHtml(label)}</option>
                `).join("")}
              </select>
            </td>
            <td>${formatInteger(profile.nonMissingCount)}</td>
            <td>${formatInteger(profile.uniqueCount)}</td>
            <td><span class="schema-example" title="${escapeHtml(exampleText)}">${escapeHtml(exampleText)}</span></td>
            <td>
              <span class="conversion-failure-count${failureCount ? " warning" : ""}" title="${requiresConversion ? "非空值中无法安全转换的数量" : "该类型不需要转换"}" aria-label="${escapeHtml(failureLabel)}">
                ${requiresConversion ? formatInteger(failureCount) : "—"}
              </span>
            </td>
          </tr>
        `;
      }).join("")}
    </tbody>
  `;
  updateFieldConfigActions();
  if (!dom.fieldConfigMessage.textContent) {
    setFieldConfigMessage("自动识别已完成。可修改类型后统一应用。", "");
  }
}

function getDraftTypeKey(profile) {
  return state.fieldTypeDrafts[profile.field] || profile.typeKey;
}

function stageFieldTypeChange(field, typeKey) {
  if (!Object.prototype.hasOwnProperty.call(FIELD_TYPE_LABELS, typeKey)) return;
  const profile = state.profiles.find((item) => item.field === field);
  if (!profile) return;

  state.fieldTypeDrafts[field] = typeKey;
  updateSchemaDraftRow(field, typeKey);
  updateFieldConfigActions();
  if (!hasFieldConfigDraftChanges()) {
    setFieldConfigMessage("当前选择与已应用的字段配置一致。", "");
    return;
  }
  const totalFailures = state.profiles.reduce((sum, item) => (
    sum + countFieldConversionFailures(item.field, getDraftTypeKey(item))
  ), 0);
  const failureText = totalFailures
    ? `；当前配置预计有 ${formatInteger(totalFailures)} 个非空值无法安全转换，应用后会从对应统计中跳过`
    : "";
  setFieldConfigMessage(`字段配置有未应用的修改${failureText}。`, "warning");
}

function updateSchemaDraftRow(field, typeKey) {
  const select = Array.from(dom.schemaTable.querySelectorAll("select[data-field-type]"))
    .find((item) => item.dataset.fieldType === field);
  if (!select) return;
  select.value = typeKey;

  const failureElement = select.closest("tr")?.querySelector(".conversion-failure-count");
  if (!failureElement) return;
  const failureCount = countFieldConversionFailures(field, typeKey);
  const requiresConversion = typeKey === "numeric" || typeKey === "date";
  failureElement.textContent = requiresConversion ? formatInteger(failureCount) : "—";
  failureElement.classList.toggle("warning", failureCount > 0);
  failureElement.title = requiresConversion
    ? "非空值中无法安全转换的数量"
    : "该类型不需要转换";
  failureElement.setAttribute(
    "aria-label",
    requiresConversion
      ? `${field} 转换失败 ${formatInteger(failureCount)} 个`
      : `${field} 当前类型无需转换`
  );
}

function applyFieldConfiguration() {
  if (!state.rows.length) return;
  state.typeOverrides = {};
  state.profiles.forEach((profile) => {
    const typeKey = getDraftTypeKey(profile);
    profile.typeKey = typeKey;
    profile.type = FIELD_TYPE_LABELS[typeKey];
    profile.conversionFailureCount = countFieldConversionFailures(profile.field, typeKey);
    if (typeKey !== profile.inferredTypeKey) state.typeOverrides[profile.field] = typeKey;
  });
  state.fieldTypeDrafts = Object.fromEntries(
    state.profiles.map((profile) => [profile.field, profile.typeKey])
  );

  const failureCount = state.profiles.reduce((sum, profile) => sum + profile.conversionFailureCount, 0);
  const message = failureCount
    ? `字段配置已应用；${formatInteger(failureCount)} 个非空值转换失败，已从对应统计和图表中跳过。`
    : "字段配置已应用，统计、图表和分析结论已重新计算。";
  rebuildAnalysisFromProfiles(message);
}

function restoreAutomaticFieldTypes() {
  if (!state.rows.length) return;
  state.typeOverrides = {};
  state.profiles.forEach((profile) => {
    profile.typeKey = profile.inferredTypeKey;
    profile.type = profile.inferredType;
    profile.conversionFailureCount = countFieldConversionFailures(profile.field, profile.inferredTypeKey);
  });
  state.fieldTypeDrafts = Object.fromEntries(
    state.profiles.map((profile) => [profile.field, profile.inferredTypeKey])
  );
  rebuildAnalysisFromProfiles("已恢复自动识别，统计、图表和分析结论已重新计算。");
}

function hasFieldConfigDraftChanges() {
  return state.profiles.some((profile) => getDraftTypeKey(profile) !== profile.typeKey);
}

function hasAutomaticTypeDifferences() {
  return state.profiles.some((profile) => (
    profile.typeKey !== profile.inferredTypeKey
      || getDraftTypeKey(profile) !== profile.inferredTypeKey
  ));
}

function updateFieldConfigActions() {
  dom.applyFieldConfig.disabled = !hasFieldConfigDraftChanges();
  dom.restoreAutoFieldTypes.disabled = !hasAutomaticTypeDifferences();
}

function setFieldConfigMessage(message, tone = "") {
  dom.fieldConfigMessage.className = `field-config-message${tone ? ` ${tone}` : ""}`;
  dom.fieldConfigMessage.textContent = message;
}

function ensureFieldConfigurationApplied(actionLabel) {
  if (!hasFieldConfigDraftChanges()) return true;
  setFieldConfigMessage(`字段配置有未应用的修改。请先点击“应用字段配置”，再${actionLabel}。`, "warning");
  return false;
}

function rebuildAnalysisFromProfiles(fieldConfigMessage = "字段配置已更新。") {
  state.numericStats = buildNumericStats();
  state.categoryStats = buildCategoryStats();
  renderOverview();
  renderInsights();
  renderSchema();
  renderQuality();
  renderNumericStats();
  renderCategoryStats();
  populateControls();
  renderCharts();
  resetV2AnalysisState();
  renderTemplateMappingForm();
  resetTemplateResults("字段类型已更新。请重新选择分析字段或模板映射。");
  setFieldConfigMessage(fieldConfigMessage, "success");
  updateFieldConfigActions();
}

function renderQuality() {
  const missingClass = state.totalMissing === 0 ? "success" : missingRateLevel();
  const duplicateClass = state.duplicateRows === 0 ? "success" : "warning";
  dom.qualitySummary.innerHTML = `
    <span class="quality-pill ${missingClass}">缺失值总数：${formatInteger(state.totalMissing)}</span>
    <span class="quality-pill ${duplicateClass}">重复行：${formatInteger(state.duplicateRows)}</span>
    <span class="quality-pill">异常值字段：${formatInteger(state.numericStats.filter((item) => item.outlierCount > 0).length)}</span>
  `;

  dom.qualityTable.innerHTML = `
    <thead>
      <tr>
        <th>字段</th>
        <th>类型</th>
        <th>缺失值数量</th>
        <th>缺失率</th>
      </tr>
    </thead>
    <tbody>
      ${state.profiles.map((profile) => `
        <tr>
          <td>${escapeHtml(profile.field)}</td>
          <td>${escapeHtml(profile.type)}</td>
          <td>${formatInteger(profile.missingCount)}</td>
          <td>${formatPercent(profile.missingRate)}</td>
        </tr>
      `).join("")}
    </tbody>
  `;

  if (!state.numericStats.length) {
    dom.outlierTable.innerHTML = emptyTable("未识别到数值字段。");
    return;
  }

  dom.outlierTable.innerHTML = `
    <thead>
      <tr>
        <th>字段</th>
        <th>异常值数量</th>
        <th>下界</th>
        <th>上界</th>
      </tr>
    </thead>
    <tbody>
      ${state.numericStats.map((stat) => `
        <tr>
          <td>${escapeHtml(stat.field)}</td>
          <td>${formatInteger(stat.outlierCount)}</td>
          <td>${formatFieldNumber(stat.field, stat.lowerBound)}</td>
          <td>${formatFieldNumber(stat.field, stat.upperBound)}</td>
        </tr>
      `).join("")}
    </tbody>
  `;
}

function renderNumericStats() {
  if (!state.numericStats.length) {
    dom.numericStatsTable.innerHTML = emptyTable("未识别到数值字段。");
    return;
  }

  dom.numericStatsTable.innerHTML = `
    <thead>
      <tr>
        <th>字段</th>
        <th>有效数量</th>
        <th>平均值</th>
        <th>中位数</th>
        <th>最大值</th>
        <th>最小值</th>
        <th>标准差</th>
      </tr>
    </thead>
    <tbody>
      ${state.numericStats.map((stat) => `
        <tr>
          <td>${escapeHtml(stat.field)}</td>
          <td>${formatInteger(stat.count)}</td>
          <td>${formatFieldNumber(stat.field, stat.mean)}</td>
          <td>${formatFieldNumber(stat.field, stat.median)}</td>
          <td>${formatFieldNumber(stat.field, stat.max)}</td>
          <td>${formatFieldNumber(stat.field, stat.min)}</td>
          <td>${formatFieldNumber(stat.field, stat.std)}</td>
        </tr>
      `).join("")}
    </tbody>
  `;
}

function renderCategoryStats() {
  if (!state.categoryStats.length) {
    dom.categoryStats.innerHTML = `<div class="empty-note">未识别到分类字段。</div>`;
    return;
  }

  dom.categoryStats.innerHTML = state.categoryStats.map((category) => `
    <div class="frequency-card">
      <h3>${escapeHtml(category.field)}</h3>
      ${category.top.map((item) => `
        <div class="frequency-row">
          <span class="category-name" title="${escapeHtml(item.name)}">${escapeHtml(item.name)}</span>
          <strong>${formatInteger(item.count)}</strong>
          <span>${formatPercent(item.ratio)}</span>
        </div>
      `).join("")}
    </div>
  `).join("");
}

function populateControls() {
  const numericProfiles = getProfilesByType("numeric")
    .filter((profile) => fieldHasNumericData(profile.field));
  const categoryProfiles = getProfilesByType("category");
  const dateProfiles = getProfilesByType("date")
    .filter((profile) => fieldHasDateData(profile.field));

  setSelectOptions(dom.numericChartField, numericProfiles.map((profile) => profile.field));
  setSelectOptions(dom.categoryChartField, categoryProfiles.map((profile) => profile.field));
  setSelectOptions(dom.dateChartField, dateProfiles.map((profile) => profile.field));
  setSelectOptions(dom.metricField, numericProfiles.map((profile) => profile.field));
  setSelectOptions(dom.groupField, categoryProfiles.map((profile) => profile.field));
  setSelectOptions(dom.v2MetricField, numericProfiles.map((profile) => profile.field), "请选择数值字段");
  setSelectOptions(dom.v2GroupField, categoryProfiles.map((profile) => profile.field), "请选择分类字段");
  setSelectOptions(dom.v2DateField, dateProfiles.map((profile) => profile.field), "不选择时间字段", true);
}

function renderCharts() {
  if (!window.Chart) {
    ["numericDistributionChart", "categoryTopChart", "dateTrendChart", "customAnalysisChart"].forEach((id) => {
      drawEmptyChart(document.getElementById(id), "图表组件加载失败，请检查网络连接");
    });
    return;
  }
  renderNumericDistributionChart();
  renderCategoryTopChart();
  renderDateTrendChart();
  renderCustomAnalysisChart();
}

function renderNumericDistributionChart() {
  const field = dom.numericChartField.value;
  const canvas = document.getElementById("numericDistributionChart");
  destroyChart("numericDistributionChart");
  if (!field) return drawEmptyChart(canvas, "没有可用的数值字段");

  const values = getNumericValues(field);
  if (!values.length) return drawEmptyChart(canvas, "该字段没有可安全转换的有效数值");
  const bins = buildHistogram(values, 10, field);
  state.charts.numericDistributionChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: bins.map((bin) => bin.label),
      datasets: [{
        label: `${field} 分布`,
        data: bins.map((bin) => bin.count),
        backgroundColor: CHART_THEME.primary
      }]
    },
    options: chartOptions("数量")
  });
}

function renderCategoryTopChart() {
  const field = dom.categoryChartField.value;
  const canvas = document.getElementById("categoryTopChart");
  destroyChart("categoryTopChart");
  if (!field) return drawEmptyChart(canvas, "没有可用的分类字段");

  const counts = Array.from(countValues(field).entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
  state.charts.categoryTopChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: counts.map(([name]) => name),
      datasets: [{
        label: `${field} Top 10`,
        data: counts.map(([, count]) => count),
        backgroundColor: CHART_THEME.secondary
      }]
    },
    options: chartOptions("频数")
  });
}

function renderDateTrendChart() {
  const field = dom.dateChartField.value;
  const canvas = document.getElementById("dateTrendChart");
  destroyChart("dateTrendChart");
  if (!field) return drawEmptyChart(canvas, "没有可用的日期字段");

  const numericField = dom.metricField.value || getProfilesByType("numeric")
    .find((profile) => fieldHasNumericData(profile.field))?.field;
  const trend = buildDateTrend(field, numericField);
  if (!trend.labels.length) return drawEmptyChart(canvas, "该字段没有可安全转换的有效日期");
  state.charts.dateTrendChart = new Chart(canvas, {
    type: "line",
    data: {
      labels: trend.labels,
      datasets: [{
        label: trend.label,
        data: trend.values,
        borderColor: CHART_THEME.primary,
        backgroundColor: CHART_THEME.primaryFill,
        tension: 0.25,
        fill: true
      }]
    },
    options: chartOptions(numericAxisLabel(numericField, trend.axisLabel), numericField)
  });
}

function renderCustomAnalysisChart() {
  const metric = dom.metricField.value;
  const group = dom.groupField.value;
  const method = dom.aggregateMethod.value;
  const canvas = document.getElementById("customAnalysisChart");
  destroyChart("customAnalysisChart");

  if (!metric || !group) {
    return drawEmptyChart(canvas, "请选择数值字段和分类字段");
  }

  const grouped = aggregateByCategory(metric, group, method).slice(0, 12);
  const methodLabel = { sum: "合计", avg: "平均值", median: "中位数" }[method];
  state.charts.customAnalysisChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: grouped.map((item) => item.name),
      datasets: [{
        label: `${group} 分组的 ${metric} ${methodLabel}`,
        data: grouped.map((item) => item.value),
        backgroundColor: CHART_THEME.warning
      }]
    },
    options: chartOptions(numericAxisLabel(metric, methodLabel), metric)
  });
}

function resetV2AnalysisState() {
  destroyChart("v2TopGroupChart");
  destroyChart("v2SelectedTrendChart");
  dom.v2AnalysisResults.classList.add("hidden");
  dom.v2FieldPrompt.className = "analysis-prompt";
  dom.v2FieldPrompt.textContent = "请选择一个数值指标字段和一个分类分组字段，再点击“重新选择字段并分析”。时间字段为可选。";
}

function renderV2Analysis() {
  const metricField = dom.v2MetricField.value;
  const groupField = dom.v2GroupField.value;
  const dateField = dom.v2DateField.value;

  destroyChart("v2TopGroupChart");
  destroyChart("v2SelectedTrendChart");

  const validationMessage = getV2ValidationMessage(metricField, groupField, dateField);
  if (validationMessage) {
    dom.v2AnalysisResults.classList.add("hidden");
    dom.v2FieldPrompt.className = "analysis-prompt warning";
    dom.v2FieldPrompt.textContent = validationMessage;
    return;
  }

  const grouped = buildV2GroupedStats(metricField, groupField);
  if (!grouped.length) {
    dom.v2AnalysisResults.classList.add("hidden");
    dom.v2FieldPrompt.className = "analysis-prompt warning";
    dom.v2FieldPrompt.textContent = "所选字段没有可计算的有效记录，请重新选择字段后再分析。";
    return;
  }

  dom.v2FieldPrompt.className = "analysis-prompt";
  dom.v2FieldPrompt.textContent = dateField
    ? "已按所选字段生成分组汇总和时间趋势。"
    : "已按所选字段生成分组汇总；未选择时间字段，因此不生成时间趋势图。";
  dom.v2AnalysisResults.classList.remove("hidden");

  renderV2SummaryCards(metricField, groupField, dateField, grouped);
  renderV2AggregationTables(metricField, groupField, grouped);
  renderV2TopGroupChart(metricField, groupField, grouped);

  if (dateField) {
    dom.v2TrendCard.classList.remove("hidden");
    renderV2TrendChart(metricField, dateField);
  } else {
    dom.v2TrendCard.classList.add("hidden");
  }
}

function getV2ValidationMessage(metricField, groupField, dateField) {
  if (!state.rows.length) return "请先上传或加载 CSV 数据。";
  if (!metricField) return "请选择一个数值字段作为分析指标。";
  if (!groupField) return "请选择一个分类字段作为分组维度。";
  if (!getProfilesByType("numeric").some((profile) => profile.field === metricField)) {
    return "分析指标必须是系统识别出的数值字段。";
  }
  if (!getProfilesByType("category").some((profile) => profile.field === groupField)) {
    return "分组维度必须是系统识别出的分类字段。";
  }
  if (dateField && !getProfilesByType("date").some((profile) => profile.field === dateField)) {
    return "时间字段必须是系统识别出的日期字段，或保持不选择。";
  }
  return "";
}

function buildV2GroupedStats(metricField, groupField) {
  return groupNumericFieldByCategory(metricField, groupField);
}

function renderV2SummaryCards(metricField, groupField, dateField, grouped) {
  const validRecordCount = grouped.reduce((total, item) => total + item.count, 0);
  const trendPointCount = dateField ? buildMetricDateTrend(dateField, metricField).labels.length : 0;
  const cards = [
    ["分析指标", metricField],
    ["分组维度", groupField],
    ["有效记录数", formatInteger(validRecordCount)],
    [dateField ? "时间点数量" : "时间字段", dateField ? formatInteger(trendPointCount) : "未选择"]
  ];

  dom.v2SummaryCards.innerHTML = cards.map(([label, value]) => `
    <div class="analysis-summary-card">
      <div class="analysis-summary-label">${escapeHtml(label)}</div>
      <div class="analysis-summary-value" title="${escapeHtml(value)}">${escapeHtml(value)}</div>
    </div>
  `).join("");
}

function renderV2AggregationTables(metricField, groupField, grouped) {
  const bySum = [...grouped].sort((a, b) => b.sum - a.sum);
  const byAvg = [...grouped].sort((a, b) => b.avg - a.avg);
  dom.v2SumTable.innerHTML = buildV2AggregationTableHtml(groupField, metricField, bySum, "sum");
  dom.v2AvgTable.innerHTML = buildV2AggregationTableHtml(groupField, metricField, byAvg, "avg");
}

function buildV2AggregationTableHtml(groupField, metricField, rows, valueKey) {
  const valueLabel = valueKey === "sum" ? "求和值" : "平均值";
  return `
    <thead>
      <tr>
        <th>${escapeHtml(groupField)}</th>
        <th>记录数</th>
        <th>${escapeHtml(metricField)} ${valueLabel}</th>
      </tr>
    </thead>
    <tbody>
      ${rows.map((row) => `
        <tr>
          <td>${escapeHtml(row.name)}</td>
          <td>${formatInteger(row.count)}</td>
          <td>${formatFieldNumber(metricField, row[valueKey])}</td>
        </tr>
      `).join("")}
    </tbody>
  `;
}

function renderV2TopGroupChart(metricField, groupField, grouped) {
  const canvas = document.getElementById("v2TopGroupChart");
  if (!window.Chart) return drawEmptyChart(canvas, "图表组件加载失败，请检查网络连接");
  const topGroups = [...grouped].sort((a, b) => b.sum - a.sum).slice(0, 10);
  state.charts.v2TopGroupChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: topGroups.map((item) => item.name),
      datasets: [{
        label: `${groupField} Top 10（${metricField} 求和）`,
        data: topGroups.map((item) => item.sum),
        backgroundColor: CHART_THEME.primary
      }]
    },
    options: chartOptions(numericAxisLabel(metricField, "求和值"), metricField)
  });
}

function renderV2TrendChart(metricField, dateField) {
  const canvas = document.getElementById("v2SelectedTrendChart");
  if (!window.Chart) return drawEmptyChart(canvas, "图表组件加载失败，请检查网络连接");
  const trend = buildMetricDateTrend(dateField, metricField);
  if (!trend.labels.length) {
    return drawEmptyChart(canvas, "所选时间字段没有可用的趋势数据");
  }

  state.charts.v2SelectedTrendChart = new Chart(canvas, {
    type: "line",
    data: {
      labels: trend.labels,
      datasets: [{
        label: `${metricField} 按${getDateGranularity(dateField) === "year" ? "年份" : "时间"}求和趋势`,
        data: trend.values,
        borderColor: CHART_THEME.secondary,
        backgroundColor: CHART_THEME.secondaryFill,
        tension: 0.25,
        fill: true
      }]
    },
    options: chartOptions(numericAxisLabel(metricField, "求和值"), metricField)
  });
}

function buildMetricDateTrend(dateField, metricField) {
  const items = [];
  state.rows.forEach((row) => {
    const date = parseFieldDate(dateField, row[dateField]);
    const value = toNumber(row[metricField]);
    if (!date || value === null) return;
    items.push({ date, value });
  });

  if (!items.length) {
    return { labels: [], values: [] };
  }

  items.sort((a, b) => a.date - b.date);
  const spanDays = (items[items.length - 1].date - items[0].date) / 86400000;
  const groupByMonth = spanDays > 120;
  const groups = new Map();

  items.forEach((item) => {
    const key = getDateTrendKey(dateField, item.date, groupByMonth);
    groups.set(key, (groups.get(key) || 0) + item.value);
  });

  const labels = Array.from(groups.keys()).sort();
  return {
    labels,
    values: labels.map((label) => groups.get(label))
  };
}

function getCurrentTemplate() {
  return TEMPLATE_DEFINITIONS[dom.scenarioTemplate.value] || TEMPLATE_DEFINITIONS.generic;
}

function renderTemplateMappingForm() {
  const template = getCurrentTemplate();
  dom.templateMappingForm.innerHTML = template.fields.map((field) => `
    <label>
      ${escapeHtml(field.label)}${field.required ? " *" : ""}
      <select id="templateField_${escapeHtml(field.key)}" data-template-field="${escapeHtml(field.key)}">
        ${buildTemplateFieldOptions(field)}
      </select>
    </label>
  `).join("");
}

function buildTemplateFieldOptions(fieldConfig) {
  const placeholder = fieldConfig.required ? "请选择字段" : "不选择";
  const options = getActiveProfiles()
    .filter((profile) => fieldConfig.expected === "any" || profile.typeKey === fieldConfig.expected)
    .filter((profile) => fieldConfig.expected !== "numeric" || fieldHasNumericData(profile.field))
    .filter((profile) => fieldConfig.expected !== "date" || fieldHasDateData(profile.field))
    .map((profile) => {
      return `<option value="${escapeHtml(profile.field)}">${escapeHtml(profile.field)}（${escapeHtml(profile.type)}）</option>`;
    }).join("");
  return `<option value="">${placeholder}</option>${options}`;
}

function resetTemplateResults(message) {
  clearTemplateCharts();
  dom.templateResults.classList.add("hidden");
  dom.templatePrompt.className = "analysis-prompt";
  dom.templatePrompt.textContent = message;
}

function clearTemplateCharts() {
  destroyChart("templateMainChart");
  destroyChart("templateSecondaryChart");
  destroyChart("templateTrendChart");
}

function getTemplateMappings(template) {
  return template.fields.reduce((mappings, field) => {
    const select = document.getElementById(`templateField_${field.key}`);
    mappings[field.key] = select ? select.value : "";
    return mappings;
  }, {});
}

function runTemplateAnalysis() {
  const templateKey = dom.scenarioTemplate.value;
  const template = getCurrentTemplate();
  const mappings = getTemplateMappings(template);
  const validationMessage = validateTemplateMappings(template, mappings);

  clearTemplateCharts();
  if (validationMessage) {
    dom.templateResults.classList.add("hidden");
    dom.templatePrompt.className = "analysis-prompt warning";
    dom.templatePrompt.textContent = validationMessage;
    return;
  }

  const result = {
    generic: analyzeGenericTemplate,
    sales: analyzeSalesTemplate,
    score: analyzeScoreTemplate,
    usedGoods: analyzeUsedGoodsTemplate,
    survey: analyzeSurveyTemplate,
    behavior: analyzeBehaviorTemplate
  }[templateKey](mappings);

  renderTemplateResult(result);
}

function validateTemplateMappings(template, mappings) {
  for (const field of template.fields) {
    const selected = mappings[field.key];
    if (field.required && !selected) {
      return `请为「${field.label.replace("（可选）", "")}」选择字段。`;
    }
    if (!selected) continue;
    const profile = getFieldProfile(selected);
    if (!profile || profile.typeKey === "ignore") {
      return `字段「${selected}」已被忽略或不再可用，请重新选择字段。`;
    }
    if (field.expected !== "any" && profile.typeKey !== field.expected) {
      return `字段「${selected}」当前配置为「${profile.type}」，与「${field.label}」要求的类型不匹配。`;
    }
    if (field.expected === "numeric" && !fieldHasNumericData(selected)) {
      return `字段「${selected}」无法解析出有效数值，请为「${field.label}」选择数值字段。`;
    }
    if (field.expected === "date" && !fieldHasDateData(selected)) {
      return `字段「${selected}」无法解析出有效日期，请为「${field.label}」选择日期字段或保持不选择。`;
    }
  }

  if (dom.scenarioTemplate.value === "survey") {
    const selectedCount = Object.values(mappings).filter(Boolean).length;
    if (selectedCount === 0) {
      return "问卷调查分析至少需要选择一个字段。";
    }
  }

  return "";
}

function renderTemplateResult(result) {
  dom.templatePrompt.className = "analysis-prompt";
  dom.templatePrompt.textContent = result.prompt;
  dom.templateResults.classList.remove("hidden");
  dom.templateSummaryCards.innerHTML = result.cards.map(([label, value]) => `
    <div class="analysis-summary-card">
      <div class="analysis-summary-label">${escapeHtml(label)}</div>
      <div class="analysis-summary-value" title="${escapeHtml(value)}">${escapeHtml(value)}</div>
    </div>
  `).join("");
  dom.templateInsights.innerHTML = result.insights.map((text) => `
    <div class="insight-item">${escapeHtml(text)}</div>
  `).join("");
  dom.templatePrimaryTable.innerHTML = buildSimpleTableHtml(result.table.headers, result.table.rows);

  renderTemplateChart("templateMainChart", dom.templateMainChartTitle, result.mainChart);
  renderTemplateChart("templateSecondaryChart", dom.templateSecondaryChartTitle, result.secondaryChart);
  if (result.trendChart) {
    dom.templateTrendCard.classList.remove("hidden");
    renderTemplateChart("templateTrendChart", dom.templateTrendChartTitle, result.trendChart);
  } else {
    dom.templateTrendCard.classList.add("hidden");
  }
}

function renderTemplateChart(chartId, titleElement, chartConfig) {
  const canvas = document.getElementById(chartId);
  titleElement.textContent = chartConfig.title;
  if (!window.Chart) {
    drawEmptyChart(canvas, "图表组件加载失败，请检查网络连接");
    return;
  }
  if (!chartConfig.labels.length) {
    drawEmptyChart(canvas, "没有足够数据生成该图表");
    return;
  }

  state.charts[chartId] = new Chart(canvas, {
    type: chartConfig.type || "bar",
    data: {
      labels: chartConfig.labels,
      datasets: [{
        label: chartConfig.label,
        data: chartConfig.values,
        borderColor: chartConfig.color || CHART_THEME.primary,
        backgroundColor: chartConfig.background || chartConfig.color || CHART_THEME.primary,
        tension: 0.25,
        fill: chartConfig.type === "line"
      }]
    },
    options: chartOptions(
      numericAxisLabel(chartConfig.valueField, chartConfig.axisLabel || "值"),
      chartConfig.valueField
    )
  });
}

function analyzeGenericTemplate(mappings) {
  const grouped = groupNumericFieldByCategory(mappings.metric, mappings.group);
  const values = getNumericValues(mappings.metric);
  const stats = summarizeNumbers(values);
  const topBySum = grouped[0];
  const topByAvg = [...grouped].sort((a, b) => b.avg - a.avg)[0];
  const trend = mappings.date ? buildTrendFromRows(mappings.date, (row) => toNumber(row[mappings.metric]), "sum") : null;
  const insights = [
    `本次通用分析使用「${mappings.metric}」作为数值指标、「${mappings.group}」作为分组维度，有效数值记录为 ${formatInteger(values.length)} 条。`,
    `「${mappings.metric}」的平均值为 ${formatFieldNumber(mappings.metric, stats.mean)}，中位数为 ${formatFieldNumber(mappings.metric, stats.median)}，最小值为 ${formatFieldNumber(mappings.metric, stats.min)}，最大值为 ${formatFieldNumber(mappings.metric, stats.max)}。`,
    topBySum ? `按「${mappings.group}」分组求和最高的是「${topBySum.name}」，求和值为 ${formatFieldNumber(mappings.metric, topBySum.sum)}。` : "所选字段没有形成可比较的分组结果。",
    topByAvg ? `按「${mappings.group}」分组平均值最高的是「${topByAvg.name}」，平均值为 ${formatFieldNumber(mappings.metric, topByAvg.avg)}。该结果只描述数据差异，不说明原因。` : "所选字段没有形成平均值比较结果。"
  ];
  if (trend) insights.push(buildTrendFactText(mappings.metric, mappings.date, trend, "求和值"));

  return {
    prompt: "已生成通用字段映射分析。",
    cards: [
      ["模板", "通用数据分析"],
      ["指标字段", mappings.metric],
      ["分组字段", mappings.group],
      ["有效记录", formatInteger(values.length)]
    ],
    insights,
    table: numericGroupTable(mappings.group, mappings.metric, grouped),
    mainChart: barChartConfig(`按 ${mappings.group} 的 ${mappings.metric} 求和 Top 10`, grouped.slice(0, 10).map((item) => item.name), grouped.slice(0, 10).map((item) => item.sum), `${mappings.metric} 求和`, "求和值", CHART_THEME.primary, mappings.metric),
    secondaryChart: histogramChartConfig(`${mappings.metric} 分布`, values, "记录数", mappings.metric),
    trendChart: trend ? lineChartConfig(`${mappings.metric} 时间趋势`, trend.labels, trend.values, `${mappings.metric} 求和`, "求和值", mappings.metric) : null
  };
}

function analyzeSalesTemplate(mappings) {
  const metricRows = buildSalesMetricRows(mappings);
  const categoryRows = metricRows
    ? groupMetricRowsByField(metricRows, mappings.category)
    : frequencyRows(mappings.category);
  const regionRows = metricRows
    ? groupMetricRowsByField(metricRows, mappings.region)
    : frequencyRows(mappings.region);
  const metricLabel = metricRows ? metricRows.label : "记录数";
  const metricFormatField = metricRows?.formatField || "";
  const categoryTop = categoryRows[0];
  const regionTop = regionRows[0];
  const trend = mappings.date
    ? buildTrendFromRows(mappings.date, metricRows ? (row) => getSalesMetricValue(row, mappings) : null, metricRows ? "sum" : "count")
    : null;

  const insights = [
    `销售模板使用「${mappings.category}」和「${mappings.region}」进行分组，共分析 ${formatInteger(state.rows.length)} 行记录。`,
    metricRows
      ? `销售数值基于「${metricLabel}」计算，有效数值记录为 ${formatInteger(metricRows.length)} 条。`
      : "未选择可计算的销售数值字段，因此仅按记录数做分组统计。",
    categoryTop ? `按品类字段「${mappings.category}」比较，${metricRows ? "求和" : "记录数"}最高的分组是「${categoryTop.name}」，值为 ${metricRows ? formatFieldNumber(metricFormatField, categoryTop.sum) : formatNumber(categoryTop.count)}。` : "品类字段没有可统计的有效分组。",
    regionTop ? `按地区字段「${mappings.region}」比较，${metricRows ? "求和" : "记录数"}最高的分组是「${regionTop.name}」，值为 ${metricRows ? formatFieldNumber(metricFormatField, regionTop.sum) : formatNumber(regionTop.count)}。该结果只反映 CSV 中的数据事实。` : "地区字段没有可统计的有效分组。"
  ];
  if (trend) insights.push(buildTrendFactText(metricLabel, mappings.date, trend, metricRows ? "求和值" : "记录数"));

  return {
    prompt: "已生成销售数据模板分析。",
    cards: [
      ["模板", "销售数据分析"],
      ["销售指标", metricLabel],
      ["品类字段", mappings.category],
      ["地区字段", mappings.region]
    ],
    insights,
    table: metricRows
      ? numericGroupTable(mappings.category, metricLabel, categoryRows, metricFormatField)
      : frequencyTable(mappings.category, categoryRows),
    mainChart: barChartConfig(`按 ${mappings.category} 的 ${metricLabel} Top 10`, topLabels(categoryRows), topValues(categoryRows, metricRows ? "sum" : "count"), metricRows ? `${metricLabel} 求和` : "记录数", metricRows ? "求和值" : "记录数", CHART_THEME.primary, metricFormatField),
    secondaryChart: barChartConfig(`按 ${mappings.region} 的 ${metricLabel} Top 10`, topLabels(regionRows), topValues(regionRows, metricRows ? "sum" : "count"), metricRows ? `${metricLabel} 求和` : "记录数", metricRows ? "求和值" : "记录数", CHART_THEME.secondary, metricFormatField),
    trendChart: trend ? lineChartConfig(`${metricLabel} 时间趋势`, trend.labels, trend.values, metricRows ? `${metricLabel} 求和` : "记录数", metricRows ? "求和值" : "记录数", metricFormatField) : null
  };
}

function analyzeScoreTemplate(mappings) {
  const values = getNumericValues(mappings.score);
  const stats = summarizeNumbers(values);
  const dimension = mappings.class || mappings.subject || "";
  const grouped = dimension ? groupNumericFieldByCategory(mappings.score, dimension) : [];
  const topAvg = grouped.length ? [...grouped].sort((a, b) => b.avg - a.avg)[0] : null;
  const lowAvg = grouped.length ? [...grouped].sort((a, b) => a.avg - b.avg)[0] : null;
  const trend = mappings.examDate ? buildTrendFromRows(mappings.examDate, (row) => toNumber(row[mappings.score]), "avg") : null;
  const insights = [
    `成绩模板使用「${mappings.score}」作为成绩字段，有效成绩记录为 ${formatInteger(values.length)} 条。`,
    `成绩平均值为 ${formatFieldNumber(mappings.score, stats.mean)}，中位数为 ${formatFieldNumber(mappings.score, stats.median)}，最低值为 ${formatFieldNumber(mappings.score, stats.min)}，最高值为 ${formatFieldNumber(mappings.score, stats.max)}。`,
    topAvg && lowAvg ? `按「${dimension}」分组时，平均值最高的是「${topAvg.name}」(${formatFieldNumber(mappings.score, topAvg.avg)})，平均值最低的是「${lowAvg.name}」(${formatFieldNumber(mappings.score, lowAvg.avg)})。这里只描述字段结果，不评价教学或个体原因。` : "未选择班级或科目字段，因此未生成分组平均值比较。"
  ];
  if (trend) insights.push(buildTrendFactText(mappings.score, mappings.examDate, trend, "平均值"));

  return {
    prompt: "已生成成绩数据模板分析。",
    cards: [
      ["模板", "成绩数据分析"],
      ["成绩字段", mappings.score],
      ["有效成绩", formatInteger(values.length)],
      ["平均值", formatFieldNumber(mappings.score, stats.mean)]
    ],
    insights,
    table: dimension ? numericGroupTable(dimension, mappings.score, grouped) : simpleStatsTable(mappings.score, stats, values.length),
    mainChart: dimension
      ? barChartConfig(`按 ${dimension} 的成绩平均值 Top 10`, topLabels(grouped, "avg"), topValues(grouped, "avg"), "成绩平均值", "平均值", CHART_THEME.primary, mappings.score)
      : histogramChartConfig(`${mappings.score} 分布`, values, "记录数", mappings.score),
    secondaryChart: histogramChartConfig(`${mappings.score} 分布`, values, "记录数", mappings.score),
    trendChart: trend ? lineChartConfig(`${mappings.score} 时间趋势`, trend.labels, trend.values, "成绩平均值", "平均值", mappings.score) : null
  };
}

function analyzeUsedGoodsTemplate(mappings) {
  const values = getNumericValues(mappings.price);
  const stats = summarizeNumbers(values);
  const dimensions = ["item", "platform", "city", "condition"].map((key) => mappings[key]).filter(Boolean);
  const primaryDimension = dimensions[0];
  const grouped = primaryDimension ? groupNumericFieldByCategory(mappings.price, primaryDimension) : [];
  const topAvg = [...grouped].sort((a, b) => b.avg - a.avg)[0];
  const trend = mappings.date ? buildTrendFromRows(mappings.date, (row) => toNumber(row[mappings.price]), "avg") : null;
  const insights = [
    `二手商品价格模板使用「${mappings.price}」作为价格字段，有效价格记录为 ${formatInteger(values.length)} 条。`,
    `价格平均值为 ${formatFieldNumber(mappings.price, stats.mean)}，中位数为 ${formatFieldNumber(mappings.price, stats.median)}，最低值为 ${formatFieldNumber(mappings.price, stats.min)}，最高值为 ${formatFieldNumber(mappings.price, stats.max)}。`,
    topAvg ? `按「${primaryDimension}」比较，平均价格最高的分组是「${topAvg.name}」，平均值为 ${formatFieldNumber(mappings.price, topAvg.avg)}。该结果只说明样本中的价格差异，不代表市场整体水平。` : "未选择商品、平台、城市或成色字段，因此仅生成价格整体统计。"
  ];
  if (trend) insights.push(buildTrendFactText(mappings.price, mappings.date, trend, "平均值"));

  return {
    prompt: "已生成二手商品价格模板分析。",
    cards: [
      ["模板", "二手商品价格分析"],
      ["价格字段", mappings.price],
      ["比较维度", primaryDimension || "未选择"],
      ["价格中位数", formatFieldNumber(mappings.price, stats.median)]
    ],
    insights,
    table: primaryDimension ? numericGroupTable(primaryDimension, mappings.price, grouped) : simpleStatsTable(mappings.price, stats, values.length),
    mainChart: primaryDimension
      ? barChartConfig(`按 ${primaryDimension} 的平均价格 Top 10`, topLabels(grouped, "avg"), topValues(grouped, "avg"), "平均价格", "平均值", CHART_THEME.primary, mappings.price)
      : histogramChartConfig(`${mappings.price} 分布`, values, "记录数", mappings.price),
    secondaryChart: histogramChartConfig(`${mappings.price} 分布`, values, "记录数", mappings.price),
    trendChart: trend ? lineChartConfig(`${mappings.price} 时间趋势`, trend.labels, trend.values, "平均价格", "平均值", mappings.price) : null
  };
}

function analyzeSurveyTemplate(mappings) {
  const selectedFields = Object.entries(mappings).filter(([, value]) => value);
  const numericSummaries = [];
  ["satisfaction", "age"].forEach((key) => {
    if (!mappings[key]) return;
    const values = getNumericValues(mappings[key]);
    if (values.length) numericSummaries.push({ key, field: mappings[key], stats: summarizeNumbers(values), count: values.length });
  });
  const categoricalFields = ["gender", "city", "choice"].map((key) => mappings[key]).filter(Boolean);
  const primaryCategory = categoricalFields[0] || "";
  const primaryRows = primaryCategory ? frequencyRows(primaryCategory) : [];
  const chartValues = numericSummaries[0] ? getNumericValues(numericSummaries[0].field) : [];
  const insights = [
    `问卷模板共选择 ${formatInteger(selectedFields.length)} 个字段，数据表总记录数为 ${formatInteger(state.rows.length)}。`,
    numericSummaries[0]
      ? `数值字段「${numericSummaries[0].field}」的平均值为 ${formatFieldNumber(numericSummaries[0].field, numericSummaries[0].stats.mean)}，中位数为 ${formatFieldNumber(numericSummaries[0].field, numericSummaries[0].stats.median)}。`
      : "未选择可解析的数值字段，因此未生成满意度或年龄的数值统计。",
    primaryRows[0]
      ? `分类字段「${primaryCategory}」中出现最多的类别是「${primaryRows[0].name}」，占比 ${formatPercent(primaryRows[0].ratio)}。该结果只描述样本分布。`
      : "未选择分类字段，因此未生成选择题或人口属性频数统计。"
  ];

  return {
    prompt: "已生成问卷调查模板分析。",
    cards: [
      ["模板", "问卷调查分析"],
      ["已选字段", formatInteger(selectedFields.length)],
      ["记录数", formatInteger(state.rows.length)],
      ["主分类字段", primaryCategory || "未选择"]
    ],
    insights,
    table: surveyTable(numericSummaries, categoricalFields),
    mainChart: primaryCategory
      ? barChartConfig(`${primaryCategory} Top 10`, topLabels(primaryRows), topValues(primaryRows, "count"), "频数", "频数", CHART_THEME.primary)
      : histogramChartConfig(numericSummaries[0]?.field || "数值字段", chartValues, "记录数", numericSummaries[0]?.field || ""),
    secondaryChart: numericSummaries[0]
      ? histogramChartConfig(`${numericSummaries[0].field} 分布`, chartValues, "记录数", numericSummaries[0].field)
      : barChartConfig("分类字段 Top 10", topLabels(primaryRows), topValues(primaryRows, "count"), "频数", "频数", CHART_THEME.secondary),
    trendChart: null
  };
}

function analyzeBehaviorTemplate(mappings) {
  const eventRows = frequencyRows(mappings.event);
  const userCounts = countNonMissingValues(mappings.userId);
  const uniqueUsers = new Set(state.rows.map((row) => displayValue(row[mappings.userId])).filter(Boolean)).size;
  const avgEventsPerUser = uniqueUsers ? userCounts / uniqueUsers : 0;
  const secondaryField = mappings.channel || mappings.device || "";
  const secondaryRows = secondaryField ? frequencyRows(secondaryField) : topUserRows(mappings.userId);
  const trend = mappings.date ? buildTrendFromRows(mappings.date, null, "count") : null;
  const insights = [
    `用户行为模板使用「${mappings.userId}」作为用户 ID、「${mappings.event}」作为事件字段，有效事件记录为 ${formatInteger(userCounts)} 条。`,
    `唯一用户数为 ${formatInteger(uniqueUsers)}，平均每个用户对应 ${formatNumber(avgEventsPerUser)} 条事件记录。`,
    eventRows[0] ? `事件字段中出现最多的是「${eventRows[0].name}」，出现 ${formatInteger(eventRows[0].count)} 次，占比 ${formatPercent(eventRows[0].ratio)}。该结果不推断用户动机。` : "事件字段没有可统计的有效类别。"
  ];
  if (secondaryField && secondaryRows[0]) {
    insights.push(`辅助维度「${secondaryField}」中记录数最高的是「${secondaryRows[0].name}」，出现 ${formatInteger(secondaryRows[0].count)} 次。`);
  }
  if (trend) insights.push(buildTrendFactText("事件记录数", mappings.date, trend, "记录数"));

  return {
    prompt: "已生成用户行为模板分析。",
    cards: [
      ["模板", "用户行为分析"],
      ["唯一用户", formatInteger(uniqueUsers)],
      ["事件记录", formatInteger(userCounts)],
      ["人均事件", formatNumber(avgEventsPerUser)]
    ],
    insights,
    table: frequencyTable(mappings.event, eventRows),
    mainChart: barChartConfig(`${mappings.event} Top 10`, topLabels(eventRows), topValues(eventRows, "count"), "事件数", "记录数", CHART_THEME.primary),
    secondaryChart: barChartConfig(`${secondaryField || mappings.userId} Top 10`, topLabels(secondaryRows), topValues(secondaryRows, "count"), "记录数", "记录数", CHART_THEME.secondary),
    trendChart: trend ? lineChartConfig("事件记录时间趋势", trend.labels, trend.values, "记录数", "记录数") : null
  };
}

function buildSalesMetricRows(mappings) {
  let label = "";
  let formatField = "";
  if (mappings.salesAmount) {
    label = mappings.salesAmount;
    formatField = mappings.salesAmount;
  }
  else if (mappings.unitPrice && mappings.quantity) label = `${mappings.unitPrice} × ${mappings.quantity}`;
  else if (mappings.unitPrice) {
    label = mappings.unitPrice;
    formatField = mappings.unitPrice;
  }
  else if (mappings.quantity) {
    label = mappings.quantity;
    formatField = mappings.quantity;
  }
  else return null;

  const rows = state.rows
    .map((row) => ({ row, value: getSalesMetricValue(row, mappings) }))
    .filter((item) => item.value !== null);
  rows.label = label;
  rows.formatField = formatField;
  return rows;
}

function getSalesMetricValue(row, mappings) {
  if (mappings.salesAmount) return toNumber(row[mappings.salesAmount]);
  if (mappings.unitPrice && mappings.quantity) {
    const price = toNumber(row[mappings.unitPrice]);
    const quantity = toNumber(row[mappings.quantity]);
    return price === null || quantity === null ? null : price * quantity;
  }
  if (mappings.unitPrice) return toNumber(row[mappings.unitPrice]);
  if (mappings.quantity) return toNumber(row[mappings.quantity]);
  return null;
}

function groupNumericFieldByCategory(metricField, groupField) {
  const groups = new Map();
  state.rows.forEach((row) => {
    const value = toNumber(row[metricField]);
    if (value === null) return;
    const groupName = displayValue(row[groupField]) || "未填写";
    if (!groups.has(groupName)) groups.set(groupName, []);
    groups.get(groupName).push(value);
  });
  return summarizeGroupedValues(groups);
}

function groupMetricRowsByField(metricRows, groupField) {
  const groups = new Map();
  metricRows.forEach((item) => {
    const groupName = displayValue(item.row[groupField]) || "未填写";
    if (!groups.has(groupName)) groups.set(groupName, []);
    groups.get(groupName).push(item.value);
  });
  return summarizeGroupedValues(groups);
}

function summarizeGroupedValues(groups) {
  return Array.from(groups.entries())
    .map(([name, values]) => ({
      name,
      count: values.length,
      sum: values.reduce((total, value) => total + value, 0),
      avg: mean(values),
      median: median(values),
      min: Math.min(...values),
      max: Math.max(...values)
    }))
    .sort((a, b) => b.sum - a.sum);
}

function buildTrendFromRows(dateField, valueGetter, aggregate) {
  const items = [];
  state.rows.forEach((row) => {
    const date = parseFieldDate(dateField, row[dateField]);
    if (!date) return;
    if (aggregate === "count") {
      items.push({ date, value: 1 });
      return;
    }
    const value = valueGetter(row);
    if (value !== null) items.push({ date, value });
  });

  if (!items.length) return { labels: [], values: [] };
  items.sort((a, b) => a.date - b.date);
  const spanDays = (items[items.length - 1].date - items[0].date) / 86400000;
  const groupByMonth = spanDays > 120;
  const groups = new Map();
  items.forEach(({ date, value }) => {
    const key = getDateTrendKey(dateField, date, groupByMonth);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(value);
  });

  const labels = Array.from(groups.keys()).sort();
  const values = labels.map((label) => {
    const bucket = groups.get(label);
    if (aggregate === "avg") return bucket.length ? mean(bucket) : 0;
    return bucket.reduce((sum, value) => sum + value, 0);
  });
  return { labels, values };
}

function buildTrendFactText(metricLabel, dateField, trend, valueLabel) {
  if (!trend.labels.length) return `已选择日期字段「${dateField}」，但没有足够的有效日期记录生成趋势。`;
  if (trend.labels.length === 1) return `日期字段「${dateField}」只有 1 个有效时间点，无法判断变化方向。`;
  const first = trend.values[0];
  const last = trend.values[trend.values.length - 1];
  const direction = last > first ? "上升" : last < first ? "下降" : "基本持平";
  return `按日期字段「${dateField}」观察，「${metricLabel}」${valueLabel}从 ${trend.labels[0]} 到 ${trend.labels[trend.labels.length - 1]} 呈${direction}变化；这只是时间序列数值变化，不说明原因。`;
}

function summarizeNumbers(values) {
  if (!values.length) {
    return { mean: 0, median: 0, min: 0, max: 0, std: 0 };
  }
  return {
    mean: mean(values),
    median: median(values),
    min: Math.min(...values),
    max: Math.max(...values),
    std: standardDeviation(values)
  };
}

function numericGroupTable(groupField, metricField, rows, formatField = metricField) {
  return {
    headers: [groupField, "记录数", `${metricField} 求和`, `${metricField} 平均值`, "最小值", "最大值"],
    rows: rows.map((row) => [
      row.name,
      formatInteger(row.count),
      formatFieldNumber(formatField, row.sum),
      formatFieldNumber(formatField, row.avg),
      formatFieldNumber(formatField, row.min),
      formatFieldNumber(formatField, row.max)
    ])
  };
}

function frequencyTable(field, rows) {
  return {
    headers: [field, "记录数", "占比"],
    rows: rows.map((row) => [row.name, formatInteger(row.count), formatPercent(row.ratio)])
  };
}

function simpleStatsTable(field, stats, count) {
  return {
    headers: ["字段", "有效数量", "平均值", "中位数", "最小值", "最大值", "标准差"],
    rows: [[field, formatInteger(count), formatFieldNumber(field, stats.mean), formatFieldNumber(field, stats.median), formatFieldNumber(field, stats.min), formatFieldNumber(field, stats.max), formatFieldNumber(field, stats.std)]]
  };
}

function surveyTable(numericSummaries, categoricalFields) {
  const rows = [];
  numericSummaries.forEach((item) => {
    rows.push([item.field, "数值字段", `有效 ${formatInteger(item.count)} 条，平均值 ${formatFieldNumber(item.field, item.stats.mean)}，中位数 ${formatFieldNumber(item.field, item.stats.median)}`]);
  });
  categoricalFields.forEach((field) => {
    const top = frequencyRows(field)[0];
    rows.push([field, "分类字段", top ? `Top 类别「${top.name}」，占比 ${formatPercent(top.ratio)}` : "没有有效类别"]);
  });
  return {
    headers: ["字段", "类型", "结果摘要"],
    rows
  };
}

function buildSimpleTableHtml(headers, rows) {
  return `
    <thead>
      <tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr>
    </thead>
    <tbody>
      ${rows.map((row) => `
        <tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>
      `).join("")}
    </tbody>
  `;
}

function frequencyRows(field, limit = 50) {
  const counts = countValues(field);
  const total = Array.from(counts.values()).reduce((sum, count) => sum + count, 0);
  return Array.from(counts.entries())
    .map(([name, count]) => ({ name, count, ratio: total ? count / total : 0 }))
    .sort((a, b) => b.count - a.count)
    .slice(0, limit);
}

function topUserRows(field) {
  return frequencyRows(field, 10);
}

function topLabels(rows, valueKey = "sum") {
  return [...rows].sort((a, b) => (b[valueKey] ?? 0) - (a[valueKey] ?? 0)).slice(0, 10).map((item) => item.name);
}

function topValues(rows, valueKey) {
  return [...rows].sort((a, b) => (b[valueKey] ?? 0) - (a[valueKey] ?? 0)).slice(0, 10).map((item) => item[valueKey] ?? 0);
}

function barChartConfig(title, labels, values, label, axisLabel, color, valueField = "") {
  return { type: "bar", title, labels, values, label, axisLabel, color, valueField };
}

function lineChartConfig(title, labels, values, label, axisLabel, valueField = "") {
  return {
    type: "line",
    title,
    labels,
    values,
    label,
    axisLabel,
    valueField,
    color: CHART_THEME.secondary,
    background: CHART_THEME.secondaryFill
  };
}

function histogramChartConfig(title, values, axisLabel, sourceField = "") {
  const bins = buildHistogram(values, 10, sourceField);
  return {
    type: "bar",
    title,
    labels: bins.map((bin) => bin.label),
    values: bins.map((bin) => bin.count),
    label: title,
    axisLabel,
    color: CHART_THEME.warning
  };
}

function fieldHasNumericData(field) {
  return getNumericValues(field).length > 0;
}

function fieldHasDateData(field) {
  return state.rows.some((row) => Boolean(parseFieldDate(field, row[field])));
}

function countNonMissingValues(field) {
  return state.rows.filter((row) => !isMissing(row[field])).length;
}

function buildTrendInsight(dateField) {
  const numericField = getProfilesByType("numeric")
    .find((profile) => fieldHasNumericData(profile.field))?.field;
  const trend = buildDateTrend(dateField, numericField);
  if (!trend.labels.length) {
    return `日期字段「${dateField}」没有可安全转换的有效日期，因此未生成趋势结论。`;
  }
  if (trend.values.length < 2) {
    return `日期字段「${dateField}」可用于趋势分析，但有效时间点不足以判断变化方向。`;
  }
  const first = trend.values[0];
  const last = trend.values[trend.values.length - 1];
  const change = first === 0 ? 0 : (last - first) / Math.abs(first);
  const direction = last > first ? "上升" : last < first ? "下降" : "基本持平";
  return `基于日期字段「${dateField}」观察，${trend.label} 从 ${trend.labels[0]} 到 ${trend.labels[trend.labels.length - 1]} 呈${direction}变化，变化幅度约 ${formatPercent(Math.abs(change))}。`;
}

function buildDateTrend(dateField, numericField) {
  const items = [];
  state.rows.forEach((row) => {
    const date = parseFieldDate(dateField, row[dateField]);
    if (!date) return;
    const value = numericField ? toNumber(row[numericField]) : null;
    if (numericField && value === null) return;
    items.push({ date, value });
  });

  if (!items.length) {
    return { labels: [], values: [], label: "趋势", axisLabel: "值" };
  }

  items.sort((a, b) => a.date - b.date);
  const spanDays = (items[items.length - 1].date - items[0].date) / 86400000;
  const groupByMonth = spanDays > 120;
  const groups = new Map();

  items.forEach((item) => {
    const key = getDateTrendKey(dateField, item.date, groupByMonth);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(item.value);
  });

  const labels = Array.from(groups.keys()).sort();
  const hasMetric = Boolean(numericField);
  const values = labels.map((label) => {
    const bucket = groups.get(label);
    if (!hasMetric) return bucket.length;
    const valid = bucket.filter((value) => value !== null);
    return valid.length ? mean(valid) : 0;
  });

  return {
    labels,
    values,
    label: hasMetric
      ? `${numericField}${getDateGranularity(dateField) === "year" ? "年度" : " "}平均值趋势`
      : `${getDateGranularity(dateField) === "year" ? "年度" : ""}记录数趋势`,
    axisLabel: hasMetric ? "平均值" : "记录数"
  };
}

function aggregateByCategory(metricField, groupField, method) {
  const groups = new Map();
  state.rows.forEach((row) => {
    const group = displayValue(row[groupField]) || "未填写";
    const value = toNumber(row[metricField]);
    if (value === null) return;
    if (!groups.has(group)) groups.set(group, []);
    groups.get(group).push(value);
  });

  return Array.from(groups.entries())
    .map(([name, values]) => {
      let value = 0;
      if (method === "avg") value = mean(values);
      else if (method === "median") value = median(values);
      else value = values.reduce((sum, item) => sum + item, 0);
      return { name, value, count: values.length };
    })
    .sort((a, b) => b.value - a.value);
}

function buildHistogram(values, binCount, field = "") {
  if (!values.length) return [];
  const min = Math.min(...values);
  const max = Math.max(...values);
  if (min === max) return [{ label: formatFieldNumber(field, min), count: values.length }];
  const width = (max - min) / binCount;
  const bins = Array.from({ length: binCount }, (_, index) => {
    const start = min + index * width;
    const end = index === binCount - 1 ? max : start + width;
    return {
      start,
      end,
      label: `${formatFieldNumber(field, start)} - ${formatFieldNumber(field, end)}`,
      count: 0
    };
  });
  values.forEach((value) => {
    const index = Math.min(Math.floor((value - min) / width), binCount - 1);
    bins[index].count += 1;
  });
  return bins;
}

function chartOptions(axisLabel, valueField = "") {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: false },
      tooltip: {
        mode: "index",
        intersect: false,
        displayColors: false,
        backgroundColor: "rgba(22, 24, 40, 0.94)",
        titleColor: "#ffffff",
        bodyColor: "#e5e7ef",
        padding: 11,
        cornerRadius: 8,
        ...(valueField ? {
          callbacks: {
            label(context) {
              const datasetLabel = context.dataset?.label ? `${context.dataset.label}: ` : "";
              const value = typeof context.parsed?.y === "number" ? context.parsed.y : Number(context.raw);
              return `${datasetLabel}${formatFieldNumber(valueField, value)}`;
            }
          }
        } : {})
      }
    },
    scales: {
      x: {
        ticks: { maxRotation: 0, minRotation: 0, color: CHART_THEME.text, autoSkip: true, maxTicksLimit: 9 },
        grid: { display: false },
        border: { display: false }
      },
      y: {
        beginAtZero: true,
        title: { display: true, text: axisLabel, color: CHART_THEME.text },
        ticks: {
          color: CHART_THEME.text,
          ...(valueField ? { callback: (value) => formatFieldNumber(valueField, Number(value)) } : {})
        },
        grid: { color: CHART_THEME.grid },
        border: { display: false }
      }
    }
  };
}

function drawEmptyChart(canvas, text) {
  const context = canvas.getContext("2d");
  context.clearRect(0, 0, canvas.width, canvas.height);
  context.save();
  context.font = '13px Inter, "Segoe UI", "Microsoft YaHei", sans-serif';
  context.fillStyle = CHART_THEME.text;
  context.textAlign = "center";
  context.fillText(text, canvas.width / 2, canvas.height / 2);
  context.restore();
}

function resetAnalysisState() {
  Object.keys(state.charts).forEach((id) => destroyChart(id));
  state.charts = {};
  state.rows = [];
  state.fields = [];
  state.profiles = [];
  state.numericStats = [];
  state.categoryStats = [];
  state.duplicateRows = 0;
  state.totalMissing = 0;
  state.parseWarnings = [];
  state.typeOverrides = {};
  state.fieldTypeDrafts = {};
  state.pendingExcel = null;

  dom.sourceName.textContent = "";
  dom.sourceName.removeAttribute("title");
  dom.overviewCards.innerHTML = "";
  dom.insights.innerHTML = "";
  dom.previewTable.innerHTML = "";
  dom.schemaTable.innerHTML = "";
  dom.fieldConfigMessage.textContent = "";
  dom.fieldConfigMessage.className = "field-config-message";
  dom.applyFieldConfig.disabled = true;
  dom.restoreAutoFieldTypes.disabled = true;
  dom.qualitySummary.innerHTML = "";
  dom.qualityTable.innerHTML = "";
  dom.outlierTable.innerHTML = "";
  dom.numericStatsTable.innerHTML = "";
  dom.categoryStats.innerHTML = "";
  dom.excelWorkbookName.textContent = "";
  dom.excelSheetSelect.innerHTML = "";
  dom.sheetSelectionMessage.textContent = "";
  dom.sheetSelectionMessage.className = "sheet-selection-message";
  dom.sheetSelectionPanel.classList.add("hidden");
  dom.v2FieldPrompt.textContent = "";
  dom.v2FieldPrompt.className = "analysis-prompt";
  dom.v2SummaryCards.innerHTML = "";
  dom.v2SumTable.innerHTML = "";
  dom.v2AvgTable.innerHTML = "";
  dom.templatePrompt.textContent = "";
  dom.templatePrompt.className = "analysis-prompt";
  dom.templateSummaryCards.innerHTML = "";
  dom.templateInsights.innerHTML = "";
  dom.templatePrimaryTable.innerHTML = "";
  dom.v2AnalysisResults.classList.add("hidden");
  dom.templateResults.classList.add("hidden");
  dom.templateMappingForm.innerHTML = "";
  dom.results.classList.add("hidden");
}

function destroyChart(id) {
  if (state.charts[id]) {
    state.charts[id].destroy();
    state.charts[id] = null;
  }
}

function countDuplicateRows(rows, fields) {
  const seen = new Set();
  let duplicates = 0;
  rows.forEach((row) => {
    const key = JSON.stringify(fields.map((field) => normalizeValue(row[field])));
    if (seen.has(key)) duplicates += 1;
    else seen.add(key);
  });
  return duplicates;
}

function countValues(field) {
  const counts = new Map();
  state.rows.forEach((row) => {
    if (isMissing(row[field])) return;
    const value = displayValue(row[field]);
    counts.set(value, (counts.get(value) || 0) + 1);
  });
  return counts;
}

function getNumericValues(field) {
  return state.rows
    .map((row) => toNumber(row[field]))
    .filter((value) => value !== null);
}

function getProfilesByType(typeKey) {
  return state.profiles.filter((profile) => profile.typeKey === typeKey);
}

function getActiveProfiles() {
  return state.profiles.filter((profile) => profile.typeKey !== "ignore");
}

function getActiveFields() {
  return getActiveProfiles().map((profile) => profile.field);
}

function getFieldProfile(field) {
  return state.profiles.find((profile) => profile.field === field);
}

function parseFieldDate(field, value) {
  return toDate(value, getFieldProfile(field)?.dateStrategy || {});
}

function getDateGranularity(field) {
  return getFieldProfile(field)?.dateStrategy?.granularity || "day";
}

function getDateTrendKey(field, date, groupByMonth = false) {
  if (getDateGranularity(field) === "year") {
    return String(date.getUTCFullYear());
  }
  if (groupByMonth) {
    return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
  }
  return date.toISOString().slice(0, 10);
}

function setSelectOptions(select, values, placeholder = "", allowEmpty = false) {
  const placeholderOption = placeholder
    ? `<option value="">${escapeHtml(placeholder)}</option>`
    : "";
  const valueOptions = values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("");
  select.innerHTML = values.length || allowEmpty
    ? `${placeholderOption}${valueOptions}`
    : `<option value="">无可用字段</option>`;
  select.disabled = values.length === 0 && !allowEmpty;
}

function missingRateLevel() {
  const denominator = state.rows.length * state.fields.length;
  const rate = denominator ? state.totalMissing / denominator : 0;
  if (rate >= 0.2) return "danger";
  if (rate > 0) return "warning";
  return "success";
}

function emptyTable(message) {
  return `
    <thead><tr><th>提示</th></tr></thead>
    <tbody><tr><td>${escapeHtml(message)}</td></tr></tbody>
  `;
}

function normalizeFieldName(field) {
  return String(field ?? "").replace(/^\uFEFF/, "").trim();
}

function normalizeParsedRow(row, fieldPairs) {
  return fieldPairs.reduce((normalized, pair) => {
    normalized[pair.clean] = Object.prototype.hasOwnProperty.call(row, pair.raw)
      ? row[pair.raw]
      : row[pair.clean];
    return normalized;
  }, {});
}

function isMissing(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

function normalizeValue(value) {
  return isMissing(value) ? "" : String(value).trim();
}

function displayValue(value) {
  return isMissing(value) ? "" : String(value).trim();
}

function isIdFieldName(field) {
  const name = String(field || "").trim();
  return /(^|[_\-\s])(id|uuid|guid|code|key|number|no)([_\-\s]|$)/i.test(name)
    || /(id|uuid|guid|code|key|number|no)$/i.test(name)
    || /[a-z](Id|ID)$/.test(name)
    || /编号|编码|代码|订单号|账号|工号|主键|用户ID|事件ID|会话ID/i.test(name)
    || /(?:学号|学籍号|考生号|准考证号|身份证号|身份证号码|证件号|手机号|手机号码|流水号|序列号|设备号|会员号|合同号|票据号|批次号|档案号)(?:$|[_\-\s（(])/i.test(name);
}

function isYearFieldName(field) {
  const name = String(field || "").trim();
  return /(年份|年度|学年|财年|会计年度|统计年度|统计年|报告年|出生年|入学年|毕业年)$/i.test(name)
    || /(^|[_\-\s])(year|fiscal_year|calendar_year|school_year|yr|fy)([_\-\s]|$)/i.test(name);
}

function buildDateStrategy(field, values) {
  let dmyEvidence = 0;
  let mdyEvidence = 0;
  let yearOnlyCount = 0;

  values.forEach((value) => {
    if (isDateObject(value)) return;
    const raw = String(value).trim();
    const yearText = raw.match(/^(\d{4})\s*年$/);
    if (isPlausibleYear(yearText ? yearText[1] : raw)) yearOnlyCount += 1;
    const match = matchDayOrMonthFirstDate(raw);
    if (!match) return;
    const first = Number(match[1]);
    const second = Number(match[3]);
    const year = Number(match[4]);
    if (first > 12 && second <= 12 && buildUtcDate(year, second, first)) dmyEvidence += 1;
    if (second > 12 && first <= 12 && buildUtcDate(year, first, second)) mdyEvidence += 1;
  });

  let order = "none";
  if (dmyEvidence && mdyEvidence) order = "conflict";
  else if (dmyEvidence) order = "dmy";
  else if (mdyEvidence) order = "mdy";

  const allowYearOnly = isYearFieldName(field)
    && values.length > 0
    && yearOnlyCount / values.length >= 0.8;

  return {
    order,
    allowYearOnly,
    granularity: allowYearOnly ? "year" : "day",
    dmyEvidence,
    mdyEvidence
  };
}

function isPlausibleYear(value) {
  if (!/^\d{4}$/.test(value)) return false;
  const year = Number(value);
  return year >= 1900 && year <= 2200;
}

function isDateObject(value) {
  return Object.prototype.toString.call(value) === "[object Date]"
    && typeof value.getTime === "function";
}

function matchYearFirstDate(raw) {
  return raw.match(new RegExp(`^(\\d{4})([-/.])(\\d{1,2})\\2(\\d{1,2})${DATE_TIME_SUFFIX_PATTERN}$`));
}

function matchDayOrMonthFirstDate(raw) {
  return raw.match(new RegExp(`^(\\d{1,2})([-/.])(\\d{1,2})\\2(\\d{4})${DATE_TIME_SUFFIX_PATTERN}$`));
}

function matchChineseDate(raw) {
  return raw.match(new RegExp(`^(\\d{4})\\s*年\\s*(\\d{1,2})\\s*月\\s*(\\d{1,2})\\s*日?${DATE_TIME_SUFFIX_PATTERN}$`));
}

function detectNumericFormat(values) {
  const convertibleValues = values.filter((value) => toNumber(value) !== null);
  if (!convertibleValues.length) return "number";
  return convertibleValues.every(hasPercentToken) ? "percent" : "number";
}

function hasPercentToken(value) {
  const raw = String(value).trim();
  const unsigned = /^\(.*\)$/.test(raw) ? raw.slice(1, -1).trim() : raw;
  return /%$/.test(unsigned);
}

function toNumber(value) {
  if (isMissing(value)) return null;
  const raw = String(value).trim();
  if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(raw)) return null;
  const isAccountingNegative = /^\(.*\)$/.test(raw);
  const unsigned = isAccountingNegative ? raw.slice(1, -1) : raw;
  const isPercent = /%$/.test(unsigned);
  const cleaned = unsigned.replace(/[$￥¥€£,\s]/g, "").replace(/%$/, "");
  if (!/^[-+]?(?:\d+\.?\d*|\.\d+)$/.test(cleaned)) return null;
  let parsed = Number(cleaned);
  if (!Number.isFinite(parsed)) return null;
  if (isAccountingNegative) parsed = -Math.abs(parsed);
  if (isPercent) parsed /= 100;
  return parsed;
}

function toDate(value, strategy = {}) {
  if (isMissing(value)) return null;
  if (isDateObject(value)) {
    if (Number.isNaN(value.getTime())) return null;
    const localMidnight = value.getHours() === 0
      && value.getMinutes() === 0
      && value.getSeconds() === 0
      && value.getMilliseconds() === 0;
    const utcMidnight = value.getUTCHours() === 0
      && value.getUTCMinutes() === 0
      && value.getUTCSeconds() === 0
      && value.getUTCMilliseconds() === 0;
    const useUtcParts = !localMidnight && utcMidnight;
    return buildUtcDate(
      useUtcParts ? value.getUTCFullYear() : value.getFullYear(),
      (useUtcParts ? value.getUTCMonth() : value.getMonth()) + 1,
      useUtcParts ? value.getUTCDate() : value.getDate()
    );
  }

  const raw = String(value).trim();
  const yearText = raw.match(/^(\d{4})\s*年$/);
  const yearOnlyValue = yearText ? yearText[1] : raw;
  if (strategy.allowYearOnly && isPlausibleYear(yearOnlyValue)) {
    return buildUtcDate(Number(yearOnlyValue), 1, 1);
  }
  if (/^[-+]?\d+(\.\d+)?$/.test(raw)) return null;
  const chineseDate = matchChineseDate(raw);
  if (chineseDate) {
    return buildUtcDate(Number(chineseDate[1]), Number(chineseDate[2]), Number(chineseDate[3]));
  }

  const yearFirst = matchYearFirstDate(raw);
  if (yearFirst) {
    return buildUtcDate(Number(yearFirst[1]), Number(yearFirst[3]), Number(yearFirst[4]));
  }

  const dayOrMonthFirst = matchDayOrMonthFirstDate(raw);
  if (dayOrMonthFirst) {
    const first = Number(dayOrMonthFirst[1]);
    const second = Number(dayOrMonthFirst[3]);
    const year = Number(dayOrMonthFirst[4]);
    if (first > 12) return buildUtcDate(year, second, first);
    if (second > 12) return buildUtcDate(year, first, second);
    if (first === second) return buildUtcDate(year, first, second);
    if (strategy.order === "dmy") return buildUtcDate(year, second, first);
    if (strategy.order === "mdy") return buildUtcDate(year, first, second);
    return null;
  }

  if (!/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/i.test(raw)) return null;
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return null;
  return buildUtcDate(parsed.getUTCFullYear(), parsed.getUTCMonth() + 1, parsed.getUTCDate());
}

function buildUtcDate(year, month, day) {
  if (year < 1000 || month < 1 || month > 12 || day < 1 || day > 31) return null;
  const parsed = new Date(Date.UTC(year, month - 1, day));
  if (parsed.getUTCFullYear() !== year || parsed.getUTCMonth() !== month - 1 || parsed.getUTCDate() !== day) {
    return null;
  }
  return parsed;
}

function mean(values) {
  if (!values.length) return 0;
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function median(values) {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
}

function standardDeviation(values, avg = mean(values)) {
  if (!values.length) return 0;
  const variance = values.reduce((sum, value) => sum + Math.pow(value - avg, 2), 0) / values.length;
  return Math.sqrt(variance);
}

function quantile(sortedValues, q) {
  if (!sortedValues.length) return 0;
  const position = (sortedValues.length - 1) * q;
  const base = Math.floor(position);
  const rest = position - base;
  if (sortedValues[base + 1] !== undefined) {
    return sortedValues[base] + rest * (sortedValues[base + 1] - sortedValues[base]);
  }
  return sortedValues[base];
}

function formatNumber(value) {
  if (!Number.isFinite(value)) return "-";
  return new Intl.NumberFormat("zh-CN", {
    maximumFractionDigits: 2
  }).format(value);
}

function formatInteger(value) {
  return new Intl.NumberFormat("zh-CN", {
    maximumFractionDigits: 0
  }).format(value || 0);
}

function formatPercent(value) {
  if (!Number.isFinite(value)) return "0%";
  return `${(value * 100).toFixed(1)}%`;
}

function getNumericFormat(field) {
  return getFieldProfile(field)?.numericFormat || "number";
}

function formatFieldNumber(field, value) {
  if (!Number.isFinite(value)) return "-";
  if (getNumericFormat(field) === "percent") {
    return new Intl.NumberFormat("zh-CN", {
      style: "percent",
      maximumFractionDigits: 2
    }).format(value);
  }
  return formatNumber(value);
}

function numericAxisLabel(field, fallbackLabel) {
  return getNumericFormat(field) === "percent"
    ? `${fallbackLabel}（百分比）`
    : fallbackLabel;
}

async function exportHtmlReport() {
  if (!state.rows.length) {
    showError("请先上传数据并生成分析报告。");
    return;
  }
  if (!ensureFieldConfigurationApplied("导出 HTML 报告")) return;

  const button = dom.exportHtmlReport;
  button.disabled = true;
  const originalLabel = button.textContent;
  button.textContent = "正在导出...";
  try {
    const report = dom.results.cloneNode(true);
    report.classList.remove("hidden");
    report.querySelectorAll(".report-actions, .nav-new-analysis, .field-config-actions").forEach((element) => element.remove());

    const originalCanvases = Array.from(dom.results.querySelectorAll("canvas"));
    const clonedCanvases = Array.from(report.querySelectorAll("canvas"));
    clonedCanvases.forEach((canvas, index) => {
      const image = document.createElement("img");
      image.className = "exported-chart-image";
      image.alt = canvas.getAttribute("aria-label") || "分析图表";
      try {
        image.src = originalCanvases[index].toDataURL("image/png");
      } catch {
        image.alt = "图表无法嵌入导出文件";
      }
      canvas.replaceWith(image);
    });

    report.querySelectorAll("select").forEach((select) => {
      const value = document.createElement("span");
      value.className = "exported-select-value";
      value.textContent = select.options[select.selectedIndex]?.text || "未选择";
      select.replaceWith(value);
    });

    let stylesheet = "";
    try {
      const response = await fetch("style.css");
      if (response.ok) stylesheet = await response.text();
    } catch {
      stylesheet = "";
    }

    const sourceLabel = dom.sourceName.textContent.replace(/^已载入\s*·\s*/, "") || "data-report";
    const safeTitle = escapeHtml(`Smart CSV Analyzer - ${sourceLabel}`);
    const html = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${safeTitle}</title>
  <style>${stylesheet}
    body { padding: 24px; background: #f5f6fa; }
    .results { display: block; max-width: 1240px; margin: 0 auto; }
    .results-nav { position: static; }
    .exported-chart-image { display: block; width: 100%; height: auto; }
    .exported-select-value { font-weight: 650; }
  </style>
</head>
<body>
  ${report.outerHTML}
  <footer class="app-footer"><p>导出时间：${escapeHtml(new Date().toLocaleString("zh-CN"))}</p></footer>
</body>
</html>`;
    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${sanitizeFileName(sourceLabel)}-分析报告.html`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch (error) {
    setStatus(`报告导出失败：${error.message}`, "error");
  } finally {
    button.disabled = false;
    button.textContent = originalLabel;
  }
}

function sanitizeFileName(value) {
  return String(value || "data-report")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80) || "data-report";
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function setStatus(message, tone = "") {
  dom.statusPanel.className = `status-panel${tone ? ` ${tone}` : ""}`;
  dom.statusPanel.textContent = message;
}

function showError(message, importId = state.activeImportId) {
  if (!isCurrentImport(importId)) return;
  resetAnalysisState();
  dom.statusPanel.className = "status-panel error";
  dom.statusPanel.textContent = message;
  dom.results.classList.add("hidden");
}

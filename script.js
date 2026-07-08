const state = {
  rows: [],
  fields: [],
  profiles: [],
  numericStats: [],
  categoryStats: [],
  duplicateRows: 0,
  totalMissing: 0,
  charts: {}
};

const dom = {
  fileInput: document.getElementById("csvFile"),
  fileEncoding: document.getElementById("fileEncoding"),
  loadSample: document.getElementById("loadSample"),
  dropZone: document.getElementById("dropZone"),
  statusPanel: document.getElementById("statusPanel"),
  results: document.getElementById("results"),
  sourceName: document.getElementById("sourceName"),
  overviewCards: document.getElementById("overviewCards"),
  insights: document.getElementById("insights"),
  previewTable: document.getElementById("previewTable"),
  schemaTable: document.getElementById("schemaTable"),
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
  templateTrendChartTitle: document.getElementById("templateTrendChartTitle")
};

const ENCODING_LABELS = {
  auto: "自动识别",
  "utf-8": "UTF-8",
  gbk: "GBK",
  gb18030: "GB18030"
};

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

dom.fileInput.addEventListener("change", (event) => {
  const file = event.target.files[0];
  if (file) parseCsvFile(file);
  event.target.value = "";
});

dom.loadSample.addEventListener("click", async () => {
  try {
    resetAnalysisState();
    setStatus("正在加载 sample-data.csv ...");
    const response = await fetch("sample-data.csv");
    if (!response.ok) throw new Error("示例数据加载失败");
    const text = await response.text();
    parseCsvText(text, "sample-data.csv");
  } catch (error) {
    showError(error.message);
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
  if (!file.name.toLowerCase().endsWith(".csv")) {
    showError("请上传 CSV 文件。");
    return;
  }
  parseCsvFile(file);
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

[dom.v2MetricField, dom.v2GroupField, dom.v2DateField].forEach((control) => {
  control.addEventListener("change", () => {
    if (state.rows.length) {
      dom.v2FieldPrompt.className = "analysis-prompt";
      dom.v2FieldPrompt.textContent = "字段已更新，点击“重新选择字段并分析”生成新的结果。";
    }
  });
});

async function parseCsvFile(file) {
  resetAnalysisState();
  try {
    const selectedEncoding = dom.fileEncoding.value || "auto";
    setStatus(`正在读取 ${file.name} ...`);
    const buffer = await readFileAsArrayBuffer(file);
    const decoded = decodeCsvBuffer(buffer, selectedEncoding);
    setStatus(`正在解析 ${file.name}（${decoded.label}）...`);
    parseCsvText(decoded.text, file.name);
  } catch (error) {
    showError(error.message);
  }
}

function parseCsvText(text, sourceName) {
  resetAnalysisState();
  Papa.parse(text, {
    header: true,
    skipEmptyLines: true,
    complete: (results) => handleParsedData(results, sourceName),
    error: (error) => showError(error.message)
  });
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

function handleParsedData(results, sourceName) {
  resetAnalysisState();
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
    showError("CSV 中没有可分析的数据。请确认文件包含表头和至少一行数据。");
    return;
  }

  state.rows = rows;
  state.fields = fields;
  state.profiles = fields.map((field) => buildColumnProfile(field, rows));
  state.duplicateRows = countDuplicateRows(rows, fields);
  state.totalMissing = state.profiles.reduce((sum, profile) => sum + profile.missingCount, 0);
  state.numericStats = buildNumericStats();
  state.categoryStats = buildCategoryStats();

  dom.sourceName.textContent = `当前文件：${sourceName}`;
  dom.results.classList.remove("hidden");
  dom.statusPanel.className = "status-panel hidden";

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
  const dateValues = nonMissingValues.map(toDate).filter(Boolean);
  const numericRatio = nonMissingCount ? numericValues.length / nonMissingCount : 0;
  const dateRatio = nonMissingCount ? dateValues.length / nonMissingCount : 0;
  const lowerName = field.toLowerCase();
  const idName = /(^|[_\-\s])(id|uuid|guid|code|key|number|no)([_\-\s]|$)/i.test(lowerName)
    || /编号|编码|代码|订单号|账号|工号|主键/.test(field);

  let type = "分类字段";
  let typeKey = "category";

  if (idName || (uniqueRatio >= 0.98 && nonMissingCount >= 12 && numericRatio < 0.7 && dateRatio < 0.7)) {
    type = "ID 字段";
    typeKey = "id";
  } else if (dateRatio >= 0.8) {
    type = "日期字段";
    typeKey = "date";
  } else if (numericRatio >= 0.8) {
    type = "数值字段";
    typeKey = "numeric";
  }

  return {
    field,
    type,
    typeKey,
    total,
    nonMissingCount,
    missingCount: total - nonMissingCount,
    missingRate: total ? (total - nonMissingCount) / total : 0,
    uniqueCount,
    uniqueRatio,
    numericRatio,
    dateRatio
  };
}

function buildNumericStats() {
  return getProfilesByType("numeric").map((profile) => {
    const values = getNumericValues(profile.field);
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

    return {
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
    };
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
  const numericCount = getProfilesByType("numeric").length;
  const categoryCount = getProfilesByType("category").length;
  const dateCount = getProfilesByType("date").length;
  const cards = [
    ["总行数", state.rows.length],
    ["总字段数", state.fields.length],
    ["数值字段数量", numericCount],
    ["分类字段数量", categoryCount],
    ["日期字段数量", dateCount],
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
  const missingRate = state.rows.length && state.fields.length
    ? state.totalMissing / (state.rows.length * state.fields.length)
    : 0;
  const topNumeric = state.numericStats
    .slice()
    .sort((a, b) => Math.abs(b.max - b.min) - Math.abs(a.max - a.min))[0];
  const firstCategory = state.categoryStats.find((item) => item.top.length);
  const dateProfile = getProfilesByType("date")[0];

  insights.push(`数据集包含 ${formatInteger(state.rows.length)} 行、${formatInteger(state.fields.length)} 个字段，其中数值字段 ${getProfilesByType("numeric").length} 个，分类字段 ${getProfilesByType("category").length} 个，日期字段 ${getProfilesByType("date").length} 个。`);
  insights.push(`共发现 ${formatInteger(state.totalMissing)} 个缺失值，整体缺失率为 ${formatPercent(missingRate)}；重复行数量为 ${formatInteger(state.duplicateRows)}。`);

  if (topNumeric) {
    insights.push(`数值字段「${topNumeric.field}」的平均值为 ${formatNumber(topNumeric.mean)}，最低值为 ${formatNumber(topNumeric.min)}，最高值为 ${formatNumber(topNumeric.max)}。`);
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
        <th>字段</th>
        <th>识别类型</th>
        <th>非空数量</th>
        <th>唯一值数量</th>
        <th>唯一率</th>
        <th>数值匹配率</th>
        <th>日期匹配率</th>
      </tr>
    </thead>
    <tbody>
      ${state.profiles.map((profile) => `
        <tr>
          <td>${escapeHtml(profile.field)}</td>
          <td><span class="type-badge ${profile.typeKey}">${escapeHtml(profile.type)}</span></td>
          <td>${formatInteger(profile.nonMissingCount)}</td>
          <td>${formatInteger(profile.uniqueCount)}</td>
          <td>${formatPercent(profile.uniqueRatio)}</td>
          <td>${formatPercent(profile.numericRatio)}</td>
          <td>${formatPercent(profile.dateRatio)}</td>
        </tr>
      `).join("")}
    </tbody>
  `;
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
          <td>${formatNumber(stat.lowerBound)}</td>
          <td>${formatNumber(stat.upperBound)}</td>
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
          <td>${formatNumber(stat.mean)}</td>
          <td>${formatNumber(stat.median)}</td>
          <td>${formatNumber(stat.max)}</td>
          <td>${formatNumber(stat.min)}</td>
          <td>${formatNumber(stat.std)}</td>
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
  const numericProfiles = getProfilesByType("numeric");
  const categoryProfiles = getProfilesByType("category");
  const dateProfiles = getProfilesByType("date");

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
  const bins = buildHistogram(values, 10);
  state.charts.numericDistributionChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: bins.map((bin) => bin.label),
      datasets: [{
        label: `${field} 分布`,
        data: bins.map((bin) => bin.count),
        backgroundColor: "#1f6feb"
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
        backgroundColor: "#17805d"
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

  const numericField = dom.metricField.value || getProfilesByType("numeric")[0]?.field;
  const trend = buildDateTrend(field, numericField);
  state.charts.dateTrendChart = new Chart(canvas, {
    type: "line",
    data: {
      labels: trend.labels,
      datasets: [{
        label: trend.label,
        data: trend.values,
        borderColor: "#1f6feb",
        backgroundColor: "rgba(31, 111, 235, 0.14)",
        tension: 0.25,
        fill: true
      }]
    },
    options: chartOptions(trend.axisLabel)
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
        backgroundColor: "#b56a00"
      }]
    },
    options: chartOptions(methodLabel)
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
          <td>${formatNumber(row[valueKey])}</td>
        </tr>
      `).join("")}
    </tbody>
  `;
}

function renderV2TopGroupChart(metricField, groupField, grouped) {
  const canvas = document.getElementById("v2TopGroupChart");
  const topGroups = [...grouped].sort((a, b) => b.sum - a.sum).slice(0, 10);
  state.charts.v2TopGroupChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: topGroups.map((item) => item.name),
      datasets: [{
        label: `${groupField} Top 10（${metricField} 求和）`,
        data: topGroups.map((item) => item.sum),
        backgroundColor: "#1f6feb"
      }]
    },
    options: chartOptions("求和值")
  });
}

function renderV2TrendChart(metricField, dateField) {
  const canvas = document.getElementById("v2SelectedTrendChart");
  const trend = buildMetricDateTrend(dateField, metricField);
  if (!trend.labels.length) {
    return drawEmptyChart(canvas, "所选时间字段没有可用的趋势数据");
  }

  state.charts.v2SelectedTrendChart = new Chart(canvas, {
    type: "line",
    data: {
      labels: trend.labels,
      datasets: [{
        label: `${metricField} 按时间求和趋势`,
        data: trend.values,
        borderColor: "#17805d",
        backgroundColor: "rgba(23, 128, 93, 0.14)",
        tension: 0.25,
        fill: true
      }]
    },
    options: chartOptions("求和值")
  });
}

function buildMetricDateTrend(dateField, metricField) {
  const items = [];
  state.rows.forEach((row) => {
    const date = toDate(row[dateField]);
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
    const key = groupByMonth
      ? `${item.date.getFullYear()}-${String(item.date.getMonth() + 1).padStart(2, "0")}`
      : item.date.toISOString().slice(0, 10);
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
  const options = state.fields.map((field) => {
    const profile = state.profiles.find((item) => item.field === field);
    const typeLabel = profile ? profile.type : "未识别类型";
    return `<option value="${escapeHtml(field)}">${escapeHtml(field)}（${escapeHtml(typeLabel)}）</option>`;
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
        borderColor: chartConfig.color || "#1f6feb",
        backgroundColor: chartConfig.background || chartConfig.color || "#1f6feb",
        tension: 0.25,
        fill: chartConfig.type === "line"
      }]
    },
    options: chartOptions(chartConfig.axisLabel || "值")
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
    `「${mappings.metric}」的平均值为 ${formatNumber(stats.mean)}，中位数为 ${formatNumber(stats.median)}，最小值为 ${formatNumber(stats.min)}，最大值为 ${formatNumber(stats.max)}。`,
    topBySum ? `按「${mappings.group}」分组求和最高的是「${topBySum.name}」，求和值为 ${formatNumber(topBySum.sum)}。` : "所选字段没有形成可比较的分组结果。",
    topByAvg ? `按「${mappings.group}」分组平均值最高的是「${topByAvg.name}」，平均值为 ${formatNumber(topByAvg.avg)}。该结果只描述数据差异，不说明原因。` : "所选字段没有形成平均值比较结果。"
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
    mainChart: barChartConfig(`按 ${mappings.group} 的 ${mappings.metric} 求和 Top 10`, grouped.slice(0, 10).map((item) => item.name), grouped.slice(0, 10).map((item) => item.sum), `${mappings.metric} 求和`, "求和值", "#1f6feb"),
    secondaryChart: histogramChartConfig(`${mappings.metric} 分布`, values, "记录数"),
    trendChart: trend ? lineChartConfig(`${mappings.metric} 时间趋势`, trend.labels, trend.values, `${mappings.metric} 求和`, "求和值") : null
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
    categoryTop ? `按品类字段「${mappings.category}」比较，${metricRows ? "求和" : "记录数"}最高的分组是「${categoryTop.name}」，值为 ${formatNumber(categoryTop.sum ?? categoryTop.count)}。` : "品类字段没有可统计的有效分组。",
    regionTop ? `按地区字段「${mappings.region}」比较，${metricRows ? "求和" : "记录数"}最高的分组是「${regionTop.name}」，值为 ${formatNumber(regionTop.sum ?? regionTop.count)}。该结果只反映 CSV 中的数据事实。` : "地区字段没有可统计的有效分组。"
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
      ? numericGroupTable(mappings.category, metricLabel, categoryRows)
      : frequencyTable(mappings.category, categoryRows),
    mainChart: barChartConfig(`按 ${mappings.category} 的 ${metricLabel} Top 10`, topLabels(categoryRows), topValues(categoryRows, metricRows ? "sum" : "count"), metricRows ? `${metricLabel} 求和` : "记录数", metricRows ? "求和值" : "记录数", "#1f6feb"),
    secondaryChart: barChartConfig(`按 ${mappings.region} 的 ${metricLabel} Top 10`, topLabels(regionRows), topValues(regionRows, metricRows ? "sum" : "count"), metricRows ? `${metricLabel} 求和` : "记录数", metricRows ? "求和值" : "记录数", "#17805d"),
    trendChart: trend ? lineChartConfig(`${metricLabel} 时间趋势`, trend.labels, trend.values, metricRows ? `${metricLabel} 求和` : "记录数", metricRows ? "求和值" : "记录数") : null
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
    `成绩平均值为 ${formatNumber(stats.mean)}，中位数为 ${formatNumber(stats.median)}，最低值为 ${formatNumber(stats.min)}，最高值为 ${formatNumber(stats.max)}。`,
    topAvg && lowAvg ? `按「${dimension}」分组时，平均值最高的是「${topAvg.name}」(${formatNumber(topAvg.avg)})，平均值最低的是「${lowAvg.name}」(${formatNumber(lowAvg.avg)})。这里只描述字段结果，不评价教学或个体原因。` : "未选择班级或科目字段，因此未生成分组平均值比较。"
  ];
  if (trend) insights.push(buildTrendFactText(mappings.score, mappings.examDate, trend, "平均值"));

  return {
    prompt: "已生成成绩数据模板分析。",
    cards: [
      ["模板", "成绩数据分析"],
      ["成绩字段", mappings.score],
      ["有效成绩", formatInteger(values.length)],
      ["平均值", formatNumber(stats.mean)]
    ],
    insights,
    table: dimension ? numericGroupTable(dimension, mappings.score, grouped) : simpleStatsTable(mappings.score, stats, values.length),
    mainChart: dimension
      ? barChartConfig(`按 ${dimension} 的成绩平均值 Top 10`, topLabels(grouped, "avg"), topValues(grouped, "avg"), "成绩平均值", "平均值", "#1f6feb")
      : histogramChartConfig(`${mappings.score} 分布`, values, "记录数"),
    secondaryChart: histogramChartConfig(`${mappings.score} 分布`, values, "记录数"),
    trendChart: trend ? lineChartConfig(`${mappings.score} 时间趋势`, trend.labels, trend.values, "成绩平均值", "平均值") : null
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
    `价格平均值为 ${formatNumber(stats.mean)}，中位数为 ${formatNumber(stats.median)}，最低值为 ${formatNumber(stats.min)}，最高值为 ${formatNumber(stats.max)}。`,
    topAvg ? `按「${primaryDimension}」比较，平均价格最高的分组是「${topAvg.name}」，平均值为 ${formatNumber(topAvg.avg)}。该结果只说明样本中的价格差异，不代表市场整体水平。` : "未选择商品、平台、城市或成色字段，因此仅生成价格整体统计。"
  ];
  if (trend) insights.push(buildTrendFactText(mappings.price, mappings.date, trend, "平均值"));

  return {
    prompt: "已生成二手商品价格模板分析。",
    cards: [
      ["模板", "二手商品价格分析"],
      ["价格字段", mappings.price],
      ["比较维度", primaryDimension || "未选择"],
      ["价格中位数", formatNumber(stats.median)]
    ],
    insights,
    table: primaryDimension ? numericGroupTable(primaryDimension, mappings.price, grouped) : simpleStatsTable(mappings.price, stats, values.length),
    mainChart: primaryDimension
      ? barChartConfig(`按 ${primaryDimension} 的平均价格 Top 10`, topLabels(grouped, "avg"), topValues(grouped, "avg"), "平均价格", "平均值", "#1f6feb")
      : histogramChartConfig(`${mappings.price} 分布`, values, "记录数"),
    secondaryChart: histogramChartConfig(`${mappings.price} 分布`, values, "记录数"),
    trendChart: trend ? lineChartConfig(`${mappings.price} 时间趋势`, trend.labels, trend.values, "平均价格", "平均值") : null
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
      ? `数值字段「${numericSummaries[0].field}」的平均值为 ${formatNumber(numericSummaries[0].stats.mean)}，中位数为 ${formatNumber(numericSummaries[0].stats.median)}。`
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
      ? barChartConfig(`${primaryCategory} Top 10`, topLabels(primaryRows), topValues(primaryRows, "count"), "频数", "频数", "#1f6feb")
      : histogramChartConfig(numericSummaries[0]?.field || "数值字段", chartValues, "记录数"),
    secondaryChart: numericSummaries[0]
      ? histogramChartConfig(`${numericSummaries[0].field} 分布`, chartValues, "记录数")
      : barChartConfig("分类字段 Top 10", topLabels(primaryRows), topValues(primaryRows, "count"), "频数", "频数", "#17805d"),
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
    mainChart: barChartConfig(`${mappings.event} Top 10`, topLabels(eventRows), topValues(eventRows, "count"), "事件数", "记录数", "#1f6feb"),
    secondaryChart: barChartConfig(`${secondaryField || mappings.userId} Top 10`, topLabels(secondaryRows), topValues(secondaryRows, "count"), "记录数", "记录数", "#17805d"),
    trendChart: trend ? lineChartConfig("事件记录时间趋势", trend.labels, trend.values, "记录数", "记录数") : null
  };
}

function buildSalesMetricRows(mappings) {
  let label = "";
  if (mappings.salesAmount) label = mappings.salesAmount;
  else if (mappings.unitPrice && mappings.quantity) label = `${mappings.unitPrice} × ${mappings.quantity}`;
  else if (mappings.unitPrice) label = mappings.unitPrice;
  else if (mappings.quantity) label = mappings.quantity;
  else return null;

  const rows = state.rows
    .map((row) => ({ row, value: getSalesMetricValue(row, mappings) }))
    .filter((item) => item.value !== null);
  rows.label = label;
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
  const groups = new Map();
  state.rows.forEach((row) => {
    const date = toDate(row[dateField]);
    if (!date) return;
    const key = date.toISOString().slice(0, 10);
    if (!groups.has(key)) groups.set(key, []);
    if (aggregate === "count") groups.get(key).push(1);
    else {
      const value = valueGetter(row);
      if (value !== null) groups.get(key).push(value);
    }
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

function numericGroupTable(groupField, metricField, rows) {
  return {
    headers: [groupField, "记录数", `${metricField} 求和`, `${metricField} 平均值`, "最小值", "最大值"],
    rows: rows.map((row) => [
      row.name,
      formatInteger(row.count),
      formatNumber(row.sum),
      formatNumber(row.avg),
      formatNumber(row.min),
      formatNumber(row.max)
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
    rows: [[field, formatInteger(count), formatNumber(stats.mean), formatNumber(stats.median), formatNumber(stats.min), formatNumber(stats.max), formatNumber(stats.std)]]
  };
}

function surveyTable(numericSummaries, categoricalFields) {
  const rows = [];
  numericSummaries.forEach((item) => {
    rows.push([item.field, "数值字段", `有效 ${formatInteger(item.count)} 条，平均值 ${formatNumber(item.stats.mean)}，中位数 ${formatNumber(item.stats.median)}`]);
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

function barChartConfig(title, labels, values, label, axisLabel, color) {
  return { type: "bar", title, labels, values, label, axisLabel, color };
}

function lineChartConfig(title, labels, values, label, axisLabel) {
  return {
    type: "line",
    title,
    labels,
    values,
    label,
    axisLabel,
    color: "#17805d",
    background: "rgba(23, 128, 93, 0.14)"
  };
}

function histogramChartConfig(title, values, axisLabel) {
  const bins = buildHistogram(values, 10);
  return {
    type: "bar",
    title,
    labels: bins.map((bin) => bin.label),
    values: bins.map((bin) => bin.count),
    label: title,
    axisLabel,
    color: "#b56a00"
  };
}

function fieldHasNumericData(field) {
  return getNumericValues(field).length > 0;
}

function fieldHasDateData(field) {
  return state.rows.some((row) => Boolean(toDate(row[field])));
}

function countNonMissingValues(field) {
  return state.rows.filter((row) => !isMissing(row[field])).length;
}

function buildTrendInsight(dateField) {
  const numericField = getProfilesByType("numeric")[0]?.field;
  const trend = buildDateTrend(dateField, numericField);
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
    const date = toDate(row[dateField]);
    if (!date) return;
    const value = numericField ? toNumber(row[numericField]) : null;
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
    const key = groupByMonth
      ? `${item.date.getFullYear()}-${String(item.date.getMonth() + 1).padStart(2, "0")}`
      : item.date.toISOString().slice(0, 10);
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
    label: hasMetric ? `${numericField} 平均值趋势` : "记录数趋势",
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

function buildHistogram(values, binCount) {
  if (!values.length) return [];
  const min = Math.min(...values);
  const max = Math.max(...values);
  if (min === max) return [{ label: formatNumber(min), count: values.length }];
  const width = (max - min) / binCount;
  const bins = Array.from({ length: binCount }, (_, index) => {
    const start = min + index * width;
    const end = index === binCount - 1 ? max : start + width;
    return {
      start,
      end,
      label: `${formatNumber(start)} - ${formatNumber(end)}`,
      count: 0
    };
  });
  values.forEach((value) => {
    const index = Math.min(Math.floor((value - min) / width), binCount - 1);
    bins[index].count += 1;
  });
  return bins;
}

function chartOptions(axisLabel) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: true, position: "bottom" },
      tooltip: { mode: "index", intersect: false }
    },
    scales: {
      x: {
        ticks: { maxRotation: 45, minRotation: 0 },
        grid: { display: false }
      },
      y: {
        beginAtZero: true,
        title: { display: true, text: axisLabel },
        grid: { color: "rgba(148, 163, 184, 0.24)" }
      }
    }
  };
}

function drawEmptyChart(canvas, text) {
  const context = canvas.getContext("2d");
  context.clearRect(0, 0, canvas.width, canvas.height);
  context.save();
  context.font = "14px Arial";
  context.fillStyle = "#657184";
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

  dom.sourceName.textContent = "";
  dom.overviewCards.innerHTML = "";
  dom.insights.innerHTML = "";
  dom.previewTable.innerHTML = "";
  dom.schemaTable.innerHTML = "";
  dom.qualitySummary.innerHTML = "";
  dom.qualityTable.innerHTML = "";
  dom.outlierTable.innerHTML = "";
  dom.numericStatsTable.innerHTML = "";
  dom.categoryStats.innerHTML = "";
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

function toNumber(value) {
  if (isMissing(value)) return null;
  const raw = String(value).trim();
  if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(raw)) return null;
  const cleaned = raw.replace(/[$￥,\s]/g, "").replace(/%$/, "");
  if (!/^[-+]?(?:\d+\.?\d*|\.\d+)$/.test(cleaned)) return null;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function toDate(value) {
  if (isMissing(value)) return null;
  const raw = String(value).trim();
  if (/^[-+]?\d+(\.\d+)?$/.test(raw)) return null;
  const looksLikeDate = /\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(raw)
    || /\d{1,2}[-/]\d{1,2}[-/]\d{2,4}/.test(raw)
    || /jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/i.test(raw);
  if (!looksLikeDate) return null;
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
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

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function setStatus(message) {
  dom.statusPanel.className = "status-panel";
  dom.statusPanel.textContent = message;
}

function showError(message) {
  resetAnalysisState();
  dom.statusPanel.className = "status-panel error";
  dom.statusPanel.textContent = message;
  dom.results.classList.add("hidden");
}

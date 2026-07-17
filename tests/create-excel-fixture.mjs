import fs from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = fileURLToPath(new URL("./fixtures/", import.meta.url));
await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const salesSheet = workbook.worksheets.add("销售数据");
salesSheet.showGridLines = false;
salesSheet.getRange("A1:D4").values = [
  ["订单ID", "订单日期", "地区", "销售额"],
  ["A-001", new Date("2026-07-15T00:00:00.000Z"), "华东", 1200],
  ["A-002", new Date("2026-07-16T00:00:00.000Z"), "华南", 980],
  ["A-003", new Date("2026-07-17T00:00:00.000Z"), "华东", 1520]
];
salesSheet.getRange("A1:D1").format = {
  fill: "#625BDE",
  font: { bold: true, color: "#FFFFFF" },
  rowHeight: 24
};
salesSheet.getRange("B2:B4").format.numberFormat = "yyyy-mm-dd";
salesSheet.getRange("D2:D4").format.numberFormat = "#,##0";
salesSheet.getRange("A1:D4").format.autofitColumns();
salesSheet.getRange("A1:D4").format.autofitRows();
salesSheet.freezePanes.freezeRows(1);

const userSheet = workbook.worksheets.add("用户数据");
userSheet.showGridLines = false;
userSheet.getRange("A1:D4").values = [
  ["用户ID", "姓名", "城市", "满意度"],
  ["U-001", "张伟", "上海", 4.8],
  ["U-002", "李娜", "广州", 4.2],
  ["U-003", "王芳", "成都", 4.6]
];
userSheet.getRange("A1:D1").format = {
  fill: "#0F9F8F",
  font: { bold: true, color: "#FFFFFF" },
  rowHeight: 24
};
userSheet.getRange("D2:D4").format.numberFormat = "0.0";
userSheet.getRange("A1:D4").format.autofitColumns();
userSheet.getRange("A1:D4").format.autofitRows();
userSheet.freezePanes.freezeRows(1);

workbook.worksheets.add("空工作表");

const noHeaderSheet = workbook.worksheets.add("无表头");
noHeaderSheet.getRange("A1:C3").values = [
  [1001, "张三", 88],
  [1002, "李四", 92],
  [1003, "王五", 85]
];
noHeaderSheet.getRange("A1:C3").format.autofitColumns();

const inspection = await workbook.inspect({
  kind: "sheet,table",
  include: "id,name,values,formulas",
  tableMaxRows: 10,
  tableMaxCols: 6,
  maxChars: 5000
});
console.log(inspection.ndjson);

for (const [sheetName, range, fileName] of [
  ["销售数据", "A1:D4", "excel-import-sales.png"],
  ["用户数据", "A1:D4", "excel-import-users.png"],
  ["无表头", "A1:C3", "excel-import-no-header.png"]
]) {
  const preview = await workbook.render({ sheetName, range, scale: 2, format: "png" });
  await fs.writeFile(`${outputDir}/${fileName}`, new Uint8Array(await preview.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/excel-import-test.xlsx`);

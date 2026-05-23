import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const templatePath = "/Users/zhangrrongjie/Library/Containers/com.tencent.xinWeChat/Data/Documents/xwechat_files/wxid_fj8pjsu4qktn12_6e37/msg/file/2026-05/海鼎商管系统数据字典提纲0422.xlsx";
const sourcePath = "/Users/zhangrrongjie/Library/Containers/com.tencent.xinWeChat/Data/Documents/xwechat_files/wxid_fj8pjsu4qktn12_6e37/msg/file/2026-05/附件9：数据资源目录.xlsx";
const outputDir = "/Users/zhangrrongjie/Documents/New project/outputs/filled_data_dictionary";
const outputPath = `${outputDir}/海鼎商管系统数据字典提纲0422_已填充.xlsx`;

function colToLetter(col) {
  let letter = "";
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - mod) / 26);
  }
  return letter;
}

function rangeAddress(sheetName, startRow, startCol, rowCount, colCount) {
  const endRow = startRow + rowCount - 1;
  const endCol = startCol + colCount - 1;
  return `${sheetName}!${colToLetter(startCol)}${startRow}:${colToLetter(endCol)}${endRow}`;
}

async function readTable(workbook, sheetName, maxRows, maxCols) {
  const sheet = workbook.worksheets.getItem(sheetName);
  return sheet.getRange(rangeAddress(sheetName, 1, 1, maxRows, maxCols).split("!")[1]).values;
}

function nonEmptyRow(row) {
  return row.some((value) => value !== null && value !== undefined && String(value).trim() !== "");
}

function normalizeRows(values) {
  const [headers, ...rows] = values;
  const headerIndex = new Map(headers.map((header, index) => [header, index]));
  return rows
    .filter(nonEmptyRow)
    .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? null])))
    .filter((row) => row["中文名称"] || row["英文名称"] || row["所属数据资源"]);
}

const template = await SpreadsheetFile.importXlsx(await FileBlob.load(templatePath));
const source = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));

const resources = normalizeRows(await readTable(source, "数据资源", 28, 19));
const rawItemsValues = await readTable(source, "数据资源信息项", 560, 9);
const itemsValues = [
  rawItemsValues[0],
  ...rawItemsValues.slice(1).filter((row) => row.slice(1).some((value) => value !== null && value !== undefined && String(value).trim() !== "")),
];
const dictValues = await readTable(source, "字典", 7, 4);

const mainSheet = template.worksheets.getItem("商业运营系统");
const mainRows = resources.map((row, index) => [
  index + 1,
  row["领域"] ?? row["分类代码"] ?? "",
  row["中文名称"] ?? "",
  row["英文名称"] ?? "",
  row["说明"] ?? "",
]);

const maxMainRows = Math.max(80, mainRows.length + 1);
mainSheet.getRange(`A1:E${maxMainRows}`).unmerge();
mainSheet.getRange(`A2:E${maxMainRows}`).clear({ applyTo: "all" });
mainSheet.getRange(`A2:E${mainRows.length + 1}`).values = mainRows;
mainSheet.getRange(`A1:E${mainRows.length + 1}`).format.wrapText = true;
mainSheet.getRange(`A2:E${mainRows.length + 1}`).format.verticalAlignment = "middle";
mainSheet.getRange(`A2:B${mainRows.length + 1}`).format.horizontalAlignment = "center";
mainSheet.getRange("A:A").format.columnWidthPx = 64;
mainSheet.getRange("B:B").format.columnWidthPx = 150;
mainSheet.getRange("C:C").format.columnWidthPx = 210;
mainSheet.getRange("D:D").format.columnWidthPx = 220;
mainSheet.getRange("E:E").format.columnWidthPx = 360;
mainSheet.freezePanes.freezeRows(1);

const itemsSheet = template.worksheets.getOrAdd("数据资源信息项");
itemsSheet.getRange("A1:I600").clear({ applyTo: "all" });
itemsSheet.getRange(`A1:I${itemsValues.length}`).values = itemsValues;
itemsSheet.getRange(`A1:I${itemsValues.length}`).format.wrapText = true;
itemsSheet.getRange("A:A").format.columnWidthPx = 64;
itemsSheet.getRange("B:B").format.columnWidthPx = 170;
itemsSheet.getRange("C:C").format.columnWidthPx = 180;
itemsSheet.getRange("D:D").format.columnWidthPx = 200;
itemsSheet.getRange("E:E").format.columnWidthPx = 260;
itemsSheet.getRange("F:I").format.columnWidthPx = 110;
itemsSheet.freezePanes.freezeRows(1);

const dictSheet = template.worksheets.getOrAdd("字典");
dictSheet.getRange("A1:D50").clear({ applyTo: "all" });
dictSheet.getRange(`A1:D${dictValues.length}`).values = dictValues;
dictSheet.getRange(`A1:D${dictValues.length}`).format.wrapText = true;
dictSheet.getRange("A:D").format.columnWidthPx = 120;

const headerFormat = {
  fill: { color: "#D9EAF7" },
  font: { bold: true, color: "#000000" },
  horizontalAlignment: "center",
  verticalAlignment: "middle",
  wrapText: true,
};
mainSheet.getRange("A1:E1").format = headerFormat;
itemsSheet.getRange("A1:I1").format = headerFormat;
dictSheet.getRange("B4:D4").format = headerFormat;

await fs.mkdir(outputDir, { recursive: true });

const check = await template.inspect({
  kind: "table",
  range: `商业运营系统!A1:E${Math.min(mainRows.length + 1, 12)}`,
  include: "values",
  tableMaxRows: 12,
  tableMaxCols: 5,
});
console.log(check.ndjson);

const errors = await template.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
});
console.log(errors.ndjson);

for (const sheetName of ["商业运营系统", "数据资源信息项", "字典"]) {
  const preview = await template.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  const bytes = new Uint8Array(await preview.arrayBuffer());
  await fs.writeFile(`${outputDir}/preview_${sheetName}.png`, bytes);
}

const output = await SpreadsheetFile.exportXlsx(template);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, resources: resources.length, items: itemsValues.length - 1 }));

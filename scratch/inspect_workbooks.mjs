import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const files = [
  "/Users/zhangrrongjie/Library/Containers/com.tencent.xinWeChat/Data/Documents/xwechat_files/wxid_fj8pjsu4qktn12_6e37/msg/file/2026-05/海鼎商管系统数据字典提纲0422.xlsx",
  "/Users/zhangrrongjie/Library/Containers/com.tencent.xinWeChat/Data/Documents/xwechat_files/wxid_fj8pjsu4qktn12_6e37/msg/file/2026-05/附件9：数据资源目录.xlsx",
];

for (const path of files) {
  console.log(`\nFILE: ${path}`);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
  const sheets = workbook.worksheets.items.map((sheet) => sheet.name);
  console.log(`SHEETS: ${sheets.join(" | ")}`);
  for (const sheetName of sheets) {
    console.log(`\nSHEET: ${sheetName}`);
    for (const range of ["A1:Z20", "A1:AZ12"]) {
      try {
        const info = await workbook.inspect({
          kind: "table",
          range: `${sheetName}!${range}`,
          include: "values,formulas",
          tableMaxRows: 20,
          tableMaxCols: 52,
        });
        console.log(info.ndjson);
        break;
      } catch (error) {
        console.log(`inspect failed ${range}: ${error.message}`);
      }
    }
  }
}

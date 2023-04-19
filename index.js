// https://github.com/exceljs/exceljs/pull/2185

// import Excel from "exceljs";
import Excel from "@nbelyh/exceljs";
const workbook = new Excel.Workbook();
await workbook.xlsx.readFile("template.xlsx");

const worksheet = workbook.getWorksheet(1);
worksheet.addRow(["John", "男", 24], "i+");
worksheet.addRow(["Jane", "女", 22], "i+");

/** @type {import("@nbelyh/exceljs").Table} */
const table  = worksheet.getTables()[0];
table.table.tableRef = "A1:C6"
table.table.autoFilterRef = "A1:C6"
table.commit();

await workbook.xlsx.writeFile("output.xlsx");

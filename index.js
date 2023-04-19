// https://github.com/exceljs/exceljs/pull/2185

// import Excel from "exceljs";
import Excel from "@nbelyh/exceljs";
const workbook = new Excel.Workbook();
await workbook.xlsx.readFile("template.xlsx");

const worksheet = workbook.getWorksheet(1);
worksheet.insertRow(["John", "男", 24], "i+");
worksheet.insertRow(["Jane", "女", 22], "i+");
await workbook.xlsx.writeFile("output.xlsx");



// const d = sheet_to_json(currentSheet);
// console.log("sheets", d);
// sheet_add_json(currentSheet, [
//         { 'A': 'John', 'B': '男', 'C': 24 },
//         { 'A': 'Jane', 'B': '女', 'C': 22 }
//     ], 
//     {
//         origin: 'A4',
//         skipHeader: true,
//         cellStyles: true,
//         sheetStubs: true
//     }
// )

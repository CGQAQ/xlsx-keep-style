// xlsx do not support styles in CE version

import {readFile, writeFileXLSX, utils } from "xlsx";
const {sheet_add_json, sheet_to_json} = utils;

const workbook = readFile("template.xlsx", {cellStyles: true, sheetStubs: true});

const [sheetName] = workbook.SheetNames;
const currentSheet = workbook.Sheets[sheetName];

// const d = sheet_to_json(currentSheet);
// console.log("sheets", d);
sheet_add_json(currentSheet, [
        { 'A': 'John', 'B': '男', 'C': 24 },
        { 'A': 'Jane', 'B': '女', 'C': 22 }
    ], 
    {
        origin: 'A4',
        skipHeader: true,
        cellStyles: true,
        sheetStubs: true
    }
)

// const data = sheet_to_json(currentSheet);
// console.log("sheets", data);

writeFileXLSX(workbook, "output.xlsx", {cellStyles: true, sheetStubs: true})
"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const ExcelJS = require("exceljs");
const settings = require("./settings.config.js");
(async () => {
    const fileName = process.argv[2];
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(fileName);
    const worksheet = workbook.getWorksheet("Export");
    deleteNoShows(worksheet);
    removeUnusedCol(worksheet);
    const subTotal = calcSubTotal(worksheet);
    const tax = calcTax(worksheet);
    insertComAndTax(worksheet, subTotal, tax);
    await workbook.xlsx.writeFile(fileName);
})();
function deleteNoShows(worksheet) {
    worksheet.eachRow(function (row, rowNumber) {
        const item = row.getCell(6);
        const status = row.getCell(12);
        if (typeof item.value === "string" &&
            item.value.includes("No Show") &&
            status.value === "unpaid")
            worksheet.spliceRows(rowNumber, 1);
    });
}
function removeUnusedCol(worksheet) {
    const colToRemove = [1, 1, 2, 2, 3, 5, 5, 5, 8, 8];
    colToRemove.forEach((col) => worksheet.spliceColumns(col, 1));
}
function calcSubTotal(worksheet) {
    const subTotalCol = worksheet.getColumn(5);
    let subTotal = 0;
    subTotalCol.eachCell(function (cell, rowNumber) {
        if (rowNumber > 1 && cell.value) {
            subTotal += +cell.value;
        }
    });
    const comSubTotal = subTotal * settings[process.argv[3]].commission;
    return comSubTotal;
}
function calcTax(worksheet) {
    const subTotalCol = worksheet.getColumn(6);
    let tax = 0;
    subTotalCol.eachCell(function (cell, rowNumber) {
        if (rowNumber > 1 && cell.value) {
            tax += +cell.value;
        }
    });
    const comTax = tax * settings[process.argv[3]].commission;
    return comTax;
}
function insertComAndTax(worksheet, subTotal, tax) {
    worksheet.addRow([]);
    worksheet.addRow(["Commission:", null, null, null, subTotal]);
    worksheet.addRow(["GST Remittance:", null, null, null, tax]);
}
function calcTotal() { }
function addGSTNum() { }

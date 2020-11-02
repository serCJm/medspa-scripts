"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const ExcelJS = require("exceljs");
(async () => {
    const fileName = process.argv[2];
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(fileName);
    const worksheet = workbook.getWorksheet("Export");
    deleteNoShows(worksheet);
    removeUnusedCol(worksheet);
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

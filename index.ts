// a script to produce practitioner statement
// usage: node index.js [filename] [practitioner] [--no-tax - OPTIONAL]

import { Worksheet } from "exceljs";

const ExcelJS = require("exceljs");
const settings = require("./settings.config.js");

const FILENAME = process.argv[2];
const PRACTITIONER = process.argv[3];
const NOGST = process.argv[4];

(async () => {
	const workbook = new ExcelJS.Workbook();
	await workbook.xlsx.readFile(FILENAME);
	const worksheet = workbook.getWorksheet("Export");

	deleteNoShows(worksheet);
	removeUnusedCol(worksheet);
	const subTotal: number = calcSubTotal(worksheet);
	const tax: number = NOGST === "--no-tax" ? 0 : calcTax(worksheet);
	insertComAndTax(worksheet, subTotal, tax);

	addGSTNum(worksheet);

	await workbook.xlsx.writeFile(FILENAME);
})();

function deleteNoShows(worksheet: Worksheet) {
	worksheet.eachRow(function (row, rowNumber) {
		const item = row.getCell(6);
		const status = row.getCell(12);

		if (
			typeof item.value === "string" &&
			item.value.includes("No Show") &&
			status.value === "unpaid"
		)
			worksheet.spliceRows(rowNumber, 1);
	});
}

function removeUnusedCol(worksheet: Worksheet) {
	const colToRemove: number[] = [1, 1, 2, 2, 3, 5, 5, 5, 8, 8];
	colToRemove.forEach((col) => worksheet.spliceColumns(col, 1));
}

function calcSubTotal(worksheet: Worksheet) {
	const subTotalCol = worksheet.getColumn(5);
	let subTotal = 0;
	subTotalCol.eachCell(function (cell, rowNumber) {
		if (rowNumber > 1 && cell.value) {
			subTotal += +cell.value;
		}
	});
	const comSubTotal = subTotal * settings[PRACTITIONER].commission;

	return comSubTotal;
}

function calcTax(worksheet: Worksheet) {
	const subTotalCol = worksheet.getColumn(6);
	let tax = 0;
	subTotalCol.eachCell(function (cell, rowNumber) {
		if (rowNumber > 1 && cell.value) {
			tax += +cell.value;
		}
	});
	const comTax = tax * settings[PRACTITIONER].commission;
	return comTax;
}

function insertComAndTax(worksheet: Worksheet, subTotal: number, tax: number) {
	worksheet.addRow([]);
	worksheet.addRow(["Commission:", null, null, null, subTotal]);
	worksheet.addRow(["GST Remittance:", null, null, null, tax]);
}

function calcTotal() {}

function addGSTNum(worksheet: Worksheet) {
	const gstNumber = settings[PRACTITIONER].gst;
	if (gstNumber) {
		worksheet.addRow([]);
		worksheet.addRow([gstNumber]);
	}
}

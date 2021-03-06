// a script to produce practitioner statement
// usage: node index.js [filename] [practitioner] [--no-tax - OPTIONAL] [--chiro - OPTIONAL]

import { CellValue, Worksheet } from "exceljs";

const ExcelJS = require("exceljs");
const settings = require("./settings.config.js");

const FILENAME = process.argv[2];
const PRACTITIONER = process.argv[3];
const NOGST = process.argv[4];
const CHIRO = process.argv[5];

(async () => {
	try {
		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.readFile(FILENAME);
		const worksheet = workbook.getWorksheet("Export");

		deleteNoShows(worksheet);
		removeUnusedCol(worksheet);

		if (CHIRO !== "--chiro") {
			const subTotal: number = calcSubTotal(worksheet);
			const tax: number = NOGST === "--no-tax" ? 0 : calcTax(worksheet);
			insertComAndTax(worksheet, subTotal, tax);

			calcTotal(worksheet, subTotal, tax);

			addGSTNum(worksheet);
		} else {
			separateICBC(worksheet);
			separateOrthotics(worksheet);
			calcChiroComm(worksheet);
		}

		await workbook.xlsx.writeFile(FILENAME);
	} catch (e) {
		console.log(e);
	}
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
	if (NOGST !== "--no-tax")
		worksheet.addRow(["GST Remittance:", null, null, null, tax]);
}

function calcTotal(worksheet: Worksheet, subTotal: number, tax: number) {
	const total = subTotal + tax;
	worksheet.addRow([]);
	worksheet.addRow(["Payment:", null, null, null, total]);
}

function addGSTNum(worksheet: Worksheet) {
	const gstNumber = settings[PRACTITIONER].gst;
	if (gstNumber) {
		worksheet.addRow([]);
		worksheet.addRow([gstNumber]);
	}
}

function separateICBC(worksheet: Worksheet) {
	type rows = CellValue[] | { [key: string]: CellValue };
	const icbcRows: rows[] = [];
	const rowNumbers: number[] = [];

	worksheet.eachRow(function (row, rowNumber) {
		const description = row.getCell(2);

		if (
			typeof description.value === "string" &&
			description.value.includes("ICBC")
		) {
			icbcRows.push(row.values);
			rowNumbers.push(rowNumber);
		}
	});
	rowNumbers.forEach((num, i) => worksheet.spliceRows(num - i, 1));
	worksheet.addRow([]);
	worksheet.addRow(["Commission:", null, null, null, null]);
	worksheet.addRow([]);
	worksheet.addRows(icbcRows);
}

function separateOrthotics(worksheet: Worksheet) {
	type rows = CellValue[] | { [key: string]: CellValue };
	const orthoRows: rows[] = [];
	const rowNumbers: number[] = [];

	worksheet.eachRow(function (row, rowNumber) {
		const description = row.getCell(2);

		if (
			typeof description.value === "string" &&
			description.value.includes("Orthotics")
		) {
			orthoRows.push(row.values);
			rowNumbers.push(rowNumber);
		}
	});
	rowNumbers.forEach((num, i) => worksheet.spliceRows(num - i, 1));
	worksheet.addRow([]);
	worksheet.addRow(["Commission:", null, null, null, null]);
	worksheet.addRow([]);
	worksheet.addRows(orthoRows);
	worksheet.addRow([]);
	worksheet.addRow(["Commission:", null, null, null, null]);
	worksheet.addRow([]);
}

function calcChiroComm(worksheet: Worksheet) {
	const subTotalCol = worksheet.getColumn(5);
	let subTotal = 0;
	let commCount = 0;
	let total = 0;
	subTotalCol.eachCell(function (cell, rowNumber) {
		if (rowNumber > 1 && cell.value) {
			subTotal += +cell.value;
		}
		if (worksheet.getRow(rowNumber).getCell(1).value === "Commission:") {
			const comSubTotal =
				subTotal * settings[PRACTITIONER].commission[commCount];
			commCount++;
			cell.value = comSubTotal;
			total += comSubTotal;
			subTotal = 0;
		}
	});
	worksheet.addRow([]);
	worksheet.addRow(["Payment:", null, null, null, total]);
}

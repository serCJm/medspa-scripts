import { Cell, Worksheet } from "exceljs";

const ExcelJS = require("exceljs");

(async () => {
	const fileName = process.argv[2];
	const workbook = new ExcelJS.Workbook();
	await workbook.xlsx.readFile(fileName);
	const worksheet = workbook.getWorksheet("Export");

	deleteNoShows(worksheet);

	await workbook.xlsx.writeFile(fileName);
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

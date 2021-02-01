"use strict";
// a script to produce practitioner statement
// usage: node index.js [filename] [practitioner] [--no-tax - OPTIONAL] [--chiro - OPTIONAL]
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
exports.__esModule = true;
var ExcelJS = require("exceljs");
var settings = require("./settings.config.js");
var FILENAME = process.argv[2];
var PRACTITIONER = process.argv[3];
var NOGST = process.argv[4];
var CHIRO = process.argv[5];
(function () { return __awaiter(void 0, void 0, void 0, function () {
    var workbook, worksheet, subTotal, tax, e_1;
    return __generator(this, function (_a) {
        switch (_a.label) {
            case 0:
                _a.trys.push([0, 3, , 4]);
                workbook = new ExcelJS.Workbook();
                return [4 /*yield*/, workbook.xlsx.readFile(FILENAME)];
            case 1:
                _a.sent();
                worksheet = workbook.getWorksheet("Export");
                deleteNoShows(worksheet);
                removeUnusedCol(worksheet);
                if (CHIRO !== "--chiro") {
                    subTotal = calcSubTotal(worksheet);
                    tax = NOGST === "--no-tax" ? 0 : calcTax(worksheet);
                    insertComAndTax(worksheet, subTotal, tax);
                    calcTotal(worksheet, subTotal, tax);
                    addGSTNum(worksheet);
                }
                else {
                    separateICBC(worksheet);
                    separateOrthotics(worksheet);
                    calcChiroComm(worksheet);
                }
                return [4 /*yield*/, workbook.xlsx.writeFile(FILENAME)];
            case 2:
                _a.sent();
                return [3 /*break*/, 4];
            case 3:
                e_1 = _a.sent();
                console.log(e_1);
                return [3 /*break*/, 4];
            case 4: return [2 /*return*/];
        }
    });
}); })();
function deleteNoShows(worksheet) {
    worksheet.eachRow(function (row, rowNumber) {
        var item = row.getCell(6);
        var status = row.getCell(12);
        if (typeof item.value === "string" &&
            item.value.includes("No Show") &&
            status.value === "unpaid")
            worksheet.spliceRows(rowNumber, 1);
    });
}
function removeUnusedCol(worksheet) {
    var colToRemove = [1, 1, 2, 2, 3, 5, 5, 5, 8, 8];
    colToRemove.forEach(function (col) { return worksheet.spliceColumns(col, 1); });
}
function calcSubTotal(worksheet) {
    var subTotalCol = worksheet.getColumn(5);
    var subTotal = 0;
    subTotalCol.eachCell(function (cell, rowNumber) {
        if (rowNumber > 1 && cell.value) {
            subTotal += +cell.value;
        }
    });
    var comSubTotal = subTotal * settings[PRACTITIONER].commission;
    return comSubTotal;
}
function calcTax(worksheet) {
    var subTotalCol = worksheet.getColumn(6);
    var tax = 0;
    subTotalCol.eachCell(function (cell, rowNumber) {
        if (rowNumber > 1 && cell.value) {
            tax += +cell.value;
        }
    });
    var comTax = tax * settings[PRACTITIONER].commission;
    return comTax;
}
function insertComAndTax(worksheet, subTotal, tax) {
    worksheet.addRow([]);
    worksheet.addRow(["Commission:", null, null, null, subTotal]);
    if (NOGST !== "--no-tax")
        worksheet.addRow(["GST Remittance:", null, null, null, tax]);
}
function calcTotal(worksheet, subTotal, tax) {
    var total = subTotal + tax;
    worksheet.addRow([]);
    worksheet.addRow(["Payment:", null, null, null, total]);
}
function addGSTNum(worksheet) {
    var gstNumber = settings[PRACTITIONER].gst;
    if (gstNumber) {
        worksheet.addRow([]);
        worksheet.addRow([gstNumber]);
    }
}
function separateICBC(worksheet) {
    var icbcRows = [];
    var rowNumbers = [];
    worksheet.eachRow(function (row, rowNumber) {
        var description = row.getCell(2);
        if (typeof description.value === "string" &&
            description.value.includes("ICBC")) {
            icbcRows.push(row.values);
            rowNumbers.push(rowNumber);
        }
    });
    rowNumbers.forEach(function (num, i) { return worksheet.spliceRows(num - i, 1); });
    worksheet.addRow([]);
    worksheet.addRow(["Commission:", null, null, null, null]);
    worksheet.addRow([]);
    worksheet.addRows(icbcRows);
}
function separateOrthotics(worksheet) {
    var orthoRows = [];
    var rowNumbers = [];
    worksheet.eachRow(function (row, rowNumber) {
        var description = row.getCell(2);
        if (typeof description.value === "string" &&
            description.value.includes("Orthotics")) {
            orthoRows.push(row.values);
            rowNumbers.push(rowNumber);
        }
    });
    rowNumbers.forEach(function (num, i) { return worksheet.spliceRows(num - i, 1); });
    worksheet.addRow([]);
    worksheet.addRow(["Commission:", null, null, null, null]);
    worksheet.addRow([]);
    worksheet.addRows(orthoRows);
    worksheet.addRow([]);
    worksheet.addRow(["Commission:", null, null, null, null]);
    worksheet.addRow([]);
}
function calcChiroComm(worksheet) {
    var subTotalCol = worksheet.getColumn(5);
    var subTotal = 0;
    var commCount = 0;
    var total = 0;
    subTotalCol.eachCell(function (cell, rowNumber) {
        if (rowNumber > 1 && cell.value) {
            subTotal += +cell.value;
        }
        if (worksheet.getRow(rowNumber).getCell(1).value === "Commission:") {
            var comSubTotal = subTotal * settings[PRACTITIONER].commission[commCount];
            commCount++;
            cell.value = comSubTotal;
            total += comSubTotal;
            subTotal = 0;
        }
    });
    worksheet.addRow([]);
    worksheet.addRow(["Payment:", null, null, null, total]);
}

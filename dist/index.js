"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.SheetQueryBuilder = exports.sheetQuery = void 0;
/**
 * Run new sheet query
 *
 * @param {Spreadsheet} activeSpreadsheet Specific spreadsheet to use, or will use SpreadsheetApp.getActiveSpreadsheet() if undefined\
 * @return {SheetQuery}
 */
function sheetQuery(activeSpreadsheet) {
    return new SheetQueryBuilder(activeSpreadsheet);
}
exports.sheetQuery = sheetQuery;
/**
 * SheetQueryBuilder class - Kind of an ORM for Google Sheets
 */
class SheetQueryBuilder {
    constructor(activeSpreadsheet) {
        this.columnNames = [];
        this._sheetHeadings = [];
        this.activeSpreadsheet = activeSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    }
    select(columnNames) {
        this.columnNames = Array.isArray(columnNames) ? columnNames : [columnNames];
        return this;
    }
    // Array of sheet names or single string sheet
    from(sheetName) {
        this.sheetName = sheetName;
        return this;
    }
    where(fn) {
        this.whereFn = fn;
        return this;
    }
    deleteRows() {
        const rows = this.getRows();
        let i = 0;
        rows.forEach((row) => {
            const deleteRowRange = this._sheet.getRange(row.__meta.row - i, 1, 1, row.__meta.cols);
            deleteRowRange.deleteCells(SpreadsheetApp.Dimension.ROWS);
            i++;
        });
        this.clearCache();
        return this;
    }
    updateRows(updateFn) {
        const rows = this.getRows();
        rows.forEach((row) => {
            const updateRowRange = this._sheet.getRange(row.__meta.row, 1, 1, row.__meta.cols);
            const updatedRow = updateFn(row);
            let arrayValues = [];
            if (updatedRow && updatedRow.__meta) {
                delete updatedRow.__meta;
                arrayValues = Object.values(updatedRow);
            }
            else {
                delete row.__meta;
                arrayValues = Object.values(row);
            }
            updateRowRange.setValues([arrayValues]);
        });
        this.clearCache();
        return this;
    }
    getSheet() {
        if (!this._sheet) {
            this._sheet = this.activeSpreadsheet.getSheetByName(this.sheetName);
        }
        return this._sheet;
    }
    getValues() {
        if (!this._sheetValues) {
            const sheet = this.getSheet();
            const numCols = sheet.getLastColumn();
            const rowValues = [];
            const sheetValues = sheet.getSheetValues(2, 1, sheet.getLastRow(), numCols);
            const numRows = sheetValues.length;
            const headings = this.getHeadings();
            for (let r = 0; r < numRows; r++) {
                const obj = { __meta: { row: r + 2, cols: numCols } }; // 2 = 0-based and heading row
                for (let c = 0; c < numCols; c++) {
                    // @ts-expect-error: Headings are set already above, so possibility of an error here is nil
                    obj[headings[c]] = sheetValues[r][c]; // @ts-ignore
                }
                rowValues.push(obj);
            }
            this._sheetValues = rowValues;
        }
        return this._sheetValues;
    }
    // Return rows with matching criteria
    getRows() {
        const sheetValues = this.getValues();
        return this.whereFn ? sheetValues.filter(this.whereFn) : sheetValues;
    }
    getHeadings() {
        if (!this._sheetHeadings || !this._sheetHeadings.length) {
            const sheet = this.getSheet();
            const numCols = sheet.getLastColumn();
            this._sheetHeadings = sheet.getSheetValues(1, 1, 1, numCols)[0];
        }
        return this._sheetHeadings;
    }
    /**
     * Insert new rows into the spreadsheet
     * Arrays of objects like { Heading: Value }
     */
    insertRows(newRows) {
        const sheet = this.getSheet();
        const headings = this.getHeadings();
        newRows.forEach(row => {
            const rowValues = headings.map(heading => {
                return row[heading] ? row[heading] : '';
            });
            sheet.appendRow(rowValues);
        });
    }
    clearCache() {
        this._sheetValues = null;
        this._sheetHeadings = [];
        SpreadsheetApp.flush();
        return this;
    }
}
exports.SheetQueryBuilder = SheetQueryBuilder;
//# sourceMappingURL=index.js.map
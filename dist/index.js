/**
 * Run new sheet query
 *
 * @param {Spreadsheet} activeSpreadsheet Specific spreadsheet to use, or will use SpreadsheetApp.getActiveSpreadsheet() if undefined\
 * @return {SheetQueryBuilder}
 */
function sheetQuery(activeSpreadsheet) {
  return new SheetQueryBuilder(activeSpreadsheet);
}

/**
 * SheetQueryBuilder class - Kind of an ORM for Google Sheets
 */
class SheetQueryBuilder {
  constructor(activeSpreadsheet) {
    this.columnNames = [];
    this.headingRow = 1;
    this._sheetHeadings = [];
    this.activeSpreadsheet = activeSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  }
  select(columnNames) {
    this.columnNames = Array.isArray(columnNames) ? columnNames : [columnNames];
    return this;
  }
  /**
   * Name of spreadsheet to perform operations on
   *
   * @param {string} sheetName
   * @param {number} headingRow
   * @return {SheetQueryBuilder}
   */
  from(sheetName, headingRow = 1) {
    this.sheetName = sheetName;
    this.headingRow = headingRow;
    return this;
  }
  /**
   * Apply a filtering function on rows in a spreadsheet before performing an operation on them
   *
   * @param {Function} fn
   * @return {SheetQueryBuilder}
   */
  where(fn) {
    this.whereFn = fn;
    return this;
  }
  /**
   * Get Sheet object that is referenced by the current query from() method
   *
   * @return {Sheet}
   */
  getSheet() {
    if (!this.sheetName) {
      throw new Error('SheetQuery: No sheet selected. Select sheet with .from(sheetName) method');
    }
    if (!this._sheet) {
      this._sheet = this.activeSpreadsheet.getSheetByName(this.sheetName);
    }
    return this._sheet;
  }
  /**
   * Get values in sheet from current query + where condition
   */
  getValues() {
    if (!this._sheetValues) {
      const zh = this.headingRow - 1;
      const sheet = this.getSheet();
      if (!sheet) {
        return [];
      }
      const rowValues = [];
      const sheetValues = sheet.getDataRange().getValues();
      const numCols = sheetValues[0] ? sheetValues[0].length : 0;
      const numRows = sheetValues.length;
      const headings = (this._sheetHeadings = sheetValues[zh] || []);
      for (let r = 0; r < numRows; r++) {
        const obj = { __meta: { row: r + 1, cols: numCols } };
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
  /**
   * Return matching rows from sheet query
   *
   * @return {RowObject[]}
   */
  getRows() {
    const sheetValues = this.getValues();
    return this.whereFn ? sheetValues.filter(this.whereFn) : sheetValues;
  }
  /**
   * Get array of headings in current sheet from()
   *
   * @return {string[]}
   */
  getHeadings() {
    if (!this._sheetHeadings || !this._sheetHeadings.length) {
      const zh = this.headingRow - 1;
      const sheet = this.getSheet();
      const numCols = sheet.getLastColumn();
      this._sheetHeadings = sheet.getSheetValues(1, 1, this.headingRow, numCols)[zh] || [];
      this._sheetHeadings = this._sheetHeadings
        .map((s) => (typeof s === 'string' ? s.trim() : ''))
        .filter(Boolean);
    }
    return this._sheetHeadings || [];
  }
  /**
   * Get all cells from a query + where condition
   * @returns {any[]}
   */
  getCells() {
    const rows = this.getRows();
    const cellArray = [];
    rows.forEach((row) => {
      cellArray.push(this._sheet.getRange(row.__meta.row, 1, 1, row.__meta.cols));
    });
    return cellArray;
  }
  /**
   * Get cells in sheet from current query + where condition and from specific header
   * @param {string} key name of the column
   * @param {Array<string>} [keys] optionnal names of columns use to select more columns than one
   * @returns {any[]} all the colum cells from the query's rows
   */
  getCellsWithHeadings(key, headings) {
    let rows = this.getRows();
    let indexColumn = 1;
    const arrayCells = [];
    for (const elem of this._sheetHeadings) {
      if (elem == key) break;
      indexColumn++;
    }
    rows.forEach((row) => {
      arrayCells.push(this._sheet.getRange(row.__meta.row, indexColumn));
    });
    //If we got more thant one param
    headings.forEach((col) => {
      let indexColumn = 1;
      for (const elem of this._sheetHeadings) {
        if (elem == col) break;
        indexColumn++;
      }
      rows.forEach((row) => {
        arrayCells.push(this._sheet.getRange(row.__meta.row, indexColumn));
      });
    });
    return arrayCells;
  }
  /**
   * Insert new rows into the spreadsheet
   * Arrays of objects like { Heading: Value }
   *
   * @param {DictObject[]} newRows - Array of row objects to insert
   * @return {SheetQueryBuilder}
   */
  insertRows(newRows) {
    const sheet = this.getSheet();
    const headings = this.getHeadings();
    newRows.forEach((row) => {
      if (!row) {
        return;
      }
      const rowValues = headings.map((heading) => {
        const val = row[heading];
        return val === undefined || val === null || val === false ? '' : val;
      });
      // appendRow() will throw if array is empty, so we check to prevent that
      if (rowValues && rowValues.length !== 0) {
        sheet.appendRow(rowValues);
      }
    });
    return this;
  }
  /**
   * Delete matched rows from spreadsheet
   *
   * @return {SheetQueryBuilder}
   */
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
    patchRows(updateFn) {
    const rows = this.getRows();
    for (let i = 0; i < rows.length; i++) {
      this.patchRow(rows[i], updateFn);
    }
    this.clearCache();
    return this;
  }
  /**
   * Patch single row
   */
  patchRow(row, updateFn) {
    const rowMeta = row.__meta;
    delete row.__meta;
    const sourceRow = {...row}; // shallow copy
    const updatedRow = updateFn(row) || row;    
    const headings = this.getHeadings();
    // Write those cells only where values were changed
    headings.forEach((heading, col) => {
      const newValue = updatedRow[heading];
      if (newValue !== sourceRow[heading]){
        const updateRowCell = this.getSheet().getRange(rowMeta.row, col + 1)
        updateRowCell.setValues([[newValue]]);
      }      
    });
    return this;
  }
  /**
   * Update matched rows in spreadsheet with provided function
   *
   * @param {UpdateFn} updateFn
   * @return {SheetQueryBuilder}
   */
  updateRows(updateFn) {
    const rows = this.getRows();
    for (let i = 0; i < rows.length; i++) {
      this.updateRow(rows[i], updateFn);
    }
    this.clearCache();
    return this;
  }
  /**
   * Update single row
   */
  updateRow(row, updateFn) {
    const updatedRow = updateFn(row) || row;
    const rowMeta = updatedRow.__meta;
    const headings = this.getHeadings();
    delete updatedRow.__meta;
    // Put new array data in order of headings in sheet
    const arrayValues = headings.map((heading) => {
      const val = updatedRow[heading];
      return val === undefined || val === null || val === false ? '' : val;
    });
    const maxCols = Math.max(rowMeta.cols, arrayValues.length);
    const updateRowRange = this.getSheet().getRange(rowMeta.row, 1, 1, maxCols);
    const rangeData = updateRowRange.getValues()[0] || [];
    // Map over old data in same index order to update it and ensure array length always matches
    const newValues = rangeData.map((value, index) => {
      const val = arrayValues[index];
      return val === undefined || val === null || val === false ? '' : val;
    });
    updateRowRange.setValues([newValues]);
    return this;
  }
  /**
   * Clear cached values, headings, and flush all operations to sheet
   *
   * @return {SheetQueryBuilder}
   */
  clearCache() {
    this._sheetValues = null;
    this._sheetHeadings = [];
    SpreadsheetApp.flush();
    return this;
  }
}


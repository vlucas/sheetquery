import { sheetQuery } from '../src/index';
import { SpreadsheetApp } from 'gasmask';
import { Spreadsheet, Sheet } from 'gasmask/dist/SpreadsheetApp';

// @ts-ignore
//global.SpreadsheetApp = SpreadsheetApp;

const SHEET_NAME = 'TestSheet';
let ss = new Spreadsheet();

let sheet = new Sheet(SHEET_NAME);
const defaultSheetData = [
  ['Date', 'Amount', 'Name', 'Category'],
  ['2021-01-01', 5.32, 'Kwickiemart', 'Shops'],
  ['2021-01-02', 72.48, 'Shopmart', 'Shops'],
  ['2021-01-03', 1.97, 'Kwickiemart', 'Shops'],
  ['2021-01-03', 43.87, 'Gasmart', 'Gas'],
  ['2021-01-04', 824.93, 'Wholepaycheck', 'Groceries'],
];

function setupSpreadsheet(sheetData: any[]) {
  if (sheet) {
    ss.deleteSheet(sheet);
  }

  ss.insertSheet(SHEET_NAME);

  sheet = ss.getSheetByName(SHEET_NAME);
  sheetData.forEach((row) => sheet.appendRow(row));
}

describe('SheetQuery', () => {
  describe('from', () => {
    it('should allow user to specify heading column number as second argument', () => {
      const customSheetData = [['Nope', 'Nope2', 'Nope3', 'Nope4'], ...defaultSheetData];

      setupSpreadsheet(customSheetData);

      const query = sheetQuery(ss).from(SHEET_NAME, 2);
      const rows = query.getRows();

      expect(Object.keys(rows[0])).toEqual(['__meta', 'Date', 'Amount', 'Name', 'Category']);
      expect(rows.length).toBe(5);
      expect(rows).toContainEqual({
        Amount: 72.48,
        Category: 'Shops',
        Date: '2021-01-02',
        Name: 'Shopmart',
        __meta: { cols: 4, row: 4 },
      });
    });
  });

  describe('getHeadings', () => {
    it('should return first heading row by default', () => {
      setupSpreadsheet(defaultSheetData);

      const query = sheetQuery(ss).from(SHEET_NAME);
      const headings = query.getHeadings();

      expect(headings).toEqual(['Date', 'Amount', 'Name', 'Category']);
    });

    it('should return custom heading row when specified', () => {
      const customSheetData = [['Nope', 'Nope2', 'Nope3', 'Nope4'], ...defaultSheetData];

      setupSpreadsheet(customSheetData);

      const query = sheetQuery(ss).from(SHEET_NAME, 2);
      const headings = query.getHeadings();

      expect(headings).toEqual(['Date', 'Amount', 'Name', 'Category']);
    });
  });

  describe('getRows', () => {
    it('should return all rows for spreadsheet', () => {
      setupSpreadsheet(defaultSheetData);

      const query = sheetQuery(ss).from(SHEET_NAME);
      const rows = query.getRows();

      expect(rows.length).toBe(5);
      expect(rows).toContainEqual({
        Amount: 72.48,
        Category: 'Shops',
        Date: '2021-01-02',
        Name: 'Shopmart',
        __meta: { cols: 4, row: 3 },
      });
    });

    it('should only return rows that match Category = Shops', () => {
      setupSpreadsheet(defaultSheetData);

      const query = sheetQuery(ss)
        .from(SHEET_NAME)
        .where((row) => row.Category === 'Shops');
      const rows = query.getRows();

      expect(rows.length).toBe(3);
      expect(rows.every((row) => row.Category === 'Shops')).toBeTruthy();
    });
  });

  describe('insertRows', () => {
    it('Should insert rows in the correct places matching column headings', () => {
      setupSpreadsheet(defaultSheetData);

      const newRows = [
        {
          Amount: -554.23,
          Name: 'BigBox, inc. __INSERT_TEST__',
        },
        {
          Amount: -29.74,
          Name: 'Fast-n-greasy Food, Inc. __INSERT_TEST__',
        },
      ];

      // Insert rows
      sheetQuery(ss).from(SHEET_NAME).insertRows(newRows);

      const query = sheetQuery(ss).from(SHEET_NAME);
      const rows = query.getRows();

      const testRows = rows.filter((row) => row.Name.includes('__INSERT_TEST__'));

      expect(testRows[0].Name).toEqual(newRows[0].Name);
      expect(testRows[0].Date).toEqual('');
      expect(testRows[1].Name).toEqual(newRows[1].Name);
      expect(testRows[1].Date).toEqual('');
    });

    it('should ignore extra columns not present in spreadsheet during insert', () => {
      setupSpreadsheet(defaultSheetData);

      const newRows = [
        {
          Amount: -554.23,
          Name: 'BigBox, inc. __INSERT_TEST__',
          XtraCol: 'whatever',
        },
        {
          Amount: -29.74,
          Name: 'Fast-n-greasy Food, Inc. __INSERT_TEST__',
          XtraCol: 'nope',
        },
      ];

      // Insert rows
      sheetQuery(ss).from(SHEET_NAME).insertRows(newRows);

      const query = sheetQuery(ss).from(SHEET_NAME);
      const rows = query.getRows();

      const testRows = rows.filter((row) => row.Name.includes('__INSERT_TEST__'));

      expect(Object.keys(testRows[0])).not.toContain('XtraCol');
      expect(Object.keys(testRows[1])).not.toContain('XtraCol');
    });
  });
});

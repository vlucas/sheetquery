import { sheetQuery } from '../src/index';
import type { DictObject } from '../src/index';
import { SpreadsheetApp, Spreadsheet, Sheet } from 'gasmask/dist/SpreadsheetApp';

// @ts-ignore
global.SpreadsheetApp = SpreadsheetApp;

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
      expect(rows.length).toBe(customSheetData.length);
      expect(rows).toContainEqual({
        Amount: 72.48,
        Category: 'Shops',
        Date: '2021-01-02',
        Name: 'Shopmart',
        __meta: { cols: 4, row: 4 },
      });
    });
  });

  describe('getRows', () => {
    it('should return correct __meta row positions', () => {
      setupSpreadsheet(defaultSheetData);

      const query = sheetQuery(ss).from(SHEET_NAME);
      const rows = query.getRows();

      expect(Object.keys(rows[0])).toEqual(['__meta', 'Date', 'Amount', 'Name', 'Category']);
      expect(rows.length).toBe(defaultSheetData.length);
      expect(rows).toContainEqual({
        Amount: 'Amount',
        Category: 'Category',
        Date: 'Date',
        Name: 'Name',
        __meta: { cols: 4, row: 1 },
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

    it('should remove empty columns and trim extra spaces', () => {
      const sheetData = [
        ['Date', 'Amount', 'Name', 'Category', ' ', ' Extra ', ''],
        ['2021-01-01', 5.32, 'Kwickiemart', 'Shops', '', ''],
      ];
      setupSpreadsheet(sheetData);

      const query = sheetQuery(ss).from(SHEET_NAME).clearCache();
      const headings = query.getHeadings();

      expect(headings).toEqual(['Date', 'Amount', 'Name', 'Category', 'Extra']);
    });
  });

  describe('getRows', () => {
    it('should return first row with 1-index', () => {
      setupSpreadsheet(defaultSheetData);

      const query = sheetQuery(ss).from(SHEET_NAME);
      const rows = query.getRows();

      expect(rows.length).toBe(defaultSheetData.length);
      expect(rows[0].__meta).toEqual({
        cols: 4,
        row: 1,
      });
    });

    it('should return all rows for spreadsheet', () => {
      setupSpreadsheet(defaultSheetData);

      const query = sheetQuery(ss).from(SHEET_NAME);
      const rows = query.getRows();

      expect(rows.length).toBe(defaultSheetData.length);
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

    it('Should insert rows with 0 numberic values', () => {
      setupSpreadsheet(defaultSheetData);

      const newRows = [
        {
          Amount: 0,
          Name: 'BigBox, inc. __INSERT_TEST_Z__',
          Date: '2023-07-28',
        },
        {
          Amount: 0,
          Name: 'Fast-n-greasy Food, Inc. __INSERT_TEST_Z__',
          Date: '2023-07-28',
        },
      ];

      // Insert rows
      sheetQuery(ss).from(SHEET_NAME).insertRows(newRows);

      const query = sheetQuery(ss).from(SHEET_NAME);
      const rows = query.getRows();

      const testRows = rows.filter((row) => row.Name.includes('__INSERT_TEST_Z__'));

      expect(testRows[0].Name).toEqual(newRows[0].Name);
      expect(testRows[0].Amount).toEqual(newRows[0].Amount);
      expect(testRows[1].Name).toEqual(newRows[1].Name);
      expect(testRows[1].Amount).toEqual(newRows[1].Amount);
    });

    it('should ignore extra columns not present in spreadsheet during insert', () => {
      setupSpreadsheet(defaultSheetData);

      const newRows = [
        {
          Date: '2021-01-02',
          Amount: -554.23,
          Name: 'BigBox, inc. __INSERT_TEST__',
          XtraCol: 'whatever',
        },
        {
          Date: '2021-01-02',
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

    it('should not error with empty spreadsheet or headings', () => {
      setupSpreadsheet([]);

      const newRows = [
        {
          Date: '2021-01-02',
          Amount: -554.23,
          Name: 'BigBox, inc. __INSERT_TEST__',
          XtraCol: 'whatever',
        },
        {
          Date: '2021-01-02',
          Amount: -29.74,
          Name: 'Fast-n-greasy Food, Inc. __INSERT_TEST__',
          XtraCol: 'nope',
        },
      ];

      // Insert rows
      sheetQuery(ss).from(SHEET_NAME).insertRows(newRows);

      const query = sheetQuery(ss).from(SHEET_NAME);
      const rows = query.getRows();

      expect(rows.length).toEqual(2);
    });

    it('should not error with no rows', () => {
      setupSpreadsheet(defaultSheetData);

      const newRows: DictObject[] = [];

      // Insert rows
      sheetQuery(ss).from(SHEET_NAME).insertRows(newRows);

      const query = sheetQuery(ss).from(SHEET_NAME);
      const rows = query.getRows();

      expect(rows.length).toEqual(defaultSheetData.length);
    });

    it('should not error with rows with no data in them', () => {
      setupSpreadsheet(defaultSheetData);

      // @ts-ignore - Obvious type error, but with runtime data... who knows?
      const newRows: DictObject[] = [null];

      // Insert rows
      sheetQuery(ss).from(SHEET_NAME).insertRows(newRows);

      const query = sheetQuery(ss).from(SHEET_NAME);
      const rows = query.getRows();

      expect(rows.length).toEqual(defaultSheetData.length);
    });
  });

  describe('updateRows', () => {
    it('should return first heading row by default', () => {
      setupSpreadsheet(defaultSheetData);

      const catName = 'TEST_CUSTOM_CATEGORY';
      const query = sheetQuery(ss).from(SHEET_NAME);

      // Update rows
      query.updateRows((row) => Object.assign(row, { Category: catName }));

      const rows = query.clearCache().getRows();
      const allNewCategory = rows.some((row) => row.Category === catName);

      expect(allNewCategory).toEqual(true);
    });
  });

  describe('updateRow', () => {
    it('should update single row using __meta info', () => {
      setupSpreadsheet(defaultSheetData);

      const query = sheetQuery(ss).from(SHEET_NAME);
      const rows = query.getRows();

      const someRow = rows.find((row) => row.Name === 'Gasmart');

      expect(someRow).not.toBeUndefined();

      // Update row
      query.updateRow(someRow, (row) => Object.assign(row, { Name: 'Gasmart Ultra' }));

      const newRows = query.clearCache().getRows();
      const someNewRow = rows.find((row) => row.Name === 'Gasmart Ultra');

      expect(someNewRow).not.toBeUndefined();
    });

    it('should update single row with numeric 0 value', () => {
      setupSpreadsheet(defaultSheetData);

      const query = sheetQuery(ss).from(SHEET_NAME);
      const rows = query.getRows();

      const someRow = rows.find((row) => row.Name === 'Gasmart');

      expect(someRow).not.toBeUndefined();

      // Update row
      query.updateRow(someRow, (row) => Object.assign(row, { Name: 'Gasmart Ultra', Amount: 0 }));

      const newRows = query.clearCache().getRows();
      const someNewRow = rows.find((row) => row.Name === 'Gasmart Ultra');

      expect(someNewRow).not.toBeUndefined();
      expect(someNewRow?.Amount).toEqual(0);
    });

    it('should update without error even with mismatching column counts', () => {
      const customSheetData = [
        ['Date', 'Amount', 'Name', 'Category'],
        ['2021-01-01', 5.32, 'Kwickiemart', 'Shops'],
        ['2021-01-02', 72.48, 'Shopmart', 'Shops'],
        ['2021-01-03', 1.97, 'Kwickiemart', 'Shops'],
        ['2021-01-03', 43.87, 'Gasmart', 'Gas', '', '', 'something here'],
        ['2021-01-04', 824.93, 'Wholepaycheck', 'Groceries'],
      ];
      setupSpreadsheet(customSheetData);

      const query = sheetQuery(ss).from(SHEET_NAME);
      const rows = query.getRows();

      const someRow = rows.find((row) => row.Name === 'Gasmart');

      expect(someRow).not.toBeUndefined();

      // Update row
      query.updateRow(someRow, (row) => Object.assign(row, { Name: 'Gasmart Ultra' }));

      const newRows = query.clearCache().getRows();
      const someNewRow = rows.find((row) => row.Name === 'Gasmart Ultra');

      expect(someNewRow).not.toBeUndefined();
    });
  });
});

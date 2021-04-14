import { sheetQuery } from '../src/index';
import { SpreadsheetApp } from 'gasmask';
import { Spreadsheet, Sheet } from 'gasmask/dist/SpreadsheetApp';

// @ts-ignore
//global.SpreadsheetApp = SpreadsheetApp;

const SHEET_NAME = 'TestSheet';
let ss = new Spreadsheet();

let sheet = new Sheet(SHEET_NAME);
const sheetData = [
  ['Date', 'Amount', 'Name', 'Category'],
  ['2021-01-01', 5.32, 'Kwickiemart', 'Shops'],
  ['2021-01-02', 72.48, 'Shopmart', 'Shops'],
  ['2021-01-03', 1.97, 'Kwickiemart', 'Shops'],
  ['2021-01-03', 43.87, 'Gasmart', 'Gas'],
  ['2021-01-04', 824.93, 'Wholepaycheck', 'Groceries'],
];

beforeEach(() => {
  if (sheet) {
    ss.deleteSheet(sheet);
  }

  ss.insertSheet(SHEET_NAME);

  sheet = ss.getSheetByName(SHEET_NAME);
  sheetData.forEach(row => sheet.appendRow(row));
});

describe('SheetQuery', () => {
  describe('getRows', () => {
    it('should return all rows for spreadsheet', () => {
      const query = sheetQuery(ss)
        .from(SHEET_NAME);
      const rows = query.getRows();

      expect(rows.length).toBe(5);
      expect(rows).toContainEqual({
        "Amount": 72.48,
        "Category": "Shops",
        "Date": "2021-01-02",
        "Name": "Shopmart",
        "__meta": {"cols": 4, "row": 3},
      });
    });

    it('should only return rows that match Category = Shops', () => {
      const query = sheetQuery(ss)
        .from(SHEET_NAME)
        .where(row => row.Category === 'Shops');
      const rows = query.getRows();

      expect(rows.length).toBe(3);
      expect(rows.every(row => row.Category === 'Shops')).toBeTruthy();
    });
  });

  describe('insertRows', () => {
    it('Should insert rows in the correct places matching column headings', () => {
      const newRows = [
        {
          Amount: -554.23,
          Name: 'BigBox, inc.'
        },
        {
          Amount: -29.74,
          Name: 'Fast-n-greasy Food, Inc.'
        },
      ];

      // Insert rows
      sheetQuery(ss)
        .from(SHEET_NAME)
        .insertRows(newRows);

      const query = sheetQuery(ss).from(SHEET_NAME);
      const rows = query.getRows();

      expect(rows).toContain(newRows[0]);
    });
  });
});


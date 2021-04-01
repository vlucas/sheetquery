import { sheetQuery } from '../src/index';
import { SpreadsheetApp } from 'gasmask';
import { Spreadsheet, Sheet } from 'gasmask/dist/SpreadsheetApp';

// @ts-ignore
//global.SpreadsheetApp = SpreadsheetApp;

let ss = new Spreadsheet();

beforeAll(() => {
  ss.insertSheet('TestSheet');
});

describe('SheetQuery', () => {
  describe('getRows', () => {
    it('should return rows for spreadsheet', () => {
      const query = sheetQuery(ss)
        .from('TestSheet');
      const rows = query.getRows();

      expect(rows).toEqual(null);
    });
  });
});


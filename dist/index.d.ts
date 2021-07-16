/**
 * Run new sheet query
 *
 * @param {Spreadsheet} activeSpreadsheet Specific spreadsheet to use, or will use SpreadsheetApp.getActiveSpreadsheet() if undefined\
 * @return {SheetQuery}
 */
export declare function sheetQuery(activeSpreadsheet?: any): SheetQueryBuilder;
export declare type DictObject = {
    [key: string]: any;
};
export declare type RowObject = {
    [key: string]: any;
    __meta: {
        row: number;
        cols: number;
    };
};
export declare type WhereFn = (row: RowObject) => boolean;
export declare type UpdateFn = (row: RowObject) => RowObject | undefined;
/**
 * SheetQueryBuilder class - Kind of an ORM for Google Sheets
 */
export declare class SheetQueryBuilder {
    activeSpreadsheet: any;
    columnNames: string[];
    sheetName: string | undefined;
    whereFn: WhereFn | undefined;
    _sheet: any;
    _sheetValues: any;
    _sheetHeadings: string[];
    constructor(activeSpreadsheet?: any);
    select(columnNames: string | string[]): this;
    from(sheetName: string): this;
    where(fn: WhereFn): this;
    deleteRows(): this;
    updateRows(updateFn: UpdateFn): this;
    getSheet(): any;
    getValues(): any;
    getRows(): RowObject[];
    getHeadings(): string[];
    /**
     * Insert new rows into the spreadsheet
     * Arrays of objects like { Heading: Value }
     */
    insertRows(newRows: DictObject[]): void;
    clearCache(): this;
}

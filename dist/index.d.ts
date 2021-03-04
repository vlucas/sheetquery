/**
 * Run new sheet query
 */
export declare function sheetQuery(activeSpreadsheet: any): SheetQuery;
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
 * SheetQuery class - Kind of an ORM for Google Sheets
 */
export declare class SheetQuery {
    activeSpreadsheet: any;
    columnNames: string[];
    sheetName: string | undefined;
    whereFn: WhereFn | undefined;
    _sheet: any;
    _sheetValues: any;
    _sheetHeadings: string[] | null;
    constructor(activeSpreadsheet: any);
    select(columnNames: string | string[]): this;
    from(sheetName: string): this;
    where(fn: WhereFn): this;
    deleteRows(): this;
    updateRows(updateFn: UpdateFn): this;
    getSheet(): any;
    getValues(): any;
    getRows(): RowObject[];
    getHeadings(): string[] | null;
    clearCache(): this;
}

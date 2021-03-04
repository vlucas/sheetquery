# SheetQuery

Query builder/ORM for easily manipulating spreadsheets with Google Apps Script for Google Sheets (SpreadsheetApp).

This library was created for [BudgetSheet](https://www.budgetsheet.net) and I thought it might be useful for others as
well.

## Installation for Apps Scripts

Create a new file named *SheetQuery.gs* in your Google Apps Script project. Copy the contents of *dist/index.js* into
that file.

## Installation for Built Projects via NPM

To use `sheetquery` via NPM, install it in your project as a dependency:

```
npm i sheetquery
```

Now you are ready to get started using sheetquery in your project via import or require().

## Requirements

SheetQuery requires a Google Sheet with a heading row (typically the first row where the columns are named). SheetQuery
will use the heading row for all other operations, and for returning row data in key/value objects.

## Usage

SheetQuery operates on a single Sheet at a time. You can start a new query with `sheetQuery().from('SheetName')`.

### Query For Data

Data is queried based on the spreadsheet name and column headings:

```javascript
const query = sheetQuery()
  .from('Transactions')
  .where(row => row.Category === 'Shops');

// query.getRows() => [{ Amount: 95, Category: 'Shops', Business: 'Walmart'}]
```


### Update Rows

Query for the rows you want to update, and then update them:

```javascript
sheetQuery()
  .from('Transactions')
  .where(row => row.Business.toLowerCase().includes('starbucks'))
  .updateRows(row => { row.Category = 'Coffee Shops' });
```

The `updateRows` method can either return nothing, or can return a row object with updated properties that will be saved
back to the spreadsheet row. If the updater function returns nothing/undefined, the row object that was passed in will
be used (along with any changed values that will be updated by reference).

### Delete Rows

Query for the rows you want to delete, and then delete them:

```javascript
sheetQuery()
  .from('Transactions')
  .where(row => row.Category === 'DELETEME')
  .deleteRows();
```

SheetQuery will keep track of row indicies, ranges, etc. even as they change while deleting rows so you don't have to.

## API

Full API coming soon...


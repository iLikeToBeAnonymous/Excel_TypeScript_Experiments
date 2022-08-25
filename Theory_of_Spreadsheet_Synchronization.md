# Anatomy of an Excel File
Excel spreadsheets are divided into two primary hierarchical components:
- **Workbook** &mdash; The Excel file itself ([ExcelScript.Workbook])
- **Worksheet** &mdash; The individual sheets within the file
    - Retrievable via [.getWorksheets()] method

```TS
let sheetNames = workbook.getWorksheets().map((sheet) => sheet.getName());
console.log(sheetNames);
```

---
Worksheets also have several components, but these are not necessarily hierarchical to each other. However, they
_ARE_ child items (or in some cases are grandchildren) of the aformentioned components.
- **Range** &mdash; A block of cells (e.g., `A1:C5`)
    - Child of [ExcelScript.Worksheet].
    - Selected cells can encompass the entire worksheet, or they can technically be a single cell.
    - Retrieving a range by address (e.g., the `A1:C5` shown above) is a method exclusive to `ExcelScript.Worksheet.getRange(address?: string)`.
    - See [ExcelScript.Range] object
    - For cool peek under the hood, see the [Github ExcelScript.Range.yml]
- **RangeAreas** &mdash; An object that is a collection of one or more rectangular ranges in the same worksheet
    - Child of [ExcelScript.Worksheet] (can be a sibling or a child of the `Range` object)?
    - See [ExcelScript.RangeAreas]
- **Table** &mdash; A specially-defined range with additional features similar to a database table
    - Child of [ExcelScript.Worksheet].
    - A worksheet can contain multiple tables
    - See [ExcelScript.Table] object and [.addTable] method.
    - Particulars about tables
        - Just as a table within a database, the header row of a table must only contain static values.
            This means that values set by formulas are invalid for a table.
        - A table cannot have header values of `ExcelScript.RangeValueType === richValue`.
        - Adding data to the row immediately after the table doesn't cause the table to automatically expand to include the new data. The 
        `.resize(Range | string)` method can be used to expand the range programmatically, but this doesn't seem to hapen automatically.
- **RangeFormat** &mdash; A format object encapsulating the formatting of the cells within a given range
    - Child of `ExcelScript.Range`.
    - See [ExcelScript.RangeFormat]
    - Noteable Methods (searchable on the "[ExcelScript.RangeFormat]" page):
        - `Range.getValueTypes()` &mdash; returns a 2D array, the contents of each index being
             [ExcelScript.RangeValueType]. 
            - Possible values are _boolean_, _double_, _empty_, _error_, _integer_, _richValue_, _string_, and _unknown_
- **TableColumn** &mdash; A specialized object representing a column of an [ExcelScript.Table] object
    - Only available on a defined table in a worksheet
    - Retrievable by calling the [.getColumn('key')] method on a Table object (wherein `key` can be either a column name or a column ID)
    - See [ExcelScript.TableColumn] object
- **Filter** &mdash; A child interface to an [ExcelScript.TableColumn] object.
    - See [ExcelScript.Filter]

---
### _Key Methods of [ExcelScript.Range]_
##### `.getFormulas()`
```TS
let debugMsg: string;
(myRange.getFormulas()).forEach((cell) => {debugMsg += cell + '\n';}); 
debugMsg = debugMsg.split(',').join('\n');
console.log(debugMsg);
```

##### `.getHidden()`
Returns:
 - `true` if all cells in a range are hidden
 - `false` if no cells in the range are hidden
 - `null` if some cells are hidden
 See also `.getColumnHidden()` and `.getRowHiden()` for similar functionality specific to columns or rows.

##### `.unmerge()` &mdash; unmerges all merged cells in a range object

---

[//]: # (HIDDEN REFERENCES)
[ExcelScript.Worksheet]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.worksheet?view=office-scripts>
[ExcelScript.Table]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.table?view=office-scripts>
[.addTable]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.worksheet?view=office-scripts#excelscript-excelscript-worksheet-addtable-member(1)>
[ExcelScript.Range]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts>
[ExcelScript.RangeAreas]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.rangeareas?view=office-scripts>
[ExcelScript.RangeFormat]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.rangeformat?view=office-scripts>
[ExcelScript.RangeValueType]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.rangevaluetype?view=office-scripts#excelscript-excelscript-rangevaluetype-richvalue-member>
[ExcelScript.Workbook]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.workbook?view=office-scripts>
[.getWorksheets()]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.workbook?view=office-scripts#excelscript-excelscript-workbook-getworksheets-member(1)>
[.getColumn('key')]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.table?view=office-scripts#excelscript-excelscript-table-getcolumn-member(1)>
[ExcelScript.TableColumn]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.tablecolumn?view=office-scripts>
[ExcelScript.Filter]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.filter?view=office-scripts>
[Github ExcelScript.Range.yml]: <https://github.com/OfficeDev/office-scripts-docs-reference/blob/main/docs/docs-ref-autogen/excel/excelscript/excelscript.range.yml>
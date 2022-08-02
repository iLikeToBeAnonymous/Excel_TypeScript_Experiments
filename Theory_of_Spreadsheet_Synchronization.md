# Anatomy of an Excel File
Excel spreadsheets are divided into two primary hierarchical components:
- **Workbook** &mdash; The Excel file itself
- **Worksheet** &mdash; The individual sheets within the file

Worksheets also have several components, but these are not necessarily hierarchical to each other. However, they
_ARE_ child items (or in some cases are grandchildren) of the aformentioned components.
- **Range** &mdash; A block of cells (e.g., `A1:C5`)
    - Child of `ExcelScript.Worksheet`.
    - Selected cells can encompass the entire worksheet, or they can technically be a single cell
    - See [ExcelScript.Range object](rangeObject)
- **RangeAreas** &mdash; An object that is a collection of one or more rectangular ranges in the same worksheet
    - Child of `ExcelScript.Worksheet` (can be a sibling or a child of the `Range` object)?
    - See [ExcelScript.RangeAreas interface]()
- **Table** &mdash; A specially-defined range with additional features similar to a database table
    - Child of `ExcelScript.Worksheet`.
    - A worksheet can contain multiple tables
    - See [ExcelScript.Table Object](tableObject)
    - Particulars about tables
        - Just as a table within a database, the header row of a table must only contain static values.
            This means that values set by formulas are invalid for a table.
        - A table cannot have header values of `ExcelScript.RangeValueType === richValue`.
- **RangeFormat** &mdash; A format object encapsulating the formatting of the cells within a given range
    - Child of `ExcelScript.Range`.
    - See [ExcelScript.RangeFormat object](rangeFormat)
    - Noteable Methods (searchable on the "[rangeFormat]" page):
        - `Range.getValueTypes()` &mdash; returns a 2D array, the contents of each index being
             `ExcelScript.RangeValueType`. 
            - Possible values are _boolean_, _double_, _empty_, _error_, _integer_, _richValue_, _string_, and _unknown_

[//]: # (HIDDEN REFERENCES)
[tableObject]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.table?view=office-scripts>
[rangeObject]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts>
[rangeAreas]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.rangeareas?view=office-scripts>
[rangeFormat]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.rangeformat?view=office-scripts>
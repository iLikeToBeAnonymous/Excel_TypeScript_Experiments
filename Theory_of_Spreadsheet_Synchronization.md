# Anatomy of an Excel File
Excel spreadsheets are divided into two primary hierarchical components:
- Workbook &mdash; The Excel file itself
- Worksheet &mdash; The individual sheets within the file

Worksheets also have several components, but these are not necessarily hierarchical:
- Range &mdash; A block of cells (e.g., `A1:C5`)
  - Selected cells can encompass the entire worksheet, or they can technically be a single cell
  - See [ExcelScript.Range object](rangeObject)
- Table &mdash; A specially-defined range with additional features similar to a database table
  - A worksheet can contain multiple tables
  - See [ExcelScript.Table Object](tableObject)

[//]: # (HIDDEN REFERENCES)
[tableObject]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.table?view=office-scripts>
[rangeObject]: <https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts>
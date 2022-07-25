function main(workbook: ExcelScript.Workbook) {
	// let selectedSheet = workbook.getActiveWorksheet();
	/* https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.workbook?view=office-scripts#excelscript-excelscript-workbook-getselectedrange-member(1) */
	let myRange = workbook.getSelectedRange();

	formatCellRange(myRange);
};

function formatCellRange(myRange: ExcelScript.Range) {
	/* Clear formats from range (just in case there is formatting applied other than what is being specified here.) 
	This also clears out any manually-selected cell fill/highlighting 
    https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-ranges-set-get-values
    https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.rangeformat?view=office-scripts
    https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.verticalalignment?view=office-scripts */
	myRange.clear(ExcelScript.ClearApplyTo.formats);

	// Set vertical alignment to ExcelScript.VerticalAlignment.center for myRange on selectedSheet
	myRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
	myRange.getFormat().setIndentLevel(0);
	// Set horizontal alignment to ExcelScript.HorizontalAlignment.left for range N61 on selectedSheet
	myRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

	// Set font name to "Consolas" for myRange on selectedSheet
	myRange.getFormat().getFont().setName("Consolas");
	// Set font size to 10 for myRange on selectedSheet
	myRange.getFormat().getFont().setSize(10);
};

function main(workbook: ExcelScript.Workbook) {
    // Set fill color to specified color for range Sheet1!A2:C2
    let myFillColor = "#99ccff"
    let mySheet = workbook.getActiveWorksheet();
    // let myRange = mySheet.getRange("A383:F383");
    // https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.workbook?view=office-scripts#excelscript-excelscript-workbook-getselectedranges-member(1)
    let myRange = workbook.getSelectedRanges();
  
    let cellFillProps = myRange.getFormat().getFill();
    //https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.conditionalrangefill?view=office-scripts
    (cellFillProps.getColor() == "#FFFFFF") ? cellFillProps.setColor(myFillColor) : cellFillProps.clear();
    console.log(cellFillProps.getColor());
  
    // Below prints a csv string of the contents of the specified cell range.
    // console.log(String(myRange.getValues()));
  
    /*let nmbrCell = mySheet.getRange('D2:E2');
    //console.log(nmbrCell.getNumberFormat()); // Doesn't work if you feed it a range of cells
    let formatList = nmbrCell.getNumberFormats()
    console.log(`Array length: ${formatList.length}\n vals: ${String(formatList)}`); // Always returns an array.
  */
  };


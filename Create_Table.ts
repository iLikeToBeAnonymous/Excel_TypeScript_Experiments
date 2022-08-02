function main(workbook: ExcelScript.Workbook) {
    // let selectedSheet = workbook.getActiveWorksheet();
    /* https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.workbook?view=office-scripts#excelscript-excelscript-workbook-getselectedrange-member(1) */
    // let myRange = workbook.getSelectedRange(); console.log('Selected Range: ' + myRange.getAddress());
    let myWorksheet = workbook.getWorksheet('Production Orders (3)');
    let myRange = myWorksheet.getRange('A4:CM458')

    let headerRange = myWorksheet.getRange('A4:CM4');
    // let debugMsg: string;
    // (headerRange.getRow(0).getValues())[0].forEach((cell)=>{debugMsg += cell + '\n';});
    // // (headerRange.getRow(0).getValueTypes())[0].forEach((cell)=>{debugMsg += cell + '\n';});
    // // (headerRange.getValueTypes()).forEach((cell) => {debugMsg += cell + '\n';}); debugMsg = debugMsg.split(',').join('\n');
    // // (headerRange.getFormulas()).forEach((cell) => {debugMsg += cell + '\n';}); debugMsg = debugMsg.split(',').join('\n');
    // console.log(debugMsg);

    //https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts#excelscript-excelscript-range-setvalues-member(1)
    /** Retrieve all values of the header row, then write them back to the headerRange, thus converting formulas to 
     *  hard-coded values suitable for a Table  */
    let headerVals = headerRange.getValues();
    headerRange.setValues(headerVals);
    // headerRange.convertDataTypeToText();
    // console.log(rowValsAsString(headerRange)); //DEBUG to print row vals
    // /* ######################################################################### */
    // /* ########################## TABLE STUFF ################################## */
    // // let myTables = myWorksheet.getTables(); // Returns an ExcelScript.Table object, 
    // // console.log('Length of array of tables: ' + myTables.length)
    // // let firstTable = myTables[0];
    // // let range = firstTable.getRangeBetweenHeaderAndTotal();
    // // let rows = range.getValues();
    // // console.log('Row 0: \n' + rows[0]);
  
    /** Check for merged areas, which will prevent table creation... 
     * https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts#excelscript-excelscript-range-getmergedareas-member(1)
     * https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.rangeareas?view=office-scripts#excelscript-excelscript-rangeareas-getareacount-member(1)
     */
    if (myRange.getMergedAreas() != null) {
      console.log('Found merged areas!');
      console.log('Number of merged areas: ' + myRange.getMergedAreas().getAreaCount());
      myRange.unmerge();
    }else{console.log('No merged areas found')};
  
    // myRange.getFilter().clear();
    console.log('Top Row Addresses: ' + myRange.getRow(0).getAddress() +
      '\n     Column Count: ' + myRange.getColumnCount() +
      '\n        Row Count: ' + myRange.getRowCount());
    
  
      /** https://docs.microsoft.com/en-us/office/dev/scripts/tutorials/excel-power-automate-trigger?source=docs
       * https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.table?view=office-scripts 
       */
    let myNewTable = workbook.addTable(myRange, true);
    myNewTable.setName('scripted_table');
  };
  
  function deTableify(myNewTable: ExcelScript.Table) {
    /** https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.table?view=office-scripts#excelscript-excelscript-table-converttorange-member(1)
     *    .convertToRange();
     *    .getWorksheet(); // Worksheet containing the current table
     */
    myNewTable.convertToRange();
  };

  function rowValsAsString(myRange: ExcelScript.Range) {
    // // //https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts#excelscript-excelscript-range-getextendedrange-member(1)
    // // https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts#excelscript-excelscript-range-getrow-member(1)
    // getRow() counts the first row as being row zero.
    /* The var myRange specified earlier is a RANGE of cells... myRange.getRow(0) retrieves the 0th row of the defined range.
       Because the Range.prototype.getValues() function is meant to retrieve all values in a range, it returnes an array of arrays
       (It anticipates the top-level array as being the row of the sheet, and the secondary array as being the cell values in the row).
       Therefore, you must specify the 0th index of the returned array to just get a single array. */
    let headerRowVals = (myRange.getRow(0).getValues())[0]; // AN ARRAY OF COLUMN HEADER NAMES
    let bigStr = ''; // An empty string which will be populated in the following loop
    headerRowVals.forEach((ele: string, indx: number) => {
        bigStr = bigStr + 'index ' + String(indx).padStart(2, '0') + ': ' + ele + '\n'
    });
    // console.log(bigStr);
    return bigStr;
};
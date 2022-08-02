function main(workbook: ExcelScript.Workbook) {
    // let selectedSheet = workbook.getActiveWorksheet();
    /* https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.workbook?view=office-scripts#excelscript-excelscript-workbook-getselectedrange-member(1) */
    // let myRange = workbook.getSelectedRange(); console.log('Selected Range: ' + myRange.getAddress());
    let myRange = workbook.getWorksheet('Production Orders').getRange('A4:CM458')

    let myTables = workbook.getWorksheet('Production Orders').getTables(); // Returns an ExcelScript.Table object, 
    console.log('Length of array of tables: ' + myTables.length)
    let firstTable = myTables[0];
    let range = firstTable.getRangeBetweenHeaderAndTotal();
    let rows = range.getValues();
    console.log('Row 0: \n' + rows[0]);
  
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
    // let myNewTable = workbook.addTable(myRange, true);
    // myNewTable.setName('scripted_table');
  };
  
  function deTableify(myNewTable: ExcelScript.Table) {
    /** https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.table?view=office-scripts#excelscript-excelscript-table-converttorange-member(1)
     *    .convertToRange();
     *    .getWorksheet(); // Worksheet containing the current table
     */
    myNewTable.convertToRange();
  }
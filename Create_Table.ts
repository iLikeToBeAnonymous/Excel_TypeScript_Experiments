function main(workbook: ExcelScript.Workbook) {
    // let selectedSheet = workbook.getActiveWorksheet();
    /* https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.workbook?view=office-scripts#excelscript-excelscript-workbook-getselectedrange-member(1) */
    // let myRange = workbook.getSelectedRange(); console.log('Selected Range: ' + myRange.getAddress());
    let myWorksheet = workbook.getWorksheet('Production Orders (3)');
    let myRange = myWorksheet.getRange('A4:CM458')
    let targetTblNm: string = 'ProductionOrders'; // This will likely be passed in by PowerAutomate later
    let targetTbl: ExcelScript.Table; // declaration with no value.
    
    // /* ######################################################################### */
    // /* ########################## TABLE STUFF ################################## */
    // let headerRange = myRange.getRow(0); //myWorksheet.getRange('A4:CM4');
    // myWorksheet.getTables()[0].convertToRange(); //DEBUGGING converts the 0th table of the worksheet back to a range
    // let debugMsg: string;
    // (headerRange.getRow(0).getValues())[0].forEach((cell)=>{debugMsg += cell + '\n';});
    // // (headerRange.getRow(0).getValueTypes())[0].forEach((cell)=>{debugMsg += cell + '\n';});
    // // (headerRange.getValueTypes()).forEach((cell) => {debugMsg += cell + '\n';}); debugMsg = debugMsg.split(',').join('\n');
    // // (headerRange.getFormulas()).forEach((cell) => {debugMsg += cell + '\n';}); debugMsg = debugMsg.split(',').join('\n');
    // console.log(debugMsg);
    
    /**FIRST, VERIFY THAT A TABLE EXISTS */
    // validateTblHeaders(headerRange);
    let myTables = myWorksheet.getTables(); // Returns an ExcelScript.Table object, 
    let tableCount = myTables.length; // like Array.length, this is the count of entities, so 0 means no tables
    if(tableCount == 0){ // If no tables exist, create one using range defined above
        console.log('No tables in this worksheet!'); // DEBUGGING
        convertRangeToTable(myWorksheet, myRange, targetTblNm); // call custom convertRangeToTable() function
    }else{ /** We know that worksheet DOES contain more than one table, but we don't know if it contains the one we're looking for */
        console.log('Worksheet contained table count: ' + tableCount + '...')
        // let targetTblNm = myWorksheet.getTables()[0].getName();
        /**Use a try-catch to see if you can acquire the table you're looking for by name */
        try{targetTbl = myWorksheet.getTable(targetTblNm)}catch{console.log(`Table "${targetTblNm}" cannot be acquired!`)};
        if(targetTbl != null){ /** .getTable returns null if table can't be acquired */
        /** ################################################### */
        /** ######### BEGIN WORKING ON VERIFIED TABLE ######### */
            targetTbl.getColumns().forEach(col=>{console.log(col.getName())});
            let range = targetTbl.getRangeBetweenHeaderAndTotal();
            let rows = range.getValues();
            // console.log('Row 0: \n' + rows[0]);

            let records: ProductionOrderData[] = [];
            for (let row of rows) {
                let [madeUpKeys, anythingGoes, useUrImagination, nameMeButDontUseMe, mebbe] = row;
                records.push({
                  theseMust: madeUpKeys as string,
                  matchThe: anythingGoes as number,
                  interfaceDec: useUrImagination as string,
                  laration: mebbe as number
                })
              }
            console.log(JSON.stringify(records,null,2))
            // console.log('The Table says its header address range is: ' + targetTbl.getHeaderRowRange().getAddress() + 
            //     '\nTable has a range of: ' + targetTbl.getRange().getAddress() + 
            //     '\nRange between header and total row: ' + targetTbl.getRangeBetweenHeaderAndTotal().getAddress()); 
        /** ########## END WORKING ON VERIFIED TABLE ########## */
        /** ################################################### */
        }else if(targetTbl == null){
            console.log(`The table "${targetTblNm}" can neither be found nor created!`);
        }else{console.log('He\'s dead, Jim!')};
    }
    /** https://docs.microsoft.com/en-us/office/dev/scripts/tutorials/excel-power-automate-trigger?source=docs
     * https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.table?view=office-scripts 
     */
    // myWorksheet.getTables()[0].convertToRange(); // DEBUGGING: TOGGLE TO REMOVE TABLE CREATED BY THIS SCRIPT    
};

interface ProductionOrderData {
    theseMust: string
    matchThe: number
    interfaceDec: string
    laration: number
};

function convertRangeToTable(myWorksheet: ExcelScript.Worksheet, myRange: ExcelScript.Range, newTblName: string){
    // SPLIT ALL MERGED AREAS BEFORE ATTEMPTING TO CONVERT TO A TABLE
    splitMergedAreas(myRange);
    // MAKE SURE THE RANGE HAS VALID HEADERS, AND CONVERT IF NECESSARY
    validateTblHeaders(myRange.getRow(0)); // SHOULD always make the header row valid...
    let myNewTable = myWorksheet.addTable(myRange, true); // "true" that it has headers
    myNewTable.setName(newTblName);

};

function deTableify(myNewTable: ExcelScript.Table) {
/** https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.table?view=office-scripts#excelscript-excelscript-table-converttorange-member(1)
 *    .convertToRange();
 *    .getWorksheet(); // Worksheet containing the current table
 */
    myNewTable.convertToRange();
};

function validateTblHeaders(headerRange: ExcelScript.Range) {
      //https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts#excelscript-excelscript-range-setvalues-member(1)
    /** Retrieve all values of the header row, then write them back to the headerRange, thus converting formulas to 
     *  hard-coded values suitable for a Table  */
    const myRegex = /=/g;
    if(headerRange.getFormulas()[0].join().search(myRegex) == -1){
        console.log('No Formulas');
        return false;
    }else{
        let headerVals = headerRange.getValues();
        console.log('Formulas found and converted!');
        headerRange.setValues(headerVals);
        return true;
    };
};

function splitMergedAreas(myRange: ExcelScript.Range){
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
};
function main(
    workbook: ExcelScript.Workbook,
    targetSheetNm: string,
    targetTblNm: string,
    targetRangeAddr: string,
    searchTerm: string,
    indicatorColNm: string
  ) {
    /* https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.workbook?view=office-scripts#excelscript-excelscript-workbook-getselectedrange-member(1) */
    
    /* ###################### BEGIN DEFAULT PARAM SETUP ######################## */
    /* NEXT FEW LINES SET UP DEFAULT PARAMS IF NONE ARE PASSED FROM THE CALLER (POWER AUTOMATE) */
    if(targetSheetNm === undefined){targetSheetNm = workbook.getFirstWorksheet(true).getName()}; // Default to first visible worksheet if none specified
    let myWorksheet = workbook.getWorksheet(targetSheetNm);
    let targetRange: ExcelScript.Range; //Declare the variable so it will be in scope for the rest of the script
    if(targetRangeAddr === undefined){
        targetRange = myWorksheet.getUsedRange(true);
    }else{
        targetRange = myWorksheet.getRange(targetRangeAddr);
    }
    console.log('1st visible worksheet of workbook: ' + targetSheetNm); // DEBUGGING
    // let targetRange = myWorksheet.getRange('A4:CM458')
    // let targetRange = myWorksheet.getUsedRange(true);
    let targetTbl: ExcelScript.Table; // Simply a declaration of the var without a value.
    /* If targetTblNm is NOT passed in from Power Automate... */
    if(targetTblNm === undefined){targetTblNm = 'Table1'}; // This should be passed in from Power Automate

    //let indicatorColNm: string = 'evntLocation'; // DELETE THIS ROW!
    /* Valid regular expressions can be passed directly to the RegExp constructor as a string. The only limitation of this
     * is that the backslash char ("\") is the escape character for strings, so any instance of a backslash must be doubled
     * for it to actually be saved in the variable. */
    if(searchTerm === undefined){searchTerm = '^[\\w\\n\\s]+.*'} // Basically, a regex for the value not being null
    // /* To account for search strings with weird or invalid chars, you must first use a regular expression
    //  * before you can convert it to a regular expression... (Yes, I wrote that correctly). The following answer
    //  * was found in a post by user "Rivenfall" on StackOverflow (https://stackoverflow.com/a/35478115). He referenced
    //  * the Github repo here: https://github.com/sindresorhus/escape-string-regexp
    //  * See also the MDN entry: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Guide/Regular_Expressions#escaping */
    // searchTerm = searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); /* A pair of backslashes can also be inserted via String.fromCharCode(92, 92)); */
    const srchRegex: RegExp = new RegExp(searchTerm, 'gi');
    console.log(`Compiled regex.source: "${srchRegex.source}"\n\tand regex: "${srchRegex}"`); //DEBUGGING
    /* ###################### END DEFAULT PARAM SETUP ########################## */
    /* ######################################################################### */
    /* ########################## TABLE STUFF ################################## */  
    /* FIRST, VERIFY THAT A TABLE EXISTS */
    // validateTblHeaders(headerRange);
    let myTables = myWorksheet.getTables(); // Returns an ExcelScript.Table object, 
    let tableCount = myTables.length; // like Array.length, this is the count of entities, so 0 means no tables
    if (tableCount == 0) { // If no tables exist, create one using range defined above
      console.log('No tables in this worksheet!'); // DEBUGGING
      convertRangeToTable(myWorksheet, targetRange, targetTblNm); return 'New table created!';// call custom convertRangeToTable() function
    } else { /** We know that worksheet DOES contain more than one table, but we don't know if it contains the one we're looking for */
      console.log('Worksheet table count: ' + tableCount + '...'); // DEBUGGING
      /* CHECK IF targetTblNm IS A VALID TABLE IN THE WORKSHEET */
      let tblsFound: Array<string | number> = myTables.map(eachTbl => eachTbl.getName()); console.log('tblsFound:\n\t'+tblsFound.join('\n\t'));
      /* IF THE WORKSHEET DOES CONTAIN TABLES, BUT NONE BY THE NAME PASSED TO THIS SCRIPT, 
       * ATTEMPT TO PROCEED BY FALLING BACK TO THE FIRST TABLE FOUND IN THE WORKSHEET...  */
      if(!tblsFound.includes(targetTblNm)){ 
        console.log(`targetTblNm '${targetTblNm}' not found...\nUsing first found table, '${tblsFound[0]}' instead...`);
        targetTblNm = String(tblsFound[0]);//myWorksheet.getTables()[0].getName();
      }; 
      // console.log('Table[0] name: ' + targetTblNm); // DEBUGGING
      /**Use a try-catch to see if you can acquire the table you're looking for by name */
      try { targetTbl = myWorksheet.getTable(targetTblNm) } catch{ console.log(`Table "${targetTblNm}" cannot be acquired!`) };
      if (targetTbl != null) { /** .getTable returns null if table can't be acquired */

        /** ################################################### */
        /** #################  FILTER AS JSON ################# */  
        let myJsonObj: Array<JSON> = rangeToJsonObj(targetTbl.getRange());
        // let searchRez: Array<JSON> = myJsonObj.filter(myRec => myRec[indicatorColNm] == 'Salt Lake City');
        if(indicatorColNm === undefined || myJsonObj[0].hasOwnProperty(indicatorColNm) == false){ //Unfortunately, the newer Object.hasOwn() doesn't work in OfficeScripts...
          console.log(`"${indicatorColNm}" is an invalid indicatorColNm value! Fallback to first key of JSON obj...`)
          indicatorColNm = Object.keys(myJsonObj[0])[0]
        };
        let searchRez: Array<JSON> = myJsonObj.filter(myRec => myRec[indicatorColNm]['Val'].match(srchRegex));//let searchRez: Array<JSON> = myJsonObj.filter(myRec => myRec[indicatorColNm].match(srchRegex));
        console.log(`For a worksheet, Row/Col refs start counting at zero.\n\tTherefore, the val of R3, C4 is: ${myWorksheet.getCell(3,4).getValue()}`);
        console.log(`Search rez not eq to undefined: ${searchRez[0]!==undefined}`);
        let frstColName: string = Object.keys(searchRez[0])[0];
        console.log(`Demonstration of address structure of filtered JSON obj: ${searchRez[0][frstColName]['Addr']}`);
        console.log(`Demonstration of .getCell() method on first returned ele of JSON obj (using 'Row' and 'Col' keys): 
            ${myWorksheet.getCell(searchRez[0][frstColName]['Row'],searchRez[0][frstColName]['Col']).getAddress()}`);
        console.log(JSON.stringify(searchRez,null,2));
        /** ############### END FILTER AS JSON ################ */
        /** ################################################### */
      
        /** ################################################### */
        /** ######### BEGIN FILTERING VERIFIED TABLE ######### */
        // let debugMsg: Array<string | number | boolean | object> = [];
        // // // targetTbl.getColumns().forEach(col => { debugMsg.push('\t'+col.getName()) }); 
        // // // console.log('Column Names: \n'+debugMsg.join('\n')); // DEBUGGING
        // // console.log('Column Names: \n\t'+targetTbl.getHeaderRowRange().getValues()[0].join('\n\t'));
        // let columnNames = targetTbl.getHeaderRowRange().getValues()[0]; console.log(`Columns found:\n\t${columnNames}`);
  
        // targetTbl.clearFilters(); /* JUST TO MAKE SURE THERE ARE NO FILTERS INTERFERING WITH THINGS. */
        // /* FIND ADDRESS OF ROWS WHEREIN A SPECIFIC COLUMN CONTAINS A SPECIFIC VALUE
        //  * Unfortunately, the below line only returns the first hit */
        // // console.log(targetTbl.getColumn(indicatorColNm).getRange().find(searchTerm,{completeMatch: false}).getAddress())
        // /* Below returns a stringified array of addresses matching the search criteria. However,
        //    the below line is "undefined" if a filter is already in-place that excludes the results from 
        //    displaying. */
        // let matchingCells = myWorksheet.findAll(searchTerm, { completeMatch: false }).getAddress().split(',')
        // debugMsg.length = 0; // clear the array for the next use
        // let tempRange: ExcelScript.Range; let indicatorColIndx = columnNames.indexOf(indicatorColNm);
        // debugMsg.push('Addresses of ALL matches in worksheet:');// console.log(matchingCells);
        // matchingCells.forEach((ele)=>{
        //   tempRange = myWorksheet.getRange(ele);
        //   if(tempRange.getColumnIndex() == indicatorColIndx){
        //     debugMsg.push('\n\tRow: ' + myWorksheet.getRange(ele).getRowIndex() +
        //     ', Col: ' + myWorksheet.getRange(ele).getColumnIndex())
        //   }
        // });
        // console.log(debugMsg.join(''));
        /** ########## END FILTERING VERIFIED TABLE ########## */
        /** ################################################### */

      } else if (targetTbl == null) {
        console.log(`The table "${targetTblNm}" can neither be found nor created!`);
      } else { console.log('He\'s dead, Jim!') };
    }
    /** https://docs.microsoft.com/en-us/office/dev/scripts/tutorials/excel-power-automate-trigger?source=docs
     * https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.table?view=office-scripts 
     */
    // myWorksheet.getTables()[0].convertToRange(); // DEBUGGING: TOGGLE TO REMOVE TABLE CREATED BY THIS SCRIPT    
  };
  
  interface EventData {
    theseMust: string
    matchThe: number
    interfaceDec: string
    laration: number
  }
  
  function convertRangeToTable(myWorksheet: ExcelScript.Worksheet, targetRange: ExcelScript.Range, newTblName: string) {
    // SPLIT ALL MERGED AREAS BEFORE ATTEMPTING TO CONVERT TO A TABLE
    splitMergedAreas(targetRange);
    // MAKE SURE THE RANGE HAS VALID HEADERS, AND CONVERT IF NECESSARY
    validateTblHeaders(targetRange.getRow(0)); // SHOULD always make the header row valid...
    let myNewTable = myWorksheet.addTable(targetRange, true); // "true" that it has headers
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
    if (headerRange.getFormulas()[0].join().search(myRegex) == -1) {
      console.log('No Formulas');
      return false;
    } else {
      let headerVals = headerRange.getValues();
      console.log('Formulas found and converted!');
      headerRange.setValues(headerVals);
      return true;
    };
  };
  
  function splitMergedAreas(targetRange: ExcelScript.Range) {
    /** Check for merged areas, which will prevent table creation... 
     * https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts#excelscript-excelscript-range-getmergedareas-member(1)
     * https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.rangeareas?view=office-scripts#excelscript-excelscript-rangeareas-getareacount-member(1)
     */
    if (targetRange.getMergedAreas() != null) {
      console.log('Found merged areas!');
      console.log('Number of merged areas: ' + targetRange.getMergedAreas().getAreaCount());
      targetRange.unmerge();
    } else { console.log('No merged areas found') };
  
    // targetRange.getFilter().clear();
    console.log('Top Row Addresses: ' + targetRange.getRow(0).getAddress() +
      '\n     Column Count: ' + targetRange.getColumnCount() +
      '\n        Row Count: ' + targetRange.getRowCount());
  };
  
  function rangeToJsonObj(targetRange: ExcelScript.Range) {
    const colOffset = targetRange.getColumnIndex();
    const rowOffset = targetRange.getRowIndex()+1; //Because indx 0 of values is actually indx 1 of overall range
    // let myWorksheet = targetRange.getWorksheet(); // gets the worksheet containing the specified range
    /* Assume that index 0 of the topmost array contains the column headers, so just pull the column headers.. */
    let headerRowVals = targetRange.getValues()[0];
    console.log(`\tHeader row range: ${targetRange.getRow(0).getAddress()}\n\tand '.getColumnIndex()' yields: ${targetRange.getColumnIndex()}`);
    
    /* Next, pull the entire range of data values, excluding the header row (assumed to be the zeroth row). By 
       using slice(), you retrieve everything BUT the header row. */
    let sheetVals = targetRange.getValues().slice(1);
  
    // console.log(sheetVals);
    // console.log(headerRowVals.join(', '));
  
    let jsonArray: Array<JSON> = []; //This is clunky, but it's the only way the compiler doesn't complain.
    let rowCtr: number = 0; // WHILE EXCEL STARTS NUMBERING ROWS AT 1, FOR OUR PURPOSES, WE'LL NUMBER AT 0
    let rowLimit: number = sheetVals.length;
    
    while (rowCtr < rowLimit) {
      jsonArray.push(JSON.parse('{}')); //PUSH AN EMPTY OBJECT ONTO THE ARRAY
      /* sheetVals[rowCtr] defines which row of the sheet range. The zero-based row number of the sheet range
         corresponds to the index of the jsonArray into which we'll insert the object representing that row. */
      sheetVals[rowCtr].forEach((cellItem: string, indx: number) => {
        /* jsonArray[rowCtr] is the index of the jsonArray at which the object for the row is stored
           headerRowVals[indx] is the column title, but the compiler doesn't like this value unless it is
           explicitly converted to a string value.  */
        // jsonArray[rowCtr][String(headerRowVals[indx])] = cellItem;
        // jsonArray[rowCtr][String(headerRowVals[indx])] = { 'Val': cellItem, 'Addr': `R${rowCtr+rowOffset}, C${indx+colOffset}`};// jsonArray[rowCtr][String(headerRowVals[indx])] = { 'Val': cellItem, 'Addr': targetRange.getCell(rowCtr, indx).getAddress()};
        jsonArray[rowCtr][String(headerRowVals[indx])] = { 'Val': cellItem, 'Addr': [(rowCtr+rowOffset),(indx+colOffset)],'Row': (rowCtr+rowOffset), 'Col':(indx+colOffset)};
        // DEBUGGING BELOW (BUT DOESN'T WORK IF targetRange DOES NOT START AT CELL A1)
        // console.log('ObjKey: ' + String(headerRowVals[indx])+
        // '\n\tRow: '+rowCtr+' ColIndx: '+indx+' Val: '+cellItem+'\n\tRetrievedAddress: '+
        // myWorksheet.getRangeByIndexes(rowCtr,indx,1,1).getAddress()
        // );
        /* BELOW LINE USES THE LOCAL ROW AND COLUMN OF THE RANGE 
         * OBJECT TO THE GLOBAL ADDRESS OF THE INDIVIDUAL ELEMENTS */
        // console.log('Addr: ' + targetRange.getCell(rowCtr,indx).getAddress());
      });
  
      rowCtr++; // Advance to the next row of the sheet.
    };
    // console.log(JSON.stringify(jsonArray, null, 2)); //DEBUGGING
    return jsonArray;
  };

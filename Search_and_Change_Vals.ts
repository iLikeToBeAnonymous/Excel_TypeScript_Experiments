function main(workbook: ExcelScript.Workbook, targetSheetNm: string, targetTblNm: string, targetRangeAddr: string) {
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
      console.log('Worksheet contained table count: ' + tableCount + '...')
    //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    //   let foundTblNm = myWorksheet.getTables()[0].getName(); //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    let dbgMsg: string = ''; myWorksheet.getTables().forEach(tblEle => {dbgMsg += (tblEle.getName()+'\n')}); console.log('Tables found:\n'+dbgMsg);
      targetTblNm = myWorksheet.getTables()[0].getName(); // Fallback if an invalid table name is initially defined
      console.log('Table[0] name: ' + targetTblNm);
      /**Use a try-catch to see if you can acquire the table you're looking for by name */
      try { targetTbl = myWorksheet.getTable(targetTblNm) } catch{ console.log(`Table "${targetTblNm}" cannot be acquired!`) };
      if (targetTbl != null) { /** .getTable returns null if table can't be acquired */
        /** ################################################### */
        /** ######### BEGIN WORKING ON VERIFIED TABLE ######### */
        let debugMsg: Array<string | number | boolean | object> = [];
        // // targetTbl.getColumns().forEach(col => { debugMsg.push('\t'+col.getName()) }); 
        // // console.log('Column Names: \n'+debugMsg.join('\n')); // DEBUGGING
        // console.log('Column Names: \n\t'+targetTbl.getHeaderRowRange().getValues()[0].join('\n\t'));
        let columnNames = targetTbl.getHeaderRowRange().getValues()[0]; console.log(columnNames);
        // // BELOW SECTION TURNS THE ENTIRE TABLE INTO A FORMATTED STRING
        // debugMsg.length = 0; // clear the array for the next use
        // debugMsg.push('Table as formatted string (range is: '+targetTbl.getRange().getAddress() + ')' );
        // targetTbl.getRange().getValues().forEach(row => {
        //     debugMsg.push('\t'+row.join(',\t\t'))
        //   }); console.log(debugMsg.join('\n'));
        // // END TABLE AS STRING
  
        // // FIND ADDRESS OF ROWS WHEREIN A SPECIFIC COLUMN CONTAINS A SPECIFIC VALUE
           // Unfortunately, the below line only returns the first hit
        // console.log(targetTbl.getColumn('evntLocation').getRange().find('salt lake',{completeMatch: false}).getAddress())
        /* Below returns a stringified array of addresses matching the search criteria. However,
           the below line is "undefined" if a filter is already in-place that excludes the results from 
           displaying. */
        let matchingCells = (myWorksheet.findAll('salt lake', { completeMatch: false }).getAddress()).split(',')
        debugMsg.length = 0; // clear the array for the next use
        console.log(matchingCells);
        matchingCells.forEach((ele)=>{
          debugMsg.push('Row: ' + myWorksheet.getRange(ele).getRowIndex() +
            ', Col: ' + myWorksheet.getRange(ele).getColumnIndex())
        });
        console.log(debugMsg.join('\n'));

        //################################################################\\
        //################# PLAYING WITH FILTERS #########################\\
        targetTbl.clearFilters(); /* JUST TO MAKE SURE THERE ARE NO FILTERS INTERFERING WITH THINGS. */
        // targetTbl.getColumn('evntLocation').getFilter().applyValuesFilter(['salt lake city']);
        // // console.log(targetTbl.getColumn('evntLocation').getFilter().getCriteria());
        // let filteredRangeView = targetTbl.getRange().getVisibleView();
        // // console.log(filteredRangeView.getCellAddresses());
        // let targetColumnIndx = targetTbl.getColumn('registrationAmount').getRange().getColumnIndex();
        // console.log('target column index: ' + targetColumnIndx);
        // // console.log(filteredRangeView.getRows());
        // // console.log(filteredRangeView.getValues()); // The filtered table values (including headers)
        // // console.log(JSON.stringify(rangeToJsonObj(filteredRangeView.getRange()),null,2));
        //############## END PLAYING WITH FILTERS ########################\\
        //################################################################\\
        // // END ADDRESS OF ROWS WHEREIN A SPECIFIC COLUMN CONTAINS A SPECIFIC VALUE

        /* BELOW LINE JUST CONVERTS THE ENTIRE TABLE TO A JSON OBJECT AND RETURNS IT TO THE
           CALLING FUNCTION (IN THIS CASE, POWER AUTOMATE)                                  */
        return (JSON.stringify(rangeToJsonObj(targetTbl.getRange()),null,2));
  
        //################################################################\\
        // let tblBodyRange = targetTbl.getRangeBetweenHeaderAndTotal();
        // // console.log(JSON.stringify(rangeToJsonObj(targetTbl.getRange()), null, 2));
        // let rows = tblBodyRange.getValues();
        // console.log('Row 0: \n' + rows[0]);
        // // targetTbl.getColumnByName('Date').setName('Better Be There!');
        // let records: EventData[] = [];
        // for (let row of rows) {
        //   /** The following "let" declaraction is an array of variables, the values of which correspond
        //    *  to the values of the individual row */
        //   // let [madeUpKeys, anythingGoes, useUrImagination, nameMeButDontUseMe, mebbe] = row;
        //   let [eventID, evntDate, evntLocation, nameMeButDontUseMe, spkrSlots] = row;
        //   debugMsg.push(row);// = row;
        //   if (evntDate == 43892){records.push({
        //     theseMust: eventID as string,
        //     matchThe: evntDate as number,
        //     interfaceDec: evntLocation as string,
        //     laration: spkrSlots as number
        //   })}
        // }
        // console.log(JSON.stringify(records,null,2))
        //################################################################\\
  
        // console.log('The Table says its header address range is: ' + targetTbl.getHeaderRowRange().getAddress() + 
        //     '\nTable has a range of: ' + targetTbl.getRange().getAddress() + 
        //     '\nRange between header and total row: ' + targetTbl.getRangeBetweenHeaderAndTotal().getAddress()); 
        /** ########## END WORKING ON VERIFIED TABLE ########## */
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
    /*  While the below JSON.stringify works, it basically just returns an array of arrays, NOT 
        a proper JSON object formatted as a string... */
  
    /* Assume that index 0 of the topmost array contains the column headers, so just pull the column headers.. */
    let headerRowVals = targetRange.getValues()[0];
  
    /* Next, pull the entire range of data values, excluding the header row (assumed to be the zeroth row). By 
       using slice(), you retrieve everything BUT the header row. */
    let sheetVals = targetRange.getValues().slice(1);
  
    // console.log(sheetVals);
    // console.log(headerRowVals.join(', '));
  
    let jsonArray: Array<string | number | object> = []; //This is clunky, but it's the only way the compiler doesn't complain.
    let rowCtr: number = 0; // WHILE EXCEL STARTS NUMBERING ROWS AT 1, FOR OUR PURPOSES, WE'LL NUMBER AT 0
    let rowLimit: number = sheetVals.length;
    // let rowLimit: number = 3; //DEBUGGING ONLY
    while (rowCtr < rowLimit) {
      jsonArray.push({}); //PUSH AN EMPTY OBJECT ONTO THE ARRAY
      /* sheetVals[rowCtr] defines which row of the sheet range. The zero-based row number of the sheet range
         corresponds to the index of the jsonArray into which we'll insert the object representing that row. */
      sheetVals[rowCtr].forEach((cellItem: string, indx: number) => {
        /* jsonArray[rowCtr] is the index of the jsonArray at which the object for the row is stored
           headerRowVals[indx] is the column title, but the compiler doesn't like this value unless it is
           explicitly converted to a string value.  */
        jsonArray[rowCtr][String(headerRowVals[indx])] = cellItem;
      });
  
      rowCtr++; // Advance to the next row of the sheet.
    };
    // console.log(JSON.stringify(jsonArray, null, 2)); //DEBUGGING
    return jsonArray;
  };

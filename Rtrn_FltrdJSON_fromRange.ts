function main(
  workbook: ExcelScript.Workbook,
  targetSheetNm: string,
  targetRangeAddr: string,
  searchTerm: string, /* Search for empty or "0" matches can be done with '^$|^0$' */
  indicatorColNm: string
) {
  /* https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.workbook?view=office-scripts#excelscript-excelscript-workbook-getselectedrange-member(1) */

  /* ###################### BEGIN DEFAULT PARAM SETUP ######################## */
  /* NEXT FEW LINES SET UP DEFAULT PARAMS IF NONE ARE PASSED FROM THE CALLER (POWER AUTOMATE) */
  if (targetSheetNm === undefined) { targetSheetNm = workbook.getFirstWorksheet(true).getName() }; // Default to first visible worksheet if none specified
  console.log('Target sheet name: "' + targetSheetNm + '"'); // DEBUGGING
  const myWorksheet = workbook.getWorksheet(targetSheetNm);
  let targetRange: ExcelScript.Range; //Declare the variable so it will be in scope for the rest of the script
  if (targetRangeAddr === undefined) {
    targetRange = myWorksheet.getUsedRange(true); console.log('No target range specified. Fallback to used range.');
  } else {
    targetRange = myWorksheet.getRange(targetRangeAddr); console.log(`Found target range ${targetRange.getAddress()}`);
  };
  console.log(`Raw targetRange.getUsedRange(): ${targetRange.getUsedRange().getAddress()}\nRaw targetRange row count: ${targetRange.getUsedRange().getRowCount()}`); // DEBUGGING
  /* Valid regular expressions can be passed directly to the RegExp constructor as a string. The only limitation of this
  * is that the backslash char ("\") is the escape character for strings, so any instance of a backslash must be doubled
  * for it to actually be saved in the variable. */
  if (searchTerm === undefined) { searchTerm = '^[\\w\\n\\s]+.*' }; // Basically, a regex for the value not being null
  // /* To account for search strings with weird or invalid chars, you must first use a regular expression
  //  * before you can convert it to a regular expression... (Yes, I wrote that correctly). The following answer
  //  * was found in a post by user "Rivenfall" on StackOverflow (https://stackoverflow.com/a/35478115). He referenced
  //  * the Github repo here: https://github.com/sindresorhus/escape-string-regexp
  //  * See also the MDN entry: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Guide/Regular_Expressions#escaping */
  // searchTerm = searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); /* A pair of backslashes can also be inserted via String.fromCharCode(92, 92)); */
  const srchRegex: RegExp = new RegExp(searchTerm, 'gi');
  console.log(`Compiled regex.source: "${srchRegex.source}"\n\t            and regex: "${srchRegex}"`); //DEBUGGING

  console.log(`Val of targetRange: ${targetRange.getAddress()}\n Zeroth used range: ${targetRange.getUsedRange().getRow(0).getAddress()}`)
  /** ################################################### */
  /** #################  FILTER AS JSON ################# */
  let myJsonObj: Array<JSON> = rangeToJsonObj(targetRange); //rangeToJsonObj(targetTbl.getRange());
  // console.log(JSON.stringify(myJsonObj,null,2)); // DEBUGGING!
  // TEST indicatorColNm TO SEE IF IT'S VALID. IF NOT, SET TO VALUE OF FIRST KEY OF JSON OBJ
  if (indicatorColNm === undefined || myJsonObj[0].hasOwnProperty(indicatorColNm) == false) { //Unfortunately, the newer Object.hasOwn() doesn't work in OfficeScripts...
    console.log(`"${indicatorColNm}" is an invalid indicatorColNm value! Fallback to first key of JSON obj...`)
    indicatorColNm = Object.keys(myJsonObj[0])[0]
  }; 
  console.log(`Number of Rows in myJsonObj: ${myJsonObj.length}\nIndicator col name: ${indicatorColNm}`) // DEBUGGING
  let debugMsg: string = 'Unfiltered Results: \n\t'; myJsonObj.forEach(ele => debugMsg += `Row: ${ele[indicatorColNm]['Row']}, ${indicatorColNm}: ${ele[indicatorColNm]['Val']}\n\t`); console.log(debugMsg);

  // PERFORM THE FILTER OPERATION (myRec[indicatorColNm]['Val'] is coerced to a string to enable Regex match to work)
  let searchRez: Array<JSON> = myJsonObj.filter(myRec => String(myRec[indicatorColNm]['Val']).match(srchRegex));//let searchRez: Array<JSON> = myJsonObj.filter(myRec => myRec[indicatorColNm].match(srchRegex));
  debugMsg = 'Filtered Results: \n\t'; searchRez.forEach(ele => debugMsg += `Row: ${ele[indicatorColNm]['Row']}, ${indicatorColNm}: ${ele[indicatorColNm]['Val']}\n\t`); console.log(debugMsg);
  console.log(`For a worksheet, Row/Col refs start counting at zero.\n\tTherefore, the val of R3, C4 is: ${myWorksheet.getCell(3, 4).getValue()}`);
  console.log(`Matches found? ${searchRez[0] !== undefined}\nNumber of matches: ${searchRez.length}`);

  if(searchRez[0] !== undefined){ // DEBUGGING AND DEMONSTRATION SECTION
    let frstColName: string = Object.keys(searchRez[0])[0];
    console.log(`Demonstration of address structure of filtered JSON obj: ${searchRez[0][frstColName]['Addr']}`);
    console.log(`Demonstration of .getCell() method on first returned ele of JSON obj (using 'Row' and 'Col' keys): 
            ${myWorksheet.getCell(searchRez[0][frstColName]['Row'], searchRez[0][frstColName]['Col']).getAddress()}`);
  };
  // console.log(JSON.stringify(searchRez, null, 2));
  return searchRez;
  /** ############### END FILTER AS JSON ################ */
  /** ################################################### */
}; /* ########################## END FUNCTION MAIN ################################## */

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
  /* Isolate the used range within the target range, then select the uppermost populated row. Next, select the used range of the uppermost row
   *  and once again get the used range. This effectively removes any populated columns without headers to the left of the first cell. 
   * Once complete, row 0 of targetRange is trimmed down so the first column header is at the top-left position of the range "headRowTrimLeft" 
   */
  let headRowTrimLeft = targetRange.getUsedRange().getRow(0).getUsedRange(); 
  /* If headRowTrimLeft has only one column, performing getExtendedRange selects all the empty cells in the row. Below ternary assignment prevents this. */
  let headRowTrimRight: ExcelScript.Range = (headRowTrimLeft.getColumnCount() > 1) ? headRowTrimLeft.getRow(0).getColumn(0).getExtendedRange(ExcelScript.KeyboardDirection.right) : headRowTrimLeft; //Essentially performs Shift + Ctrl + RightArrow from the first cell
  let rowCount = targetRange.getRowCount(); // targetRange.getLastCell().getRowIndex()-rowOffset //.getColumn(<index>) pulls using the local index in the range
  // let colCount = headRowTrimRight.getColumnCount(); //headerRowVals.filter(word => word.toString().length > 0).length; // Number of populated cells in the header row
  let colCount = (headRowTrimLeft.getColumnCount() > headRowTrimRight.getColumnCount()) ? headRowTrimRight.getColumnCount() : headRowTrimLeft.getColumnCount();
  let resizedRange = headRowTrimRight.getAbsoluteResizedRange(rowCount, colCount).getUsedRange();
  console.log(` headRowTrimLeft: ${headRowTrimLeft.getAddress()}\nheadRowTrimRight: ${headRowTrimRight.getAddress()}\nRow Count: ${rowCount}\nRow Index: ${targetRange.getLastCell().getRowIndex()}\nCol Count:${colCount}\nResized range: ${resizedRange.getAddress()}`); //\nLast cell addr: ${rowCount.getAddress()}`);
  const colOffset = resizedRange.getColumnIndex();
  const rowOffset = resizedRange.getRowIndex() + 1; //Because indx 0 of values is actually indx 1 of overall range
  // let myWorksheet = targetRange.getWorksheet(); // gets the worksheet containing the specified range
  /* Assume that index 0 of the topmost array contains the column headers, so just pull the column headers.. */
  console.log(`\tHeader row range: ${targetRange.getRow(0).getAddress()}\n\tand '.getColumnIndex()' yields: ${targetRange.getColumnIndex()}`);
  let headerRowVals = resizedRange.getValues()[0];
  /* Next, pull the entire range of data values, excluding the header row (assumed to be the zeroth row). By 
     using slice(), you retrieve everything BUT the header row. */
  let sheetVals = resizedRange.getValues().slice(1);

  // console.log(sheetVals);
  // console.log(headerRowVals.join(', '));

  let jsonArray: Array<JSON> = []; //This is clunky, but it's the only way the compiler doesn't complain.
  let rowCtr: number = 0; // WHILE EXCEL STARTS NUMBERING ROWS AT 1, FOR OUR PURPOSES, WE'LL NUMBER AT 0
  let rowLimit: number = sheetVals.length;
  let jsonCtr: number = 0;

  while (rowCtr < rowLimit) {
    /* sheetVals[rowCtr] defines which row of the sheet range. The zero-based row number of the sheet range
    corresponds to the index of the jsonArray into which we'll insert the object representing that row. */
    if (sheetVals[rowCtr].join('').trim() != '') { // Filters out empty rows
      jsonArray.push(JSON.parse('{}')); //PUSH AN EMPTY OBJECT ONTO THE ARRAY
      sheetVals[rowCtr].forEach((cellItem: string, indx: number) => {
        /* jsonArray[rowCtr] is the index of the jsonArray at which the object for the row is stored
          headerRowVals[indx] is the column title, but the compiler doesn't like this value unless it is
          explicitly converted to a string value.  */
        // jsonArray[rowCtr][String(headerRowVals[indx])] = cellItem;
        // jsonArray[rowCtr][String(headerRowVals[indx])] = { 'Val': cellItem, 'Addr': `R${rowCtr+rowOffset}, C${indx+colOffset}`};// jsonArray[rowCtr][String(headerRowVals[indx])] = { 'Val': cellItem, 'Addr': resizedRange.getCell(rowCtr, indx).getAddress()};
        jsonArray[jsonCtr][String(headerRowVals[indx])] = {
          'Val': cellItem, 'Addr': [(rowCtr + rowOffset), (indx + colOffset)],
          'Row': (rowCtr + rowOffset), 'Col': (indx + colOffset)
        };

      }); jsonCtr++;
    }

    rowCtr++; // Advance to the next row of the sheet.
  };
  // console.log(JSON.stringify(jsonArray, null, 2)); //DEBUGGING
  return jsonArray;
};
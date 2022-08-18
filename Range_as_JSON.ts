function main(workbook: ExcelScript.Workbook) {
    let mySheet = workbook.getActiveWorksheet();

    //https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.worksheet?view=office-scripts#excelscript-excelscript-worksheet-getusedrange-member(1)
    //https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts
    // NOTE: THIS CAUSES PROBLEMS IF THE TOPMOST-ROW OF THE POPULATED RANGE IS NOT THE HEADER ROW
    let usedRange = mySheet.getUsedRange(true);
    console.log('Used Range: ' + usedRange.getAddress());
    // // // console.log('usedRange is of type: ' + typeof usedRange) // It's an "Object"

    console.log(columnHeadersListString(usedRange));
    printRangeAsJson(usedRange);
    // let headerRowVals = (usedRange.getRow(0).getValues())[0]; // AN ARRAY OF COLUMN HEADER NAMES
    // // console.log(JSON.stringify(headerRowVals,null,2)); //calling JSON.stringify() works!
};

function printRangeAsJson(myRange: ExcelScript.Range) {
    /*  While the below JSON.stringify works, it basically just returns an array of arrays, NOT 
        a proper JSON object formatted as a string... */

    /* Pull the entire sheet of data values, excluding the header row. By using slice(), 
       you retrieve everything BUT the header row. */
    let sheetVals = myRange.getValues().slice(1);

    /* Next, assume that index 0 of the topmost array contains the column headers... */
    let headerRowVals = myRange.getValues()[0];
    
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
    console.log(JSON.stringify(jsonArray, null, 2));
};

function rangeToJsonObj(myRange: ExcelScript.Range) {
    /*  While the below JSON.stringify works, it basically just returns an array of arrays, NOT 
        a proper JSON object formatted as a string... */

    /* Assume that index 0 of the topmost array contains the column headers, so just pull the column headers.. */
    let headerRowVals = myRange.getValues()[0];

    /* Next, pull the entire range of data values, excluding the header row (assumed to be the zeroth row). By 
       using slice(), you retrieve everything BUT the header row. */
    let sheetVals = myRange.getValues().slice(1);
    
    // console.log(sheetVals);
    // console.log(headerRowVals.join(', '));

    let jsonArray: Array<JSON> = []; //This is clunky, but it's the only way the compiler doesn't complain.
    let rowCtr: number = 0; // WHILE EXCEL STARTS NUMBERING ROWS AT 1, FOR OUR PURPOSES, WE'LL NUMBER AT 0
    let rowLimit: number = sheetVals.length;
    // let rowLimit: number = 3; //DEBUGGING ONLY
    while (rowCtr < rowLimit) {
        jsonArray.push(JSON.parse('{}')); //PUSH AN EMPTY OBJECT ONTO THE ARRAY
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

function columnHeadersListString(usedRange: ExcelScript.Range) {
    // // let headerRowStr = usedRange.getAddress().split(':');
    // // console.log(headerRowStr[0]);
    // // //https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts#excelscript-excelscript-range-getextendedrange-member(1)
    // // let firstRow = mySheet.getRange("A1").getExtendedRange(ExcelScript.KeyboardDirection.right); //mySheet.getRange("A1").getEntireRow() 
    // // console.log(firstRow.getAddress());
    // // let columnCount = usedRange.getColumnCount();
    // // https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts#excelscript-excelscript-range-getrow-member(1)
    // getRow() counts the first row as being row zero.
    console.log('Top Row Addresses: ' + usedRange.getRow(0).getAddress() +
        '\n     Column Count: ' + usedRange.getColumnCount() +
        '\n        Row Count: ' + usedRange.getRowCount());
    /* The var usedRange specified earlier is a RANGE of cells... usedRange.getRow(0) retrieves the 0th row of the defined range.
       Because the Range.prototype.getValues() function is meant to retrieve all values in a range, it returnes an array of arrays
       (It anticipates the top-level array as being the row of the sheet, and the secondary array as being the cell values in the row).
       Therefore, you must specify the 0th index of the returned array to just get a single array. */
    let headerRowVals = (usedRange.getRow(0).getValues())[0]; // AN ARRAY OF COLUMN HEADER NAMES
    // console.log(JSON.stringify(headerRowVals,null,2)); //calling JSON.stringify() works!
    // console.log(headerRowVals); 
    //headerRowVals.forEach(ele => console.log(ele)); // THIS WORKS WELL!
    let bigStr = ''; // An empty string which will be populated in the following loop
    headerRowVals.forEach((ele: string, indx: number) => {
        bigStr = bigStr + 'index ' + String(indx).padStart(2, '0') + ': ' + ele + '\n'
    });
    // console.log(bigStr);
    return bigStr;
};
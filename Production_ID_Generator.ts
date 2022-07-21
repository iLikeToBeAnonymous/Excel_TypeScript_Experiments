function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getActiveWorksheet();
  // let cleanedDt = String(genDateTime().match(/\d*/g).join(''));

  // // let myPid = convertToBase(cleanedDt, 32);
  // let myPid = shortenTimestamp(cleanedDt);
  // // workbook.getSelectedRange().setValue(myPid);
  // workbook.getSelectedRange().setValue(myPid);

  //https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.worksheet?view=office-scripts#excelscript-excelscript-worksheet-getusedrange-member(1)
  //https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts
  let usedRange = mySheet.getUsedRange(true);
  console.log('Used Range: ' + usedRange.getAddress());
  // // // console.log('usedRange is of type: ' + typeof usedRange) // It's an "Object"
  // // console.log(Object.entries(usedRange));
  console.log(columnHeadersListString(usedRange));
  printRangeAsJson(usedRange);
  // let headerRowVals = (usedRange.getRow(0).getValues())[0]; // AN ARRAY OF COLUMN HEADER NAMES
  // // console.log(JSON.stringify(headerRowVals,null,2)); //calling JSON.stringify() works!
};

function printRangeAsJson(myRange: ExcelScript.Range){
  /*  While the below JSON.stringify works, it basically just returns an array of arrays, NOT 
      a proper JSON object formatted as a string... */
  // console.log(JSON.stringify(myRange.getValues(),null,2));
  /* Pull the entire sheet of data values, excluding the header row. By using slice(), 
     you retrieve everything BUT the header row. */
  let sheetVals = myRange.getValues().slice(1);
  let headerRowVals = myRange.getValues()[0];
  /* Next, assume that index 0 of the topmost array contains the column headers... */
  console.log(sheetVals);
  console.log(headerRowVals.join(', '));

  let jsonArray: Array<string|number|object> = []; //This is clunky, but the compiler doesn't complain.
  let rowCtr: number = 0; 
  let rowLimit: number = sheetVals.length;
  while (rowCtr < rowLimit){
    jsonArray.push({});
    sheetVals[rowCtr].forEach((cellItem: string, indx: number) => {
      // console.log(headerRowVals[indx] + ': ' + cellItem);
      jsonArray[rowCtr][headerRowVals[indx]] = cellItem;
    });

    rowCtr++;
  };

  // //let jsonArray = [{}]// Array(); // It doesn't like the Array() constructor...
  // let jsonArray: Array<string|number|object> = []; //This is clunky, but the compiler doesn't complain.
  // let myCtr = 0;
  // //console.log('Counter: ' + myCtr);
  // jsonArray.push({});
  // jsonArray[myCtr]['alpha'] = 'first';
  // jsonArray[myCtr]['beta'] = 'second';
  // myCtr += 1;
  // //console.log('Counter: ' + myCtr);
  // jsonArray.push({});
  // jsonArray[myCtr]['alpha'] = 'third';
  // jsonArray[myCtr]['beta'] = 'fourth';
  console.log(JSON.stringify(jsonArray,null,2));

};

function columnHeadersListString(usedRange: ExcelScript.Range){
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
  headerRowVals.forEach((ele: string,indx: number) => {
    bigStr = bigStr + 'index ' + String(indx).padStart(2,'0') + ': ' + ele + '\n'
  });
  // console.log(bigStr);
  return bigStr;
};

function genDateTime() {
  var now = new Date();
  return now.toISOString(); // must use "val" instead of "text" since it's an input box.
};

function convertToBase(originalNumber: string, targetBaseSystem: number) {
    var convertedNumber = ""; //targetBaseSystem = 32;
    var extraNumeralTable = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghikjlmnopqrstuvwxyz";
    var chineseNumerals = '一二三四五六七八九十'; // These are the Chinese characters for 1 through 10 (the plus-looking thing is 10)
    var monthTable = "123456789OND"; //specialized base12 using 'O' for October, 'N' for November, 'D' for December.
    var dayTable = "123456789ABCDEFGHIJKLMNOPQRSTUV";

    while (Number(originalNumber) > 0) {
        /* The javaScript "remainder" method fails due to the shorcomings
        / of floating point numbers. Therefore, a function needs to be created instead.*/
        var returnedModulo = modulo(originalNumber, String(targetBaseSystem)); // call to custom modulo function
        // modulo from loop: "+ loopRemainder);
        //rightDigit = returnedModulo;
        if (returnedModulo > 9) {
            var rightDigit = extraNumeralTable[Number(returnedModulo) - 10];
        } else { rightDigit = String(returnedModulo); };
        // console.log("rightDigit: " + rightDigit);
        //console.log("originalNumber before flooring: " + longDivision(originalNumber,targetBaseSystem));
        //originalNumber = Math.floor(originalNumber / targetBaseSystem); //this is still introducing error.
        originalNumber = (longDivision(Number(originalNumber), targetBaseSystem)).match(/\d{1,}/g)[0];
        /*The line above extracts the substring left of the decimal using match() method with a regular expression
          • See 'https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String/match'
        */
        // console.log("originalNumber at end of loop: " + Number(originalNumber)); // BigInt trims the leading zeroes, but isn't necessary for this to work
        //convertedNumber = String(rightDigit) + convertedNumber;
        convertedNumber = String(rightDigit).concat(convertedNumber); //
    }
    console.log(convertedNumber)
    return convertedNumber;
};

function modulo(divident: string, divisor: string) { // when passing, make sure the divident is either already a string or is a BigInt.
    divident = typeof (divident) != 'string' ? String(divident) : divident; //per MDN: "String() converts anything to a string, safer than toString()"
    divisor = typeof (divisor) != 'string' ? String(divisor) : divisor;

    var shallowCopyArray = Array.from(divident); //creates a new, shallow-copied Array instance from an array-like or iterable object.
    /*Next, use the Array.prototype.map() method to iterate through the shallow copy array and perform a function on each value.
      The var "eachIndex" is the contents at each index in the shallowCopyArray.
      The contents at each index is parsed into an int.*/
    var anArrayOfInts = shallowCopyArray.map(eachIndex => parseInt(eachIndex, 10)); // The "10" is just to make sure it parses to base 10
    console.log(anArrayOfInts.toString()); // equivalent value to divident/divisor rounded down to an int.
    //The var "remainder" gets used as the accumulator in the reduce method below
    var myAccumulator = anArrayOfInts.reduce((remainder, value) => (remainder * 10 + value) % divisor, 0); // Mod is calculated on no more than two digits at a time this way
    console.log("Accumulator Value: " + myAccumulator);
    return myAccumulator;
};

function longDivision(myNumerator: number, myDenominator: number) {
    var num = myNumerator + '',
        numLength = num.length,
        remainder = 0,
        answer = '',
        i = 0; //the index of var "num".

    while (i < numLength + 3) { //Why did I put "+ 3" here???
        // Here, parseInt(num[i]) just seems to be converting the string back into an int.
        var digit = i < numLength ? parseInt(num[i]) : 0; //If i < numLength{digit = parseInt(num[i])} else{digit = 0}

        if (i == numLength) {
            answer = answer + ".";
        }
        //answer = itself appended with the whole-number-only quotient of each digit (times 10) and the passed denominator
        // REMEMBER! var answer is a STRING!
        answer = answer + Math.floor((digit + (remainder * 10)) / myDenominator);
        remainder = (digit + (remainder * 10)) % myDenominator;
        i++;
    }
    return answer;
};

function monthCode(rawNmbr: number) {
  rawNmbr = Math.floor(Math.abs(rawNmbr) * 1); //make sure it's a positive integer
  var monthTable = '123456789OND'; //specialized base12 using 'O' for October, 'N' for November, 'D' for December.
  if (rawNmbr < 10) { return String(rawNmbr) }
  else { return monthTable[rawNmbr - 1] };
};

function shortenTimestamp(rawNmbr: string) {

  /* See "https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/slice"
     The slice() method RETURNS a shallow copy of A PORTION OF AN ARRAY into a new array object SELECTED FROM START TO END
     (end not included) where start and end represent the index of items in that array. The original array will not be modified.
     • arr.slice(fromIndex,toIndex) // returns what falls between these indices (includes 1st index, excludes 2nd index). If both are negative, it extracts from the end instead of the beginning.
     • arr.slice(i)
       • If positive, this removes i elements from the beginning of the index and returns the rest.
       • If negative, this returns i elements from the end and discards the rest.*/
  //var milSec = rawNmbr.slice(rawNmbr.length-3); // returns the last 3 of the array

  //There will never be more than 999 ms, and in base32, this can be represented with only two places. Pad to two if there's only 1.
  var milSec = rawNmbr.slice(-3);
  milSec = milSec.padStart(3, '0'); // See "https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String/padStart"

  var theScnds = rawNmbr.slice(-5, -3); // max seconds > 32 & < 100, so conversion doesn't save any space.

  var allMilSec = rawNmbr.slice(-5); //var allMilSec = theScnds.concat(milSec);
  var hndrthsSec = Math.round(`${allMilSec.slice(0, 4)}.${allMilSec.slice(-1)}`); //basically divides "allMilSec" by 10 and 
  console.log('Notice me, senpai! ' + hndrthsSec);
  var theMinutes = rawNmbr.slice(-7, -5); // max minutes > 32 & < 100, so conversion doesn't save any space.

  var theHrs = rawNmbr.slice(-9, -7); //max hours < 32 & > 9, so conversion DOES save space. DO NOT PAD.
  // console.log(convertToBase(theDays,32).padStart(1,'0') + '-' + convertToBase(theHrs,32).padStart(1,'0')); //9999.31556925.999
  var theDays = rawNmbr.slice(-11, -9); // max days per month < 32 & > 9, so conversion DOES save space. DO NOT PAD.

  var theMonth = rawNmbr.slice(-13, -11); // max month < 32 & > 9, so conversion DOES save space. DO NOT PAD.

  // Year is long, so perhaps restrict to 3 digits and just assume the 1st digit will be a "2" for the life the the business?
  var theYear = rawNmbr.slice(-17, -13); //4 digits can be reduced to 3, so only pad 3

  var allTogether = [convertToBase(theYear, 36).padStart(3, '0'),
  '-' + monthCode(theMonth), //convertToBase(theMonth,12).padStart(1,'0'), 
  convertToBase(theDays, 32).padStart(1, '0'),
  '-' + convertToBase(theHrs, 24).padStart(1, '0'),
  theMinutes.padStart(2, '0'),
  '-' + convertToBase(hndrthsSec, 36).padStart(3, '0')]; // you gain nothing by representing seconds and ms separately. w' base 36, max ms can be cut to 4 digits.

  //var abbrvAsGroup = [theYear, theMinutes, theScnds, milSec];
  //var abbrvSeparately = [theMonth, theDays, theHrs]; // 59:59.999 = 3,599,999 ms 5LS9V vs 3DRJV
  var shortenedTimestamp = allTogether.join('');//[allTogether.join('')]; //Puts them all together in the right order.

  // If an invalid date (i.e., one having more than 12 months, more than 31 days, etc.) is passed, the function throws an "undefined"

  // console.log(shortenedTimestamp.join('.').split('.'));
  return shortenedTimestamp; // joins all the elements of the array together with periods as a separator.
};
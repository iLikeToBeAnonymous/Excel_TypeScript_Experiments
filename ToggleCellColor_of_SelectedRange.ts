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

function itWorks(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();
  // Set range D101 on selectedSheet
  // selectedSheet.getRange("D101").setValue("HUMKBAGTKA3");
  workbook.getSelectedRange().setValue("HUMKBAGTKA3");
};

function main(workbook: ExcelScript.Workbook) {
    let mySheet = workbook.getActiveWorksheet();
    let cleanedDt = String(genDateTime().match(/\d*/g).join(''));
    // let myPid = convertToBase('20220718194020675',32);
    let myPid = convertToBase(cleanedDt, 32);
    // workbook.getSelectedRange().setValue(myPid);
    workbook.getSelectedRange().setValue(myPid);
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
          //loopRemainder= originalNumber % targetBaseSystem; // The "%" is the "modulus" operator (returns the remainder of dividing first num by second num)
          // console.log("originalNumber at start of loop: " + BigInt(originalNumber));
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
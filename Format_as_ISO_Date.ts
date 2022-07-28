function main(workbook: ExcelScript.Workbook) {
	// let selectedSheet = workbook.getActiveWorksheet();
	/* https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.workbook?view=office-scripts#excelscript-excelscript-workbook-getselectedrange-member(1) */
	let myRange = workbook.getSelectedRange();

	formatAsIsoDate(myRange);
};

// https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.numberformatcategory?view=office-scripts#examples
function formatAsIsoDate(myRange: ExcelScript.Range) {
	let cellObj: ExcelScript.Range
    const rangeContents = myRange.getValues();
	const nmbrFormatCats = myRange.getNumberFormatCategories();
    // let nmbrFormatLocalCats = myRange.getNumberFormatsLocal(); // myRange.getNumberFormatLocal();
    const nmbrFormatCodes = myRange.getNumberFormats();
    let debugMsg: string;
	rangeContents.forEach((row, rowIndx) => {
		row.forEach((cellVal, cellIndx) => {
            cellObj = myRange.getCell(rowIndx, cellIndx);
			// https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts#excelscript-excelscript-range-getvaluetype-member(1)
            debugMsg = 'Row ' + rowIndx + ', Col ' + cellIndx + ': ' + cellVal + 
            '\n\thas numberFormatCategory of ' + nmbrFormatCats[rowIndx][cellIndx] +
            '\n\tand nmbrFormatCode of ' + nmbrFormatCodes[rowIndx][cellIndx] + 
            '\n\tand RangeVAlueType is: ' + cellObj.getValueType();
            // https://github.com/OfficeDev/office-scripts-docs/blob/main/docs/resources/samples/excel-samples.md#dates
            // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Date/toISOString
                if(nmbrFormatCats[rowIndx][cellIndx] == ExcelScript.NumberFormatCategory.date){
                    debugMsg += '\n\tFORMAT IS "DATE"'; //console.log('format is date')
                    // debugMsg += '\n\t   valAsNmbr: ' + Date.parse(cellObj.getValue());
                    if(String(cellObj.getValueType())=='String'){
                        // debugMsg += '\n\t  ' + (new Date(Date.parse(cellObj.getValue() + ' GMT'))).toISOString();
                        // debugMsg += '\n\t  ' + (new Date(Date.parse(cellObj.getValue() + ' GMT'))).toISOString();
                        // cellObj.setValue(microsoftToUnix((new Date(Date.parse(cellVal + ' GMT'))).valueOf()));
                        cellObj.setValue((((new Date(Date.parse(cellVal + ' GMT'))).valueOf()/(86400000))+25569));
                    };
                    formatRangeAsCode(cellObj);
                    cellObj.setNumberFormat("yyyy-mm-dd;@");
                }
                // Beginning of Unix Epoc is 1970-01-01, which in Microsoft time is 25569.
                else if(((nmbrFormatCats[rowIndx][cellIndx] == ExcelScript.NumberFormatCategory.custom) || (nmbrFormatCats[rowIndx][cellIndx] == ExcelScript.NumberFormatCategory.time)) && cellVal > 25569){
                    debugMsg += '\n\tCUSTOM FORMAT IS LIKELY A DATE'; //console.log('Custom format is likely a date');
                    formatRangeAsCode(cellObj);
                    cellObj.setNumberFormat("yyyy-mm-dd;@");
                }
                else {
                    debugMsg += '\n\tVALUE CANNOT BE COERCED INTO A DATE'; //console.log('Value cannot be coerced into a date')
                };

            console.log(debugMsg);
            // if(nmbrFormatCats[rowIndx][cellIndx] == ExcelScript.NumberFormatCategory.date){
            //     console.log('Correcting format for cellVal');
            //     cellObj.setNumberFormatLocal("yyyy-mm-dd;@");
            // }
            // // else if (nmbrFormatCats[rowIndx][cellIndx] == ExcelScript.NumberFormatCategory.custom){
            // //     console.log('nmbrFormatLocalCats is ' + nmbrFormatLocalCats[rowIndx][cellIndx]);
            // // }
            // else{console.log('Cell is not a date')}
		});
	});

};

function microsoftToUnix(myNmbr: number){
	return ((myNmbr/(86400000))+25569);
};

// function stringToDateCode(myRange: ExcelScript.Range){
//     /**Date.parse() returns a number value, but NOT a date object.
//      * If a Date.parse() value is wrapped in a new Date() constructor,
//      * a date object is returned instead.
//      */
//     const mdntGmt = new Date(Date.parse('01 Jan 1999 00:00:00 GMT'));
//     //const mdntNyc = Date.parse('04 Dec 1995 00:12:00 GMT');
//     const mdntNyc = new Date(Date.parse('JAN/1/1999' + ' GMT')); //'1 Jan 1999 GMT -5'
//     const fromNmbr = new Date(915148800000);
//     //console.log(mdntGmt);
//     console.log(mdntGmt.valueOf());
//     console.log(mdntNyc.valueOf());
// 	console.log(fromNmbr.valueOf());//console.log(fromNmbr.toISOString());
//     // expected output: 818035920000
//     //#####################################################################
//     //const event = Date.parse('JAN/1/1999'); //new Date('JAN/1/1999');
//     const event = new Date('JAN/1/1999');
//     //const event = new Date(Date.UTC(2012, 11, 20, 3, 0, 0));
//     console.log('valueOf(): ' + (event.valueOf()));
//     console.log('toISOString(): ' + event.toISOString());
//     /*
//     let excelDateValue = 36161;
//     let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
//     console.log('toISOString(): ' + javaScriptDate.toISOString());
//     //console.log(javaScriptDate.valueOf());

//     //https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Date/parse
//     const unixTimeZero = Date.parse('01 Jan 1999 00:00:00 GMT');
//     const javaScriptRelease = Date.parse('Jan/1/99');

//     console.log(unixTimeZero); console.log(javaScriptRelease);
//     */
// };

function formatRangeAsCode(myRange: ExcelScript.Range) {
    /* Clear formats from range (just in case there is formatting applied other than what is being specified here.) 
    This also clears out any manually-selected cell fill/highlighting 
    https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-ranges-set-get-values
    https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.rangeformat?view=office-scripts
    https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.verticalalignment?view=office-scripts */
    myRange.clear(ExcelScript.ClearApplyTo.formats);

    // Set vertical alignment to ExcelScript.VerticalAlignment.center for myRange on selectedSheet
    myRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    myRange.getFormat().setIndentLevel(0);
    // Set horizontal alignment to ExcelScript.HorizontalAlignment.left for range N61 on selectedSheet
    myRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);

    // Set font name to "Consolas" for myRange on selectedSheet
    myRange.getFormat().getFont().setName("Consolas");
    // Set font size to 10 for myRange on selectedSheet
    myRange.getFormat().getFont().setSize(10);
};

	// myRange.setNumberFormatLocal("yyyy-mm-dd;@");
	// let cellContents = myRange.getValues();
	// cellContents.forEach((row,rowIndx) => {
	//     // console.log('row val: ' + row); // Prints each row as a csv string
	//     row.forEach((cellVal,cellIndx) => {
	//         console.log('Row ' + rowIndx + ', Col ' + cellIndx + ' = ' + cellVal)
	//     });
	// });

	// // let numberFormatCategories = myRange.getNumberFormatCategories();
	// // console.log(numberFormatCategories);
	// // numberFormatCategories.forEach(row => {
	// // 		row.forEach(cellVal => {console.log(cellVal)	
	// // 	});
	// // });
	// // // numberFormatCategories.forEach((category, index) => {
	// // // 	console.log('category[0]: ' + String(category[0]))
	// // // 	// if (category[0] != ExcelScript.NumberFormatCategory.currency) {
	// // // 	// 	costColumnRange.getCell(index, 0).getFormat().getFill().setColor("red");
	// // // 	// };
	// // // }); 


	// // // let selectedSheet = workbook.getActiveWorksheet();
	// // // // Set range U393 on selectedSheet
	// // // selectedSheet.getRange("U393").setFormulaLocal("=now()");
	// // // // Set number format for range U393 on selectedSheet
	// // // selectedSheet.getRange("U393").setNumberFormatLocal("yyyy-mm-dd;@");


/**
 * d-mmm-yy
 * [h]:mm:ss
 * m/d/yy
 */
function main(workbook: ExcelScript.Workbook) {
	// let selectedSheet = workbook.getActiveWorksheet();
	/* https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.workbook?view=office-scripts#excelscript-excelscript-workbook-getselectedrange-member(1) */
	let myRange = workbook.getSelectedRange();

	formatAsIsoDate(myRange);
};

// https://docs.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.numberformatcategory?view=office-scripts#examples
function formatAsIsoDate(myRange: ExcelScript.Range) {
	let rangeContents = myRange.getValues();
	let nmbrFormatCats = myRange.getNumberFormatCategories();
    // let nmbrFormatLocalCats = myRange.getNumberFormatsLocal(); // myRange.getNumberFormatLocal();
    let nmbrFormatCodes = myRange.getNumberFormats();
    let debugMsg: string;
	rangeContents.forEach((row, rowIndx) => {
		row.forEach((cellVal, cellIndx) => {
			// console.log('Cell val: ' + cellVal + ' has format of ' + nmbrFormatCats[rowIndx][cellIndx]);
            // console.log('Row ' + rowIndx + ', Col ' + cellIndx + ' has\n\tnumberFormatCategory of ' + nmbrFormatCats[rowIndx][cellIndx] +
            //     '\n\tand nmbrFormatLocalCats val of ' + nmbrFormatLocalCats[rowIndx][cellIndx] + 
            //     '\n\tand nmbrFormatCode of ' + nmbrFormatCodes[rowIndx][cellIndx]);
            // console.log('Row ' + rowIndx + ', Col ' + cellIndx + ' has\n\tnumberFormatCategory of ' + nmbrFormatCats[rowIndx][cellIndx] +
            //     '\n\tand nmbrFormatCode of ' + nmbrFormatCodes[rowIndx][cellIndx]);
            debugMsg = 'Row ' + rowIndx + ', Col ' + cellIndx + ': ' + cellVal + 
            '\n\thas numberFormatCategory of ' + nmbrFormatCats[rowIndx][cellIndx] +
            '\n\tand nmbrFormatCode of ' + nmbrFormatCodes[rowIndx][cellIndx];
                if(nmbrFormatCats[rowIndx][cellIndx] == ExcelScript.NumberFormatCategory.date){
                    debugMsg += '\n\tFORMAT IS "DATE"'; //console.log('format is date')
                    myRange.getCell(rowIndx, cellIndx).setNumberFormatLocal("yyyy-mm-dd;@");
                }
                // Beginning of Unix Epoc is 1970-01-01, which in Microsoft time is 25569.
                else if((nmbrFormatCats[rowIndx][cellIndx] == ExcelScript.NumberFormatCategory.custom) && cellVal > 25569){
                    debugMsg += '\n\tCUSTOM FORMAT IS LIKELY A DATE'; //console.log('Custom format is likely a date');
                    myRange.getCell(rowIndx, cellIndx).setNumberFormatLocal("yyyy-mm-dd;@");
                }
                else {
                    debugMsg += '\n\tVALUE CANNOT BE COERCED INTO A DATE'; //console.log('Value cannot be coerced into a date')
                };

            console.log(debugMsg);
            // if(nmbrFormatCats[rowIndx][cellIndx] == ExcelScript.NumberFormatCategory.date){
            //     console.log('Correcting format for cellVal');
            //     myRange.getCell(rowIndx, cellIndx).setNumberFormatLocal("yyyy-mm-dd;@");
            // }
            // // else if (nmbrFormatCats[rowIndx][cellIndx] == ExcelScript.NumberFormatCategory.custom){
            // //     console.log('nmbrFormatLocalCats is ' + nmbrFormatLocalCats[rowIndx][cellIndx]);
            // // }
            // else{console.log('Cell is not a date')}
		});
	});

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
};

/**
 * d-mmm-yy
 * [h]:mm:ss
 * m/d/yy
 */
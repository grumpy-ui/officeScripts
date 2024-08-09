function main(workbook: ExcelScript.Workbook) {
	const date = new Date();
	const monthIndex = date.getMonth();
	const currentYear = date.getFullYear();
	const sourceWs = workbook.getWorksheet("Summary");
	const tempWs = workbook.addWorksheet("temp");
	const mainWsName = `0${monthIndex + 1}.${currentYear}`
	const mainWs = workbook.addWorksheet(mainWsName);
	const months = [
		"January",
		"February",
		"March",
		"April",
		"May",
		"June",
		"July",
		"August",
		"September",
		"October",
		"November",
		"December",
	];
	const currentMonth = months[monthIndex];
	tempWs
		.getRange("A1")
		.copyFrom(sourceWs.getUsedRange(), ExcelScript.RangeCopyType.values);
	const lastRowTemp = tempWs.getUsedRange().getRowCount();
	const uacvTotal = tempWs.getRange(`T${lastRowTemp + 1}`).getValue();
	const colDArr =  tempWs.getRange(`D1:D${lastRowTemp}`).getValues()
	const nonBlankElements = colDArr.filter((row) => {
		const value = row[0];
		return typeof value === "string" && value.trim() !== "" && value !== `Total Invoice\n(€)` && value !== 'Program Group'
	}).map((val) => val[0]).filter((el,index, self) => {
		return self.indexOf(el) === index
	}).sort()



	for (let i = 0; i < nonBlankElements.length; i++) {
		let mainWsRngCell = mainWs.getRange(`A${i + 1}`);
		mainWsRngCell.setValue(nonBlankElements[i]);
	}

	//Delete first 3 rotempWs
	mainWs.getRange("1:2").insert(ExcelScript.InsertShiftDirection.down);
	mainWs.getRange("B2").setFormula(`=SUMIF(temp!D:D,${mainWsName}!A2,temp!J:J)`);
	mainWs.getRange("C2").setFormula(`=SUMIF(temp!D:D,${mainWsName}!A2,temp!P:P)`);
	mainWs.getRange("B2").autoFill();
	mainWs.getRange("C2").autoFill();
	mainWs
		.getRange("B:C")
		.setNumberFormatLocal('_( #,##0.00_);_( (#,##0.00);_( ""-""??_);_(@_)');

	mainWs.getRange("A2").setValue("Program");
	mainWs.getRange("B2").setValue("Total Invoice €");
	mainWs.getRange("C2").setValue("Total Invoice UACA €");

	//Formatting
	mainWs.getRange("A1:C1").clear(ExcelScript.ClearApplyTo.contents);
	mainWs.getRange("A1:C1").merge(false);
	mainWs.getRange("A1").setValue(`${currentYear} ${currentMonth} Invoice`);
	mainWs
		.getRange("A1:C1")
		.getFormat()
		.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	mainWs.getRange("A1:C1").getFormat().getFont().setName("Calibri");
	mainWs.getRange("A1:C1").getFormat().getFont().setSize(14);
	mainWs.getRange("A1:C1").getFormat().getFont().setColor("#000000");
	mainWs.getRange("A1:C2").getFormat().getFont().setBold(true);
	mainWs.getRange("A1:C1").getFormat().getFill().setColor("#fce4d6");
	mainWs.getRange("A2:C2").getFormat().getFill().setColor("#203764");
	mainWs.getRange("A2:C2").getFormat().getFont().setColor("#FFFFFF");
	mainWs.getRange().getFormat().autofitColumns();

	let lastRowMain = mainWs.getUsedRange().getRowCount();

	mainWs.getRange(`A${lastRowMain + 1}`).setValue("Total UACE");
	mainWs.getRange(`A${lastRowMain + 2}`).setValue("Total UACV");
	mainWs.getRange(`A${lastRowMain + 3}`).setValue("Grand Total");
	mainWs.getRange(`B${lastRowMain + 1}`).setValue(`=SUM(B3:B${lastRowMain})`);
	mainWs.getRange(`B${lastRowMain + 2}`).setValue(uacvTotal);
	mainWs
		.getRange(`B${lastRowMain + 3}`)
		.setFormula(`=SUM(B${lastRowMain + 1}:B${lastRowMain + 2})`);
	mainWs.getRange(`C${lastRowMain + 1}`).setFormula(`=SUM(C3:C${lastRowMain})`);

	mainWs
		.getRange(`B3:C${lastRowMain}`)
		.copyFrom(`B3:C${lastRowMain}`, ExcelScript.RangeCopyType.values);

	mainWs
		.getRange(`B${lastRowMain + 1}:C${lastRowMain + 3}`)
		.setNumberFormatLocal(
			'_([$€-x-euro2] * #,##0.00_);_([$€-x-euro2] * (#,##0.00);_([$€-x-euro2] * ""-""??_);_(@_)'
		);
	mainWs
		.getRange(`A${lastRowMain + 1}:C${lastRowMain + 3}`)
		.getFormat()
		.getFont()
		.setBold(true);
	mainWs
		.getRange(`A1:C${lastRowMain + 3}`)
		.getFormat()
		.getRangeBorder(ExcelScript.BorderIndex.diagonalDown)
		.setStyle(ExcelScript.BorderLineStyle.none);
	mainWs
		.getRange(`A1:C${lastRowMain + 3}`)
		.getFormat()
		.getRangeBorder(ExcelScript.BorderIndex.diagonalUp)
		.setStyle(ExcelScript.BorderLineStyle.none);
	mainWs
		.getRange(`A1:C${lastRowMain + 3}`)
		.getFormat()
		.getRangeBorder(ExcelScript.BorderIndex.edgeLeft)
		.setStyle(ExcelScript.BorderLineStyle.continuous);
	mainWs
		.getRange(`A1:C${lastRowMain + 3}`)
		.getFormat()
		.getRangeBorder(ExcelScript.BorderIndex.edgeLeft)
		.setWeight(ExcelScript.BorderWeight.thin);
	mainWs
		.getRange(`A1:C${lastRowMain + 3}`)
		.getFormat()
		.getRangeBorder(ExcelScript.BorderIndex.edgeTop)
		.setStyle(ExcelScript.BorderLineStyle.continuous);
	mainWs
		.getRange(`A1:C${lastRowMain + 3}`)
		.getFormat()
		.getRangeBorder(ExcelScript.BorderIndex.edgeTop)
		.setWeight(ExcelScript.BorderWeight.thin);
	mainWs
		.getRange(`A1:C${lastRowMain + 3}`)
		.getFormat()
		.getRangeBorder(ExcelScript.BorderIndex.edgeBottom)
		.setStyle(ExcelScript.BorderLineStyle.continuous);
	mainWs
		.getRange(`A1:C${lastRowMain + 3}`)
		.getFormat()
		.getRangeBorder(ExcelScript.BorderIndex.edgeBottom)
		.setWeight(ExcelScript.BorderWeight.thin);
	mainWs
		.getRange(`A1:C${lastRowMain + 3}`)
		.getFormat()
		.getRangeBorder(ExcelScript.BorderIndex.edgeRight)
		.setStyle(ExcelScript.BorderLineStyle.continuous);
	mainWs
		.getRange(`A1:C${lastRowMain + 3}`)
		.getFormat()
		.getRangeBorder(ExcelScript.BorderIndex.edgeRight)
		.setWeight(ExcelScript.BorderWeight.thin);
	mainWs
		.getRange(`A1:C${lastRowMain + 3}`)
		.getFormat()
		.getRangeBorder(ExcelScript.BorderIndex.insideVertical)
		.setStyle(ExcelScript.BorderLineStyle.continuous);
	mainWs
		.getRange(`A1:C${lastRowMain + 3}`)
		.getFormat()
		.getRangeBorder(ExcelScript.BorderIndex.insideVertical)
		.setWeight(ExcelScript.BorderWeight.thin);
	mainWs
		.getRange(`A1:C${lastRowMain + 3}`)
		.getFormat()
		.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal)
		.setStyle(ExcelScript.BorderLineStyle.continuous);
	mainWs
		.getRange(`A1:C${lastRowMain + 3}`)
		.getFormat()
		.getRangeBorder(ExcelScript.BorderIndex.insideHorizontal)
		.setWeight(ExcelScript.BorderWeight.thin);

	tempWs.delete()
}

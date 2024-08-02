async function main(workbook: ExcelScript.Workbook) {
    let secondarySheet = workbook.addWorksheet();
    let temp = workbook.addWorksheet();
    let nc = workbook.addWorksheet();
    let bol = workbook.addWorksheet();
    let invoice = workbook.addWorksheet();

    secondarySheet.setName("Secondary Sheet");
    nc.setName("Centralizator NC");
    temp.setName("temp");
    bol.setName("Centralizator BOL");
    invoice.setName("Centralizator facturi");

    let sourceWorksheet = workbook.getWorksheet("Get_AllShipments");
    let rangeSource = sourceWorksheet.getUsedRange();
    let lastRowSource = rangeSource.getRowCount() + 2;

    let columnG = sourceWorksheet.getRange("G:G");
    columnG.delete(ExcelScript.DeleteShiftDirection.left);
    let columnX = sourceWorksheet.getRange("X:X");
    columnX.delete(ExcelScript.DeleteShiftDirection.left);
    let columnS = sourceWorksheet.getRange("S:S");
    columnS.delete(ExcelScript.DeleteShiftDirection.left);

    secondarySheet
        .getRange("A1")
        .copyFrom(
            sourceWorksheet.getRange(`8:${lastRowSource}`),
            ExcelScript.RangeCopyType.values,
            false,
            false
        );
    secondarySheet.getRange().getFormat().autofitColumns();

    //CENTRALIZATOR BOL//
    let rangeSecondary = secondarySheet.getUsedRange();
    let lastRowSecondary = rangeSecondary.getRowCount();

    temp
        .getRange("A1")
        .copyFrom(
            secondarySheet.getRange(`V1:V${lastRowSecondary}`),
            ExcelScript.RangeCopyType.values,
            false,
            false
        );
    temp
        .getRange("B1")
        .copyFrom(
            secondarySheet.getRange(`AY1:AY${lastRowSecondary}`),
            ExcelScript.RangeCopyType.values,
            false,
            false
        );
    temp.getRange().getFormat().autofitColumns();
    //Sort columns ascending
    temp.getAutoFilter().apply(temp.getRange("A1"));
    temp
        .getAutoFilter()
        .getRange()
        .getSort()
        .apply([{ key: 0, ascending: true }], false, true);

    consolidateData(temp);

    let tempRange = temp.getUsedRange();
    let tempLastRow = tempRange.getRowCount();

    //Get gross weight in centralizator BOL

    temp
        .getRange("C2")
        .setFormulaLocal("=VLOOKUP(@A:A,'Secondary Sheet'!V:AZ,31,0)");

    if (tempLastRow > 2) {
        temp
            .getRange("C2")
            .autoFill(`C2:C${tempLastRow}`, ExcelScript.AutoFillType.fillDefault);
    }

    //Insert columns

    insertColumns(2, temp, "B:B");

    //Get BOL issue date
    temp
        .getRange("B2")
        .setFormulaLocal("=VLOOKUP(@A:A,'Secondary Sheet'!V:X,3,0)");

    if (tempLastRow > 2) {
        temp
            .getRange("B2")
            .autoFill(`B2:B${tempLastRow}`, ExcelScript.AutoFillType.fillDefault);
    }
    temp.getRange(`B:B`).setNumberFormatLocal("m/d/yyyy");

    //Get BOL No.
    temp
        .getRange("C2")
        .setFormulaLocal("=VLOOKUP(@A:A,'Secondary Sheet'!V:Y,4,0)");

    if (tempLastRow > 2) {
        temp
            .getRange("C2")
            .autoFill(`C2:C${tempLastRow}`, ExcelScript.AutoFillType.fillDefault);
    }

    bol
        .getRange("A1")
        .copyFrom(
            temp.getRange(`C1:E${tempLastRow}`),
            ExcelScript.RangeCopyType.values,
            false,
            false
        );

    consolidateData(bol);
    insertColumns(1, bol, "B:B");
    let bolRange = bol.getUsedRange();
    let bolLastRow = bolRange.getRowCount();
    temp.getRange("F2").setFormulaLocal("=B2");

    if (tempLastRow > 2) {
        temp
            .getRange("F2")
            .autoFill(`F2:F${tempLastRow}`, ExcelScript.AutoFillType.fillDefault);
    }
    bol.getRange("B2").setFormulaLocal("=VLOOKUP(@A:A,'temp'!C:F,4,0)");

    if (bolLastRow > 2) {
        bol
            .getRange("B2")
            .autoFill(`B2:B${bolLastRow}`, ExcelScript.AutoFillType.fillDefault);
    }

    bol.getRange("B:B").setNumberFormatLocal("m/d/yyyy");
    bol
        .getRange("B:B")
        .copyFrom(
            bol.getRange("B:B"),
            ExcelScript.RangeCopyType.values,
            false,
            false
        );

    bol.getRange("A1").setValue("BOL");
    bol.getRange("B1").setValue("BOL Issue Date");
    bol.getRange("C1").setValue("Net Weight (kg)");
    bol.getRange("D1").setValue("Gross Weight (kg)");
    sumValues(bol, bolLastRow, "C");
    sumValues(bol, bolLastRow, "D");

    bol.getRange().getFormat().autofitColumns();
    bol.getRange(`B${bolLastRow + 1}`).setValue("Total:");
    bol
        .getRange(`B${bolLastRow + 1}`)
        .getFormat()
        .setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);

    setColor("A1:D1", bol, "#000000", "#FFFFFF", true);
    setColor(
        `B${bolLastRow + 1}:D${bolLastRow + 1}`,
        bol,
        "#000000",
        "#FFFFFF",
        true
    );
    bol.getFreezePanes().freezeRows(1);
    temp.delete();

    //CENTRALIZATOR FACTURI//

    insertColumns(1, secondarySheet, "Z:Z");
    secondarySheet.getRange("Z2").setFormulaLocal(`=LEFT(AA2,FIND("/",AA2) -1)`);
    secondarySheet
        .getRange("Z2")
        .autoFill(`Z2:Z${lastRowSecondary}`, ExcelScript.AutoFillType.fillDefault);
    secondarySheet
        .getRange("Z2")
        .copyFrom(
            secondarySheet.getRange(`Z2:Z${lastRowSecondary}`),
            ExcelScript.RangeCopyType.values,
            false,
            false
        );
    secondarySheet.getRange("Z:Z").setNumberFormatLocal("0");
    secondarySheet.getRange("BF2").setFormulaLocal(`=A2`);
    secondarySheet
        .getRange("BF2")
        .autoFill(
            `BF2:BF${lastRowSecondary}`,
            ExcelScript.AutoFillType.fillDefault
        );
    invoice
        .getRange("A1")
        .copyFrom(
            secondarySheet.getRange(`Z1:Z${lastRowSecondary}`),
            ExcelScript.RangeCopyType.values,
            false,
            false
        );
    invoice
        .getRange("B1")
        .copyFrom(
            secondarySheet.getRange(`AG1:AG${lastRowSecondary}`),
            ExcelScript.RangeCopyType.values,
            false,
            false
        );


    consolidateData(invoice);

    //TODO: modify the consolidateData function to avoid doing this operation
    let columnToDelete = invoice.getRange("C:C");
    columnToDelete.delete(ExcelScript.DeleteShiftDirection.left);
    insertColumns(1, invoice, "B:B");

    let invoiceRange = invoice.getUsedRange();
    let invoiceLastRow = invoiceRange.getRowCount();

    let rangeInvNo = secondarySheet.getRange(`Z2:Z${lastRowSecondary}`);

    //This part converts cells value type to number. This is necessary for VLOOKUP to work.
    rangeInvNo.getValues().forEach((row, index) => {
        let cell = row[0];
        let value = cell.toString(); // Convert cell value to string
        let numberValue = parseFloat(value); // Convert string to number

        if (!isNaN(numberValue)) {
            let cellAddress = `Z${index + 2}`; // Get cell address
            let cellToUpdate = secondarySheet.getRange(cellAddress); // Get the cell to update
            cellToUpdate.setValue(numberValue); // Set the number value to the cell
        }
    });

    invoice
        .getRange("B2")
        .setFormulaLocal(`=VLOOKUP(@A:A,'Secondary Sheet'!Z:AB,3,0)`);
    invoice
        .getRange("D2")
        .setFormulaLocal(`=VLOOKUP(@A:A,'Secondary Sheet'!Z:BF,33,0)`);
    if (invoiceLastRow > 2) {
        invoice
            .getRange("D2")
            .autoFill(`D2:D${invoiceLastRow}`, ExcelScript.AutoFillType.fillDefault);
        invoice
            .getRange("B2")
            .autoFill(`B2:B${invoiceLastRow}`, ExcelScript.AutoFillType.fillDefault);
    }

    invoice.getRange("B:B").setNumberFormatLocal("m/d/yyyy");
    invoice.getRange("A1").setValue("Invoice Number");
    invoice.getRange("B1").setValue("Invoice Date");
    invoice.getRange("C1").setValue("Invoice Amount USD");
    invoice.getRange("D1").setValue("Customer");
    invoice.getRange(`B${invoiceLastRow + 1}`).setValue("Total:");
    sumValues(invoice, invoiceLastRow, "C");
    invoice.getRange().getFormat().autofitColumns();
    invoice.getRange(`C${invoiceLastRow + 1}`).setNumberFormatLocal("$#,##0.00");
    invoice
        .getRange(`B${invoiceLastRow + 1}`)
        .getFormat()
        .setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
    invoice
        .getRange(`A2:D${invoiceLastRow}`)
        .copyFrom(
            invoice.getRange(`A2:D${invoiceLastRow}`),
            ExcelScript.RangeCopyType.values,
            false,
            false
        );
    setColor("A1:D1", invoice, "#000000", "#FFFFFF", true);
    setColor(
        `B${invoiceLastRow + 1}:C${invoiceLastRow + 1}`,
        invoice,
        "#000000",
        "#FFFFFF",
        true
    );
    invoice.getFreezePanes().freezeRows(1);


    //CENTRALIZATOR NC

    nc.getRange("A1").copyFrom(
        secondarySheet.getRange(`1:${lastRowSecondary}`),
        ExcelScript.RangeCopyType.values,
        false,
        false
    );
    nc.getRange().getFormat().autofitColumns();
    nc.getRange("A:B").delete(ExcelScript.DeleteShiftDirection.left);
    nc.getRange("B:E").delete(ExcelScript.DeleteShiftDirection.left);
    nc.getRange("D:K").delete(ExcelScript.DeleteShiftDirection.left);
    nc.getRange("F:G").delete(ExcelScript.DeleteShiftDirection.left);
    nc.getRange("G:G").delete(ExcelScript.DeleteShiftDirection.left);
    nc.getRange("I:N").delete(ExcelScript.DeleteShiftDirection.left);
    nc.getRange("K:Z").delete(ExcelScript.DeleteShiftDirection.left);
    nc.getRange("O:S").delete(ExcelScript.DeleteShiftDirection.left);


    nc.getRange("O:O").getFormat().setColumnWidth(14.11);
    nc.getAutoFilter().apply(nc.getRange("1:1"));
    nc.getAutoFilter()
        .getRange()
        .getSort()
        .apply([{ key: 5, ascending: true }], false, true);
    nc.getRange("O2").setFormulaLocal('=IF(F2=F1," ",N2)');
    nc.getRange("O2").autoFill(
        `O2:O${lastRowSecondary}`,
        ExcelScript.AutoFillType.fillDefault
    );
    nc.getRange(`O${lastRowSecondary + 1}`).setFormulaLocal(
        `=SUM(O2:O${lastRowSecondary})`
    );
    nc.getRange('O1').setValue('Gross Weight(kg)')
    nc.getRange("O:O").copyFrom(
        nc.getRange("O:O"),
        ExcelScript.RangeCopyType.values,
        false,
        false
    );
    nc.getRange("N:N").delete(ExcelScript.DeleteShiftDirection.left);
    nc.getRange(`M${lastRowSecondary + 1}`).setFormulaLocal(
        `=SUM(M2:M${lastRowSecondary})`
    );
    nc.getFreezePanes().freezeRows(1);

    nc.getRange(`J${lastRowSecondary + 1}`).setFormulaLocal(
        `=SUM(J2:J${lastRowSecondary})`
    );
    nc.getRange(`B${lastRowSecondary + 1}`).setFormulaLocal(
        `=SUM(B2:B${lastRowSecondary})`
    );
    nc.getRange("G:G").setNumberFormatLocal("m/d/yyyy");

    let conditionalFormatting: ExcelScript.ConditionalFormat;

    /////////////////////////////////////////
    /////////POTENTIAL BUG BELOW////////////
    ///////////////////////////////////////

    conditionalFormatting = nc
        .getRange(`N2:N${lastRowSecondary}`)
        .addConditionalFormat(ExcelScript.ConditionalFormatType.presetCriteria);
    conditionalFormatting.getPreset().setRule({
        criterion: ExcelScript.ConditionalFormatPresetCriterion.blanks,
    });
    conditionalFormatting.getPreset().getFormat().getFill().setColor("#dce6f1");
    conditionalFormatting.setStopIfTrue(false);
    conditionalFormatting.setPriority(0);
    nc.getRange("M:N").setNumberFormatLocal("0.00");
    nc.getRange(`J${lastRowSecondary + 1}`).setNumberFormatLocal("$#,##0.00");
    nc.getRange().getFormat().autofitColumns();
    nc.getFreezePanes().freezeRows(1);
    setColor("A1:N1", nc, "#000000", "#FFFFFF", true);
    setColor(`B${lastRowSecondary + 1}`, nc, "#000000", "#FFFFFF", true);
    setColor(`J${lastRowSecondary + 1}`, nc, "#000000", "#FFFFFF", true);
    setColor(
        `M${lastRowSecondary + 1}:N${lastRowSecondary + 1}`,
        nc,
        "#000000",
        "#FFFFFF",
        true
    );
    secondarySheet.delete()

    //Add styling to a range
    function setColor(
        range: string,
        sheet: ExcelScript.Worksheet,
        fill: string,
        fontColor: string,
        bold: boolean
    ) {
        sheet.getRange(`${range}`).getFormat().getFill().setColor("#000000");
        sheet.getRange(`${range}`).getFormat().getFont().setColor("#FFFFFF");
        sheet.getRange(`${range}`).getFormat().getFont().setBold(true);
    }

    //Sum values from a specified column
    function sumValues(
        sheet: ExcelScript.Worksheet,
        lastRow: number,
        column: string
    ) {
        sheet
            .getRange(`${column}${lastRow + 1}`)
            .setFormulaLocal(`=SUM(${column}2:${column}${lastRow})`);
    }

    //This function inserts a specified number of columns at a specified range in a specified sheet
    function insertColumns(
        int: number,
        sheet: ExcelScript.Worksheet,
        range: string
    ) {
        for (let i = 0; i < int; i++) {
            sheet.getRange(range).insert(ExcelScript.InsertShiftDirection.right);
        }
    }

    //This function consolidates data
    function consolidateData(sheet: ExcelScript.Worksheet) {
        let usedRange = sheet.getUsedRange();
        let columnCount = usedRange.getColumnCount();

        let valuesA = usedRange.getColumn(0).getValues();
        let valuesB = usedRange.getColumn(1).getValues();
        let valuesC = columnCount > 2 ? usedRange.getColumn(2).getValues() : [];

        let dictionary: { [key: string]: { valueB: number; valueC: number } } = {};

        for (let i = 1; i < valuesA.length; i++) {
            let valueA = valuesA[i][0].toString();
            let valueB = Number(valuesB[i][0]);
            let valueC = columnCount > 2 ? Number(valuesC[i][0]) : 0;

            if (dictionary.hasOwnProperty(valueA)) {
                dictionary[valueA].valueB += valueB;
                dictionary[valueA].valueC += valueC;
            } else {
                dictionary[valueA] = { valueB, valueC };
            }
        }

        let range = sheet.getRange(`A2:C${usedRange.getRowCount()}`);
        range.clear();

        let outputRange = sheet.getRangeByIndexes(
            0,
            0,
            Object.keys(dictionary).length + 1,
            3
        );
        let outputData: (string | number)[][] = [
            ["Column A", "Column B", "Column C"],
        ];

        for (let key in dictionary) {
            outputData.push([key, dictionary[key].valueB, dictionary[key].valueC]);
        }

        outputRange.setValues(outputData);
    }
}

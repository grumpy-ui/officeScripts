function main(workbook: ExcelScript.Workbook) {
    //Init
    const today = new Date();
    const day = today.getDate();
    const month = today.getMonth() + 1;
    const year = today.getFullYear();
    const sourceSheet = workbook.getWorksheet("Query1");
    const rangeSource = sourceSheet.getUsedRange();
    const lastRowSource = rangeSource.getRowCount();
    const worksheets = workbook.getWorksheets();
    const destinationSheetName = `${day}.${month}.${year}`;
    let destinationSheet = workbook.getWorksheet(destinationSheetName);

    
    worksheets.forEach(sheet => {
        const sheetName = sheet.getName();
        const checkName = sheetName == 'Query1' || sheetName == 'Blocked' ||  sheetName == `LoB.${destinationSheetName}`;
        if (!checkName) {
            sheet.delete()
        }

    })
    destinationSheet = workbook.addWorksheet();
    destinationSheet.setName(destinationSheetName);

    destinationSheet
        .getRange("A1")
        .copyFrom(
            sourceSheet.getRange(`A1:AF${lastRowSource}`),
            ExcelScript.RangeCopyType.values,
            false,
            false
        );
    destinationSheet.getRange().getFormat().autofitColumns();

    let destinationRange = destinationSheet.getUsedRange();
    let destinationTable = destinationSheet.addTable(destinationRange, true);

    const asnFilter: ExcelScript.Filter = destinationTable
        .getColumn(6)
        .getFilter();

    asnFilter.applyValuesFilter(["0"]);
    destinationSheet
        .getRange("2:2")
        .getExtendedRange(ExcelScript.KeyboardDirection.down)
        .delete(ExcelScript.DeleteShiftDirection.up);
    asnFilter.clear();

    const onHandFilter: ExcelScript.Filter = destinationTable
        .getColumn(16)
        .getFilter();
    onHandFilter.applyValuesFilter([""]);
    destinationSheet
        .getRange("2:2")
        .getExtendedRange(ExcelScript.KeyboardDirection.down)
        .delete(ExcelScript.DeleteShiftDirection.up);
    onHandFilter.clear();

    destinationTable.getSort().apply([{ key: 6, ascending: true }]);

    destinationSheet
        .getRange("R:R")
        .insert(ExcelScript.InsertShiftDirection.right);
    destinationSheet.getRange("R1").setValue("To be shipped");
    destinationSheet
        .getRange("S:S")
        .insert(ExcelScript.InsertShiftDirection.right);
    destinationSheet.getRange("S1").setValue("To be packed");

    const lastRow = destinationRange.getRowCount();
    destinationSheet
        .getRange("S2")
        .setFormulaLocal(
            "=IF(N2-O2=0,R2,IF(OR(R2-(N2-O2)>R2,R2-(N2-O2) < 0),R2,R2-(N2-O2)))"
        );

    //For some reason, using autofill doesnt fill in all the cells
    // destinationSheet
    //     .getRange("S2")
    //     .autoFill(`S2:S${lastRow}`, ExcelScript.AutoFillType.fillDefault);
    destinationSheet
        .getRange("N:N")
        .insert(ExcelScript.InsertShiftDirection.right);
    destinationSheet.getRange("N1").setValue("First critical");
    destinationSheet
        .getRange("O:O")
        .insert(ExcelScript.InsertShiftDirection.right);
    destinationSheet.getRange("O1").setValue("First critical DAY");
    destinationSheet
        .getRange("P:P")
        .insert(ExcelScript.InsertShiftDirection.right);
    destinationSheet.getRange("P1").setValue("Ship via");

    destinationSheet
        .getRange("U:U")
        .insert(ExcelScript.InsertShiftDirection.right);
    destinationSheet.getRange("U1").setValue("Blocked");
    destinationSheet
        .getRange(`U1:U${lastRow}`)
        .copyFrom(
            destinationSheet.getRange(`U1:U${lastRow}`),
            ExcelScript.RangeCopyType.values,
            false,
            false
        );

    destinationSheet
        .getRange("U2")
        .setFormulaLocal("=XLOOKUP(@G:G,Blocked!A:A,Blocked!B:B,0)");

    destinationSheet
        .getRange("U2")
        .autoFill(`U2:U${lastRow}`, ExcelScript.AutoFillType.fillDefault);


    const epicorPartNo = destinationRange.getColumn(6).getValues();
    const partNumOccurences = new Map<string, number>();

    const onHandQty = destinationRange.getColumn(18).getValues();
    const asnQty = destinationRange.getColumn(5).getValues();
    const openDemand = destinationRange.getColumn(23).getValues();
    const blocked = destinationRange.getColumn(20).getValues();

    for (let i = 1; i < epicorPartNo.length; i++) {
        const currPartNo = epicorPartNo[i][0] as string;
        if (partNumOccurences.has(currPartNo)) {
            partNumOccurences.set(currPartNo, partNumOccurences.get(currPartNo)! + 1);
        } else {
            partNumOccurences.set(currPartNo, 1);
        }
    }
    let row = 1;
    // @ts-ignore
    for (let [part, occurence] of partNumOccurences) {
        let toBeShipped: number;
        let onHand = onHandQty[row][0] as number;
        let blockedQTY = blocked[row][0] as number;

        for (let i = 0; i < occurence; i++) {
            const toBeShippedCell = destinationSheet.getRange(`V${row + 1}`);
            let asn = asnQty[row][0] as number;
            let demand = openDemand[row][0] as number;


            if (onHand < 1 || blockedQTY >= onHand) {
                toBeShippedCell.setValue(0);
                toBeShipped = 0;
            } else {
                let value = Math.min(onHand - blockedQTY, asn, demand);
                toBeShippedCell.setValue(value);

                if (blockedQTY > 0) {
                    blockedQTY += value
                } else {
                    onHand -= value;
                }
            }
            row++
        }
    }

    destinationSheet
        .getRange("N2")
        .setFormulaLocal(
            `=XLOOKUP(E2,LoB.${day}.${month}.${year}!A:A,LoB.${day}.${month}.${year}!AC:AC,"Not critical"`
        );
    destinationSheet
        .getRange("O2")
        .setFormulaLocal(
            `=IF(OR(ISBLANK(N2),N2=""),"NOT CRITICAL",IF(N2="Current","Current",IF(N2="Not critical","Not critical",IFNA(DATE(IF(ISNUMBER(SEARCH(",", N2)),VALUE(RIGHT(N2, 4)),IF(DATEVALUE(LEFT(N2, 3) & " " & MID(N2, 5, 2) & ", " & YEAR(TODAY())) < TODAY(),2025,IF(DATEVALUE(LEFT(N2, 3) & " " & MID(N2, 5, 2) & ", " & YEAR(TODAY())) > TODAY(),2024,YEAR(TODAY())))),MATCH(LEFT(N2,3), {"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"}, 0),MID(N2, 5, 2))-7,"Not critical"))))`
        );
    destinationSheet.getRange("O:O").setNumberFormat("dd/mm/yyyy");
    destinationSheet
        .getRange("P2")
        .setFormulaLocal(
            `=IFS(O2="Not critical","SEA",AND(O2="Current",Y2<=2600),"UPS",AND(O2="Current",Y2>2600),"AIR",O2-TODAY()>=45,"SEA",AND(O2-TODAY()<45,Y2>2600),"AIR",AND(O2-TODAY()<45,Y2<=2600,O2>TODAY()),"UPS",O2<TODAY(),"SEA")`
        );

    // destinationSheet
    //     .getRange("P2")
    //     .autoFill(`P2:P${lastRow}`, ExcelScript.AutoFillType.fillDefault);





    //This line makes sure no #REF error is being thrown out after deleting a referenced sheet
    // destinationSheet
    //     .getRange(`N1:P${lastRow}`)
    //     .copyFrom(
    //         destinationSheet.getRange(`N1:P${lastRow}`),
    //         ExcelScript.RangeCopyType.values,
    //         false,
    //         false
    //     );



        destinationSheet
        .getRange("N2")
        .autoFill(`N2:N${lastRow}`, ExcelScript.AutoFillType.fillDefault);


        destinationSheet
        .getRange("O2")
        .autoFill(`O2:O${lastRow}`, ExcelScript.AutoFillType.fillDefault);


    destinationSheet
        .getRange("W2")
        .autoFill(`W2:W${lastRow}`, ExcelScript.AutoFillType.fillDefault);

 

    //CONDITIONAL FORMATTING
    let conditionalFormatting: ExcelScript.ConditionalFormat;

    conditionalFormatting = destinationSheet
        .getRange("P:P")
        .addConditionalFormat(ExcelScript.ConditionalFormatType.containsText);
    conditionalFormatting.getTextComparison().setRule({
        operator: ExcelScript.ConditionalTextOperator.contains,
        text: "UPS",
    });
    conditionalFormatting
        .getTextComparison()
        .getFormat()
        .getFill()
        .setColor("#ffeb9c");
    conditionalFormatting
        .getTextComparison()
        .getFormat()
        .getFont()
        .setColor("#9c5700");
    conditionalFormatting.setStopIfTrue(false);
    conditionalFormatting.setPriority(0);

    conditionalFormatting = destinationSheet
        .getRange("P:P")
        .addConditionalFormat(ExcelScript.ConditionalFormatType.containsText);
    conditionalFormatting.getTextComparison().setRule({
        operator: ExcelScript.ConditionalTextOperator.contains,
        text: "AIR",
    });
    conditionalFormatting
        .getTextComparison()
        .getFormat()
        .getFill()
        .setColor("#ffc7ce");
    conditionalFormatting
        .getTextComparison()
        .getFormat()
        .getFont()
        .setColor("#9c0006");
    conditionalFormatting.setStopIfTrue(false);
    conditionalFormatting.setPriority(0);

    conditionalFormatting = destinationSheet
        .getRange("P:P")
        .addConditionalFormat(ExcelScript.ConditionalFormatType.containsText);
    conditionalFormatting.getTextComparison().setRule({
        operator: ExcelScript.ConditionalTextOperator.contains,
        text: "SEA",
    });
    conditionalFormatting
        .getTextComparison()
        .getFormat()
        .getFill()
        .setColor("#c6efce");
    conditionalFormatting
        .getTextComparison()
        .getFormat()
        .getFont()
        .setColor("#006100");
    conditionalFormatting.setStopIfTrue(false);
    conditionalFormatting.setPriority(0);

    // Create cell value criteria for range U:U on selectedSheet
    conditionalFormatting = destinationSheet
        .getRange("U:U")
        .addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
    conditionalFormatting.getCellValue().setRule({
        operator: ExcelScript.ConditionalCellValueOperator.greaterThan,
        formula1: "=0",
    });
    conditionalFormatting
        .getCellValue()
        .getFormat()
        .getFill()
        .setColor("#ffc7ce");
    conditionalFormatting
        .getCellValue()
        .getFormat()
        .getFont()
        .setColor("#9c0006");
    conditionalFormatting.setStopIfTrue(false);
    conditionalFormatting.setPriority(0);

}
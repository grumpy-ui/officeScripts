function calculateShippingList(workbook: ExcelScript.Workbook) {
  //Init
  const today = new Date();
  const day = today.getDate();
  const month = today.getMonth() + 1;
  const year = today.getFullYear();
  const sourceSheet = workbook.getWorksheet("Query1");
  const rangeSource = sourceSheet.getUsedRange();
  const lastRowSource = rangeSource.getRowCount();
  const toBeShippedDetailedSheetName = `${day}.${month}.${year}`;
  let toBeShippedSheet = workbook.addWorksheet("ToBeShipped");
  let toBeShippedDetailedSheet = workbook.addWorksheet();
  toBeShippedDetailedSheet.setName(toBeShippedDetailedSheetName);

  //Move values from Query1
  copyValues(
    toBeShippedDetailedSheet,
    "A1",
    sourceSheet,
    `A1:AF${lastRowSource}`
  );
  toBeShippedDetailedSheet.getRange().getFormat().autofitColumns();

  //Format sheet as table so filters can work on it
  let toBeShippedDetRange = toBeShippedDetailedSheet.getUsedRange();
  let toBeShippedDetTable = toBeShippedDetailedSheet.addTable(
    toBeShippedDetRange,
    true
  );

  //Delete empty or 0 on hand values
  deleteFilteredValues(toBeShippedDetTable, 6, "0", toBeShippedDetailedSheet);
  deleteFilteredValues(toBeShippedDetTable, 16, "", toBeShippedDetailedSheet);

  toBeShippedDetTable.getSort().apply([{ key: 6, ascending: true }]);

  //Insert additional columns
  const columnNames = {
    R: "To be shipped",
    S: "To be packed",
    N: "First critical",
    O: "First critical DAY",
    P: "Ship via",
    U: "Blocked",
  };

  const insertColumns = function (
    sheet: ExcelScript.Worksheet,
    column: string,
    columnName: string
  ) {
    insertNamedColumn(sheet, column, columnName);
  };

  for (const [column, columnName] of Object.entries(columnNames)) {
    insertColumns(toBeShippedDetailedSheet, column, columnName);
  }
  const lastRow = toBeShippedDetRange.getRowCount();

  copyValues(
    toBeShippedDetailedSheet,
    `U1:U${lastRow}`,
    toBeShippedDetailedSheet,
    `U1:U${lastRow}`
  );

  const setToBeShippedFormulas = function () {
    const lastRow = toBeShippedDetailedSheet.getUsedRange().getRowCount();
    const formulas = {
      W: "=IF(Q2-R2=0,V2,IF(OR(V2-(Q2-R2)>V2,V2-(Q2-R2) < 0),V2,V2-(Q2-R2)))",
      U: "=XLOOKUP(@G:G,Blocked!A:A,Blocked!B:B,0)",
      N: `=XLOOKUP(E2,'Line Of Balance Weekly'!A:A,'Line Of Balance Weekly'!AC:AC,"Not critical",,-1`,
      O: `=IF(Z2="Spirit Firm Serial PO",TEXT(M2,"m/d/yyyy"),IF(OR(ISBLANK(N2),N2=""),"NOT CRITICAL",IF(N2="Current","Current",IF(N2="Not critical","Not critical",IFNA(DATE(IF(ISNUMBER(SEARCH(",", N2)),VALUE(RIGHT(N2, 4)),IF(DATEVALUE(LEFT(N2, 3) & " " & MID(N2, 5, 2) & ", " & YEAR(TODAY())) < TODAY(),2025,IF(DATEVALUE(LEFT(N2, 3) & " " & MID(N2, 5, 2) & ", " & YEAR(TODAY())) > TODAY(),2024,YEAR(TODAY())))),MATCH(LEFT(N2,3), {"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"}, 0),MID(N2, 5, 2))-7,"Not critical")))))`,
      P: `=IFS(O2="Not critical","SEA",AND(O2="Current",Y2<=2600),"UPS",AND(O2="Current",Y2>2600),"AIR",O2-TODAY()>=55,"SEA",AND(O2-TODAY()<55,Y2>2600),"AIR",AND(O2-TODAY()<55,Y2<=2600,O2>TODAY()),"UPS",AND(O67<TODAY(),Y2 > 2600),"AIR",AND(O2<TODAY(),Y2<=2600),"UPS")`,
    };

    Object.entries(formulas).forEach(([cell, formula]) => {
      const cellRange = `${cell}2`;
      toBeShippedDetailedSheet.getRange(cellRange).setFormula(formula);
      toBeShippedDetailedSheet
        .getRange(cellRange)
        .autoFill(
          `${cellRange}:${cellRange[0]}${lastRow}`,
          ExcelScript.AutoFillType.fillDefault
        );
    });
  };

  setToBeShippedFormulas();

  //This part calculates the appropriate to be shipped Qty
  const epicorPartNo = toBeShippedDetRange.getColumn(6).getValues();
  const partNumOccurences = new Map<string, number>();

  const onHandQty = toBeShippedDetRange.getColumn(18).getValues();
  const asnQty = toBeShippedDetRange.getColumn(5).getValues();
  const openDemand = toBeShippedDetRange.getColumn(23).getValues();
  const blocked = toBeShippedDetRange.getColumn(20).getValues();

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
      const toBeShippedCell = toBeShippedDetailedSheet.getRange(`V${row + 1}`);
      let asn = asnQty[row][0] as number;
      let demand = openDemand[row][0] as number;

      if (onHand < 1 || blockedQTY >= onHand) {
        toBeShippedCell.setValue(0);
        toBeShipped = 0;
      } else {
        let value = Math.min(onHand - blockedQTY, asn, demand);
        toBeShippedCell.setValue(value);

        if (blockedQTY > 0) {
          blockedQTY += value;
        } else {
          onHand -= value;
        }
      }
      row++;
    }
  }

  toBeShippedDetailedSheet.getRange("O:O").setNumberFormat("dd/mm/yyyy");

  copyValues(toBeShippedSheet, "A1", toBeShippedDetailedSheet, "E:E");
  copyValues(toBeShippedSheet, "B1", toBeShippedDetailedSheet, "V:V");
}

function main(workbook: ExcelScript.Workbook) {
  //init
  calculateShippingList(workbook);
  let sourceSheet = workbook.getWorksheet("ASN");
  let overviewSheet = workbook.addWorksheet("Overview");
  let demandRO = workbook.getWorksheet("SummarisationRO");
  let demandVN = workbook.getWorksheet("SummarisationVN");
  let toBeShipped = workbook.getWorksheet("ToBeShipped");
  let asnReportSheet = workbook.addWorksheet("ASN Report");
  let summary = workbook.addWorksheet("Summary");

  let demandRORange = demandRO.getUsedRange();
  let lastRowDemandRO = demandRORange.getRowCount();
  let demandVNRange = demandVN.getUsedRange();
  let lastRowDemandVN = demandVNRange.getRowCount();

  demandVN.getRange(`E2:E${lastRowDemandVN}`).setValues("UACV");
  demandRO.getRange(`E2:E${lastRowDemandRO}`).setValues("UACE");

  let destinationRO = `A${lastRowDemandRO + 1}`;
  let sourceVNRange = `A2:E${lastRowDemandVN}`;

  copyValues(demandRO, destinationRO, demandVN, sourceVNRange);
  let temp = workbook.addWorksheet();
  temp.setName("temp");

  let lastRowDemandROVN = lastRowDemandRO + lastRowDemandVN;
  let sourceROVN = `A1:E${lastRowDemandROVN}`;

  copyValues(temp, "A1", demandRO, sourceROVN);

  let deleteB = temp.getRange("B:B");
  let deleteC = temp.getRange("C:C");
  deleteB.delete(ExcelScript.DeleteShiftDirection.left);
  deleteC.delete(ExcelScript.DeleteShiftDirection.left);

  let tempRange = temp.getUsedRange();
  let tempLastRow = tempRange.getRowCount();

  consolidateDataString(temp, "CustPart#", "Available", "Ships from");
  temp.getRange("D2").setFormulaLocal('=IF(C2="UACEUACV","Both",C2)');
  temp
    .getRange("D2")
    .autoFill(`D2:D${tempLastRow}`, ExcelScript.AutoFillType.fillDefault);
  temp
    .getRange(`D2:D${tempLastRow}`)
    .copyFrom(
      temp.getRange(`D2:D${tempLastRow}`),
      ExcelScript.RangeCopyType.values,
      false,
      false
    );
  deleteC = temp.getRange("C:C");
  deleteC.delete(ExcelScript.DeleteShiftDirection.left);

  //calculate last row
  const rangeSource = sourceSheet.getUsedRange();
  const lastRowSource = rangeSource.getRowCount();

  //copy values from source to ASN report
  asnReportSheet
    .getRange("A1")
    .copyFrom(
      sourceSheet.getRange(`A1:H${lastRowSource}`),
      ExcelScript.RangeCopyType.values,
      false,
      false
    );
  asnReportSheet.getRange().getFormat().autofitColumns();
  asnReportSheet.getRange(`H:H`).setNumberFormatLocal("m/d/yyyy");

  let asnHeader = {
    A1: "Supplier",
    B1: "PO",
    C1: "PO Line",
    D1: "CustPart#",
    E1: "Report Date",
    F1: "PO Status",
    G1: "ASN Qty [pcs]",
    H1: "Date",
  };

  let overviewHeader = {
    A1: "CustPart#",
    B1: "ASN Qty [pcs]",
    C1: "Epicor Demand",
    D1: "Ships From",
    E1: "On Hand UACE",
    F1: "On Hand CW",
    G1: "Available to ship UACE",
    H1: "Available to ship CW",
    I1: "Unit price",
    J1: "Available to invoice UACE",
    K1: "Available to invoice CW",
    L1: "Total ASN Value",
    M1: "On Hand value UACE",
    N1: "On Hand value CW",
  };

  Object.entries(asnHeader).forEach(([cell, value]) =>
    setValues(asnReportSheet, cell, value)
  );

  copyValues(overviewSheet, "A1", asnReportSheet, "D:D");
  copyValues(overviewSheet, "B1", asnReportSheet, "G:G");

  consolidateData(overviewSheet, "CustPart#", "ASN Qty [pcs]", "_");
  let rangeOverview = overviewSheet.getUsedRange();
  let lastRowOverview = rangeOverview.getRowCount();
  Object.entries(overviewHeader).forEach(([cell, value]) =>
    setValues(overviewSheet, cell, value)
  );

  overviewSheet
    .getRange("C2")
    .setFormulaLocal("=IFNA(VLOOKUP(@A:A,temp!A:B,2,0),0)");
  overviewSheet
    .getRange("C2")
    .autoFill(`C2:C${lastRowOverview}`, ExcelScript.AutoFillType.fillDefault);

  overviewSheet
    .getRange("D2")
    .setFormulaLocal('=IFNA(VLOOKUP(@A:A,temp!A:C,3,0),"UACE")');
  overviewSheet
    .getRange("D2")
    .autoFill(`D2:D${lastRowOverview}`, ExcelScript.AutoFillType.fillDefault);

  overviewSheet
    .getRange("E2")
    .setFormulaLocal("=XLOOKUP(@A:A,OnHandRO!A:A,OnHandRO!B:B,0)");
  overviewSheet
    .getRange("F2")
    .setFormulaLocal("=XLOOKUP(@A:A,OnHandCW!A:A,OnHandCW!B:B,0)");

  overviewSheet
    .getRange("E2")
    .autoFill(`E2:E${lastRowOverview}`, ExcelScript.AutoFillType.fillDefault);

  overviewSheet
    .getRange("F2")
    .autoFill(`F2:F${lastRowOverview}`, ExcelScript.AutoFillType.fillDefault);

  consolidateData(toBeShipped, "Part", "To be shipped", "_");
  overviewSheet
    .getRange("G2")
    .setFormulaLocal("=IFNA(VLOOKUP(@A:A,ToBeShipped!A:B,2,0),0)");
  overviewSheet
    .getRange("G2")
    .autoFill(`G2:G${lastRowOverview}`, ExcelScript.AutoFillType.fillDefault);

  overviewSheet.getRange("H2").setFormulaLocal("=MIN(F2,C2)");
  overviewSheet
    .getRange("H2")
    .autoFill(`H2:H${lastRowOverview}`, ExcelScript.AutoFillType.fillDefault);

  overviewSheet
    .getRange("I2")
    .setFormulaLocal("=VLOOKUP(@A:A,PriceList!A:B,2,0)");
  overviewSheet
    .getRange("I2")
    .autoFill(`I2:I${lastRowOverview}`, ExcelScript.AutoFillType.fillDefault);

  //Values can be copied at the end for all columns
  overviewSheet
    .getRange(`C2:I${lastRowOverview}`)
    .copyFrom(
      overviewSheet.getRange(`C2:I${lastRowOverview}`),
      ExcelScript.RangeCopyType.values,
      false,
      false
    );

  overviewSheet.getRange("J2").setFormulaLocal("=G2*I2");
  overviewSheet.getRange("K2").setFormulaLocal("=H2*I2");
  overviewSheet.getRange("L2").setFormulaLocal("=B2*I2");
  overviewSheet.getRange("M2").setFormulaLocal("=E2*I2");
  overviewSheet.getRange("N2").setFormulaLocal("=F2*I2");

  overviewSheet
    .getRange("J2")
    .autoFill(`J2:J${lastRowOverview}`, ExcelScript.AutoFillType.fillDefault);
  overviewSheet
    .getRange("K2")
    .autoFill(`K2:K${lastRowOverview}`, ExcelScript.AutoFillType.fillDefault);
  overviewSheet
    .getRange("L2")
    .autoFill(`L2:L${lastRowOverview}`, ExcelScript.AutoFillType.fillDefault);
  overviewSheet
    .getRange("M2")
    .autoFill(`M2:M${lastRowOverview}`, ExcelScript.AutoFillType.fillDefault);
  overviewSheet
    .getRange("N2")
    .autoFill(`N2:N${lastRowOverview}`, ExcelScript.AutoFillType.fillDefault);

  overviewSheet.getRange().getFormat().autofitColumns();
  overviewSheet
    .getRange("I:N")
    .setNumberFormatLocal('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)');

  overviewSheet.getRange("A1:N1").getFormat().getFill().setColor("5B9BD5");
  overviewSheet.getRange("A1:N1").getFormat().getFont().setColor("FFFFFF");
  overviewSheet.getRange("A1:N1").getFormat().getFont().setBold(true);

  asnReportSheet.getRange().getFormat().autofitColumns();

  asnReportSheet.getRange("A1:H1").getFormat().getFill().setColor("70AD47");
  asnReportSheet.getRange("A1:H1").getFormat().getFont().setColor("FFFFFF");
  asnReportSheet.getRange("A1:H1").getFormat().getFont().setBold(true);

  //Add data in the Summary sheet
  let asnReportSheetUsedRange = asnReportSheet.getUsedRange();
  let lastRowAsnReportSheet = asnReportSheetUsedRange.getRowCount();

  //Make this a function
  summary
    .getRange("B1")
    .setFormulaLocal(`=SUM(Overview!J2:J${lastRowAsnReportSheet})`);
  summary
    .getRange("B2")
    .setFormulaLocal(`=SUM(Overview!G2:G${lastRowAsnReportSheet})`);
  summary
    .getRange("B3")
    .setFormulaLocal(`=SUM(Overview!N2:N${lastRowAsnReportSheet})`);
  summary
    .getRange("B4")
    .setFormulaLocal(`=SUM(Overview!M2:M${lastRowAsnReportSheet})`);
  summary
    .getRange("B5")
    .setFormulaLocal(`=SUM(Overview!B2:B${lastRowAsnReportSheet})`);
  summary
    .getRange("B6")
    .setFormulaLocal(`=SUMIF(Overview!D:D,"UACE",Overview!L:L)`);
  summary
    .getRange("B7")
    .setFormulaLocal(`=SUM(Overview!L2:L${lastRowAsnReportSheet})`);

  //Get unit price from demand and delete #N/A values

  rangeOverview = overviewSheet.getUsedRange();
  const table = overviewSheet.addTable(rangeOverview, true);
  const filterUnitPrice: ExcelScript.Filter = table.getColumn(9).getFilter();

  filterUnitPrice.applyValuesFilter(["#N/A"]);

  let columnIRange = overviewSheet.getUsedRange();
  let visibleCells = columnIRange.getVisibleView().getCellAddresses();
  let formula = `=XLOOKUP(@A:A,PriceListFromOpenDemand!A:A,PriceListFromOpenDemand!B:B)`;

  for (let i = 1; i < visibleCells.length; i++) {
    const cell = visibleCells[i][8];
    overviewSheet.getRange(cell).setFormula(formula);
  }

  filterUnitPrice.applyValuesFilter(["#N/A"]);
  overviewSheet
    .getRange("2:2")
    .getExtendedRange(ExcelScript.KeyboardDirection.down)
    .delete(ExcelScript.DeleteShiftDirection.up);
  filterUnitPrice.clear();

  lastRowOverview = overviewSheet.getUsedRange().getRowCount();
  copyValues(
    overviewSheet,
    `I2:I${lastRowOverview}`,
    overviewSheet,
    `I2:I${lastRowOverview}`
  );

  const currFormatRanges = ["B1", "B3", "B4", "B6", "B7"];

  currFormatRanges.forEach((range) => {
    summary.getRange(range).setNumberFormat("$0,00");
  });

  summary.getRange("B2").setNumberFormat("0,00");
  summary.getRange("B5").setNumberFormat("0,00");

  setValues(summary, "A1", "Available to invoice");
  setValues(summary, "A2", "Available to ship (PCS)");
  setValues(summary, "A3", "On Hand value CW");
  setValues(summary, "A4", "On Hand value RO");
  setValues(summary, "A5", "Total ASN Quantity (PCS)");
  setValues(summary, "A6", "Total ASN Value UACE");
  setValues(summary, "A7", "Total ASN Value");

  summary.getRange().getFormat().autofitColumns();

  summary.getRange("A1:A4").getFormat().getFill().setColor("DDEBF7");
  summary.getRange("A5:A7").getFormat().getFill().setColor("FFF2CC");
  summary.getRange("A1:A7").getFormat().getFont().setBold(true);

  summary
    .getRange("A1:B7")
    .getFormat()
    .getRangeBorder(ExcelScript.BorderIndex.diagonalDown)
    .setStyle(ExcelScript.BorderLineStyle.none);
  // Set border for range A1:B7 on selectedSheet
  summary
    .getRange("A1:B7")
    .getFormat()
    .getRangeBorder(ExcelScript.BorderIndex.diagonalUp)
    .setStyle(ExcelScript.BorderLineStyle.none);
  // Set border for range A1:B7 on selectedSheet
  summary
    .getRange("A1:B7")
    .getFormat()
    .getRangeBorder(ExcelScript.BorderIndex.edgeLeft)
    .setStyle(ExcelScript.BorderLineStyle.continuous);
  summary
    .getRange("A1:B7")
    .getFormat()
    .getRangeBorder(ExcelScript.BorderIndex.edgeLeft)
    .setWeight(ExcelScript.BorderWeight.thin);
  // Set border for range A1:B7 on selectedSheet
  summary
    .getRange("A1:B7")
    .getFormat()
    .getRangeBorder(ExcelScript.BorderIndex.edgeTop)
    .setStyle(ExcelScript.BorderLineStyle.continuous);
  summary
    .getRange("A1:B7")
    .getFormat()
    .getRangeBorder(ExcelScript.BorderIndex.edgeTop)
    .setWeight(ExcelScript.BorderWeight.thin);
  // Set border for range A1:B7 on selectedSheet
  summary
    .getRange("A1:B7")
    .getFormat()
    .getRangeBorder(ExcelScript.BorderIndex.edgeBottom)
    .setStyle(ExcelScript.BorderLineStyle.continuous);
  summary
    .getRange("A1:B7")
    .getFormat()
    .getRangeBorder(ExcelScript.BorderIndex.edgeBottom)
    .setWeight(ExcelScript.BorderWeight.thin);
  // Set border for range A1:B7 on selectedSheet
  summary
    .getRange("A1:B7")
    .getFormat()
    .getRangeBorder(ExcelScript.BorderIndex.edgeRight)
    .setStyle(ExcelScript.BorderLineStyle.continuous);
  summary
    .getRange("A1:B7")
    .getFormat()
    .getRangeBorder(ExcelScript.BorderIndex.edgeRight)
    .setWeight(ExcelScript.BorderWeight.thin);
  // Set border for range A1:B7 on selectedSheet
  summary
    .getRange("A1:B7")
    .getFormat()
    .getRangeBorder(ExcelScript.BorderIndex.insideVertical)
    .setStyle(ExcelScript.BorderLineStyle.continuous);
  summary
    .getRange("A1:B7")
    .getFormat()
    .getRangeBorder(ExcelScript.BorderIndex.insideVertical)
    .setWeight(ExcelScript.BorderWeight.thin);
  // Set border for range A1:B7 on selectedSheet
  summary
    .getRange("A1:B7")
    .getFormat()
    .getRangeBorder(ExcelScript.BorderIndex.insideHorizontal)
    .setStyle(ExcelScript.BorderLineStyle.continuous);
  summary
    .getRange("A1:B7")
    .getFormat()
    .getRangeBorder(ExcelScript.BorderIndex.insideHorizontal)
    .setWeight(ExcelScript.BorderWeight.thin);

  //Delete unnecesarry sheets

  const worksheets = workbook.getWorksheets();
  const mainWorksheets = ["Overview", "ASN Report", "Summary"];

  worksheets.forEach((sheet) => {
    const sheetName = sheet.getName();
    if (!mainWorksheets.includes(sheetName)) {
      sheet.delete();
    }
  });

  console.log("SCRIPTUL A RULAT CU SUCCES!");
}

//Helper functions

function setValues(sheet: ExcelScript.Worksheet, range: string, value: string) {
  sheet.getRange(range).setValue(value);
}

function copyValues(
  destination: ExcelScript.Worksheet,
  destinationCell: string,
  source: ExcelScript.Worksheet,
  sourceRange: string
) {
  destination
    .getRange(destinationCell)
    .copyFrom(
      source.getRange(sourceRange),
      ExcelScript.RangeCopyType.values,
      false,
      false
    );
}

function consolidateData(
  sheet: ExcelScript.Worksheet,
  columnA: string,
  columnB: string,
  columnC: string
) {
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
  let outputData: (string | number)[][] = [[columnA, columnB, columnC]];

  for (let key in dictionary) {
    outputData.push([key, dictionary[key].valueB, dictionary[key].valueC]);
  }

  outputRange.setValues(outputData);
}

function consolidateDataString(
  sheet: ExcelScript.Worksheet,
  columnA: string,
  columnB: string,
  columnC: string
) {
  let usedRange = sheet.getUsedRange();
  let columnCount = usedRange.getColumnCount();

  let valuesA = usedRange.getColumn(0).getValues();
  let valuesB = usedRange.getColumn(1).getValues();
  let valuesC = usedRange.getColumn(2).getValues();

  let dictionary: {
    [key: string]: { valueB: number; valueC: number | string };
  } = {};

  for (let i = 1; i < valuesA.length; i++) {
    let valueA = valuesA[i][0].toString();
    let valueB = Number(valuesB[i][0]);
    let valueC = valuesC[i][0];

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
  let outputData: (string | number)[][] = [[columnA, columnB, columnC]];

  for (let key in dictionary) {
    outputData.push([key, dictionary[key].valueB, dictionary[key].valueC]);
  }

  outputRange.setValues(outputData);
}

function deleteFilteredValues(
  destTable: ExcelScript.Table,
  column: number,
  filterValues: string,
  sheet: ExcelScript.Worksheet
) {
  const filter: ExcelScript.Filter = destTable.getColumn(column).getFilter();
  filter.applyValuesFilter([filterValues]);
  sheet
    .getRange("2:2")
    .getExtendedRange(ExcelScript.KeyboardDirection.down)
    .delete(ExcelScript.DeleteShiftDirection.up);
  filter.clear();
}

function insertNamedColumn(
  sheet: ExcelScript.Worksheet,
  column: string,
  name: string
) {
  sheet
    .getRange(`${column}:${column}`)
    .insert(ExcelScript.InsertShiftDirection.right);
  setValues(sheet, `${column}1`, name);
}

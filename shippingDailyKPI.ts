function main(workbook: ExcelScript.Workbook) {
  //init
  let sourceWorksheet = workbook.getWorksheet("Get_AllOpenOrdersShipping");
  let invoiceProjection = workbook.getWorksheets()[1];
  let invoiceProjectionName = invoiceProjection.getName();
  let pivotSheet = workbook.addWorksheet("Summary");
  let firstFourRows = sourceWorksheet.getRange("1:4");
  firstFourRows.delete(ExcelScript.DeleteShiftDirection.up);

  //initial value of T1 doesn't work with the deleteFilteredRows() function
  sourceWorksheet.getRange("T1").setValue("OnHand Distributed");

  let usedRange: ExcelScript.Range = sourceWorksheet.getUsedRange();
  let lastRow = usedRange.getRowCount();

  //Filter parts not in stock
  const range = sourceWorksheet.getRange(`A1:BQ${lastRow}`);
  let table = sourceWorksheet.addTable(range, true);
  table.setPredefinedTableStyle("TableStyleMedium2");

  deleteFilteredRows(["OnHand Distributed"], [["", "0"]], table);
  deleteFilteredRows(
      ["Customer Name"],
      [
          ["Spirit AeroSystems Inc.", "Spirit AeroSystem Oklahoma"],
      ],
      table
  );

  let columnB = sourceWorksheet.getRange(`B1:B${lastRow}`)
  let values: string[][] = columnB.getValues() as string[][]

  let programs: { [key: string]: string } = {
    "Sirius" : "SIRIUS", "PAG Profile Parts": "PAG Profile Parts & SA_LR", "AIRBUS A350 Crash Rod": "AIRBUS Crash Rod", "BOMBARDIER SHORTS BEAM ASSY": "BOMBARDIER", "Deharde A321 QSF-B Couplings": "DEHARDE", "DEHARDE A340 LR Parts": "DEHARDE", "Orizon  Aerostructures-Chanute INC": "SPIRIT RO", "SPIRIT SUNSHINE": "SPIRIT RO", "Pilatus PC-12 Seat Rails": "Pilatus", "SONACA C-Series Trailing Edge": "Sonaca", "Sonaca Embraer Trailing Edge": "Sonaca", "Sonaca Girders": "Sonaca", "Sonaca Trailing Edge": "Sonaca", "Telair International": "Telair", "TE A350 Raceways": "Tyco", "SPIRIT BOEING": "SPIRIT RO", "SPIRIT-BOEING-777 CFB": "SPIRIT RO", "BOMBARDIER CHALLENGER STRING.": "BOMBARDIER", "Deharde Maschinenbau": "DEHARDE", "Magellan A350 Struts": "Magellan", "MECACHROME Fittings": "MECACHROME", "Pilatus PC - 24 Seat Rail.": 'Pilatus', "Premium Aerotec Cross Beam": "PAG Cross Beams FQT", "Spirit Composite": "SPIRIT RO", "SPIRIT KEEL BEAM": "SPIRIT RO", "Triumph A330 Leading Edge": "Triumph A330", "Sogerma A400M Channels": "Sogerma A400M", "Mecachrome": "MECACHROME", "AIRBUS A350 PLR": "AIRBUS - A350 PLR"
  }


  values.forEach((row, index) => {
      let program = row[0];
      let cell = sourceWorksheet.getRange(`B${index + 1}`)
      if (programs[program]) {
          cell.setValue(programs[program])
      }
  })

  let columnsToDelete = [
      "J:K",
      "L:M",
      "M:O",
      "N:Q",
      "N:R",
      "S:U",
      "U:U",
      "W:X",
      "Z:Z",
      "AA:AB",
      "AD:AG",
      "AF:AH",
      "AG:AK",
  ];

  deleteColumns(columnsToDelete);

  insertFormulaColumn("K", "=MIN(N2,M2")
  insertFormulaColumn(
      "AF",
      `=IFERROR(VALUE(MID(AH2,FIND("-",AE2),IFERROR(FIND(",",AE2)-FIND("-",AE2),5))),0)`
  );
  insertFormulaColumn("L", "=IF(H2 < TODAY()+AG2,MIN(N2,O2),0)");
  insertFormulaColumn("P", "=U2/N2");
  insertFormulaColumn("S", "=P2*L2");
  insertFormulaColumn("S", "=P2*K2");

  lastRow = usedRange.getRowCount();

  sourceWorksheet.getRange("S1").setValue("Total value (by Demand)");
  sourceWorksheet.getRange("K1").setValue("Available to ship (by Demand)");
  sourceWorksheet.getRange("L1").setValue("Available to ship (by Need date)");
  sourceWorksheet.getRange("P1").setValue("Unit Price EUR");
  sourceWorksheet.getRange("T1").setValue("Total Value (by Need date)");
  sourceWorksheet.getRange("AJ1").setValue("Days pre need-date");

  
  //create pivot table
  let table1 = workbook.getTable("Table1");
  let newPivotTable = workbook.addPivotTable(
      "PivotTable1",
      table1,
      pivotSheet.getRange("A1")
  );
  newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Program"));
  newPivotTable.addDataHierarchy(
      newPivotTable.getHierarchy("Available to ship (by Need date)")
  );
  newPivotTable.addDataHierarchy(
      newPivotTable.getHierarchy("Available to ship (by Demand)")
  );
  newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Total Value (by Need date)"));
  newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Total Value (by Demand)"));
  pivotSheet.getRange("D:E").setNumberFormatLocal("[$€-x-euro2] #,##0.00");

  let summaryRange = pivotSheet.getUsedRange();
  let lastRowSummary = summaryRange.getRowCount();

  pivotSheet.getRange('F2').setFormulaLocal(`=XLOOKUP(A2,'${invoiceProjectionName}'!A:A,'${invoiceProjectionName}'!B:B),"Please enter value manually"`)
  pivotSheet.getRange('F2').autoFill(pivotSheet.getRange(`F2:F${lastRowSummary - 1}`),ExcelScript.AutoFillType.fillDefault)
  pivotSheet.getRange('F1').setValue('Target');


  //FUNCTIONS

  //This function filters and deletes rows that contain specified values in specified columns.
  //Works with multiple columns/values, but the AND logical operator is used

  function deleteFilteredRows(
      columnName: string[],
      value: string[][],
      table: ExcelScript.Table
  ) {
      for (let i = 0; i < columnName.length; i++) {
          table
              .getColumnByName(columnName[i])
              .getFilter()
              .applyValuesFilter(value[i]);
          let visibleRows = table
              .getRangeBetweenHeaderAndTotal()
              .getVisibleView()
              .getRows();
          let firstVisibleRow = visibleRows[0].getRange().getRowIndex() + 1;
          let lastVisibleRow =
              visibleRows[visibleRows.length - 1].getRange().getRowIndex() + 1;
          sourceWorksheet
              .getRange(`${firstVisibleRow}:${lastVisibleRow}`)
              .delete(ExcelScript.DeleteShiftDirection.up);
          table.getColumnByName(columnName[i]).getFilter().clear();
      }
  }

  //This function deletes specified columns (ranges)
  function deleteColumns(range: string[]) {
      for (let i = 0; i < range.length; i++)
          sourceWorksheet
              .getRange(range[i])
              .delete(ExcelScript.DeleteShiftDirection.left);
  }

  //This function inserts a column with a specified formula
  function insertFormulaColumn(column: string, formula: string) {
      const lastRowFunc = usedRange.getRowCount();
      sourceWorksheet
          .getRange(`${column}:${column}`)
          .insert(ExcelScript.InsertShiftDirection.right);
      sourceWorksheet.getRange(`${column}2`).setFormula(formula);
      sourceWorksheet
          .getRange(`${column}2`)
          .autoFill(
              `${column}2:${column}${lastRowFunc}`,
              ExcelScript.AutoFillType.fillDefault
          );
  }
}

function main(workbook: ExcelScript.Workbook) {
  const mainWs = workbook.getWorksheet("Get_AllOpenOrdersShipping");
  const worksheets = workbook.getWorksheets();

  worksheets.forEach((ws) => {
     if (ws.getName() === "AllOpenOrders") {
        ws.delete();
     }
  });
  let workWs = workbook.addWorksheet("AllOpenOrders");
  const lastRowMain = mainWs.getUsedRange().getRowCount();
  const lastColumnMain = mainWs.getUsedRange().getColumnCount();
  workWs
     .getRange("A1")
     .copyFrom(
        mainWs.getRange(`5:${lastRowMain}`),
        ExcelScript.RangeCopyType.values
     );


  const columnsToDelete = [
     "K:K",
     "G:G",
     "I:I",
     "J:J",
     "N:N",
     "U:U",
     "V:V",
     "W:W",
     "Y:AC",
     "AJ:AJ",
     "AK:AK",
     "AN:ANM",
     "BJ:BJ",
     "BK:BK",
     "BN:BN",
     "BP:BQ",
  ];

  deleteColumns(workWs, columnsToDelete);
  workWs.getRange("F:G").setNumberFormat("dd/mm/yyyy")
  const allOpenOrdersTable = workWs.addTable(workWs.getUsedRange(), true)

  allOpenOrdersTable.getSort().apply([{key: 6, ascending: true}, {key: 7, ascending: true}])
  
  workWs.getRange("K:K").insert(ExcelScript.InsertShiftDirection.right);
  workWs.getRange("K1").setValue("WIP Qty");
  workWs.getRange("L:L").insert(ExcelScript.InsertShiftDirection.right);
  workWs.getRange("L1").setValue("WIP Qty Distributed");
  workWs.getRange("M:M").insert(ExcelScript.InsertShiftDirection.right);
  workWs.getRange("M1").setValue("Available to ship");
  const allWIPjobs = workWs.getRange(`N2:N${workWs.getUsedRange().getRowCount()}`).getValues();
  const allWIPqty: number[][] = [];

  allWIPjobs.forEach(wip => {
     const wipQty = extractWIPQty(wip[0])
     allWIPqty.push([wipQty])
  })


  workWs.getRange(`K2:K${ workWs.getUsedRange().getRowCount() }`).setValues(allWIPqty)
  deleteFilteredRows(
    ["Customer Name"],
    [
        ["Spirit AeroSystems Inc.", "Spirit AeroSystem Oklahoma"],
    ],
    allOpenOrdersTable
);
  workWs.getRange().getFormat().autofitColumns();
 //  Functions
  function deleteColumns(sheet: ExcelScript.Worksheet, ranges: string[]) {
     ranges.forEach((range) =>
        sheet.getRange(range).delete(ExcelScript.DeleteShiftDirection.left)
     );
  }

  function extractWIPQty(input: string): number {
     const quantities = input.match(/\d+\s*pcs/g);

     if (!quantities) return 0;

     const total = quantities.reduce((sum, quantity) => {
        const num = parseInt(quantity);
        return sum + num;
     }, 0);
     return total;
  }

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
        workWs
            .getRange(`${firstVisibleRow}:${lastVisibleRow}`)
            .delete(ExcelScript.DeleteShiftDirection.up);
        table.getColumnByName(columnName[i]).getFilter().clear();
    }
  }
}


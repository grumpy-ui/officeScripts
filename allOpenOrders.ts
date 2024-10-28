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
  workWs.getRange("F:G").setNumberFormat("dd/mm/yyyy");
  const allOpenOrdersTable = workWs.addTable(workWs.getUsedRange(), true);

  allOpenOrdersTable.getSort().apply([
    { key: 7, ascending: true },
    { key: 6, ascending: true },
  ]);

  deleteFilteredRows(
    ["Customer Name"],
    [["Spirit AeroSystems Inc.", "Spirit AeroSystem Oklahoma"]],
    allOpenOrdersTable
  );

  const lastRowWs = workWs.getUsedRange().getRowCount();
  workWs.getRange("K:K").insert(ExcelScript.InsertShiftDirection.right);
  workWs.getRange("K1").setValue("WIP Qty");
  workWs.getRange("L:L").insert(ExcelScript.InsertShiftDirection.right);
  workWs.getRange("L1").setValue("WIP Value");
  workWs.getRange("M:M").insert(ExcelScript.InsertShiftDirection.right);
  workWs.getRange("M1").setValue("Available to ship");
  workWs.getRange("L:L").insert(ExcelScript.InsertShiftDirection.right);
  workWs.getRange("L1").setValue("WIP Qty Distributed");

  const allWIPjobs = workWs
    .getRange(`O2:O${workWs.getUsedRange().getRowCount()}`)
    .getValues();
  const allWIPqty: number[][] = [];
  const wipValueArr: number[][] = [];
  const onHandDistrArr = workWs
    .getRange(`S2:S${lastRowWs}`)
    .getValues()
    .map((oh) => (oh[0] === "" ? [0] : oh));
  const qtyNotShipped = workWs.getRange(`P2:P${lastRowWs}`).getValues();
  const totalLinePrice = workWs.getRange(`AB2:AB${lastRowWs}`).getValues();
  const orderQty = workWs.getRange(`I2:I${lastRowWs}`).getValues();

  allWIPjobs.forEach((wip) => {
    const wipQty = extractWIPQty(wip[0]);
    allWIPqty.push([wipQty]);
  });

  const wipDistArr: number[][] = []

  workWs
    .getRange(`K2:K${workWs.getUsedRange().getRowCount()}`)
    .setValues(allWIPqty);

  const epicorPnArr = workWs
    .getRange(`H2:H${workWs.getUsedRange().getRowCount()}`)
    .getValues();
  const epicorPnObj = {};

  //Save all part numbers and their occurences in an object
  epicorPnArr.forEach((pn, i) => {
    if (epicorPnObj[pn]) {
      epicorPnObj[pn].occurence++;
    } else {
      epicorPnObj[pn] = {
        occurence: 1,
        totalWip: allWIPqty[i][0],
      };
    }
  });

  let row = 1;
  for (const pn in epicorPnObj) {
    let totalWipCell: number = epicorPnObj[pn].totalWip;
    const partOccurence: number = epicorPnObj[pn].occurence;

    for (let i = 0; i < partOccurence; i++) {
      let wipDist: number;
      let wipValue: number;
      const currOnHandDist = onHandDistrArr[row - 1][0];
      const currQtyNotShipped = qtyNotShipped[row - 1][0];
      const currTotalLinePrice = totalLinePrice[row - 1][0];
      const currOrderQty = orderQty[row - 1][0];
      if (currQtyNotShipped <= currOnHandDist || totalWipCell === 0) {
        wipDist = 0;
      } else {
        if (totalWipCell <= currQtyNotShipped) {
          wipDist = totalWipCell;
          totalWipCell = 0;
        } else {
          wipDist =
            Math.min(totalWipCell - currQtyNotShipped, currQtyNotShipped) -
            currOnHandDist;
          totalWipCell -= wipDist;
        }
      }
      wipDistArr.push([wipDist]);
      wipValue = (currTotalLinePrice / currOrderQty) * wipDist;
      wipValueArr.push([wipValue]);
      row++;
    }
  }
  workWs.getRange(`L2:L${lastRowWs}`).setValues(wipDistArr);
  workWs.getRange(`M2:M${lastRowWs}`).setValues(wipValueArr);
  workWs.getRange().getFormat().autofitColumns();

  //Functions

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

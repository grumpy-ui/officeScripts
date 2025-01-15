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
    "AN:AN",
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

  const lastRowWs = workWs.getUsedRange().getRowCount();
  workWs.getRange("K:K").insert(ExcelScript.InsertShiftDirection.right);
  workWs.getRange("K1").setValue("Available to Ship");

  const onHand = workWs
    .getRange(`S2:S${lastRowWs}`)
    .getValues()
    .map((value) => (value[0] === "" ? 0 : Number(value[0])));
  const orderQty = workWs
    .getRange(`I2:I${lastRowWs}`)
    .getValues()
    .map((value) => (value[0] === "" ? 0 : Number(value[0])));
  const wipQty = workWs
    .getRange(`P2:P${lastRowWs}`)
    .getValues()
    .map((value) => (value[0] === "" ? 0 : Number(value[0])));

  const availableToShip = onHand.map((onHandQty, index) => {
    const remainingWIP = wipQty[index] || 0;
    const neededQty = orderQty[index] || 0;
    return Math.min(onHandQty + remainingWIP, neededQty);
  });

  workWs
    .getRange(`K2:K${lastRowWs}`)
    .setValues(availableToShip.map((value) => [value]));

  workWs.getRange().getFormat().autofitColumns();

  // Helper Functions
  function deleteColumns(sheet: ExcelScript.Worksheet, ranges: string[]) {
    ranges.forEach((range) =>
      sheet.getRange(range).delete(ExcelScript.DeleteShiftDirection.left)
    );
  }
}

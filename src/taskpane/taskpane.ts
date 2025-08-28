/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("formatTransactions").onclick = formatTable;
  }
});

export async function formatTable() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      let dispTable = sheet.tables.add(range, true);

      await context.sync();

      for (let column of ["pi", "pn", "mogroup", "mophysn_na", "mo_qty",
        "give_units", "dose_text", "route", "start_dati", "stop_dati",
        "give_amt", "dose_max", "mas_desc", "ymi_grp", "bs_id", "bs_name",
        "bs_suffix"]) {
        console.log(column);
        dispTable.columns.getItem(column).delete();
      }

      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        dispTable.getRange().format.autofitRows();
        dispTable.getRange().format.autofitColumns();
      }
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("formatTransactions").onclick = formatTable;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function formatTable() {
  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange("A1:A1000");
      range.load("text");
      await context.sync();
      let contents = range.text[0];
      var tableEnd = "";
      for (let i = 1; i < 1000; i++) {
        if (contents[i] === "") {
          tableEnd = `AV${i + 1}`;
          break;
        }
      }
      let dispTable = worksheet.tables.add(`A1:${tableEnd}`, true);

      // worksheet.activate();

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

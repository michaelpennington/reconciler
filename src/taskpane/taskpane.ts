/* global console, document, Excel, Office */

import { MTFile } from "./parser"

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    var appBody = document.getElementById("app-body");
    var formatTransactions = document.getElementById("formatTransactions");
    var importButton = document.getElementById("importButton");
    if (!appBody || !formatTransactions || !importButton) {
      throw Error("Failed to find necessary html components!");
    }
    appBody.style.display = "flex";
    formatTransactions.onclick = formatTable;
    importButton.onclick = importData;
  }
});

export async function formatTable() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      let dispTable = sheet.tables.add(range, true);

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

export async function importData() {
  const fileInput = document.getElementById("fileInput") as HTMLInputElement;

  if (fileInput.files?.length === 0) {
    console.error("No file selected.");
    return;
  }

  const file = (fileInput.files as FileList)[0];

  const mtFile = new MTFile(file);
  try {
    for await (const line of mtFile) {
      console.log(line);
    }
    console.log("File Processing complete.");
  } catch (error) {
    console.error("An error occured while reading the file: ", error);
  }
}

/* Copyright © 2025 Michael Pennington - All Rights Reserved */

/* global Excel, Office, console, window, URL */
import { analyzeData, formatTable, processImportData, handleSheetAdded } from "../controller";

Office.onReady(async () => {
  // If needed, Office.js is ready to be called.
  await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.onAdded.add(handleSheetAdded);
    context.runtime.load("enableEvents");
    await context.sync();
    context.runtime.enableEvents = true;
    await context.sync();
  });
});

let dialog: Office.Dialog;

function importData(event: Office.AddinCommands.Event) {
  const url = new URL("/import-dialog.html", window.location.origin);
  Office.context.ui.displayDialogAsync(
    url.href,
    { height: 25, width: 25, displayInIframe: true },
    (result: Office.AsyncResult<Office.Dialog>) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error(result.error.message);
      } else {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg: any) => {
          dialog.close();

          if (arg.message === "error") {
            console.error("Error received from dialog during file read.");
          } else {
            try {
              await processImportData(arg.message);
            } catch (e) {
              console.error("Error processing file:", e);
              // Optionally, you could show a notification to the user here.
            }
          }

          event.completed();
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          event.completed();
        });
      }
    }
  );
}

async function callFormatTable(event: Office.AddinCommands.Event) {
  await formatTable();
  event.completed();
}

async function callAnalyzeData(event: Office.AddinCommands.Event) {
  await analyzeData();
  event.completed();
}

// Register the functions with Office.
Office.actions.associate("importData", importData);
Office.actions.associate("formatTable", callFormatTable);
Office.actions.associate("analyzeData", callAnalyzeData);

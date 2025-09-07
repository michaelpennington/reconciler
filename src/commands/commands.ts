/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, console */
import { analyzeData, formatTable } from "../controller";

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

let dialog: Office.Dialog;

function importData(event: Office.AddinCommands.Event) {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/import-dialog.html",
    { height: 25, width: 25, displayInIframe: true },
    (result: Office.AsyncResult<Office.Dialog>) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error(result.error.message);
      } else {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
          dialog.close();
          event.completed();
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg: any) => {
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

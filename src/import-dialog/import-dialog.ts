/* global console, document, Excel, Office */

import { MTFile } from "../taskpane/parser";
import { processImportData } from "../controller";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const importButton = document.getElementById("importButton") as HTMLButtonElement;
    const fileNameDisplay = document.getElementById("fileName") as HTMLElement;
    const fileInput = document.getElementById("fileInput") as HTMLInputElement;

    fileInput.addEventListener("change", () => {
      if (fileInput.files && fileInput.files.length > 0) {
        fileNameDisplay.textContent = fileInput.files[0].name;
        importButton.disabled = false;
      } else {
        fileNameDisplay.textContent = "No file chosen";
        importButton.disabled = true;
      }
    });

    importButton.onclick = importData;
  }
});

export async function importData() {
  try {
    const fileInput = document.getElementById("fileInput") as HTMLInputElement;

    if (fileInput.files?.length === 0) {
      console.error("No file selected.");
      return;
    }

    const file = (fileInput.files as FileList)[0];

    const mtFile = new MTFile(file);
    await processImportData(mtFile);
    Office.context.ui.messageParent("success");
  } catch (error) {
    console.error("An error occured while reading the file: ", error);
    Office.context.ui.messageParent("error");
  }
}

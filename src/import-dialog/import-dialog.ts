/* global console, document, Office, FileReader, HTMLInputElement */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const fileInput = document.getElementById("fileInput") as HTMLInputElement;
    const browseButton = document.getElementById("browse-button");

    browseButton.addEventListener("click", () => {
      fileInput.click();
    });

    fileInput.addEventListener("change", () => {
      const file = fileInput.files[0];
      if (file) {
        const reader = new FileReader();

        reader.onload = (event) => {
          const fileContent = event.target.result as string;
          Office.context.ui.messageParent(fileContent);
        };

        reader.onerror = (event) => {
          console.error("File could not be read! Code " + event.target.error.code);
          // Let the parent know something went wrong.
          Office.context.ui.messageParent("error");
        };

        reader.readAsText(file, "latin1");
      }
    });
  }
});

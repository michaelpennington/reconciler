/* global console, document, Excel, Office */

import { MTFile } from "./parser";
import { fromOCDate, toOADate } from "./parser_utils";
import { ignoredMeds } from "./ignore";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const appBody = document.getElementById("app-body");
    const formatTransactions = document.getElementById("formatTransactions");
    const importButton = document.getElementById("importButton") as HTMLButtonElement;
    const fileNameDisplay = document.getElementById("fileName") as HTMLElement;
    const fileInput = document.getElementById("fileInput") as HTMLInputElement;
    const analyzeDataButton = document.getElementById("analyzeData") as HTMLButtonElement;
    const aggregateDataButton = document.getElementById("aggregateData") as HTMLButtonElement;
    if (!appBody || !formatTransactions) {
      throw Error("Failed to find necessary html components!");
    }
    appBody.style.display = "flex";
    formatTransactions.onclick = formatTable;
    importButton.onclick = importData;
    analyzeDataButton.onclick = analyzeData;
    aggregateDataButton.onclick = aggregateData;

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

type UnifiedRecord = {
  ptID: string;
  rxNumber: string;
  medication: string;
  mnemonic: string;
  time: number;
  dispensed: number;
  given: number;
  returned: number;
  wasted: number;
};

export async function analyzeData() {
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const sheet = sheets.add("AdminVDisp");
      sheet.name = "AdminVDisp";
      let avdTable = sheet.tables.add("A1:I1", true);
      avdTable.name = "AdminsVDispenses";

      avdTable.load("headerRowRange");
      await context.sync();

      avdTable.getHeaderRowRange().values = [
        [
          "PtID",
          "RxNumber",
          "Medication",
          "Mnemonic",
          "Time",
          "NumberDispensed",
          "NumberGiven",
          "NumberReturned",
          "NumberWasted",
        ],
      ];
      const adminsTable = context.workbook.worksheets
        .getItem("Admins")
        .tables.getItem("Admins")
        .getDataBodyRange()
        .load("values");
      const disposTable = context.workbook.worksheets
        .getItem("Dispenses")
        .tables.getItem("Dispenses")
        .getDataBodyRange()
        .load("values");
      await context.sync();
      const adminsData = adminsTable.values;
      const disposData = disposTable.values;
      let newRecords: Array<UnifiedRecord> = [];

      for (const admin of adminsData) {
        let mnemonic = "";
        if (admin[AdminsColumns.Given]) {
          for (const disp of disposData) {
            if (disp[DisposColumns.RxNumber] === admin[AdminsColumns.RxNumber]) {
              mnemonic = disp[DisposColumns.Mnemonic];
              break;
            }
          }
          newRecords.push({
            ptID: admin[AdminsColumns.PtID],
            rxNumber: admin[AdminsColumns.RxNumber],
            medication: admin[AdminsColumns.Medication],
            mnemonic,
            time: admin[AdminsColumns.AdminTime],
            dispensed: 0,
            given: admin[AdminsColumns.NumberGiven],
            returned: 0,
            wasted: 0,
          });
        }
      }

      for (const dispo of disposData) {
        let medName;
        for (const admin of adminsData) {
          if (admin[AdminsColumns.RxNumber] == dispo[DisposColumns.RxNumber]) {
            medName = admin[AdminsColumns.Medication];
            break;
          }
        }
        let transType = dispo[DisposColumns.TransactionType];
        const qty: number = dispo[DisposColumns.Quantity];
        let wasteQty: number = 0;
        let issueQty = 0;
        let returnQty = 0;
        if (transType === "I") {
          issueQty = qty;
        } else if (transType === "R") {
          returnQty = -qty;
        } else if (transType === "W") {
          wasteQty = qty;
        } else {
          console.log(`unknown transaction type ${transType}`);
          continue;
        }
        newRecords.push({
          ptID: dispo[DisposColumns.PtID],
          rxNumber: dispo[DisposColumns.RxNumber],
          medication: medName ?? dispo[DisposColumns.RxName],
          mnemonic: dispo[DisposColumns.Mnemonic],
          time: dispo[DisposColumns.TransactionDate],
          dispensed: issueQty,
          given: 0,
          returned: returnQty,
          wasted: wasteQty,
        });
      }
      outer: for (const record of newRecords) {
        for (const ignored of ignoredMeds) {
          if (record.medication.startsWith(ignored)) {
            continue outer;
          }
        }
        avdTable.rows.add(
          undefined,
          [
            [
              record.ptID,
              record.rxNumber,
              record.medication,
              record.mnemonic,
              record.time,
              record.dispensed,
              record.given,
              record.returned,
              record.wasted,
            ],
          ],
          true
        );
      }
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        avdTable.getRange().format.autofitRows();
        avdTable.getRange().format.autofitColumns();
      }
      avdTable.getDataBodyRange().sort.apply([
        { key: 0, ascending: true },
        { key: 1, ascending: true },
        { key: 4, ascending: true },
      ]);
      avdTable.columns.getItem("Time").getDataBodyRange().numberFormat = [
        ["[$-409]m/d/yy h:mm AM/PM;@"],
      ];
      sheet.activate();
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function aggregateData() {
  try {
    await Excel.run(async (context) => {
      const avd = context.workbook.worksheets
        .getItem("AdminVDisp")
        .tables.getItem("AdminsVDispenses")
        .getDataBodyRange()
        .load("values");
      const sheets = context.workbook.worksheets;
      const sheet = sheets.add("Aggregate");
      sheet.name = "Aggregate";
      let aggTable = sheet.tables.add("A1:H1", true);
      aggTable.name = "Aggregate";
      aggTable.load("headerRowRange");
      await context.sync();

      aggTable.getHeaderRowRange().values = [
        [
          "PtID",
          "RxNumber",
          "Medication",
          "Mnemonic",
          "NumberDispensed",
          "NumberGiven",
          "NumberReturned",
          "NumberWasted",
        ],
      ];

      let nextItem = {
        ptID: undefined,
        rxNumber: undefined,
        medication: undefined,
        mnemonic: undefined,
        numberDispensed: undefined,
        numberGiven: undefined,
        numberReturned: undefined,
        numberWasted: undefined,
      };

      let needNewItem = false;
      let rowCount = 0;
      for (const row of avd.values) {
        if (!nextItem.rxNumber) {
          nextItem.rxNumber = row[1];
          needNewItem = true;
        }
        if (nextItem.rxNumber !== row[1]) {
          rowCount++;
          aggTable.rows.add(
            undefined,
            [
              [
                nextItem.ptID!,
                nextItem.rxNumber!,
                nextItem.medication!,
                nextItem.mnemonic!,
                nextItem.numberDispensed!,
                nextItem.numberGiven!,
                nextItem.numberReturned!,
                nextItem.numberWasted!,
              ],
            ],
            true
          );
          needNewItem = true;
        }

        if (needNewItem) {
          nextItem.ptID = row[0];
          nextItem.rxNumber = row[1];
          nextItem.medication = row[2];
          nextItem.mnemonic = row[3];
          nextItem.numberDispensed = row[5];
          nextItem.numberGiven = row[6];
          nextItem.numberReturned = row[7];
          nextItem.numberWasted = row[8];
          needNewItem = false;
        } else {
          nextItem.numberDispensed += row[5];
          nextItem.numberGiven += row[6];
          nextItem.numberReturned += row[7];
          nextItem.numberWasted += row[8];
        }
      }

      let varianceColumn = [];
      varianceColumn.push(["Variance"]);
      for (let i = 0; i < rowCount; i++) {
        varianceColumn.push([
          "=[@NumberDispensed]-[@NumberGiven]-[@NumberReturned]-[@NumberWasted]",
        ]);
      }
      console.log(varianceColumn);
      console.log(rowCount);
      await context.sync();
      aggTable.columns.add(undefined, varianceColumn);

      sheet.activate();
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        aggTable.getRange().format.autofitRows();
        aggTable.getRange().format.autofitColumns();
      }
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function formatTable() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const dispSheet = sheet.copy(Excel.WorksheetPositionType.end);
      dispSheet.name = "Dispenses";
      const range = dispSheet.getUsedRange();
      let dispTable = dispSheet.tables.add(range, true);
      dispTable.name = "Dispenses";

      const dateColumn = dispTable.columns.getItem("xact_dati").getDataBodyRange();
      dateColumn.load("values");
      await context.sync();
      console.log(dateColumn);
      let dates = Array.from(
        dateColumn.values.map((date) => [toOADate(fromOCDate(date[0] as string))])
      );
      console.log(dates);
      dateColumn.values = dates;
      dateColumn.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];
      dispTable.getHeaderRowRange().values = [
        [
          "PtID",
          "Omnicell",
          "Mnemonic",
          "TransactionDate",
          "ChargeID",
          "ChargeType",
          "TransactionType",
          "TransactionSubtype",
          "Quantity",
          "Countback",
          "MOOverride",
          "IssuedAfterDischarge",
          "QuantityOnHand",
          "UnitOfIssue",
          "UserName",
          "WitnessName",
          "WasteQuantity",
          "OmnicellName",
          "ItemName",
          "RxSuffix",
          "RxName",
          "PtName",
          "NullType",
          "ReconcileDose",
          "QtyZ",
          "WasteQtyZ",
          "MedStrengthUnits",
          "CaseId",
          "RxNumber",
          "MasDesc",
          "MRNumber",
          "User",
          "UserType",
          "WitnessID",
        ],
      ];
      dispSheet.activate();
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        dispTable.getRange().format.autofitRows();
        dispTable.getRange().format.autofitColumns();
      }
      const ptIDColumn = dispTable.columns.getItem("PtID");
      const rxNumColumn = dispTable.columns.getItem("RxNumber");
      const timeColumn = dispTable.columns.getItem("TransactionDate");
      ptIDColumn.load("index");
      rxNumColumn.load("index");
      timeColumn.load("index");
      await context.sync();
      dispTable.getDataBodyRange().sort.apply([
        { key: ptIDColumn.index, ascending: true },
        { key: rxNumColumn.index, ascending: true },
        { key: timeColumn.index, ascending: true },
      ]);
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function importData() {
  try {
    const fileInput = document.getElementById("fileInput") as HTMLInputElement;

    if (fileInput.files?.length === 0) {
      console.error("No file selected.");
      return;
    }

    const file = (fileInput.files as FileList)[0];

    const mtFile = new MTFile(file);
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const sheet = sheets.add("Admins");
      sheet.name = "Admins";
      let adminsTable = sheet.tables.add("A1:S1", true);
      adminsTable.name = "Admins";

      adminsTable.load("headerRowRange");
      await context.sync();

      adminsTable.getHeaderRowRange().values = [
        [
          "RxNumber",
          "PtName",
          "PtID",
          "Medication",
          "AdminTime",
          "FiledTime",
          "SchedTime",
          "User",
          "Given",
          "RxScanned",
          "PtScanned",
          "DoseAmount",
          "Units",
          "AdminDoseAmt",
          "AdminUnits",
          "MedStrength",
          "MedStrengthUnits",
          "NumberPerDose",
          "NumberGiven",
        ],
      ];
      for await (const line of mtFile) {
        console.log(line);
        adminsTable.rows.add(
          undefined,
          [
            [
              line.rxNum,
              line.ptName,
              line.ptId,
              line.medication,
              toOADate(line.adminTime) ?? "UNKNOWN",
              toOADate(line.filedTime) ?? "UNKNOWN",
              toOADate(line.schedTime) ?? "PRN",
              line.user,
              line.given,
              line.rxScanned,
              line.ptScanned,
              line.doseAmt,
              line.units,
              line.adminDoseAmt,
              line.adminUnits,
              line.medStrength ?? "UNKNOWN",
              line.medStrengthUnits ?? "UNKNOWN",
              line.countPerDose ?? "UNKNOWN",
              line.countGiven ?? "UNKNOWN",
            ],
          ],
          true
        );
      }
      console.log("File Processing complete.");
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        adminsTable.getRange().format.autofitRows();
        adminsTable.getRange().format.autofitColumns();
      }
      adminsTable.columns.getItem("AdminTime").getDataBodyRange().numberFormat = [
        ["[$-409]m/d/yy h:mm AM/PM;@"],
      ];
      adminsTable.columns.getItem("FiledTime").getDataBodyRange().numberFormat = [
        ["[$-409]m/d/yy h:mm AM/PM;@"],
      ];
      adminsTable.columns.getItem("SchedTime").getDataBodyRange().numberFormat = [
        ["[$-409]m/d/yy h:mm AM/PM;@"],
      ];
      const ptIDColumn = adminsTable.columns.getItem("PtID");
      const rxNumColumn = adminsTable.columns.getItem("RxNumber");
      const timeColumn = adminsTable.columns.getItem("AdminTime");
      ptIDColumn.load("index");
      rxNumColumn.load("index");
      timeColumn.load("index");
      await context.sync();
      adminsTable.getDataBodyRange().sort.apply([
        { key: ptIDColumn.index, ascending: true },
        { key: rxNumColumn.index, ascending: true },
        { key: timeColumn.index, ascending: true },
      ]);

      sheet.activate();
      await context.sync();
    });
  } catch (error) {
    console.error("An error occured while reading the file: ", error);
  }
}

const enum AdminsColumns {
  RxNumber,
  PtName,
  PtID,
  Medication,
  AdminTime,
  FiledTime,
  SchedTime,
  User,
  Given,
  RxScanned,
  PtScanned,
  DoseAmount,
  Units,
  AdminDoseAmt,
  AdminUnits,
  MedStrength,
  MedStrengthUnits,
  NumberPerDose,
  NumberGiven,
}

const enum DisposColumns {
  PtID,
  Omnicell,
  Mnemonic,
  TransactionDate,
  ChargeID,
  ChargeType,
  TransactionType,
  TransactionSubtype,
  Quantity,
  Countback,
  MOOverride,
  IssuedAfterDischarge,
  QuantityOnHand,
  UnitOfIssue,
  UserName,
  WitnessName,
  WasteQuantity,
  OmnicellName,
  ItemName,
  RxSuffix,
  RxName,
  PtName,
  NullType,
  ReconcileDose,
  QtyZ,
  WasteQtyZ,
  MedStrengthUnits,
  CaseId,
  RxNumber,
  MasDesc,
  MRNumber,
  User,
  UserType,
  WitnessID,
}

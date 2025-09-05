/* global console, document, Excel, Office */

import { MTFile } from "./parser";
import { fromOCDate, toOADate } from "./parser_utils";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const appBody = document.getElementById("app-body");
    const formatTransactions = document.getElementById("formatTransactions");
    const importButton = document.getElementById("importButton") as HTMLButtonElement;
    const fileNameDisplay = document.getElementById("fileName") as HTMLElement;
    const fileInput = document.getElementById("fileInput") as HTMLInputElement;
    const analyzeDataButton = document.getElementById("analyzeData") as HTMLButtonElement;
    if (!appBody || !formatTransactions) {
      throw Error("Failed to find necessary html components!");
    }
    appBody.style.display = "flex";
    formatTransactions.onclick = formatTable;
    importButton.onclick = importData;
    analyzeDataButton.onclick = analyzeData;

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

export async function analyzeData() {
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const sheet = sheets.add("AdminVDisp");
      sheet.name = "AdminVDisp";
      let avdTable = sheet.tables.add("A1:H1", true);
      avdTable.name = "AdminsVDispenses";

      avdTable.load("headerRowRange");
      await context.sync();

      avdTable.getHeaderRowRange().values = [
        [
          "PtID",
          "RxNumber",
          "Medication",
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
      let admins = adminsData.values();
      let dispos = disposData.values();
      let nextAdmin = admins.next();
      let nextDispo = dispos.next();
      let currentMnemonic;
      let currentRxNumber;
      let currentPtID;
      let currentDrugName;
      let next;
      while (
        (nextAdmin && !nextAdmin.done && nextAdmin.value) ||
        (nextDispo && !nextDispo.done && nextDispo.value)
      ) {
        let admin = nextAdmin.value;
        let dispo = nextDispo.value;
        if (!nextAdmin || nextAdmin.done || !admin) {
          next = "dispo";
        } else if (!nextDispo || nextDispo.done || !dispo) {
          next = "admin";
        } else {
          let adminPtID = admin[AdminsColumns.PtID];
          let adminRxNumber = admin[AdminsColumns.RxNumber];
          let adminTime = admin[AdminsColumns.AdminTime];
          let adminMedication = admin[AdminsColumns.Medication];
          let dispoPtID = dispo[DisposColumns.PtID];
          let dispoRxNumber = dispo[DisposColumns.RxNumber];
          let dispoTime = dispo[DisposColumns.TransactionDate];
          let dispoMnemonic = dispo[DisposColumns.Mnemonic];
          if (adminRxNumber !== currentRxNumber) {
            if (dispoRxNumber !== currentRxNumber) {
              if (adminPtID !== currentPtID) {
                if (dispoPtID !== currentPtID) {
                  if (adminPtID.localeCompare(dispoPtID) < 0) {
                    currentMnemonic = undefined;
                    currentRxNumber = adminRxNumber;
                    currentPtID = adminPtID;
                    currentDrugName = adminMedication;
                    next = "admin";
                  } else if (adminPtID.localeCompare(dispoPtID) > 0) {
                    currentPtID = dispoPtID;
                    currentRxNumber = dispoRxNumber;
                    currentMnemonic = dispoMnemonic;
                    currentDrugName = dispo[DisposColumns.RxName];
                    next = "dispo";
                  } else if (dispoTime < adminTime) {
                    currentRxNumber = dispoRxNumber;
                    currentMnemonic = dispoMnemonic;
                    currentPtID = dispoPtID;
                    next = "dispo";
                  } else {
                    currentRxNumber = adminRxNumber;
                    // current;
                    next = "admin";
                  }
                } else {
                }
              }
            }
          }
        }
        // if (!admin[AdminsColumns.Given]) {
        //   continue;
        // }
        // if (admin[AdminsColumns.Rx])
      }

      for (const admin of adminsData) {
        if (admin[AdminsColumns.Given]) {
          avdTable.rows.add(
            undefined,
            [
              [
                admin[AdminsColumns.PtID],
                admin[AdminsColumns.RxNumber],
                admin[AdminsColumns.Medication],
                admin[AdminsColumns.AdminTime],
                0,
                admin[AdminsColumns.NumberGiven] ?? 0,
                0,
                0,
              ],
            ],
            true
          );
        }
      }
      for (const disp of disposData) {
        if (!disp[DisposColumns.RxName].startsWith("*PATIENT SPECIFIC BIN")) {
          let transType = disp[DisposColumns.TransactionType];
          const qty: number = disp[DisposColumns.Quantity];
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
          avdTable.rows.add(
            undefined,
            [
              [
                disp[DisposColumns.PtID],
                disp[DisposColumns.RxNumber],
                disp[DisposColumns.RxName],
                disp[DisposColumns.TransactionDate],
                issueQty,
                0,
                returnQty,
                wasteQty,
              ],
            ],
            true
          );
        }
      }
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        avdTable.getRange().format.autofitRows();
        avdTable.getRange().format.autofitColumns();
      }
      avdTable.getDataBodyRange().sort.apply([
        { key: 0, ascending: true },
        { key: 1, ascending: true },
        { key: 3, ascending: true },
      ]);
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

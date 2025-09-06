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
  userID: string;
  orderType: string;
  schedule: string;
  location: string;
  prnReason: string;
};

export async function analyzeData() {
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const sheet = sheets.add("AdminVDisp");
      sheet.name = "AdminVDisp";
      let avdTable = sheet.tables.add("A1:P1", true);
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
          "Total",
          "UserID",
          "OrderType",
          "Location",
          "Schedule",
          "PRNReason",
          "PtID+Rx+Medication",
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
            userID: admin[AdminsColumns.User],
            orderType: admin[AdminsColumns.OrderType],
            schedule: admin[AdminsColumns.Schedule],
            location: admin[AdminsColumns.Unit],
            prnReason: admin[AdminsColumns.PRNReason],
          });
        }
      }

      for (const dispo of disposData) {
        let medName;
        let orderType = "";
        let schedule = "";
        let prnReason = "";
        for (const admin of adminsData) {
          if (admin[AdminsColumns.RxNumber] == dispo[DisposColumns.RxNumber]) {
            medName = admin[AdminsColumns.Medication];
            prnReason = admin[AdminsColumns.PRNReason];
            orderType = admin[AdminsColumns.OrderType];
            schedule = admin[AdminsColumns.Schedule];
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
          userID: dispo[DisposColumns.User],
          orderType,
          location: dispo[DisposColumns.Omnicell],
          schedule,
          prnReason,
        });
      }
      outer: for (const record of newRecords) {
        for (const ignored of ignoredMeds) {
          if (record.medication.startsWith(ignored)) {
            continue outer;
          }
        }
        if (record.location.startsWith("BR")) {
          record.location = record.location.slice(2);
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
              "=[@NumberWasted]+[@NumberReturned]+[@NumberGiven]-[@NumberDispensed]",
              record.userID,
              record.orderType,
              record.location,
              record.schedule,
              record.prnReason,
              '=[@PtID] & " - " & [@RxNumber] & " - " ' +
                '& [@Schedule] & " " & [@OrderType] & " - " & [@Medication]',
            ],
          ],
          true,
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
      await context.sync();

      let auditData = sheets.add("AuditData");
      let avtPivotTable = auditData.pivotTables.add(
        "AVDAUDIT",
        avdTable,
        auditData.getRange("A1:C18"),
      );
      avtPivotTable.rowHierarchies.add(avtPivotTable.hierarchies.getItem("PtID+Rx+Medication"));
      avtPivotTable.rowHierarchies.getItem("PtID+Rx+Medication").position = 0;
      avtPivotTable.dataHierarchies.add(avtPivotTable.hierarchies.getItem("NumberDispensed"));
      avtPivotTable.dataHierarchies.getItem("Sum of NumberDispensed").position = 0;
      avtPivotTable.dataHierarchies.add(avtPivotTable.hierarchies.getItem("NumberGiven"));
      avtPivotTable.dataHierarchies.getItem("Sum of NumberGiven").position = 1;
      avtPivotTable.dataHierarchies.add(avtPivotTable.hierarchies.getItem("NumberReturned"));
      avtPivotTable.dataHierarchies.getItem("Sum of NumberReturned").position = 2;
      avtPivotTable.dataHierarchies.add(avtPivotTable.hierarchies.getItem("NumberWasted"));
      avtPivotTable.dataHierarchies.getItem("Sum of NumberWasted").position = 3;
      avtPivotTable.dataHierarchies.add(avtPivotTable.hierarchies.getItem("Total"));
      avtPivotTable.dataHierarchies.getItem("Sum of Total").position = 4;
      avtPivotTable.dataHierarchies.load("no-properties-needed");
      await context.sync();
      avtPivotTable.dataHierarchies.items[0].name = "NumDispensed";
      avtPivotTable.dataHierarchies.items[1].name = "NumGiven";
      avtPivotTable.dataHierarchies.items[2].name = "NumReturned";
      avtPivotTable.dataHierarchies.items[3].name = "NumWasted";
      avtPivotTable.dataHierarchies.items[4].name = "Sum";

      const conditionalFormat = auditData
        .getRange("F2:F3000")
        .conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
      conditionalFormat.cellValue.format.font.color = "#9C5700";
      conditionalFormat.cellValue.format.fill.color = "#FFEB9C";
      conditionalFormat.cellValue.rule = { formula1: "=0", operator: "GreaterThan" };
      const conditionalFormat2 = auditData
        .getRange("F2:F3000")
        .conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
      conditionalFormat2.cellValue.format.font.color = "#9C0006";
      conditionalFormat2.cellValue.format.fill.color = "#FFC7CE";
      conditionalFormat2.cellValue.rule = { formula1: "=0", operator: "LessThan" };

      // auditData.getRange("F1:F1").conditionalFormats.clearAll();
      auditData.freezePanes.freezeAt(auditData.getRange("1:1"));

      auditData.activate();

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
        dateColumn.values.map((date) => [toOADate(fromOCDate(date[0] as string))]),
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
      let adminsTable = sheet.tables.add("A1:W1", true);
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
          "Schedule",
          "OrderType",
          "Unit",
          "PRNReason",
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
              line.schedule,
              line.orderType,
              line.location,
              line.prnReason,
            ],
          ],
          true,
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
  Schedule,
  OrderType,
  Unit,
  PRNReason,
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

/*global console*/

import { Field, getField } from "./parser_utils";

function parseMTDate(date: string): Date {
  const month = parseInt(date.slice(0, 2), 10) - 1;
  const day = parseInt(date.slice(3, 5), 10);
  const year = 2000 + parseInt(date.slice(6, 8), 10);
  if (date.includes("-")) {
    const hours = parseInt(date.slice(9, 11), 10);
    const minutes = parseInt(date.slice(11, 13), 10);
    return new Date(year, month, day, hours, minutes);
  } else {
    return new Date(year, month, day);
  }
}

export class EMARLineItem {
  rxNum: string;
  ptName: string;
  ptId: string;
  medication: string;
  adminTime: Date;
  filedTime: Date;
  schedTime: Date | undefined;
  user: string;
  given: boolean;
  rxScanned: boolean;
  ptScanned: boolean;
  doseAmt: number;
  units: string;
  adminDoseAmt: number;
  adminUnits: string;
  medStrength: number | undefined;
  medStrengthUnits: string | undefined;
  countPerDose: number | undefined;
  countGiven: number | undefined;
  schedule: string;
  orderType: string;
  location: string;
  prnReason: string;
  refReason: string | undefined;

  constructor(
    rxNum: string,
    ptName: string,
    ptId: string,
    medication: string,
    adminTime: Date,
    filedTime: Date,
    schedTime: Date | undefined,
    user: string,
    given: boolean,
    rxScanned: boolean,
    ptScanned: boolean,
    doseAmt: number,
    units: string,
    adminDoseAmt: number,
    adminUnits: string,
    medStrength: number | undefined,
    medStrengthUnits: string | undefined,
    countPerDose: number | undefined,
    countGiven: number | undefined,
    schedule: string,
    orderType: string,
    location: string,
    prnReason: string,
    refReason: string | undefined
  ) {
    this.rxNum = rxNum;
    this.ptName = ptName;
    this.ptId = ptId;
    this.medication = medication;
    this.adminTime = adminTime;
    this.filedTime = filedTime;
    this.schedTime = schedTime;
    this.user = user;
    this.given = given;
    this.rxScanned = rxScanned;
    this.ptScanned = ptScanned;
    this.doseAmt = doseAmt;
    this.units = units;
    this.adminDoseAmt = adminDoseAmt;
    this.adminUnits = adminUnits;
    this.medStrength = medStrength;
    this.medStrengthUnits = medStrengthUnits;
    this.countPerDose = countPerDose;
    this.countGiven = countGiven;
    this.schedule = schedule;
    this.orderType = orderType;
    this.location = location;
    this.prnReason = prnReason;
    this.refReason = refReason;
  }
}

export async function* mtLineParser(lines: string[]): AsyncGenerator<EMARLineItem> {
  let currentPtName: string | null = null;
  let currentPtId: string | null = null;
  let currentRx: string | null = null;
  let lineNo = 0;
  let currentMedication: string = "";
  let justSawPt = false;
  let readingMedication = false;
  let readingAdmins = false;
  let adminStack: Array<AdminDetails> = [];
  let currentDoseAmt = 0.0;
  let currentDoseUnits = "";
  let currentMedStrength = undefined;
  let currentMedStrengthUnits = undefined;
  let currentSchedule = "";
  let currentOrderType = "";
  let currentLocation = "";
  let lookingForPRNReason = false;
  let doseAmt;
  let units;
  let currentDosePerUnits = undefined;
  let currentPRNReason = "";
  for (const line of lines) {
    let readyToSubmit = false;
    if (line.startsWith("Patient")) {
      currentPtName = getField(line, Field.PtName).trim().replace(",", ", ");
      justSawPt = true;
    } else if (justSawPt) {
      currentPtId = getField(line, Field.PtId).trim();
      currentLocation = getField(line, Field.Location).trim();
      justSawPt = false;
    } else if (line.startsWith("Z") || line.startsWith("U")) {
      const number = /[0-9]/;
      if (number.test(line.charAt(1))) {
        if (!currentPtName || !currentPtId) {
          throw new Error(`Rx ${getField(line, Field.RxNum)} found outside context of patient!`);
        }
        currentRx = getField(line, Field.RxNum);
        currentMedication = getField(line, Field.Medication).trim();
        readingMedication = true;
        readingAdmins = true;
      }
    } else if (line.startsWith("       Dose")) {
      readingAdmins = false;
      let doseStr = getField(line, Field.Dose).trim();
      currentSchedule = getField(line, Field.Schedule).trim();
      currentOrderType = getField(line, Field.Prn).trim();
      doseAmt = undefined;
      units = undefined;
      if (doseStr.startsWith("See Taper")) {
        doseAmt = currentDoseAmt;
        units = currentDoseUnits;
      } else {
        let doseStrs = doseStr.split(" ");
        doseAmt = parseFloat(doseStrs[0].replace(",", ""));
        units = doseStrs[1] ?? "NF";
      }
      if (!currentRx || !currentPtName || !currentPtId || (!doseAmt && doseAmt != 0) || !units) {
        console.log(currentRx, currentPtName, currentPtId, doseAmt, units);
        throw new Error(`Found Dose line before finding an Rx or Pt! on line ${lineNo}`);
      }
      if (
        (doseAmt !== currentDoseAmt || units !== currentDoseUnits) &&
        doseAmt !== 0 &&
        units !== "NF"
      ) {
        console.log(currentRx, currentPtName, currentPtId, doseAmt, units);
        throw new Error(`Dose amount is not equal! on line ${lineNo}`);
      }
      currentDosePerUnits = undefined;
      if (
        currentMedStrength &&
        currentMedStrengthUnits &&
        currentMedStrengthUnits === currentDoseUnits
      ) {
        currentDosePerUnits = currentDoseAmt / currentMedStrength;
      }
      if (adminStack.length > 0 && adminStack[0].schedTime === undefined) {
        console.log(adminStack.slice());
        lookingForPRNReason = true;
      } else {
        readyToSubmit = true;
        currentPRNReason = "";
      }
    } else if (readingMedication) {
      let newMedString = getField(line, Field.Medication);
      if (newMedString.startsWith("  ")) {
        readingMedication = false;
        let lastParen = currentMedication.lastIndexOf("(");
        if (lastParen === -1) {
          console.log(currentMedication);
          throw Error(`Found medication string without ending dose! On line ${lineNo}`);
        }
        let dose = currentMedication.slice(lastParen);
        currentMedication = currentMedication.slice(0, lastParen - 1);
        if (!dose.startsWith("(") || !dose.endsWith(")")) {
          throw Error(`Malformed dose ${dose} at and of medication]`);
        }
        let doseStrs = dose.slice(1, -1).trim().split(" ");
        currentDoseAmt = parseFloat(doseStrs[0].replace(",", ""));
        currentDoseUnits = doseStrs[1];
        let strength = readStrength(currentMedication);
        if (
          currentDoseUnits === "EACH" ||
          currentDoseUnits === "ea" ||
          currentDoseUnits === "EA" ||
          currentDoseUnits === "each"
        ) {
          currentMedStrength = 1;
          currentMedStrengthUnits = currentDoseUnits;
        } else if (strength) {
          ({ amount: currentMedStrength, units: currentMedStrengthUnits } = strength);
        } else {
          currentMedStrength = undefined;
          currentMedStrengthUnits = undefined;
        }
      } else {
        currentMedication += " " + newMedString.trim();
      }
    } else if (lookingForPRNReason) {
      if (line.startsWith(" PRN Reason")) {
        currentPRNReason = getField(line, Field.PrnReason).trim();
      }
      lookingForPRNReason = false;
      readyToSubmit = true;
    }
    if (readingAdmins) {
      let admin = readAdminDetails(line);
      if (admin !== null) {
        adminStack.push(admin);
      }
    }
    if (readyToSubmit) {
      while (adminStack.length > 0) {
        let admin = adminStack.shift() as AdminDetails;
        let countGiven = undefined;
        if (
          currentMedStrength &&
          currentMedStrengthUnits &&
          currentMedStrengthUnits === admin.adminUnits &&
          admin.given
        ) {
          countGiven = admin.adminDoseAmt / currentMedStrength;
        } else if (!admin.given) {
          countGiven = 0;
        }
        if (!currentRx || !currentPtName || !currentPtId || (!doseAmt && doseAmt != 0) || !units) {
          console.log(currentRx, currentPtName, currentPtId, doseAmt, units);
          throw new Error(`Found Dose line before finding an Rx or Pt! on line ${lineNo}`);
        }
        yield new EMARLineItem(
          currentRx,
          currentPtName,
          currentPtId,
          currentMedication,
          admin.adminTime,
          admin.filedTime,
          admin.schedTime,
          admin.user,
          admin.given,
          admin.rxScanned,
          admin.ptScanned,
          doseAmt,
          units,
          admin.adminDoseAmt,
          admin.adminUnits,
          currentMedStrength,
          currentMedStrengthUnits,
          currentDosePerUnits,
          countGiven,
          currentSchedule,
          currentOrderType,
          mapUnit(currentLocation),
          currentPRNReason,
          admin.refReason
        );
      }
    }
    lineNo++;
  }
}

class AdminDetails {
  adminTime: Date;
  filedTime: Date;
  schedTime: Date | undefined;
  user: string;
  given: boolean;
  rxScanned: boolean;
  ptScanned: boolean;
  adminDoseAmt: number;
  adminUnits: string;
  refReason: string | undefined;

  constructor(
    adminTime: Date,
    filedTime: Date,
    schedTime: Date | undefined,
    user: string,
    given: boolean,
    rxScanned: boolean,
    ptScanned: boolean,
    adminDoseAmt: number,
    adminUnits: string,
    refReason: string | undefined
  ) {
    this.adminTime = adminTime;
    this.filedTime = filedTime;
    this.schedTime = schedTime;
    this.user = user;
    this.given = given;
    this.rxScanned = rxScanned;
    this.ptScanned = ptScanned;
    this.adminDoseAmt = adminDoseAmt;
    this.adminUnits = adminUnits;
    this.refReason = refReason;
  }
}

/**
 * Checks if a given value is a valid Date object.
 * @param value The value to check.
 * @returns True if the value is a valid Date, false otherwise.
 */
function isValidDate(value: any): boolean {
  // A value is a valid date if it's an instance of Date
  // and its internal time value is not Not-a-Number.
  return value instanceof Date && !isNaN(value.getTime());
}

const UNITS: Array<string> = [
  "MG",
  "mg",
  "MCG",
  "mcg",
  "G",
  "g",
  "GR",
  "gr",
  "PUFF",
  "EACH",
  "each",
  "puff",
  "mL",
  "ML",
  "ml",
  "GM",
  "gm",
  "unit",
  "UNIT",
  "MG/ML",
  "mg/mL",
  "mg/ml",
];

function readStrength(medication: string): { amount: number; units: string } | undefined {
  let seenOpenParen = false;
  let lastWord = "";
  let units: string | undefined = undefined;
  let amount: number | undefined = undefined;
  for (const word of medication.split(" ")) {
    if (word === "") {
      continue;
    }
    if (word.startsWith("(")) {
      seenOpenParen = true;
    }
    if (!seenOpenParen) {
      for (const unit of UNITS) {
        if (word === unit) {
          if (unit === "MG/ML") {
            units = "MG";
          } else if (unit === "mg/mL" || unit === "mg/ml") {
            units = "mg";
          } else {
            units = unit;
          }
          amount = parseFloat(lastWord.replace(",", ""));
          break;
        } else if (word.endsWith(unit)) {
          if (unit === "MG/ML") {
            units = "MG";
          } else if (unit === "mg/mL" || unit === "mg/ml") {
            units = "mg";
          } else {
            units = unit;
          }
          units = unit;
          let unitStart = word.lastIndexOf(unit);
          amount = parseFloat(word.slice(0, unitStart).trim().replace(",", ""));
          break;
        }
      }
      lastWord = word;
    }
    if (word.endsWith(")")) {
      seenOpenParen = false;
    }
  }
  if (!units || !amount) {
    return undefined;
  } else {
    return { units, amount };
  }
}

function mapUnit(unit: string): string {
  switch (unit) {
    case "IP.CHILD":
      return "OSGOOD1";
    case "IP.ADOL1":
      return "OSGOOD2";
    case "IP.ADOL2":
      return "OSGOOD3";
    case "IP.ALGBT":
      return "TYLER1";
    case "IP.ADULT":
      return "TYLER2";
    case "IP.AIU2":
      return "TYLER3";
    case "IP.AIU1":
      return "TYLER4";
  }
  return "UNKNOWNUNIT";
}

function mapRefReason(reason: string): string {
  switch (reason) {
    case "PAA":
      return "Patient is sleeping";
    case "REF":
      return "Patient refused";
    default:
      return `Unknown refusal reason ${reason}`;
  }
}

function readAdminDetails(line: string): AdminDetails | null {
  const schedTimeText = getField(line, Field.SchedTime);
  let schedTime = undefined;
  if (schedTimeText !== "NON-SCHEDULED") {
    schedTime = parseMTDate(schedTimeText);
    if (
      !isValidDate(schedTime) ||
      (getField(line, Field.Given) !== "Y" && getField(line, Field.Given) !== "N")
    ) {
      return null;
    }
  }
  let adminTime = parseMTDate(getField(line, Field.AdminTime));
  let filedTime = parseMTDate(getField(line, Field.FiledTime));
  let user = getField(line, Field.User).trim();
  let given = getField(line, Field.Given) === "Y";
  let rxScanned = getField(line, Field.RxScanned) === "Y";
  let ptScanned = getField(line, Field.PtScanned) === "Y";
  let doseStrs = getField(line, Field.AdminDose).trim().split(" ");
  let adminDoseAmt = parseFloat(doseStrs[0].replace(",", ""));
  let adminUnits = doseStrs[1];
  let refReason;
  if (!given) {
    refReason = mapRefReason(getField(line, Field.RefReason).trim());
  }
  return new AdminDetails(
    adminTime,
    filedTime,
    schedTime,
    user,
    given,
    rxScanned,
    ptScanned,
    adminDoseAmt,
    adminUnits,
    refReason
  );
}

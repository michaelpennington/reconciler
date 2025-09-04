import { readFileLines } from "./filelinesreader";

import { Field, getField } from "./parser_utils";

export class MTFile implements AsyncIterable<EMARLineItem> {
  private fromDate: Date | undefined = undefined;
  private thruDate: Date | undefined = undefined;
  private lineParser: AsyncGenerator<EMARLineItem> | null = null;
  private readonly lineReader: AsyncGenerator<string>;

  constructor(file: File) {
    this.lineReader = readFileLines(file);
  }

  [Symbol.asyncIterator](): AsyncIterator<EMARLineItem, any, any> {
    return this;
  }

  async initialize() {
    if (!this.lineParser) {
      let counter = 0;
      let result = await this.lineReader.next();
      while (!result.done) {
        if (result.value.includes("From Date-Time")) {
          this.fromDate = parseMTDate(result.value.slice(34, 47));
        } else if (result.value.includes("Thru Date-Time")) {
          this.thruDate = parseMTDate(result.value.slice(34, 47));
        }

        if (this.fromDate && this.thruDate) {
          break;
        }
        if (counter > 10) {
          throw new Error("Are you sure you have the right file?");
        }
        counter++;
        result = await this.lineReader.next();
      }
      console.log("Data from: ", this.fromDate);
      console.log("Data thru: ", this.thruDate);
      this.lineParser = mtLineParser(this.lineReader);
    }
  }

  async next(): Promise<IteratorResult<EMARLineItem, any>> {
    if (!this.lineParser) {
      await this.initialize();
    }
    // SAFETY: We called initialize if it was null, and only make it
    // thus far if initialization was successful
    return (this.lineParser as AsyncGenerator<EMARLineItem>).next();
  }

  public async getThruDate(): Promise<Date> {
    if (!this.lineParser) {
      await this.initialize();
    }
    // SAFETY: We called initialize if it was null, and only make it
    // thus far if initialization was successful
    return this.thruDate as Date;
  }

  public async getFromDate(): Promise<Date> {
    if (!this.lineParser) {
      await this.initialize();
    }
    // SAFETY: We called initialize if it was null, and only make it
    // thus far if initialization was successful
    return this.fromDate as Date;
  }
}

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
    countGiven: number | undefined
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
  }
}

async function* mtLineParser(lineReader: AsyncGenerator<string>): AsyncGenerator<EMARLineItem> {
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
  for await (let line of lineReader) {
    if (line.startsWith("Patient")) {
      currentPtName = getField(line, Field.PtName).trim().replace(",", ", ");
      justSawPt = true;
    } else if (justSawPt) {
      currentPtId = getField(line, Field.PtId).trim();
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
      let doseAmt;
      let units;
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
      let dosePerUnits = undefined;
      if (
        currentMedStrength &&
        currentMedStrengthUnits &&
        currentMedStrengthUnits === currentDoseUnits
      ) {
        dosePerUnits = currentDoseAmt / currentMedStrength;
      }
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
          dosePerUnits,
          countGiven
        );
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
    }
    if (readingAdmins) {
      let admin = readAdminDetails(line);
      if (admin !== null) {
        adminStack.push(admin);
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

  constructor(
    adminTime: Date,
    filedTime: Date,
    schedTime: Date | undefined,
    user: string,
    given: boolean,
    rxScanned: boolean,
    ptScanned: boolean,
    adminDoseAmt: number,
    adminUnits: string
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
          units = unit;
          amount = parseFloat(lastWord.replace(",", ""));
          break;
        } else if (word.endsWith(unit)) {
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

function readAdminDetails(line: string): AdminDetails | null {
  const schedTimeText = getField(line, Field.SchedTime);
  let schedTime = undefined;
  if (schedTimeText !== "NON-SCHEDULED") {
    schedTime = parseMTDate(schedTimeText);
    if (!isValidDate(schedTime)) {
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
  return new AdminDetails(
    adminTime,
    filedTime,
    schedTime,
    user,
    given,
    rxScanned,
    ptScanned,
    adminDoseAmt,
    adminUnits
  );
}

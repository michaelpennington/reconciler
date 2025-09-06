export const enum Field {
  RxNum,
  PtName,
  PtId,
  Medication,
  AdminTime,
  FiledTime,
  SchedTime,
  User,
  Given,
  RxScanned,
  PtScanned,
  Dose,
  AdminDose,
  Schedule,
  Prn,
  Location,
  PrnReason,
}

const fields: Array<Array<number>> = [
  [0, 9],
  [9, 39],
  [16, 29],
  [13, 63],
  [92, 105],
  [106, 119],
  [73, 86],
  [120, 130],
  [131, 132],
  [133, 134],
  [135, 136],
  [12, 32],
  [139, 151],
  [44, 64],
  [65, 72],
  [40, 59],
  [13, 53],
] as const;

export function getField(line: string, field: Field): string {
  return line.slice(fields[field][0], fields[field][1]);
}

const SECONDS_IN_MINUTE: number = 60 as const;
const MINUTES_IN_HOUR: number = 60 as const;
const HOURS_IN_DAY: number = 24 as const;
const SECONDS_IN_HOUR: number = SECONDS_IN_MINUTE * MINUTES_IN_HOUR;
const SECONDS_IN_DAY: number = SECONDS_IN_MINUTE * MINUTES_IN_HOUR * HOURS_IN_DAY;

export function toOADate(date: Date | undefined): number | undefined {
  if (!date) {
    return undefined;
  }
  let oaDate = getNumDaysSince(date);
  let hours = date.getHours();
  let minutes = date.getMinutes();
  let seconds = date.getSeconds();
  if (hours) {
    let totalSeconds = SECONDS_IN_HOUR * hours + SECONDS_IN_MINUTE * minutes + seconds;
    if (oaDate < 0) {
      oaDate = -(-oaDate + totalSeconds / SECONDS_IN_DAY);
    } else {
      oaDate += totalSeconds / SECONDS_IN_DAY;
    }
  }
  return oaDate;
}

export function fromOCDate(date: string): Date {
  let year = parseInt(date.slice(0, 4));
  let month = parseInt(date.slice(4, 6)) - 1;
  let day = parseInt(date.slice(6, 8));
  let hours = parseInt(date.slice(8, 10));
  let minutes = parseInt(date.slice(10, 12));
  let seconds = parseInt(date.slice(12, 14));
  if (seconds >= 30) {
    if (minutes === 59) {
      if (hours === 23) {
        if (month === 1 && isLeapYear(year) && day === 29) {
          month = 2;
          day = 1;
        } else if (day === DAYS_IN_MONTH[month] && (month !== 1 || !isLeapYear(year))) {
          if (month === 11) {
            year++;
            month = 0;
          } else {
            month++;
          }
          day = 1;
        } else {
          day++;
        }
        hours = 0;
      } else {
        hours++;
      }
      minutes = 0;
    } else {
      minutes++;
    }
  }
  return new Date(year, month, day, hours, minutes);
}

function isLeapYear(year: number): boolean {
  return year % 400 === 0 || (year % 100 !== 0 && year % 4 === 0);
}

function dayOfYear(date: Date): number {
  let days = 0;
  for (const m of Array.from({ length: date.getMonth() }, (_, index) => index)) {
    days += DAYS_IN_MONTH[m];
    if (m === 1 && isLeapYear(date.getFullYear())) {
      days += 1;
    }
  }
  days += date.getDate();
  return days;
}

function getNumDaysSince(date: Date): number {
  let year = date.getFullYear();
  return 365 * (year - 1900) + countLeapYears(year) + dayOfYear(date) + 1;
}

function countLeapYears(year: number): number {
  return Math.ceil(year / 400) - Math.ceil(year / 100) + Math.ceil(year / 4) - 461;
}

const DAYS_IN_MONTH: Array<number> = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31] as const;

import { readFileLines } from "./filelinesreader";
export class MTFile implements AsyncIterable<LineItem> {
  private fromDate: Date | undefined = undefined;
  private thruDate: Date | undefined = undefined;
  private lineParser: AsyncGenerator<LineItem> | null = null;
  private readonly lineReader: AsyncGenerator<string>;

  constructor(file: File) {
    this.lineReader = readFileLines(file);
  }

  [Symbol.asyncIterator](): AsyncIterator<LineItem, any, any> {
    return this;
  }

  async initialize() {
    if (!this.lineParser) {
      var counter = 0;
      var { value: line, done } = await this.lineReader.next();
      while (!done) {
        if (line.includes("From Date-Time")) {
          this.fromDate = parseMTDate(line.slice(34, 47));
        } else if (line.includes("Thru Date-Time")) {
          this.thruDate = parseMTDate(line.slice(34, 47));
        }

        if (this.fromDate && this.thruDate) {
          break;
        }
        if (counter > 10) {
          throw new Error("Are you sure you have the right file?");
        }
        var { value: line, done } = await this.lineReader.next();
        counter++;

      }
      console.log("Data from: ", this.fromDate);
      console.log("Data thru: ", this.thruDate);
      this.lineParser = mtLineParser(this.lineReader);
    }
  }

  async next(): Promise<IteratorResult<LineItem, any>> {
    if (!this.lineParser) {
      await this.initialize();
    }
    // SAFETY: We called initialize if it was null, and only make it
    // thus far if initialization was successful
    return (this.lineParser as AsyncGenerator<LineItem>).next();
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
  const month = parseInt(date.slice(0, 2)) - 1;
  const day = parseInt(date.slice(3, 5));
  const year = 2000 + parseInt(date.slice(6, 8));
  if (date.includes("-")) {
    const hours = parseInt(date.slice(9, 11));
    const minutes = parseInt(date.slice(11, 13));
    return new Date(year, month, day, hours, minutes);
  } else {
    return new Date(year, month, day);
  }
}

export class LineItem {
  rxNum: number;
  ptName: string;
  ptId: number;
  medication: string;
  adminTime: Date;
  filedTime: Date;
  user: string;
  given: boolean;
  rxScanned: boolean;
  ptScanned: boolean;
  doseAmt: number;
  amtUnits: string;
  givenDoseAmt: number;
  givenAmtUnits: string;
}

async function* mtLineParser(lineReader: AsyncGenerator<string>): AsyncGenerator<LineItem> {
  var currentPtName: string | null = null;
  var currentPtId: number | null = null;
  var currentRx: number | null = null;
  var just_saw_pt = false;
  var just_saw_rx = false;
  for await (var line of lineReader) {
    if (line.startsWith("Patient")) {
      currentPtName = line.slice(9, 39).trim();
      just_saw_pt = true;
    } else if (line.startsWith("Z")) {
      if (!currentPtName || !currentPtId) {
        throw new Error(`Rx ${line.slice(0, 9)} found outside context of patient!`);
      }
      currentRx = parseInt(line.slice(1, 9));
    } else if (just_saw_pt) {
      currentPtId = parseInt(line.slice(18, 28));
      just_saw_pt = false;
    }
  }
  return {};
}

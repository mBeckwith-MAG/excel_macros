const STORE_NAME: string = "BOSM";
const GETS_RETRO: Array<number> = [];

enum Color {
  GREY = "#C0C0C0",
  LIGHT_GREY = "#E6E6E6",
  GREEN = "#4CC273",
  WHITE = "#FFFFFF",
  YELLOW = "#FAE49D",
  RED = "#FF1515",
  BLUE_GREY = "#DAEEF3",
  PEACH = "#FDE9D9"
}

enum NumberFormat {
  NUMBER = "#.#",
  ACCOUNTING = "_($* #,##0.00_);[Red]_($* -#,##0.00;_($* \" - \"??_);_(@_)",
  CURRENCY = "$#,##0.00;[Red]$#,##0.00",
  DATE = "mm/dd/yy",
  PERCENT = "0.00%"
}

enum Total {
  UNITS,
  COMM_GROSS,
  FI_GROSS,
  COMM_AMOUNT
}

interface TableTotals {
  totals: number,
  header: number,
  start: number
}

abstract class Person {
  private name: string;
  private number: number;
  constructor(name: string, number: number) {
    this.name = name;
    this.number = number;
  }
  getName(): string {
    return this.name;
  }
  getNumber(): number {
    return this.number;
  }
  getNameSplit(): string {
    let splitName = this.name.split(" ");
    return splitName.length > 2 ? [splitName[splitName.length], splitName[0]].join(", ") : [splitName[1], splitName[0]].join(", ");
  }
}

abstract class Sheet {
  sheet: ExcelScript.Worksheet;
  constructor(workbook: ExcelScript.Workbook, name: string) {
    this.sheet = workbook.addWorksheet(name);
  }
  build() { }
  format() { }
  mergeCells(range: string) {
    this.sheet.getRange(range).merge();
  }
  write_horizontal(data: Array<string | number | boolean | Deal>, row: number, col: number = 0) {
    data.forEach((d, i) => {
      this.sheet.getCell(row, i + col).setValue(d);
    });
  }
  write_vertical(data: Array<string | number | boolean | Deal>, col: number, row: number = 0) {
    data.forEach((d, i) => {
      this.sheet.getCell(i + row, col).setValue(d);
    });
  }
  write_cell(data: string | number, row: number, col: number) {
    this.sheet.getCell(row, col).setValue(data);
  }
  resizeColumns(width: number = 0) {
    let range = this.sheet.getRange("1:1");
    width > 0 ? range.getFormat().setColumnWidth(width) : range.getFormat().autofitColumns();
  }
  resizeRows(height: number = 0) {
    let range = this.sheet.getRange("A:A");
    height > 0 ? range.getFormat().setColumnWidth(height) : range.getFormat().autofitColumns();
  }
  changeFontColor(range: ExcelScript.Range, color: string) {
    range.getFormat().getFont().setColor(color);
  }
  changeCellColor(range: ExcelScript.Range, color: string) {
    range.getFormat().getFill().setColor(color);
  }
  conditionalFormatting(range: string, minColor: Color, midColor: Color, maxColor: Color) {
    let conditionalFmt = this.sheet.getRange(range).addConditionalFormat(ExcelScript.ConditionalFormatType.colorScale);
    conditionalFmt.getColorScale().setCriteria({
      minimum: {
        color: minColor,
        type: ExcelScript.ConditionalFormatColorCriterionType.lowestValue
      },
      midpoint: {
        color: midColor,
        formula: '=40',
        type: ExcelScript.ConditionalFormatColorCriterionType.percentile
      },
      maximum: {
        color: maxColor,
        type: ExcelScript.ConditionalFormatColorCriterionType.highestValue
      }
    });
  }
  set_bold(ranges: Array<string>) {
    ranges.forEach(range => {
      this.sheet.getRange(range).getFormat().getFont().setBold(true);
    });
  }
  set_font_color(ranges: Array<string>, color: Color) {
    ranges.forEach(range => {
      this.sheet.getRange(range).getFormat().getFont().setColor(color);
    });
  }
  set_cell_color(ranges: Array<string>, color: Color) {
    ranges.forEach(range => {
      this.sheet.getRange(range).getFormat().getFill().setColor(color);
    });
  }
  set_col_width(ranges: Array<string>, width: number) {
    ranges.forEach(range => {
      this.sheet.getRange(range).getFormat().setColumnWidth(width);
    });
  }
  text_set_center(range: ExcelScript.Range) {
    range.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    range.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  }
}

class Report {
    private name: string;
    private rawData: Array<string | number | boolean>[];
    constructor(name: string, rawData: Array<string | number | boolean>[]) {
        this.name = name;
        this.rawData = rawData;
    }
    getName(): string {
        return this.name;
    }
    getHeaders(): Array<string | number | boolean> {
        return this.rawData[0];
    }
    getDataSet(): Array<string | number | boolean>[] {
        return this.rawData.splice(1, this.rawData.length);
    }
}

class Customer extends Person {
    constructor(name: string, number: number) {
        super(name, number);
    }
}

class Employee extends Person {
    private deals: Array<Deal>;
    private bucket: number;
    private spiffs: number;
    private stars: boolean;
    constructor(name: string, number: number) {
        super(name, number);
        this.bucket = 0;
        this.spiffs = 0;
        this.stars = false;
        this.deals = [];
    }
    addDeal(deal: Deal) {
        this.deals.push(deal);
    }
    getDeals(): Array<Deal> {
      return this.deals;
    }
    getDealCount(): number {
      return this.deals.length;
    }
    getTotalUnits(): number{
      let count = 0;
      this.deals.forEach(deal => {
        count += deal.getUnits()
      });
      return count;
    }
    getTotalCommGross(): number{
      let amount = 0;
      this.deals.forEach(deal => {
        amount += deal.getCommGross()
      });
      return amount;
    }
    getTotalFiGross(): number{
      let amount = 0;
      this.deals.forEach(deal => {
        amount += deal.getFiGross()
      });
      return amount;
    }
    getTotalCommAmount(): number{
      let amount = 0;
      this.deals.forEach(deal => {
        amount += deal.getCommAmnt()
      });
      return amount;
    }
    getBucket(): number {
        return this.bucket;
    }
    setBucket(amount: number) {
        this.bucket = amount;
    }
    getStars(): boolean {
      return this.stars;
    }
    setStars(reportTxt: string) {
      this.stars = reportTxt === 'YES' ? true : false;
    }
}

class Vehicle {
    private prefix: string;
    private stockNumber: string;
    private description: string;
    private saleType: string;
    constructor(prefix: string, stockNumber: string, description: string, saleType: string) {
        this.prefix = prefix;
        this.stockNumber = stockNumber;
        this.description = description;
        this.saleType = saleType;
    }
    getPrefix(): string {
        return this.prefix;
    }
    getStockNumber(): string {
        return this.stockNumber;
    }
    getDescription(): string {
        return this.description;
    }
    getSaleType(): string {
        return this.saleType;
    }
    getYear(): string {
      return `${20}` + `${this.description.split(",")[0]}`;
    }
    getMake(): string {
      return this.description.split(",")[1];
    }
    getModel(): string {
      return this.description.split(",")[2];
    }
    getData(): Array<string|number> {
      return [
        this.getYear(),
        this.getMake(),
        this.getModel(),
        this.saleType,
        this.stockNumber
      ];
    }
}

class Deal {
    private date: string;
    private number: string;
    private customer: Customer;
    private vehicle: Vehicle;
    private units: number;
    private commGross: number;
    private fiGross: number;
    private commAmnt: number;
    constructor(date: string, number: string, customer: Customer, vehicle: Vehicle, units: number, commGross: number, fiGross: number, commAmnt: number) {
        this.date = date;
        this.number = number;
        this.customer = customer;
        this.vehicle = vehicle;
        this.units = units;
        this.commGross = commGross;
        this.fiGross = fiGross;
        this.commAmnt = commAmnt;
    }
    getDate(): string {
        return this.date;
    }
    getNumber(): string {
        return this.number;
    }
    getCustomerName(): string {
      return this.customer.getName();
    }
    getVehicle(): Vehicle {
      return this.vehicle;
    }
    getUnits(): number {
      return this.units;
    }
    getCommGross(): number {
      return this.commGross;
    }
    getFiGross(): number {
      return this.fiGross;
    }
    getCommAmnt(): number {
      return this.commAmnt;
    }
    getData(): Array<string|number> {
      return [
        this.date,
        this.number,
        this.customer.getName(),
        ...this.vehicle.getData(),
        this.units,
        this.commGross,
        this.fiGross,
        this.commAmnt
      ];
    }
}

// This is a Test Class to see if we can make the sections easier to manage
// abstract class Section {
//   sheet: Sheet;
//   data: Object;
//   address: Object;
//   constructor(sheet: Sheet, data: Object, address: Object) {
//     this.sheet = sheet;
//     this.data = data;
//     this.address = address;
//   }
//   build(has_row_numbers:boolean = true, total_on_top:boolean = false) {
//     if(total_on_top){
//       let row = this.address.row + 1;
//       this.sheet.write_horizontal(this.data.totals, this.address.row, this.address.column);
//       this.sheet.write_horizontal(this.data.headers, row, this.address.column);
//       this.data.rows.forEach((data, index) => {
//         let row_number: number = index + 1;
//         has_row_numbers ? this.sheet.write_horizontal([row_number, ...data], row + row_number, this.address.column) : this.sheet.write_horizontal(data, row + row_number, this.address.column);
//       });
//     } else {
//       this.sheet.write_horizontal(this.data.headers, this.address.row, this.address.column);
//       this.data.rows.forEach((data, index) => {
//         let row_number: number = index + 1;
//         has_row_numbers ? this.sheet.write_horizontal([row_number, ...data], this.address.row + row_number, this.address.column) : this.sheet.write_horizontal(data, this.address.row + row_number, this.address.column);
//       });
//     }
//   }
//   format() {}
// }

// class SalesSummary extends Section {
//   constructor(sheet: ExcelScript.Worksheet, data: Object) {
//     super(sheet, data);
//   }

// }

class SalesSheet extends Sheet {
  private employee: Employee;
  private tableRows: TableTotals;
  constructor(workbook: ExcelScript.Workbook, employee: Employee) {
    super(workbook, employee.getName());
    this.employee = employee;
    this.tableRows = {
      totals: 9,
      header: 10,
      start: 11
    }
  }
  build() {
    let excelDate: number = Number(this.employee.getDeals()[0].getDate());
    let date = new Date(Math.round((excelDate - 25569) * 86400 * 1000));
    let month = date.toLocaleDateString("default", {month: "long", year: "numeric"});
    this.write_cell(this.employee.getNumber(), 0, 1);
    this.write_horizontal([
      this.employee.getName(),
      STORE_NAME
    ], 0, 3);
    this.write_cell(month, 0, 6);
    this.write_horizontal([
      "STARS",
      this.employee.getStars()
    ], 0, 8);
    this.write_cell("CSI Deduction", 0, 11);
    this.write_cell(``, 0, 13); // TODO: Link to created CSI Page

    // TODO: Put in the totals info...

    this.write_cell("TOTALS:", this.tableRows.totals, 7);
    this.write_horizontal([
      this.employee.getTotalUnits(),
      this.employee.getTotalCommGross(),
      this.employee.getTotalFiGross(),
      this.employee.getTotalCommAmount()
    ], this.tableRows.totals, 9);
    this.write_horizontal([
      "Date",
      "Deal #",
      "Customer",
      "Year",
      "Make",
      "Model",
      "N/U",
      "Stock #",
      "Unit Count",
      "Vehicle Gross",
      "Finance Gross",
      "Commission Amount"
    ], this.tableRows.header, 1);
    this.employee.getDeals().forEach((deal, index) => {
      let dealNumber: number = index + 1;
      let row: number = dealNumber + this.tableRows.header;
      this.write_horizontal([
        dealNumber,
        ...deal.getData()
      ], row);
    });
  }
  format() {
    let headRangeFmt: ExcelScript.RangeFormat = this.sheet.getRange("A1:M10").getFormat();
    headRangeFmt.getFont().setBold(true);
    headRangeFmt.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

    let dealsLastRow = this.employee.getDealCount() + (this.tableRows.start + 1);
    this.sheet.getRange("A:A").getFormat().getFont().setBold(true);
    this.mergeCells("E1:F1");
    this.mergeCells("L1:M1");
    this.sheet.getRange(`A${this.tableRows.header + 1}:M${this.tableRows.header + 1}`).getFormat().getFont().setBold(true);
    this.sheet.getRange(`A${this.tableRows.header + 1}:M${this.tableRows.header + 1}`).getFormat().setTextOrientation(90);
    this.mergeCells(`H${this.tableRows.totals + 1}:I${this.tableRows.totals + 1}`);
    this.sheet.getCell(0, 4).getFormat().setColumnWidth(50);
    this.set_cell_color([`A${this.tableRows.header + 1}:M${this.tableRows.header + 1}`],Color.BLUE_GREY)
    this.employee.getDeals().forEach((deal, index) => {
      let row = index + (this.tableRows.totals + 2);
      if (row % 2 === 1) {
        this.sheet.getRange(`A${row}:M${row}`).getFormat().getFill().setColor(Color.BLUE_GREY);
      }
    });
    this.sheet.getRange(`B8:B${dealsLastRow}`).setNumberFormat(NumberFormat.DATE);
    this.text_set_center(this.sheet.getRange(`J8:J${dealsLastRow}`));
    this.sheet.getRange(`K6:M${dealsLastRow}`).setNumberFormat(NumberFormat.ACCOUNTING);

    this.sheet.getRange(`A${this.tableRows.header + 1}:M${this.tableRows.header + 1}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    this.sheet.getRange(`A${this.tableRows.header + 1}:M${this.tableRows.header + 1}`).getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.bottom);

    this.sheet.getRange("1:1").getFormat().autofitColumns();
  }
}

function main(workbook: ExcelScript.Workbook) {
  let sheets = workbook.getWorksheets();
  let reports: Array<Report> = [];
  let all_employees: Array<Employee> = [];
  let all_sheets: Array<Sheet> = [];

  // RESET / RELOAD the workbook
  if (sheets.length > 4) {
    clearSheets(sheets);
    sheets = workbook.getWorksheets();
  }
  
  sheets.forEach(sheet => {
    reports.push(new Report(sheet.getName(), sheet.getUsedRange().getValues()));
  });
  reports.forEach(report => {
    let reportName = report.getName();
    let data = report.getDataSet();
    switch(reportName) {
        case '0432': 
          data.forEach(row => {
            if(all_employees.length > 0) {
              if (all_employees.filter(employee => employee.getName() === String(row[3])).length > 0) {
                all_employees.forEach(employee => {
                  if (String(row[3]) === employee.getName()) {
                    employee.addDeal(createDeal(row));
                  }
                });
              } else {
                all_employees.push(createEmployee(row));
              }
            } else {
              all_employees.push(createEmployee(row));
            }
          });
        break;
        case 'SPIFF SCH': 
          data.forEach(row => {
            all_employees.forEach(employee => {
              if (employee.getNumber() === row[0]) {
                employee.setBucket(Number(row[4]));
              }
            });
          });
        break;
        case 'SLS COMM SCH': 
          data.forEach(row => {
            all_employees.forEach(employee => {
              if (employee.getNumber() === row[0]) {
                employee.setBucket(Number(row[4]));
              }
            });
          });
        break;
        case 'STARS': 
          data.forEach(row => {
            all_employees.forEach(employee => {
              if(employee.getNumber() === row[0]) {
                employee.setStars(String(row[2]));
              }
            });
          });
        break;
        default: console.log("Not Built: ", report.getName());
        break;
    }
  });

  all_employees.forEach(employee => {
    all_sheets.push(new SalesSheet(workbook, employee));
    
  });

  all_sheets.forEach(sheet => {
    sheet.build();
    sheet.format();
  });
}

function createEmployee(data: Array<string|number|boolean>): Employee {
  let emp = new Employee(String(data[3]), Number(data[2]));
  let cust = new Customer(String(data[5]), Number(data[4]));
  let veh = new Vehicle(String(data[6]), String(data[7]), String(data[8]), String(data[9]));
  let deal = new Deal(String(data[0]), String(data[1]), cust, veh, Number(data[11]), Number(data[10]), Number(data[12]), Number(data[13]));
    emp.addDeal(deal);
  return emp;
}

function createDeal(data: Array<string | number | boolean>): Deal {
  let cust = new Customer(String(data[5]), Number(data[4]));
  let veh = new Vehicle(String(data[6]), String(data[7]), String(data[8]), String(data[9]));
  return new Deal(String(data[0]), String(data[1]), cust, veh, Number(data[11]), Number(data[10]), Number(data[12]), Number(data[13]));
}





// This is used while testing the Workbook, Clears the created pages so they can be rebuilt with the new code
// Can be left in to reload a workbook while working on it, or if the report changes, etc.
// NOTE:
// If you reload, it will erase all created sheets and create them again. So if you've added to a sheet that 
// is NOT a report, this will be overwritten when the sheet recreates.
function clearSheets(sheets: Array<ExcelScript.Worksheet>): void {
  let length = sheets.length - 1;
  for(let i=4; i<=length; i++) {
    sheets[i].delete()
  }
}
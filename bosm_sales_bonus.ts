// Reports needed for Script: 0432 (deals), 3213 (bucket), Spiff Schedule
function main(workbook: ExcelScript.Workbook) {
  const initialSheets: Array<ExcelScript.Worksheet> = workbook.getWorksheets();

  let data_0432: Array<string | number | boolean>[] = []
  let data_90day: Array<string | number | boolean>[] = []
  let data_3213: Array<string | number | boolean>[] = []
  let data_spiffs: Array<string | number | boolean>[] = []
  let data_nps: Array<string | number | boolean>[] = []
  let data_lut: Array<string | number | boolean>[] = []

  initialSheets.forEach(sheet => {
    switch (sheet.getName()) {
      case '0432': data_0432 = sheet.getUsedRange().getValues();
        break;
      case '90': data_90day = sheet.getUsedRange().getValues();
        break;
      case '3213': data_3213 = sheet.getUsedRange().getValues();
        break;
      case 'SPIFFS': data_spiffs = sheet.getUsedRange().getValues();
        break;
      case 'NPS Sheet': data_nps = sheet.getUsedRange().getValues();
        break;
      case 'Look Up Table': data_lut = sheet.getUsedRange().getValues();
        break;
      default: sheet.delete()
        break;
    }
  })

  const store = new Store(data_lut);
  store.createEmployees(data_0432);

  new NpsSheet(workbook, store.employees);

  new PaySummarySheet(workbook, store.employees);
  new JvSheet(workbook, store.employees);

  store.employees.forEach(employee => {
    new SalesSheet(workbook, employee)
  });
}


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


// --------   PAGES   -------- \\
class NpsSheet {
  constructor(workbook: ExcelScript.Workbook, employeeList: Array<Employee>) {
    const sheet = workbook.addWorksheet('NPS');
    sheet.getRange("A1:H1").setValues([
      ["Regional Score", "Employee #", "Employee Name", "# of Surveys", "Monthly Score", "90 Day Score", "NPS Score for Bonus", "CSI Outcome"]
    ])

    employeeList.forEach((employee, index) => {
      const row = index + 2;
      sheet.getRange(`B${row}:C${row}`).setValues([[employee.id, employee.name]])
      sheet.getRange(`D${row}:G${row}`).setValues([[
        `=IFERROR(VLOOKUP(B${row},'NPS Sheet'!\$B\$4:\$AC\$45,5,FALSE),0)`,
        `=IFERROR(VLOOKUP(B${row},'NPS Sheet'!\$B\$4:\$AC\$45,8,FALSE),0)`,
        `=IFERROR(VLOOKUP(B${row},'NPS Sheet'!\$B\$4:\$AC\$45,23,FALSE),0)`,
        `=IF(E${row}>F${row},E${row},F${row})`
      ]])
      sheet.getCell(row - 1, 7).setValue(`=IF(G${row}>A2+3%,"3P",IF(G${row}=A2,"A",IF(G${row}<A2,"B")))`)
    })
    sheet.getRange("1:1").getFormat().getFill().setColor("lightgrey")
    sheet.getRange("1:1").getFormat().getFont().setBold(true)
    sheet.getRange("1:1").getFormat().setColumnWidth(120)
    sheet.getRange("1:1").getFormat().setRowHeight(50)

    sheet.getRange("A2:A3").merge()
    sheet.getCell(1, 0).getFormat().getFill().setColor("yellow")
    sheet.getCell(1, 0).setNumberFormat("0.0%")

    sheet.getRange("E2:G100").setNumberFormat('0%')

    sheet.getRange("A1:H100").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center)
    sheet.getRange("A1:H100").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center)
  }
}

class PaySummarySheet {
  constructor(workbook: ExcelScript.Workbook, employeeList: Array<Employee>) {
    const sheet = workbook.addWorksheet('Pay Summary');
    sheet.getRange("A1:P1").setValues([
      ["Employee #", "Employee Name", "Total Units", "Rank", "Spiffs to Pay", "Commission 3120", "Retro Commission", "F&I Commission", "Month End Bonus 3122", "Total EOM Bonus 8328", "Draw 3121", "Total Commission", "YTD Bucket", "Deposit Gross", "Check Column - Should be Zero", "Draw to Take"]
    ]);

    employeeList.forEach((employee, index) => {
      const row = index + 2;
      sheet.getRange(`A${row}:G${row}`).setValues([[
        employee.id,
        employee.name,
        employee.getTotalUnits(),
        `=IF(C${row}<15, "", RANK.EQ(C${row}, C:C, 0))`,
        `=IFERROR(VLOOKUP(A${row}, 'SPIFFS'!A:H, 8, 0), 0)`,
        employee.getTotalCommission('grossAmount'),
        `=IFERROR('${employee.name}'!M${employee.getResultRow(5)}, 0)`
      ]]);
    });
  }
}

class JvSheet {
  constructor(workbook: ExcelScript.Workbook, employeeList: Array<Employee>) {
    const sheet = workbook.addWorksheet('JV Posting');
    sheet.getRange("A1:M1").setValues([
      ["Employee #", "Employee Name", "Draw", "Commission", "Retro Commission", "F&I Commission", "Bonus", "Spiffs", "Total Comm/Bonus", "Total Due/Owed", "YTD Bucket", "Expense 1", "Expense 2"]
    ]);

    employeeList.forEach((employee, index) => {
      const row = index + 2;
      sheet.getRange(`A${row}:B${row}`).setValues([[employee.id, employee.name]])
    });
  }
}

class SalesSheet {
  constructor(workbook: ExcelScript.Workbook, employee: Employee) {
    if (employee.deals.length == 0) return;
    const sheet = workbook.addWorksheet(employee.name);

    const data_lastRow = employee.getReportResultRow()
    const report_lastRow = employee.getReportResultRow() - 1;
    const results_row1 = employee.getResultRow(1);
    const results_row2 = employee.getResultRow(2);
    const results_row3 = employee.getResultRow(3);
    const results_row4 = employee.getResultRow(4);
    const results_row5 = employee.getResultRow(5);
    const results_row6 = employee.getResultRow(6);
    const results_row7 = employee.getResultRow(7);
    const results_row8 = employee.getResultRow(8);
    const results_row9 = employee.getResultRow(9);
    const results_row10 = employee.getResultRow(10);
    const results_row11 = employee.getResultRow(11);
    const results_row12 = employee.getResultRow(12);
    const results_row13 = employee.getResultRow(13);
    const results_row14 = employee.getResultRow(14);
    const results_row15 = employee.getResultRow(15);
    const results_row16 = employee.getResultRow(16);

    const headerRange: ExcelScript.Range = sheet.getRange("A1:B6");
    const colHeaderRange: ExcelScript.Range = sheet.getRange("A7:P7"); headerRange;
    const reportRange: ExcelScript.Range = sheet.getRange(`A8:R${data_lastRow}`);

    let reportData: Array<string | number>[] = [];

    employee.deals.forEach(deal => {
      reportData.push([deal.date, deal.id, deal.customer.id, deal.customer.name, deal.vehicle.id, deal.vehicle.year, deal.vehicle.make, deal.vehicle.model, deal.vehicle.salesType, deal.commission.gross, deal.commission.grossPercentage, deal.unitCount, deal.commission.amount]);
    });

    headerRange.setValues([
      ["Name", employee.name],
      ["Employee Number", employee.id],
      ["90 Day Rolling Average #", `=VLOOKUP(B2, '90'!A:E, 5, 0)`],
      ["CSI", `=IFERROR(VLOOKUP(B2, 'NPS'!B:H, 7, 0), 0)`],
      ["# of Surveys", `=IFERROR(VLOOKUP(B2, 'NPS'!B:H, 3, 0), 0)`],
      ["Retro Percentage", `=VLOOKUP(${employee.getTotalUnits()}, 'Look Up Table'!A:B, 2, TRUE)`]
    ]);
    headerRange.getFormat().getFont().setBold(true);
    sheet.getRange("B1:B6").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right)
    sheet.getCell(5, 1).setNumberFormat("0%")

    colHeaderRange.setValues([["Date", "Reference #", "Customer #", "Customer Name", "Stock #", "Year", "Make", "Model", "Sale Type", "Commission F&I", "Commission Gross", "Units", "Commission Amount", "Retro Mini", "Retro Owed", "Retro Commission Payout"]]);
    colHeaderRange.getFormat().getFill().setColor("lightgrey");
    colHeaderRange.getFormat().setRowHeight(50);
    colHeaderRange.getFormat().getFont().setBold(true);
    colHeaderRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    colHeaderRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);

    reportData.forEach((row, index) => {
      const lineNumber = index + 8;
      sheet.getRange(`A${lineNumber}:M${lineNumber}`).setValues([row]);
      sheet.getRange(`N${lineNumber}:P${lineNumber}`).setValues([[
        `=IF(M${lineNumber}<=251, VLOOKUP(B3, 'Look Up Table'!I:J, 2, TRUE) * L${lineNumber}, 0)`,
        `=IF(N${lineNumber}>0, N${lineNumber} - M${lineNumber}, 0)`,
        `=IF(N${lineNumber} = 0, K${lineNumber} * B6, 0)`
      ]])
    })

    sheet.getRange(`J${data_lastRow}:P${data_lastRow}`).setValues([[
      `=SUM(J8:J${report_lastRow})`,
      `=SUM(K8:K${report_lastRow})`,
      `=SUM(L8:L${report_lastRow})`,
      `=SUM(M8:M${report_lastRow})`,
      `=SUM(N8:N${report_lastRow})`,
      `=SUM(O8:O${report_lastRow})`,
      `=SUM(P8:P${report_lastRow})`,
    ]]);

    sheet.getRange(`J${results_row1}:M${results_row1}`).setValues([[
      "Prior Draw Balance", '', '', `=-VLOOKUP(B2, '3213'!A:G, 7, 0)`
    ]]);

    sheet.getRange(`J${results_row2}:M${results_row2}`).setValues([[
      "Commission", 0.18, '', `=M${data_lastRow}`
    ]]);
    sheet.getRange(`K${results_row2}`).setNumberFormat("0%")

    sheet.getRange(`J${results_row3}:M${results_row3}`).setValues([[
      "Retro Commission", '=B6', '', `=P${data_lastRow}`
    ]]);

    sheet.getRange(`J${results_row4}:M${results_row4}`).setValues([[
      "Retro MINI", '', '', `=O${data_lastRow}`
    ]]);

    sheet.getRange(`J${results_row5}:M${results_row5}`).setValues([[
      "Total Retro Commission", '', '', `=SUM(P${data_lastRow}, O${data_lastRow})`
    ]]);

    sheet.getRange(`J${results_row6}:M${results_row6}`).setValues([[
      "Total F&I", '', '', `=J${data_lastRow}`
    ]]);

    sheet.getRange(`J${results_row7}:M${results_row7}`).setValues([[
      "25% Reserve F&I", -0.25, '', `=K${results_row7} * M${results_row6}`
    ]]);
    sheet.getRange(`K${results_row7}`).setNumberFormat("0%");

    sheet.getRange(`J${results_row8}:M${results_row8}`).setValues([[
      "Total F&I Payable Gross", '', '', `=M${results_row6} + M${results_row7}`
    ]]);

    sheet.getRange(`J${results_row9}:M${results_row9}`).setValues([[
      "Total F&I Payout", 0.05, '', `=K${results_row9} * M${results_row8}`
    ]]);
    sheet.getRange(`K${results_row9}`).setNumberFormat("0%");

    sheet.getRange(`J${results_row10}:M${results_row10}`).setValues([[
      "Top Salesman Bonus", "=VLOOKUP(B2,'Pay Summary'!A:D,4,0)", '', `=IF(K${results_row10} = 1, 500, 0)`
    ]]);

    sheet.getRange(`J${results_row11}:M${results_row11}`).setValues([[
      "Unit Bonus", `=L${data_lastRow}`, '', `=VLOOKUP(K${results_row11}, 'Look Up Table'!E:F, 2, TRUE)`
    ]]);

    sheet.getRange(`J${results_row12}:O${results_row12}`).setValues([[
      "CSI", '=B4', '', `=IF(B5>=3,IF(B4="3P",L${data_lastRow}*50,IF(B4="A",0,IF(B4="B",L${data_lastRow}*-50))),0)`, "Rolling 90 Day", `='NPS Sheet'!X51`
    ]]);
    sheet.getRange(`K${results_row12}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);

    sheet.getRange(`J${results_row13}:M${results_row13}`).setValues([[
      "Total Bonus", '', '', `=SUM(M${results_row11}, M${results_row12}, M${results_row10})`
    ]]);

    sheet.getRange(`J${results_row14}:M${results_row14}`).setValues([[
      "Spiff", '', '', `=IFERROR(VLOOKUP(B2,'SPIFFS'!A:H,8,0),0)`
    ]]);

    sheet.getRange(`J${results_row15}:M${results_row15}`).setValues([[
      "Total Pay", '', '', `=SUM(M${results_row2}, M${results_row5}, M${results_row9}, M${results_row13}, M${results_row1}, M${results_row14})`
    ]]);

    sheet.getRange(`J${results_row16}:M${results_row16}`).setValues([[
      "Bucket Total YTD", '', '', `=IF(M${results_row15}<0, SUM(M${results_row1}, M${results_row13}, M${results_row9}, M${results_row5}), 0)`
    ]]);

    sheet.getRange(`J${results_row1}:J${results_row16}`).getFormat().getFont().setBold(true);
    sheet.getRange(`M${results_row1}:M${results_row16}`).setNumberFormat(NumberFormat.ACCOUNTING);
    sheet.getRange(`N${results_row12}:O${results_row12}`).getFormat().getFont().setBold(true);

    const employeeSignatureRow = data_lastRow + 6;
    const managerSignatureRow = employeeSignatureRow + 6;

    sheet.getRange(`A${employeeSignatureRow}:C${employeeSignatureRow}`).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
    sheet.getRange(`D${employeeSignatureRow}`).setValue("Employee");

    sheet.getRange(`A${managerSignatureRow}:C${managerSignatureRow}`).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
    sheet.getRange(`D${managerSignatureRow}`).setValue("Manager");

    sheet.getRange(`J8:K${data_lastRow}`).setNumberFormat(NumberFormat.ACCOUNTING);
    sheet.getRange(`M8:P${data_lastRow}`).setNumberFormat(NumberFormat.ACCOUNTING);
    sheet.getRange(`A${data_lastRow}:P${data_lastRow}`).getFormat().getFill().setColor(Color.LIGHT_GREY);

    sheet.getRange("A:A").setNumberFormat(NumberFormat.DATE);
    reportRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    colHeaderRange.getFormat().autofitColumns();
  }
}


// --------  CLASSES  -------- \\
class Store {
  public name: string;
  public abbr: string;
  public employees: Array<Employee>;
  constructor(public lookupData: Array<string | number | boolean>[]) {
    this.name = "BMW of South Miami";
    this.abbr = "BOSM";
    this.lookupData = lookupData;
    this.employees = [];
  }

  createEmployees(data_0432: Array<string | number | boolean>[]) {
    data_0432.forEach((row, index) => {
      const [date, deal_id, emp_id, emp_name, cust_id, cust_name, prefix, veh_id, veh_desc, sale_type, comm_fni, comm_front, unit_count, comm_amount] = row

      if (index > 0) {
        let employee = this.employees.filter(emp => emp.id == emp_id)[0];

        if (!employee) {
          employee = this.addEmployee(Number(emp_id), String(emp_name).toLocaleUpperCase());
        }
        if (Number(unit_count) > 0) {
          employee.addDeal(String(deal_id), Number(date), Number(cust_id), String(cust_name), String(veh_id), String(veh_desc), String(sale_type), Number(unit_count), Number(comm_fni), Number(comm_front), Number(comm_amount));
        }
      }
    })
  }

  addEmployee(id: number, name: string) {
    const employee = new Employee(id, name);
    this.employees.push(employee);
    return employee;
  }
}

abstract class Person {
  constructor(public id: number, public name: string) {
    this.id = id;
    this.name = name;
  }
}

class Customer extends Person {
  constructor(public id: number, public name: string) {
    super(id, name);
  }
}

class Employee extends Person {
  public deals: Array<Deal>;
  constructor(public id: number, public name: string) {
    super(id, name);

    this.deals = [];
  }

  addDeal(id: string, date: number, custId: number, custName: string, vehId: string, vehDesc: string, salesType: string, unitCount: number, commGross: number, grossPercent: number, commAmount: number) {
    const customer = new Customer(custId, custName);
    const vehicle = new Vehicle(vehId, vehDesc, salesType);
    const commission = new Commission(commGross, grossPercent, commAmount);
    this.deals.push(new Deal(id, date, customer, vehicle, unitCount, commission));
  }

  getReportResultRow() {
    return this.deals.length + 8;
  }

  getResultRow(rowNumber) {
    return this.getReportResultRow() + (2 * rowNumber);
  }

  getTotalUnits(filter?: string): number {
    let units = { 'new': 0, 'used': 0, 'total': 0 };
    this.deals.forEach(deal => {
      const t = deal.vehicle.salesType.toLowerCase();
      const count = deal.unitCount;
      units[t] += count;
      units.total += count;
    });
    return filter ? units[filter] : units['total'];
  }

  getTotalCommission(filter: string): number {
    let comms = { gross: 0, grossPercentage: 0, grossAmount: 0 };
    this.deals.forEach(deal => {
      comms.gross += deal.commission.gross;
      comms.grossPercentage += deal.commission.grossPercentage;
      comms.grossAmount += deal.commission.amount;
    });
    return comms[filter];
  }
}

class Commission {
  constructor(public gross: number, public grossPercentage: number, public amount: number) {
    this.gross = gross;
    this.grossPercentage = grossPercentage;
    this.amount = amount;
  }
}

class Vehicle {
  public year: string;
  public make: string;
  public model: string;
  public description: string;
  constructor(public id: string, public vehicleDescription: string, public salesType: string) {
    this.id = id;
    [this.year, this.make, this.model, this.description] = vehicleDescription.split(',');
    this.salesType = salesType;
  }
}

class Deal {
  constructor(public id: string, public date: number, public customer: Customer, public vehicle: Vehicle, public unitCount: number, public commission: Commission) {
    this.id = id;
    this.date = date;
    this.customer = customer;
    this.vehicle = vehicle;
    this.unitCount = unitCount;
    this.commission = commission;
  }

  calculateRetroMini() {
    if(this.commission.amount > 251) return 0;
    
  }
}


// -------- FUNCTIONS -------- \\
function calculateRetroBonus(input: number): number {
  if (input >= 16) return 7;
  if (input >= 12 && input < 16) return 4;
  return 0;
}

function calculateUnitBonus(input: number): number {
  if (input >= 10 && input < 12) return 375;
  if (input >= 12 && input < 16) return 750;
  if (input >= 16 && input < 20) return 1500;
  if (input >= 20 && input < 24) return 2500;
  if (input >= 24) return 3000;
  return 0;
}

function calculateRollingMini(input: number): number {
  if (input >= 12 && input < 16) return 250;
  if (input >= 16 && input < 20) return 300;
  if (input >= 20 && input < 24) return 350;
  if (input >= 24) return 400;
  return 200;
}

function calculate90DayUnitBonus(input: number): number {
  if (input >= 12 && input < 16) return 250;
  if (input >= 16 && input < 20) return 300;
  if (input >= 20 && input < 24) return 350;
  if (input >= 24) return 400;
  return 200;
}

function calculateTotalEOMBonus8328(retroCommission: number, fniCommission: number, monthEndBonus3122: number, spiffsToPay: number): number {
  return retroCommission + fniCommission + monthEndBonus3122 + spiffsToPay;
}

function calculateTotalCommissions(commission3120: number, totalEOMBonus8328: number): number {
  return commission3120 + totalEOMBonus8328;
}

function calculateYTDBucket(totalCommissions: number, draw3121: number, spiffsToPay: number): number {
  if (totalCommissions - draw3121 > 0) return 0;
  return totalCommissions - draw3121 - spiffsToPay;
}

function calculateDepositGross(totalCommissions: number, draw3121: number, ytdBucket: number): number {
  return totalCommissions - draw3121 - ytdBucket;
}

function calculateDrawToTake(commission3120: number, monthEndBonus3122: number, draw3121: number): number {
  if (commission3120 + monthEndBonus3122 >= draw3121) return draw3121;
  return commission3120 + monthEndBonus3122;
}

function getNPSScore(individualMonthlyScore: number, individual90DayScore: number): number {
  if (individualMonthlyScore > individual90DayScore) return individualMonthlyScore;
  return individual90DayScore;
}

function calculateCSIOutcome(npsScore: number, regionalScore: number): string {
  const plusThreePercent = regionalScore + (regionalScore * 0.03);
  if (npsScore > plusThreePercent) return "3P";
  if (npsScore == regionalScore) return "A";
  return "B";
}

function calculateNPSPercentage(promoterValue: number, passiveValue: number, detractorValue: number): number {
  return ((promoterValue - detractorValue) / (promoterValue + passiveValue + detractorValue));
}

function calculateJVCommissionBonus(commission3120: number, retroCommission: number, fniCommission: number, monthEndBonus3122: number, spiffsToPay: number): number {
  return commission3120 + retroCommission + fniCommission + monthEndBonus3122 + spiffsToPay;
}

function calculateJVTotalDue(draw3121: number, jvCommissionBonus: number): number {
  return jvCommissionBonus - draw3121;
}

function calculateUnitTotals(newUnitCount: number, usedUnitCount: number): number {
  return newUnitCount + usedUnitCount;
}

function getPercentage(value: number, totalAmount: number): number {
  return value / totalAmount;
}

function calculateExpense(jvCommissionBonus: number, commission3120: number, percentage: number): number {
  return ((jvCommissionBonus - commission3120) * percentage);
}
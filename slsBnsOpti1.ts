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

class Sheet {
    public sheet: ExcelScript.Worksheet;
    public reportStartRow: number;
    constructor(workbook: ExcelScript.Workbook, sheetName: string, columnHeadRng: string, columnHeadValues: Array<string | number>[]) {
        this.sheet = workbook.addWorksheet(sheetName);
        
        const columnHeadRow = Number(columnHeadRng.split('')[1]);
        this.reportStartRow = columnHeadRow + 1;

        const columnHeadRange = this.sheet.getRange(columnHeadRng);
        columnHeadRange.setValues(columnHeadValues);
        columnHeadRange.getFormat().getFill().setColor(Color.LIGHT_GREY);
        columnHeadRange.getFont().setBold(true);
        columnHeadRange.setRowHeight(50);
    }
}

class NpsSheet extends Sheet {
    constructor(workbook: ExcelScript.Workbook, employees: Array<Employee>) {
        super(workbook, "NPS", "A1:H1", [["Regional Score", "Employee #", "Employee Name", "# of Surveys", "Monthly Score", "90 Day Score", "NPS Score for Bonus", "CSI Outcome"]]);

        employees.forEach((employee, index) => {
            const row = this.reportStartRow + index;
            const rowRange = this.sheet.getRange(`B${row}:H${row}`);
            rowRange.setValues([[
                employee.id, employee.name,
                `=IFERROR(VLOOKUP(B${row},'NPS Sheet'!\$B\$4:\$AC\$45,5,FALSE),0)`,
                `=IFERROR(VLOOKUP(B${row},'NPS Sheet'!\$B\$4:\$AC\$45,8,FALSE),0)`,
                `=IFERROR(VLOOKUP(B${row},'NPS Sheet'!\$B\$4:\$AC\$45,23,FALSE),0)`,
                `=IF(E${row}>F${row},E${row},F${row})`,
                `=IF(ISBLANK(A2), '', IF(G${row}>A2+3%,"3P",IF(G${row}=A2,"A",IF(G${row}<A2,"B"))))`
            ]]);
        });

        this.sheet.getRange("A2:A3").merge();
        this.sheet.getCell(1, 0).getFormat().getFill().setColor("yellow");
        this.sheet.getCell(1, 0).setNumberFormat("0.0%");
        this.sheet.getRange("E:G").setNumberFormat("0.0%");
        this.sheet.getRange("A:H").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
        this.sheet.getRange("A:H").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    }
}

class PaySummarySheet extends Sheet {
    constructor(workbook: ExcelScript.Workbook, employees: Array<Employee>) {
        super(workbook, "Pay Summary", "A1:P1", [["Employee #", "Employee Name", "Total Units", "Rank", "Spiffs to Pay", "Commission 3120", "Retro Commission", "F&I Commission", "Month End Bonus 3122", "Total EOM Bonus 8328", "Draw 3121", "Total Commission", "YTD Bucket", "Deposit Gross", "Check Column - Should be Zero", "Draw to Take"]]);
    
        employees.forEach((employee, index) => {
            const row = this.reportStartRow + index;
            const rowRange = this.sheet.getRange(`A${row}:G${row}`);
            rowRange.setValues([[
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

class JvSheet extends Sheet {
    constructor(workbook: ExcelScript.Workbook, employees: Array<Employee>) {
        super(workbook, "JV Posting", "A1:M1", [["Employee #", "Employee Name", "Draw", "Commission", "Retro Commission", "F&I Commission", "Bonus", "Spiffs", "Total Comm/Bonus", "Total Due/Owed", "YTD Bucket", "Expense 1", "Expense 2"]]);
    
        employees.forEach((employee, index) => {
            const row = this.reportStartRow + index;
            const rowRange = this.sheet.getRange(`A${row}:B${row}`);
            rowRange.setValues([[
                employee.id,
                employee.name
            ]]);
        });
    }
}

class SalesSheet extends Sheet {
    constructor(workbook: ExcelScript.Workbook, employee: Employee) {
        super(workbook, employee.name, "A7:P7", [["Date", "Reference #", "Customer #", "Customer Name", "Stock #", "Year", "Make", "Model", "Sale Type", "Commission F&I", "Commission Gross", "Units", "Commission Amount", "Retro Mini", "Retro Owed", "Retro Commission Payout"]]);
    
        const reportTotalsRow = this.reportStartRow + (employee.deals.length + 1);
        const results_row1 = reportTotalsRow + 2;
        const results_row2 = results_row1 + 2;
        const results_row3 = results_row2 + 2;
        const results_row4 = results_row3 + 2;
        const results_row5 = results_row4 + 2;
        const results_row6 = results_row5 + 2;
        const results_row7 = results_row6 + 2;
        const results_row8 = results_row7 + 2;
        const results_row9 = results_row8 + 2;
        const results_row10 = results_row9 + 2;
        const results_row11 = results_row10 + 2;
        const results_row12 = results_row11 + 2;
        const results_row13 = results_row12 + 2;
        const results_row14 = results_row13 + 2;
        const results_row15 = results_row14 + 2;
        const results_row16 = results_row15 + 2;
        const employeeSignatureRow = reportTotalsRow + 6;
        const managerSignatureRow = employeeSignatureRow + 6;

        const employeeDataRange: ExcelScript.Range = this.sheet.getRange("A1:B6");
        const reportRange: ExcelScript.Range = this.sheet.getRange(`A${this.reportStartRow}:R${reportTotalsRow - 1}`);

        employeeDataRange.setValues([
            ["Name", employee.name],
            ["Employee Number", employee.id],
            ["90 Day Rolling Average #", `=VLOOKUP(B2, '90'!A:E, 5, 0)`],
            ["CSI", `=IFERROR(VLOOKUP(B2, 'NPS'!B:H, 7, 0), 0)`],
            ["# of Surveys", `=IFERROR(VLOOKUP(B2, 'NPS'!B:H, 3, 0), 0)`],
            ["Retro Percentage", `=VLOOKUP(${employee.getTotalUnits()}, 'Look Up Table'!A:B, 2, TRUE)`]
        ]);

        employee.deals.forEach((deal, index) => {
            const lineNumber = index + this.reportStartRow;
        });
    }
}
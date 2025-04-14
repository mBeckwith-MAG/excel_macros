// Reports needed for Script: 0432 (deals), 3213 (bucket), Spiff Schedule


function main(workbook: ExcelScript.Workbook) {
    const initialSheets: Array<ExcelScript.Worksheet> = workbook.getWorksheets();

    let data_0432: Array<string | number | boolean>[][] = []
    let data_90day: Array<string | number | boolean>[][] = []
    let data_3213: Array<string | number | boolean>[][] = []

    initialSheets.forEach(sheet => {
        switch (sheet.getName()) {
            case '0432': data_0432.push(sheet.getUsedRange().getValues())
                break;
            case '90': data_90day.push(sheet.getUsedRange().getValues())
                break;
            case '3213': data_3213.push(sheet.getUsedRange().getValues())
                break;
            default: console.log(sheet.getName())
                break;
        }
    })

    const store = new Store()
    const allEmployees = getEmployees(data_0432[0])

    console.log(allEmployees)
}


class Store {
    public name: string;
    public abbr: string;
    public employees: Array<Employee>
    constructor() {
        this.name = "BMW of South Miami";
        this.abbr = "BOSM";
        this.employees = [];
    }

    addEmployee(employee: Employee) {
        this.employees.push(employee)
    }
}

class Customer {
    public id: number;
    public name: string;
    constructor(id, name) {
        this.id = id;
        this.name = name;
    }
}

class Employee {
    public id: number;
    public name: string;
    public deals: Array<Deal>
    constructor(id, name) {
        this.id = id;
        this.name = name;
        this.deals = []
    }

    addDeal(deal: Deal) {
        this.deals.push(deal)
    }
}

class Deal {
    public id: string;
    public date: string;
    public customer: Customer;
    public vehicle: Vehicle;
    public comm_fni: number;
    public comm_front: number;
    public unit_count: number;
    public comm_amount: number;
    constructor(id, date, customer, vehicle, comm_fni, comm_front, unit_count, comm_amount) {
        this.id = id;
        this.date = date;
        this.customer = customer;
        this.vehicle = vehicle;
        this.comm_fni = comm_fni;
        this.comm_front = comm_front;
        this.unit_count = unit_count;
        this.comm_amount = comm_amount;
    }
}

class Vehicle {
    public id: string;
    public year: number;
    public make: string;
    public model: string;
    public description: string;
    public sale_type: string;
    constructor(id: string, vehicleDescription: string, sale_type: string) {
        this.id = id
        // [this.year, this.make, this.model, this.description] = vehicleDescription.split(',')
        this.sale_type = sale_type
    }
}



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

function getEmployees(data_0432: Array<string | number | boolean>[]) {
    let employeeNumbers: Array<number> = [];
    let employees: Array<Employee> = []

    data_0432.forEach((row, index) => {
        const [date, deal_id, emp_id, emp_name, cust_id, cust_name, prefix, veh_id, veh_desc, sale_type, comm_fni, comm_front, unit_count, comm_amount] = row

        if (index > 0) {
            const customer = new Customer(Number(cust_id), String(cust_name))
            const vehicle = new Vehicle(String(veh_id), String(veh_desc), String(sale_type))

            if (Number(unit_count) > 0) {
                if (!employeeNumbers.includes(Number(emp_id))) {
                    const employee = new Employee(Number(emp_id), String(emp_name))
                    employeeNumbers.push(Number(emp_id))
                    employee.addDeal(new Deal(String(deal_id), String(date), customer, vehicle, Number(comm_fni), Number(comm_front), Number(unit_count), Number(comm_amount)))
                } else {
                    const employee = employees.filter(employee => employee.id == emp_id)[0]
                    employee.addDeal(new Deal(String(deal_id), String(date), customer, vehicle, Number(comm_fni), Number(comm_front), Number(unit_count), Number(comm_amount)))
                }
            }
        }
    })
    return employees
}
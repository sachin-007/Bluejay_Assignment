const XLSX = require('xlsx');
const moment = require('moment');

const workbook = XLSX.readFile('Timecard.xlsx');
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(sheet);

const processedEmployees = new Set();

function excelDateToDateString(excelDate) {
    return moment(new Date((excelDate - 25569) * 86400 * 1000)).format('MM/DD/YYYY hh:mm A');
}

function calculateTimeDifference(start, end) {
    const startTime = moment(excelDateToDateString(start), 'MM/DD/YYYY hh:mm A');
    const endTime = moment(excelDateToDateString(end), 'MM/DD/YYYY hh:mm A');
    return endTime.diff(startTime, 'hours', true);
}

function checkConsecutiveDays(employeeData) {
    const sortedEntries = employeeData.sort((a, b) => moment(excelDateToDateString(a.Time), 'MM/DD/YYYY hh:mm A') - moment(excelDateToDateString(b.Time), 'MM/DD/YYYY hh:mm A'));

    let consecutiveDaysCount = 1;

    for (let i = 1; i < sortedEntries.length; i++) {
        const currentEntry = sortedEntries[i];
        const previousEntry = sortedEntries[i - 1];

        const currentDate = moment(excelDateToDateString(currentEntry.Time), 'MM/DD/YYYY');
        const previousDate = moment(excelDateToDateString(previousEntry.Time), 'MM/DD/YYYY');

        if (currentDate.diff(previousDate, 'days') === 1) {
            consecutiveDaysCount++;
        } else {
            consecutiveDaysCount = 1;
        }

        if (consecutiveDaysCount === 7) {
            return true;
        }
    }

    return false;
}

function checkLessThan10HoursBetweenShifts(sortedEntries) {
    for (let i = 1; i < sortedEntries.length; i++) {
        const currentEntry = sortedEntries[i];
        const previousEntry = sortedEntries[i - 1];

        const hoursBetweenShifts = calculateTimeDifference(previousEntry.Time_Out, currentEntry.Time);

        if (hoursBetweenShifts > 1 && hoursBetweenShifts < 10) {
            return true;
        }
    }

    return false;
}

function checkMoreThan14HoursInSingleShift(employeeData) {
    for (let i = 0; i < employeeData.length; i++) {
        const entry = employeeData[i];
        
        if (entry['Timecard_Hours_(as_Time)'] && parseFloat(entry['Timecard_Hours_(as_Time)'].replace(':', '.')) > 14) {
            return true;
        }
    }

    return false;
}

data.forEach(employee => {
    if (processedEmployees.has(employee.Employee_Name)) {
        return;
    }

    const employeeData = data.filter(entry => entry.Employee_Name === employee.Employee_Name);
    const sortedEntries = employeeData.sort((a, b) => moment(excelDateToDateString(a.Time), 'MM/DD/YYYY hh:mm A') - moment(excelDateToDateString(b.Time), 'MM/DD/YYYY hh:mm A'));

    const employeePosition = `${employee.Employee_Name}, ${employee.Position_ID}`;

    if (checkConsecutiveDays(sortedEntries)) {
        console.log(`${employeePosition} worked for 7 consecutive days.`);
        processedEmployees.add(employee.Employee_Name);
    }

    if (checkLessThan10HoursBetweenShifts(sortedEntries)) {
        console.log(`${employeePosition} has less than 10 hours between shifts.`);
        processedEmployees.add(employee.Employee_Name);
    }

    if (checkMoreThan14HoursInSingleShift(employeeData)) {
        console.log(`${employeePosition} worked for more than 14 hours in a single shift.`);
        processedEmployees.add(employee.Employee_Name);
    }
});
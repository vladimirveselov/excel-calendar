const ExcelJS = require('exceljs');

function getMonthName(date) {
    return date.toLocaleString('ru', { month: 'long' });
}
function getWeekdayName(date) {
    return date.toLocaleString('ru', { weekday: 'long' });
}
function getWeekdayMon(weekdaySun) {
    return weekdaySun === 0 ? 6 : weekdaySun - 1;
}
function daysFromLastMonth(date) {
    return getWeekdayMon(date.getDay());
}
function createEmptyArray(raws, columns) {
    const array = new Array(raws);
    for (let i = 0; i < raws; i++) {
        array[i] = new Array(columns);
        for (let j = 0; j < columns; j++) {
            array[i][j] = '';
        }
    }
    return array;
}
function createWeeksArray(daysBeforFirstFromMonday, daysInMonth, weeksInMonth) {
    const array = createEmptyArray(weeksInMonth, 7);
    for (let n = daysBeforFirstFromMonday; n < daysBeforFirstFromMonday + daysInMonth; n++) {
        const i = Math.floor(n / 7);
        const j = n - i * 7;
        array[i][j] = (n - daysBeforFirstFromMonday + 1).toString();
    }
    return array;
}
function getWeeks(year, month) {
    const firstDate = new Date(year, month, 1);
    const lastDate = new Date(year, month + 1, 0);
    const daysInMonth = lastDate.getDate() - firstDate.getDate() + 1;
    const daysBeforFirstFromMonday = daysFromLastMonth(firstDate);
    const monthName = getMonthName(firstDate);
    const firstWeekday = getWeekdayName(firstDate);
    const weeksInMonth = Math.ceil((daysInMonth + daysBeforFirstFromMonday) / 7);
    const weeks = createWeeksArray(daysBeforFirstFromMonday, daysInMonth, weeksInMonth);
    return weeks;
}

const workbook = new ExcelJS.Workbook();

workbook.creator = 'program';
workbook.created = new Date();

const today = new Date();
const year = process.argv[2] || today.getFullYear();
console.log(`Printing calendar for year ${year}`);
for (let month = 0; month < 12; month ++) {
    const firstDate = new Date(year, month, 1);
    const monthName = getMonthName(firstDate);
    const worksheet = workbook.addWorksheet( monthName, {
        pageSetup: {
            paperSize: undefined, orientation: 'landscape'
        }
    });
    worksheet.properties.defaultColWidth = 17;
    const firstRow = worksheet.getRow(1);
    firstRow.height = 20;
    firstRow.getCell(1).value = `${monthName} ${year}`;
    firstRow.getCell(1).font =  {
        name: 'Arial Black',
        family: 2,
        size: 16,
        bold: false
      };
    firstRow.getCell(1).alignment = { vertical: 'top', horizontal: 'left' };
    const secondRow = worksheet.getRow(2);
    secondRow.height = 20;
    for (let day = 0; day < 7; day ++) {
        const tmpDate = new Date(year, month, day);
        const wdMon = getWeekdayMon(tmpDate.getDay());
        const cell = secondRow.getCell(1 + wdMon);
        cell.value = getWeekdayName(tmpDate);
        cell.alignment = { vertical: 'top', horizontal: 'left' };
        cell.border = {
            top: {style:'thin'},
            left: {style:'thin'},
            bottom: {style:'thin'},
            right: {style:'thin'}
        };
        cell.font =  {
            name: 'Arial Black',
            family: 2,
            size: 12,
            bold: false
        }; 
    }

    const weeks = getWeeks(year, month);
    for (let i = 0; i < weeks.length; i++) {
        const week = weeks[i];
        const row = worksheet.getRow(3 + i);
        row.height = 80;
        for (let j = 0; j < week.length; j++) {
            const cell = row.getCell(1 + j);
            cell.value = week[j];
            cell.alignment = { vertical: 'top', horizontal: 'left' };
            cell.border = {
                top: {style:'thin'},
                left: {style:'thin'},
                bottom: {style:'thin'},
                right: {style:'thin'}
            };
            cell.font =  {
                name: 'Arial Black',
                family: 2,
                size: 12,
                bold: false
            }; 
        }
    }
    worksheet.pageSetup.printArea = `A1:G${2 + weeks.length}`;
}

const timeStamp = new Date().toISOString();
const fileName = `./printed/calendar-${timeStamp}.xlsx`;

workbook.xlsx.writeFile(fileName);

console.log(`Saved to ${fileName}`);
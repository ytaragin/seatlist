const xlsxFile = require('read-excel-file/node');
//const createCsvWriter = require('csv-writer').createObjectCsvWriter;
const Excel = require('exceljs');
const { fromPairs } = require('lodash');


const rowNames = new Set(['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט', 'י', 'יא', 'יב', 'יג' ]);


// const WORKDIR = '/d/WebDrives/Dropbox/Personal/shul/Seating/5782 Seating';
const WORKDIR = '/mnt/c/Users/taragin/Temp/YN';

function addSeat(seatmap, name, seatlabel) {
    if (!seatlabel) {
        return;
    }

    const { rowname, seat } = seatlabel;

    seatmap.getOrInsertComputed(name, () => new Map())
           .getOrInsertComputed(rowname, () => [])
           .push(seat);
}

function isSeatNumber(val) {
    if (Number.isInteger(val)) {
        return true;
    }
    let c = val.charAt(0)
    return c >= '0' && c <= '9';
}

function isRowName(val) {
    return (val && rowNames.has(val));
}

function isSpecialField(val) {
    if (!val) {
        return false;
    }

    const v = val.toUpperCase();

    return (v === 'בימה')
        || (v === 'ארון קודש')
        || v.startsWith('ראש השנה')
        || v.startsWith('קהילת אהבת')
        || v.startsWith('מקומות')
        || v.startsWith('יום כיפור')
        || v.startsWith('מעבר')
        || v.startsWith('ROSH')
        || v.startsWith('YOM');
}


function isName(val) {
    return val 
        && !isSeatNumber(val)
        && !isSpecialField(val) 
        && !isRowName(val)
}

async function getSheets() {
    let sheets = await xlsxFile('./Mens\ 5780.xlsx', { getSheets: true });
       
    sheets.forEach((obj)=>{
         console.log(obj.name);
     })
}


function getSeatLabel(rows, rownum, colnum){
    let labelrow = rows[rownum-1]
    if (!labelrow) {
        console.log(`No label row above row:${rownum} col: ${colnum}`);
        return null;
    }
    let seat =  labelrow[colnum];
    let currspot = colnum-1;
    let rowname = null;
    while((currspot>=0) && !rowname) {
        if (isRowName(labelrow[currspot])) {
            rowname = labelrow[currspot];
        }
        currspot--;
    }

    if (!rowname) {
        console.log(`Error with row:${rownum} col: ${colnum}`)
    }

    return {rowname, seat};
}



async function getRows(file, sheet, seatmap = new Map()) {
    // let rows = await xlsxFile('/d/Mens\ 5782.xlsx', { sheet: 'MenRH' });
    let rows = await xlsxFile(file, { sheet});
       
    rows.forEach((row, rownum)=>{
//         console.log(row);
        row.forEach((cell, cellnum) => {
            if (isName(cell)) {
                addSeat(seatmap, cell, getSeatLabel(rows, rownum, cellnum));
            }
        })
     });

     console.log(seatmap)
     return seatmap;
}

async function seatsToCsv(seatmap) {
    let names = [...seatmap.keys()].sort();

    //console.log(sorted);

    let data = [];

    names.forEach(n=> {
        let seats = seatmap.get(n);
        let rows = [...seats.keys()].sort();
        rows.forEach(r => {
            let items = seats.get(r).sort()
            let range = `${items[0]}`
            if (items.length > 1) {
                range += `-${items[items.length-1]}`
            }
            data.push({
                name: n,
                row: r,
                seats: range
            });
        })

    })



    const csvWriter = createCsvWriter({
        path: '/d/out.csv',
        header: [
          {id: 'name', title: 'שם'},
          {id: 'row', title: 'שורה'},
          {id: 'seats', title: 'כיסא'}
        ]
      });
      
              
      csvWriter
          .writeRecords(data)
          .then(()=> console.log('The CSV file was written successfully'));


}

async function seatsToExcel(seatmap) {
    let names = [...seatmap.keys()].sort();

    //console.log(sorted);

    let data = [];

    names.forEach(n=> {
        let seats = seatmap.get(n);
        let rows = [...seats.keys()].sort();
        rows.forEach(r => {
            let items = seats.get(r).sort()
            let range = `${items[0]}`
            if (items.length > 1) {
                range += `-${items[items.length-1]}`
            }
            data.push({
                name: n,
                row: r,
                seats: range
            });
        })

    })

    let workbook = new Excel.Workbook();
    let worksheet = workbook.addWorksheet('Seats');



    const MAXROWS = 23;
    let colCount = Math.ceil(data.length/MAXROWS);

    let headerRow
    let columns = [];

    for (let i=0; i<colCount; i++ ){
        columns.push({header: 'שם', key: `name${i}`});
        columns.push({header: 'שורה', key: `row${i}`});
        columns.push({header: 'כיסא', key: `seats${i}`});
        columns.push({header: '', key: `blank${i}`});
    }

    worksheet.columns = columns;

    worksheet.views = [
        {rightToLeft: true}
    ];    

    let maxDone = 0;

    for (let i=0; (i<MAXROWS); i++) {// && maxDone<data.length); i++ ){
        let e = {};
        for (let j=0; j<colCount; j++ ){
            let curSpot = j*MAXROWS + i;
            let d = {};
            if (curSpot < data.length) {
                d = data[curSpot];
            } else {
                d = {
                    name: "",
                    row: "",
                    seats: "",
                    blank: ""
                };
            }
            e[`name${j}`] = d.name;
            e[`row${j}`] = d.row;
            e[`seats${j}`] = d.seats;
            e[`blank${j}`] = "";            

            maxDone = Math.max(maxDone, curSpot);
        }
        worksheet.addRow(e)
    }
    
    // data.forEach((e) => {
    //     worksheet.addRow(e)
    // });

    workbook.xlsx.writeFile(`${WORKDIR}/seats.xlsx`)

    // const csvWriter = createCsvWriter({
    //     path: '/d/out.csv',
    //     header: [
    //       {id: 'name', title: 'שם'},
    //       {id: 'row', title: 'שורה'},
    //       {id: 'seats', title: 'כיסא'}
    //     ]
    //   });
      
              
    //   csvWriter
    //       .writeRecords(data)
    //       .then(()=> console.log('The CSV file was written successfully'));


}





async function getSheetRows(file, sheets) {
    const seatmap = new Map();
    for (const sheet of sheets) {
        await getRows(file, sheet, seatmap);
    }
    return seatmap;
}

const MEN_RH_SHEETS = ['MenRH'];
const MEN_YK_SHEETS = ['MenYK'];

const WOMEN_RH_SHEETS = [
    'Downstairs',
    'Upstairs',
    'Annexe',
];

const WOMEN_YK_SHEETS = [
    'Downstairs',
    'Upstairs',
    'Annexe',
];

async function genList(file, sheets) {
    const seatmap = await getSheetRows(file, sheets);
    seatsToExcel(seatmap);
}


genList(`${WORKDIR}/Seating 5787.xlsx`, WOMEN_RH_SHEETS);
//genList(`${WORKDIR}/Women KAT seats 5782 YK.xlsx`, WOMEN_YK_SHEETS);
// genList(`${WORKDIR}/Mens YK 5782.xlsx`, MEN_YK_SHEETS);


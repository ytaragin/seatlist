const xlsxFile = require('read-excel-file/node');
//const createCsvWriter = require('csv-writer').createObjectCsvWriter;
const Excel = require('exceljs');
const { fromPairs } = require('lodash');


const rowNames = new Set(['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט', 'י', 'יא', 'יב', 'יג' ]);


const seatmap = {};
const WORKDIR = '/d/WebDrives/Dropbox/Personal/shul/Seating/5782 Seating';

function addSeat(name, seatlabel) {
    if (!seatlabel) {
        return;
    }

    if (!seatmap.hasOwnProperty(name)) {
        seatmap[name] = {};
    }
    if (!seatmap[name][seatlabel.rowname]) {
        seatmap[name][seatlabel.rowname] = [];
    }
    seatmap[name][seatlabel.rowname].push(seatlabel.seat)
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
    return val && 
           ((val === 'בימה') 
            || (val === 'ארון קודש') 
            || val.startsWith('ראש השנה') 
            || val.startsWith('קהילת אהבת') 
            || val.startsWith('מקומות')
            || val.startsWith('יום כיפור')
            || val.startsWith('מעבר')
            || val.startsWith('ROSH')
            || val.startsWith('YOM')
           );
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



async function getRows(file, sheet) {
    // let rows = await xlsxFile('/d/Mens\ 5782.xlsx', { sheet: 'MenRH' });
    let rows = await xlsxFile(file, { sheet});
       
    rows.forEach((row, rownum)=>{
//         console.log(row);
        row.forEach((cell, cellnum) => {
            if (isName(cell)) {
                addSeat(cell, getSeatLabel(rows, rownum, cellnum));
            }
        })
     });

     console.log(seatmap)
    // return seatmap;     
}

async function seatsToCsv(seatmap) {
    let names = Object.keys(seatmap).sort();

    //console.log(sorted);

    let data = [];

    names.forEach(n=> {
        let seats = seatmap[n];
        let rows = Object.keys(seats).sort();
        rows.forEach(r => {
            let items = seats[r].sort()
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
    let names = Object.keys(seatmap).sort();

    //console.log(sorted);

    let data = [];

    names.forEach(n=> {
        let seats = seatmap[n];
        let rows = Object.keys(seats).sort();
        rows.forEach(r => {
            let items = seats[r].sort()
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





async function genList() {
    await getRows(`${WORKDIR}/Mens\ YK\ 5782.xlsx`, 'MenYK');
    await getRows(`${WORKDIR}/Mens\ YK\ 5782.xlsx`, 'MenYK_Downstairs');
    seatsToExcel(seatmap);
}

async function genListWRH() {
    await getRows('${WORKDIR}/Women KAT seats 5782 YK.xlsx', 'Women Downstairs RH');
    await getRows('/d/Womens\ 5782.xlsx', 'Women upstairs-BM-RH');
    await getRows('/d/Womens\ 5782.xlsx', 'Annex RH');
    await getRows('/d/Womens\ 5782.xlsx', 'Women Hall RH');

    
    seatsToExcel(seatmap);
}

async function genListWYK() {
    await getRows(`${WORKDIR}/Women KAT seats 5782 YK.xlsx`, 'Women Downstairs YK');
    await getRows(`${WORKDIR}/Women KAT seats 5782 YK.xlsx`, 'Women upstairs-BM-YK');
    await getRows(`${WORKDIR}/Women KAT seats 5782 YK.xlsx`, 'Annex YK');
    await getRows(`${WORKDIR}/Women KAT seats 5782 YK.xlsx`, 'Women Hall YK');

    
    seatsToExcel(seatmap);
}


//genListWYK();
genList();


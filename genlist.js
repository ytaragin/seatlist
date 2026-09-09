const xlsxFile = require('read-excel-file/node');
//const createCsvWriter = require('csv-writer').createObjectCsvWriter;
const Excel = require('exceljs');
const fs = require('fs');
const { fromPairs } = require('lodash');


// const WORKDIR = '/d/WebDrives/Dropbox/Personal/shul/Seating/5782 Seating';
const CONFIG = {
    workdir: '/mnt/c/Users/taragin/Temp/YN',
    csvFile: '/d/out.csv',
    worksheetName: 'Seats',

    // Column packing: aim for columnGoal columns, but never exceed maxRows per column.
    maxRows: 24,
    columnGoal: 4,

    headers: {
        name: 'שם',
        row: 'שורה',
        seats: 'כיסא',
    },

    rowNames: new Set(['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט', 'י', 'יא', 'יב', 'יג']),

    specialFields: ['בימה', 'ארון קודש'],
    specialFieldPrefixes: [
        'ראש השנה',
        'קהילת אהבת',
        'מקומות',
        'יום כיפור',
        'מעבר',
        'ROSH',
        'YOM',
    ],

    jobs: {
        menRH: {
            input: 'Mens 5787.xlsx',
            sheets: ['MenRH'],
            output: 'seats men RH.xlsx',
        },
        womenRH: {
            input: 'Women Seating RH 5787.xlsx',
            sheets: ['Downstairs', 'Upstairs', 'Annexe'],
            output: 'seats women RH.xlsx',
        },
        womenYK: {
            input: 'Women KAT seats 5782 YK.xlsx',
            sheets: ['Downstairs', 'Upstairs', 'Annexe'],
            output: 'seats women YK.xlsx',
        },
    },
};

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
    return (val && CONFIG.rowNames.has(val));
}

function isSpecialField(val) {
    if (!val) {
        return false;
    }

    const v = val.toUpperCase();

    return CONFIG.specialFields.includes(v)
        || CONFIG.specialFieldPrefixes.some(prefix => v.startsWith(prefix));
}


function isName(val) {
    return val 
        && !isSeatNumber(val)
        && !isSpecialField(val) 
        && !isRowName(val)
}

async function getSheets(file) {
    let sheets = await xlsxFile(file, { getSheets: true });
       
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
        path: CONFIG.csvFile,
        header: [
          {id: 'name', title: CONFIG.headers.name},
          {id: 'row', title: CONFIG.headers.row},
          {id: 'seats', title: CONFIG.headers.seats}
        ]
      });
      
              
      csvWriter
          .writeRecords(data)
          .then(()=> console.log('The CSV file was written successfully'));


}

// Packs the data into columnGoal columns if it fits within maxRows, otherwise adds columns.
function calcLayout(itemCount, maxRows = CONFIG.maxRows, columnGoal = CONFIG.columnGoal) {
    if (itemCount <= 0) {
        return { rowCount: 0, colCount: 0 };
    }

    let rowCount = Math.min(maxRows, Math.ceil(itemCount / columnGoal));
    let colCount = Math.ceil(itemCount / rowCount);

    return { rowCount, colCount };
}

async function seatsToExcel(seatmap, outputPath) {
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
    let worksheet = workbook.addWorksheet(CONFIG.worksheetName);



    const { rowCount, colCount } = calcLayout(data.length);

    let headerRow
    let columns = [];

    for (let i=0; i<colCount; i++ ){
        columns.push({header: CONFIG.headers.name, key: `name${i}`});
        columns.push({header: CONFIG.headers.row, key: `row${i}`});
        columns.push({header: CONFIG.headers.seats, key: `seats${i}`});
        columns.push({header: '', key: `blank${i}`});
    }

    worksheet.columns = columns;

    worksheet.views = [
        {rightToLeft: true}
    ];    

    let maxDone = 0;

    for (let i=0; (i<rowCount); i++) {// && maxDone<data.length); i++ ){
        let e = {};
        for (let j=0; j<colCount; j++ ){
            let curSpot = j*rowCount + i;
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

    await workbook.xlsx.writeFile(outputPath)
    console.log(`Wrote ${outputPath}`)

    return { entryCount: data.length, rowCount, colCount };

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

async function genList(inputPath, sheets, outputPath) {
    const seatmap = await getSheetRows(inputPath, sheets);
    return seatsToExcel(seatmap, outputPath);
}

function printSummary(results) {
    console.log('\nSummary:');
    for (const r of results) {
        console.log(`  ${r.name}: ${r.status}`);
        console.log(`      in:  ${r.input}`);
        console.log(`      out: ${r.output ?? '(none)'}`);
        if (r.stats) {
            console.log(`      entries: ${r.stats.entryCount}, columns: ${r.stats.colCount}, rows per column: ${r.stats.rowCount}`);
        }
    }
}

async function genAll() {
    const results = [];

    for (const [name, job] of Object.entries(CONFIG.jobs)) {
        const inputPath = `${CONFIG.workdir}/${job.input}`;
        const outputPath = `${CONFIG.workdir}/${job.output}`;

        // A missing input only skips its own job; the remaining jobs still run.
        if (!fs.existsSync(inputPath)) {
            console.log(`Skipping ${name}: input file not found`);
            results.push({ name, input: inputPath, output: null, status: 'SKIPPED (input not found)' });
            continue;
        }

        try {
            const stats = await genList(inputPath, job.sheets, outputPath);
            results.push({ name, input: inputPath, output: outputPath, status: 'OK', stats });
        } catch (err) {
            console.log(`Failed ${name}: ${err.message}`);
            results.push({ name, input: inputPath, output: null, status: `FAILED (${err.message})` });
        }
    }

    printSummary(results);
}

genAll();


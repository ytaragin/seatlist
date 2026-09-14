const xlsxFile = require('read-excel-file/node');
//const createCsvWriter = require('csv-writer').createObjectCsvWriter;
const Excel = require('exceljs');
const fs = require('fs');
const { createSeatLayout, writeToSource } = require('./sourceWorkbook');
const { fromPairs } = require('lodash');


// const WORKDIR = '/d/WebDrives/Dropbox/Personal/shul/Seating/5782 Seating';
const CONFIG = {
    workdir: '/mnt/c/Users/taragin/Temp/YN',
    csvFile: '/d/out.csv',
    worksheetName: 'Seats',

    // Column packing: use outputRange height as the row limit, or maxRows without a range.
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
            outputSheet: 'RH_Names',
            outputRange: 'B7:P32',

        },
        menYK: {
            input: 'Mens 5787.xlsx',
            sheets: ['MenYK'],
            output: 'seats men YK.xlsx',
            outputSheet: 'YK_Names',
            outputRange: 'B7:P32',
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

function seatmapToEntries(seatmap) {
    const entries = [];

    for (const name of [...seatmap.keys()].sort()) {
        const seatsByRow = seatmap.get(name);
        for (const row of [...seatsByRow.keys()].sort()) {
            const seats = seatsByRow.get(row).sort((first, second) => first - second);
            const range = seats.length > 1
                ? `${seats[0]}-${seats[seats.length - 1]}`
                : `${seats[0]}`;
            entries.push({ name, row, seats: range });
        }
    }

    return entries;
}

function createSeatColumns(colCount) {
    const columns = [];

    for (let group = 0; group < colCount; group++) {
        columns.push({header: CONFIG.headers.name, key: `name${group}`});
        columns.push({header: CONFIG.headers.row, key: `row${group}`});
        columns.push({header: CONFIG.headers.seats, key: `seats${group}`});
        columns.push({header: '', key: `blank${group}`});
    }

    return columns;
}

async function seatsToExcel(seatmap, outputPath, destination) {
    const entries = seatmapToEntries(seatmap);
    const layout = destination
        ? await writeToSource(entries, destination.inputPath, destination.sheetName, destination.range, CONFIG)
        : createSeatLayout(entries, CONFIG.maxRows, CONFIG.columnGoal);

    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet(CONFIG.worksheetName);
    const { rowCount, colCount } = layout;

    worksheet.columns = createSeatColumns(colCount);
    worksheet.views = [{rightToLeft: true}];
    for (const values of layout.rows) {
        worksheet.addRow(values.map(value => value ?? ''));
    }

    await workbook.xlsx.writeFile(outputPath);
    console.log(`Wrote ${outputPath}`);

    return { entryCount: entries.length, rowCount, colCount };
}





async function getSheetRows(file, sheets) {
    const seatmap = new Map();
    for (const sheet of sheets) {
        await getRows(file, sheet, seatmap);
    }
    return seatmap;
}

async function genList(inputPath, sheets, outputPath, job = {}) {
    const hasDestination = job.outputSheet !== undefined || job.outputRange !== undefined;
    if (hasDestination && (!job.outputSheet || !job.outputRange)) {
        throw new Error('Both outputSheet and outputRange are required for source workbook output');
    }
    if (hasDestination && sheets.includes(job.outputSheet)) {
        throw new Error('The output sheet must not be one of the input sheets');
    }
    const seatmap = await getSheetRows(inputPath, sheets);
    return seatsToExcel(seatmap, outputPath, hasDestination ? {
        inputPath,
        sheetName: job.outputSheet,
        range: job.outputRange,
    } : undefined);
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
            const stats = await genList(inputPath, job.sheets, outputPath, job);
            results.push({ name, input: inputPath, output: outputPath, status: 'OK', stats });
        } catch (err) {
            console.log(`Failed ${name}: ${err.message}`);
            results.push({ name, input: inputPath, output: null, status: `FAILED (${err.message})` });
        }
    }

    printSummary(results);
}

if (require.main === module) {
    genAll();
}

module.exports = { genList };


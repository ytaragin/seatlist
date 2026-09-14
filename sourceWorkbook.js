const Excel = require('exceljs');
const fs = require('fs');
const path = require('node:path');
const { randomUUID } = require('node:crypto');
const JSZip = require('jszip');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

function calcLayout(itemCount, maxRows, columnGoal) {
    if (itemCount <= 0) {
        return { rowCount: 0, colCount: 0 };
    }

    let rowCount = Math.min(maxRows, Math.ceil(itemCount / columnGoal));
    let colCount = Math.ceil(itemCount / rowCount);

    return { rowCount, colCount };
}

function createSeatLayout(entries, maxRows, columnGoal) {
    const { rowCount, colCount } = calcLayout(entries.length, maxRows, columnGoal);
    const rows = Array.from({ length: rowCount }, () => Array(colCount * 4).fill(null));

    entries.forEach((entry, index) => {
        const row = index % rowCount;
        const column = Math.floor(index / rowCount) * 4;
        rows[row][column] = entry.name;
        rows[row][column + 1] = entry.row;
        rows[row][column + 2] = entry.seats;
    });

    return { rowCount, colCount, rows };
}

async function saveSourceRange(original, inputPath, worksheet, start, end) {
    const archive = await JSZip.loadAsync(original);
    const namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const relationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const parser = new DOMParser({
        onError(level, message) {
            throw new Error(`Invalid workbook XML: ${message}`);
        },
    });
    const readXml = async name => {
        const entry = archive.file(name);
        if (!entry) {
            throw new Error(`Missing workbook part: ${name}`);
        }
        return parser.parseFromString(await entry.async('string'), 'application/xml');
    };
    const workbookXml = await readXml('xl/workbook.xml');
    const sheet = Array.from(workbookXml.getElementsByTagNameNS(namespace, 'sheet'))
        .find(element => element.getAttribute('name') === worksheet.name);
    const relationships = await readXml('xl/_rels/workbook.xml.rels');
    const relationship = Array.from(relationships.getElementsByTagNameNS('*', 'Relationship'))
        .find(element => element.getAttribute('Id') === sheet.getAttributeNS(relationshipNamespace, 'id'));
    if (!relationship || relationship.getAttribute('TargetMode') === 'External') {
        throw new Error(`Invalid worksheet relationship: ${worksheet.name}`);
    }
    const target = relationship.getAttribute('Target');
    const sheetPath = target.startsWith('/') ? target.slice(1) : path.posix.join('xl', target);
    const document = await readXml(sheetPath);
    const sheetData = document.getElementsByTagNameNS(namespace, 'sheetData')[0];
    const children = (parent, name) => Array.from(parent.childNodes)
        .filter(node => node.namespaceURI === namespace && node.localName === name);
    const createElement = name => document.createElementNS(namespace,
        document.documentElement.prefix ? `${document.documentElement.prefix}:${name}` : name);

    for (let rowNumber = start.row; rowNumber <= end.row; rowNumber++) {
        const rows = children(sheetData, 'row');
        let row = rows.find(element => Number(element.getAttribute('r')) === rowNumber);
        if (!row) {
            row = createElement('row');
            row.setAttribute('r', String(rowNumber));
            sheetData.insertBefore(row, rows.find(element => Number(element.getAttribute('r')) > rowNumber) || null);
        }
        for (let column = start.col; column <= end.col; column++) {
            const valueCell = worksheet.getCell(rowNumber, column);
            const cells = children(row, 'c');
            let cell = cells.find(element => element.getAttribute('r') === valueCell.address);
            if (!cell && valueCell.value === null) {
                continue;
            }
            if (!cell) {
                cell = createElement('c');
                cell.setAttribute('r', valueCell.address);
                row.insertBefore(cell, cells.find(element =>
                    worksheet.getCell(element.getAttribute('r')).col > column) || null);
            }
            const formula = children(cell, 'f')[0];
            if (formula && (formula.hasAttribute('t') || archive.file('xl/calcChain.xml'))) {
                throw new Error(`Cannot safely replace linked formula in ${valueCell.address}`);
            }
            for (const name of ['f', 'v', 'is']) {
                for (const child of children(cell, name)) {
                    cell.removeChild(child);
                }
            }
            cell.removeAttribute('t');
            if (valueCell.value !== null) {
                cell.setAttribute('t', 'inlineStr');
                const inlineString = createElement('is');
                const text = createElement('t');
                text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
                text.appendChild(document.createTextNode(String(valueCell.value)));
                inlineString.appendChild(text);
                cell.insertBefore(inlineString, cell.firstChild);
            }
        }
    }

    archive.file(sheetPath, new XMLSerializer().serializeToString(document), { createFolders: false });
    const updated = await archive.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    const temporaryPath = `${inputPath}.${randomUUID()}.tmp`;
    try {
        await fs.promises.writeFile(temporaryPath, updated, { flag: 'wx', mode: (await fs.promises.stat(inputPath)).mode });
        if (!(await fs.promises.readFile(inputPath)).equals(original)) {
            throw new Error('Source workbook changed during generation; refusing to overwrite it');
        }
        await fs.promises.copyFile(inputPath, `${inputPath}.bak`, fs.constants.COPYFILE_EXCL)
            .catch(error => {
                if (error.code !== 'EEXIST') {
                    throw error;
                }
            });
        await fs.promises.rename(temporaryPath, inputPath);
    } finally {
        await fs.promises.rm(temporaryPath, { force: true });
    }
}

async function writeToSource(data, inputPath, sheetName, range, { columnGoal }) {
    const match = typeof range === 'string'
        && range.trim().match(/^([A-Z]{1,3}[1-9]\d{0,6})\s*[:-]\s*([A-Z]{1,3}[1-9]\d{0,6})$/i);
    if (!match) {
        throw new Error('outputRange must be a cell range such as B6:P31');
    }

    const workbook = new Excel.Workbook();
    const original = await fs.promises.readFile(inputPath);
    await workbook.xlsx.load(original);
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
        throw new Error(`Output sheet not found: ${sheetName}`);
    }

    const start = worksheet.getCell(match[1].toUpperCase());
    const end = worksheet.getCell(match[2].toUpperCase());
    const height = end.row - start.row + 1;
    const width = end.col - start.col + 1;
    if (height < 1 || width < 3 || end.row > 1048576 || end.col > 16384) {
        throw new Error(`Invalid output range: ${range}`);
    }

    const availableColumns = Math.floor((width + 1) / 4);
    const layout = createSeatLayout(
        data,
        height,
        Math.min(columnGoal, availableColumns)
    );
    if (layout.colCount > availableColumns) {
        throw new Error(`Output has ${data.length} entries and does not fit in ${sheetName}!${range}`);
    }

    for (let row = start.row; row <= end.row; row++) {
        for (let column = start.col; column <= end.col; column++) {
            const cell = worksheet.getCell(row, column);
            if (cell.isMerged) {
                throw new Error(`Output range contains merged cell ${cell.address}`);
            }
            cell.value = null;
        }
    }

    layout.rows.forEach((values, rowOffset) => {
        values.forEach((value, columnOffset) => {
            if (value !== null) {
                worksheet.getCell(start.row + rowOffset, start.col + columnOffset).value = value;
            }
        });
    });

    await saveSourceRange(original, inputPath, worksheet, start, end);
    console.log(`Updated ${inputPath}: ${sheetName}!${range}`);
    return layout;
}

module.exports = { createSeatLayout, writeToSource };
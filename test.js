const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const Excel = require('exceljs');
const JSZip = require('jszip');
const { genList } = require('./genlist');

async function fixture(context, count = 25) {
	const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'seatlist-'));
	context.after(() => fs.rm(directory, { recursive: true, force: true }));
	const inputPath = path.join(directory, 'source.xlsx');
	const outputPath = path.join(directory, 'output.xlsx');
	const workbook = new Excel.Workbook();
	const source = workbook.addWorksheet('MenRH');
	for (let index = 0; index < count; index++) {
		source.addRow(['\u05d0', 1]);
		source.addRow([null, `Name ${String(index).padStart(3, '0')}`]);
	}
	const target = workbook.addWorksheet('RH_Names');
	target.views = [{ rightToLeft: true }];
	target.getColumn(2).width = 28;
	target.getRow(6).height = 30;
	target.getCell('B5').value = 'Existing header';
	target.getCell('A6').value = 'Outside left';
	target.getCell('Q6').value = { formula: '1+1', result: 2 };
	target.getCell('B32').value = 'Outside below';
	for (let row = 6; row <= 31; row++) {
		for (let column = 2; column <= 16; column++) {
			const cell = target.getCell(row, column);
			cell.value = 'Stale';
			cell.font = { bold: true, color: { argb: 'FF123456' } };
			cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
			cell.border = { bottom: { style: 'thin' } };
			cell.alignment = { horizontal: 'right' };
			cell.numFmt = '@';
		}
	}
	await workbook.xlsx.writeFile(inputPath);
	return { inputPath, outputPath, workbook, target };
}

const destination = { outputSheet: 'RH_Names', outputRange: 'B6:P31' };

test('sorts seats numerically across the single-digit boundary in both outputs', async context => {
	const { inputPath, outputPath, workbook } = await fixture(context, 0);
	const source = workbook.addWorksheet('MenYK');
	const name = '\u05d9\u05e8\u05e1';
	source.addRow(['\u05d0', 8, 9, 10, 11]);
	source.addRow([null, name, name, name, name]);
	await workbook.xlsx.writeFile(inputPath);
	await genList(inputPath, ['MenYK'], outputPath, destination);
	const updated = await new Excel.Workbook().xlsx.readFile(inputPath);
	const standalone = await new Excel.Workbook().xlsx.readFile(outputPath);
	assert.equal(updated.getWorksheet('RH_Names').getCell('D6').value, '8-11');
	assert.equal(standalone.getWorksheet('Seats').getCell('C2').value, '8-11');
});

test('preserves every workbook part except the destination worksheet', async context => {
	const { inputPath, outputPath } = await fixture(context, 1);
	const original = await JSZip.loadAsync(await fs.readFile(inputPath));
	original.file('xl/printerSettings/printerSettings1.bin', Buffer.from([0, 1, 2, 255]));
	for (const [name, entry] of Object.entries(original.files)) {
		if (entry.dir) {
			delete original.files[name];
		}
	}
	const originalBytes = await original.generateAsync({ type: 'nodebuffer' });
	await fs.writeFile(inputPath, originalBytes);
	await genList(inputPath, ['MenRH'], outputPath, destination);
	const updated = await JSZip.loadAsync(await fs.readFile(inputPath));
	assert.deepEqual(Object.keys(updated.files).sort(), Object.keys(original.files).sort());
	for (const [name, entry] of Object.entries(original.files)) {
		if (!entry.dir && name !== 'xl/worksheets/sheet2.xml') {
			assert.deepEqual(await updated.file(name).async('nodebuffer'), await entry.async('nodebuffer'), name);
		}
	}
	assert.deepEqual(await fs.readFile(`${inputPath}.bak`), originalBytes);
	await genList(inputPath, ['MenRH'], outputPath, destination);
	assert.deepEqual(await fs.readFile(`${inputPath}.bak`), originalBytes);
});

test('a failed replacement leaves the source intact and removes the temporary file', async context => {
	const { inputPath, outputPath } = await fixture(context, 1);
	const original = await fs.readFile(inputPath);
	context.mock.method(fs, 'rename', async () => {
		throw Object.assign(new Error('Workbook is locked'), { code: 'EACCES' });
	});
	await assert.rejects(genList(inputPath, ['MenRH'], outputPath, destination), /Workbook is locked/);
	assert.deepEqual(await fs.readFile(inputPath), original);
	assert.deepEqual(await fs.readFile(`${inputPath}.bak`), original);
	assert.deepEqual((await fs.readdir(path.dirname(inputPath))).sort(), ['source.xlsx', 'source.xlsx.bak']);
});

test('writes sorted groups, clears unused cells, and preserves template formatting', async context => {
	const { inputPath, outputPath } = await fixture(context);
	const before = await new Excel.Workbook().xlsx.readFile(inputPath);
	const stats = await genList(inputPath, ['MenRH'], outputPath, destination);
	assert.deepEqual(stats, { entryCount: 25, rowCount: 7, colCount: 4 });
	const after = await new Excel.Workbook().xlsx.readFile(inputPath);
	const target = after.getWorksheet('RH_Names');
	for (let row = 6; row <= 31; row++) {
		for (let column = 2; column <= 16; column++) {
			const offset = column - 2;
			const index = Math.floor(offset / 4) * 7 + row - 6;
			const values = [`Name ${String(index).padStart(3, '0')}`, '\u05d0', '1'];
			const expected = row < 13 && index < 25 && offset % 4 < 3 ? values[offset % 4] : null;
			assert.equal(target.getCell(row, column).value, expected);
			assert.deepEqual(target.getCell(row, column).style,
				before.getWorksheet('RH_Names').getCell(row, column).style);
		}
	}
	for (const address of ['B5', 'A6', 'Q6', 'B32']) {
		assert.deepEqual(target.getCell(address).value, before.getWorksheet('RH_Names').getCell(address).value);
	}
	assert.equal(target.getColumn(2).width, 28);
	assert.equal(target.getRow(6).height, 30);
	assert.equal(target.views[0].rightToLeft, true);
	assert.deepEqual(after.getWorksheet('MenRH').getSheetValues(), before.getWorksheet('MenRH').getSheetValues());
	const standalone = await new Excel.Workbook().xlsx.readFile(outputPath);
	const seatsSheet = standalone.getWorksheet('Seats');
	assert.equal(seatsSheet.views[0].rightToLeft, true);
	assert.equal(seatsSheet.rowCount, stats.rowCount + 1);
	assert.equal(seatsSheet.columnCount, stats.colCount * 4);
	const headers = ['\u05e9\u05dd', '\u05e9\u05d5\u05e8\u05d4', '\u05db\u05d9\u05e1\u05d0', ''];
	for (let group = 0; group < stats.colCount; group++) {
		for (let field = 0; field < 4; field++) {
			const column = group * 4 + field + 1;
			assert.equal(seatsSheet.getCell(1, column).value, headers[field]);
			for (let rowIndex = 0; rowIndex < stats.rowCount; rowIndex++) {
				const index = group * stats.rowCount + rowIndex;
				const values = index < stats.entryCount
					? [`Name ${String(index).padStart(3, '0')}`, '\u05d0', '1', '']
					: ['', '', '', ''];
				assert.equal(seatsSheet.getCell(rowIndex + 2, column).value, values[field]);
			}
		}
	}
});

test('clears the entire destination when there are no entries', async context => {
	const { inputPath, outputPath } = await fixture(context, 0);
	await genList(inputPath, ['MenRH'], outputPath, destination);
	const workbook = await new Excel.Workbook().xlsx.readFile(inputPath);
	for (let row = 6; row <= 31; row++) {
		for (let column = 2; column <= 16; column++) {
			assert.equal(workbook.getWorksheet('RH_Names').getCell(row, column).value, null);
		}
	}
});

test('packs within smaller ranges and accepts spaced range notation', async context => {
	const { inputPath, outputPath } = await fixture(context, 6);
	const stats = await genList(inputPath, ['MenRH'], outputPath, { ...destination, outputRange: 'B6 - H8' });
	assert.deepEqual(stats, { entryCount: 6, rowCount: 3, colCount: 2 });
	const workbook = await new Excel.Workbook().xlsx.readFile(inputPath);
	const target = workbook.getWorksheet('RH_Names');
	assert.equal(target.getCell('B8').value, 'Name 002');
	assert.equal(target.getCell('F8').value, 'Name 005');
	assert.equal(target.getCell('E8').value, null);
	assert.equal(target.getCell('B9').value, 'Stale');
});

test('rejects overflow and invalid destinations without changing the source file', async context => {
	const { inputPath, outputPath, workbook, target } = await fixture(context, 105);
	const before = await fs.readFile(inputPath);
	for (const [job, error] of [
		[destination, /does not fit/],
		[{ outputSheet: 'RH_Names' }, /Both outputSheet and outputRange/],
		[{ ...destination, outputSheet: 'Missing' }, /Output sheet not found/],
		[{ ...destination, outputSheet: 'MenRH' }, /must not be one of the input sheets/],
		[{ ...destination, outputRange: 'invalid' }, /outputRange must be/],
		[{ ...destination, outputRange: 'P31:B6' }, /Invalid output range/],
	]) {
		await assert.rejects(genList(inputPath, ['MenRH'], outputPath, job), error);
		assert.deepEqual(await fs.readFile(inputPath), before);
	}
	target.mergeCells('B6:C6');
	await workbook.xlsx.writeFile(inputPath);
	const mergedBefore = await fs.readFile(inputPath);
	await assert.rejects(genList(inputPath, ['MenRH'], outputPath, {
		...destination, outputRange: 'B6:T31',
	}), /merged cell/);
	assert.deepEqual(await fs.readFile(inputPath), mergedBefore);
});

test('jobs without a destination leave the source workbook unchanged', async context => {
	const { inputPath, outputPath } = await fixture(context, 104);
	const before = await fs.readFile(inputPath);
	const stats = await genList(inputPath, ['MenRH'], outputPath);
	assert.deepEqual(stats, { entryCount: 104, rowCount: 24, colCount: 5 });
	assert.deepEqual(await fs.readFile(inputPath), before);
	await fs.access(outputPath);
});

for (const { range, count, rowCount, colCount } of [
	{ range: 'B6:P31', count: 104, rowCount: 26, colCount: 4 },
	{ range: 'B6:T7', count: 10, rowCount: 2, colCount: 5 },
	{ range: 'B6 - H8', count: 5, rowCount: 3, colCount: 2 },
	{ range: 'b6:d10', count: 5, rowCount: 5, colCount: 1 },
]) {
	test(`uses the same range-constrained layout for both outputs in ${range}`, async context => {
		const { inputPath, outputPath } = await fixture(context, count);
		const stats = await genList(inputPath, ['MenRH'], outputPath, { ...destination, outputRange: range });
		assert.deepEqual(stats, { entryCount: count, rowCount, colCount });
		const source = await new Excel.Workbook().xlsx.readFile(inputPath);
		const standalone = await new Excel.Workbook().xlsx.readFile(outputPath);
		const seatsSheet = standalone.getWorksheet('Seats');
		assert.equal(seatsSheet.rowCount, rowCount + 1);
		assert.equal(seatsSheet.columnCount, colCount * 4);
		for (let index = 0; index < count; index++) {
			const rowOffset = index % rowCount;
			const columnOffset = Math.floor(index / rowCount) * 4;
			const name = `Name ${String(index).padStart(3, '0')}`;
			assert.equal(source.getWorksheet('RH_Names').getCell(6 + rowOffset, 2 + columnOffset).value, name);
			assert.equal(seatsSheet.getCell(2 + rowOffset, 1 + columnOffset).value, name);
		}
		for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
			for (let columnOffset = 0; columnOffset < colCount * 4 - 1; columnOffset++) {
				const sourceValue = source.getWorksheet('RH_Names')
					.getCell(6 + rowOffset, 2 + columnOffset).value;
				assert.equal(seatsSheet.getCell(2 + rowOffset, 1 + columnOffset).value, sourceValue ?? '');
			}
			assert.equal(seatsSheet.getCell(2 + rowOffset, colCount * 4).value, '');
		}
	});
}

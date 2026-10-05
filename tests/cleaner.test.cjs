const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const ExcelJS = require('../docs/vendor/exceljs.min.js');

async function clean(buffer) {
  const elements = new Map();
  let blob;
  const document = { querySelector(id) {
    if (!elements.has(id)) elements.set(id, { hidden: true, classList: { add() {} },
      events: {}, addEventListener(name, fn) { this.events[name] = fn; } });
    return elements.get(id);
  } };
  const context = vm.createContext({ document, ExcelJS, Blob, structuredClone, console,
    URL: { createObjectURL(value) { blob = value; return 'blob:test'; }, revokeObjectURL() {} } });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../docs/app.js'), 'utf8'), context);
  context.file = { name: 'schedule.xlsx', size: buffer.length, arrayBuffer: async () => buffer };
  vm.runInContext('selectFile(file)', context);
  await document.querySelector('#process').events.click();
  assert.equal(document.querySelector('#error').hidden, true);
  assert.equal(document.querySelector('#result').hidden, false);
  assert.equal(document.querySelector('#download').download, 'schedule_クライアント用.xlsx');
  const output = new ExcelJS.Workbook();
  await output.xlsx.load(await blob.arrayBuffer());
  return output;
}

test('removes prefix spaces, preserves names and merged bottom borders after XLSX round trip', async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('番組表');
  sheet.getRow(3).values = ['放映時間', '担当', '会社名・番組名', '秒数', '終了日', '備考'];
  sheet.mergeCells('C4:C5');
  sheet.getCell('C4').value = 'グレイド枠　 \t渋谷　店舗CM　Aタイプ';
  sheet.getCell('C4').alignment = { horizontal: 'center', indent: 2, wrapText: true };
  sheet.getCell('C5').style = { border: { bottom: { style: 'thin', color: { argb: 'FF000000' } } } };
  sheet.getCell('C6').value = 'そのまま　残す';
  sheet.getCell('D4').value = 30;
  const result = (await clean(await workbook.xlsx.writeBuffer())).worksheets[0];
  assert.equal(result.getCell('B4').value, '渋谷　店舗CM　Aタイプ');
  assert.equal(result.getCell('B4').alignment.horizontal, 'left');
  assert.equal(result.getCell('B4').alignment.indent || 0, 0);
  assert.equal(result.getCell('B4').alignment.wrapText, true);
  assert.equal(result.getCell('B5').border.bottom.style, 'thin');
  assert.equal(result.getCell('B6').value, 'そのまま　残す');
  assert.equal(result.getCell('C4').value, 30);
  assert.ok(result.model.merges.includes('B4:B5'));
});

// Optional local real-file regression; source files stay outside the repository.
for (const file of process.argv.slice(2).filter(name => name.endsWith('.xlsx'))) {
  test(`real workbook: ${path.basename(file)}`, async () => {
    const data = fs.readFileSync(file);
    const before = new ExcelJS.Workbook();
    await before.xlsx.load(data);
    const after = await clean(data);
    before.eachSheet((sheet, id) => {
      const result = after.getWorksheet(id);
      for (let row = 4; row <= sheet.rowCount; row++) {
        for (const [source, target] of [[1, 1], [3, 2], [4, 3]]) {
          const original = sheet.getCell(row, source);
          const actual = result.getCell(row, target);
          let expected = original.value;
          if (source === 3 && typeof expected === 'string' && expected.includes('グレイド枠')) {
            expected = expected.replace(/グレイド枠[\s\u3000]*/gu, '').trimStart();
          }
          assert.deepEqual(actual.value, expected, `value ${original.address}`);
          for (const edge of ['top', 'bottom', 'left', 'right']) {
            if (original.border[edge]) assert.deepEqual(actual.border[edge], original.border[edge], `border ${original.address} ${edge}`);
          }
          assert.deepEqual(actual.font, original.font, `font ${original.address}`);
          assert.equal(actual.numFmt, original.numFmt, `format ${original.address}`);
        }
      }
    });
  });
}

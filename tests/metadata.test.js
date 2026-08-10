/* Header detection against awkward real-world workbooks. node tests/metadata.test.js */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const ExcelJS = require(path.join(ROOT, 'vendor/exceljs.min.js'));

const sandbox = { window: {}, console, URL, URLSearchParams, TextEncoder, TextDecoder, ExcelJS };
sandbox.global = sandbox;
vm.createContext(sandbox);
for (const f of ['js/core.js', 'js/metadata.js']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), sandbox, { filename: f });
}
const META = sandbox.window.MetadataMerge;

let pass = 0, fail = 0;
const check = (label, actual, expected) => {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a === e) { pass++; console.log('  ok   ' + label); }
  else { fail++; console.log('  FAIL ' + label + '\n         got      ' + a + '\n         expected ' + e); }
};

const HEADERS = ['Participant Name As Per ICF', 'Participant Email', 'Expressions',
  'For Non Neutral, Please Select', 'Birthdate', 'Environment', 'Glasses', 'Hat',
  'Mustache', 'Makeup', 'Freckles', 'Skintone', 'Country of Collection'];

async function sheetWith(build) {
  const wb = new ExcelJS.Workbook();
  build(wb.addWorksheet('MetaData_v1'));
  const buf = await wb.xlsx.writeBuffer();
  const back = new ExcelJS.Workbook();
  await back.xlsx.load(buf);
  return back;
}

(async () => {
  console.log('\nHeader detection');

  // plain sheet, headers on row 1
  let wb = await sheetWith(ws => {
    ws.addRow(HEADERS);
    ws.addRow(['ABRAHAM OTIENO', 'a@x.com', 'Neutral', null, '2012/6', 'Indoor', 'x']);
  });
  let found = META.findHeaderRow(wb.getWorksheet('MetaData_v1'));
  check('headers on row 1 found', found && found.rowNumber, 1);
  check('fields + attributes mapped', found && found.map.matched >= 10, true);

  // title and instruction rows above the headers
  wb = await sheetWith(ws => {
    ws.addRow(['Centaurus Metadata Collection']);
    ws.addRow(['Please complete every column. Do not reorder.']);
    ws.addRow([]);
    ws.addRow(HEADERS);
    ws.addRow(['ASNET KALAI', 'b@x.com', 'Non-Neutral', 'Angry', '07/1992', 'Outdoor']);
  });
  found = META.findHeaderRow(wb.getWorksheet('MetaData_v1'));
  check('headers on row 4 found', found && found.rowNumber, 4);

  const read = META.readMetadataWorkbook(wb, 'Asnet Kalai');
  check('rows read past the title block', read.ok && read.rows.length, 1);
  check('name carried to column F', read.rows[0].values[META.OUT.NAME], 'ASNET KALAI');
  check('folder carried to column D', read.rows[0].values[META.OUT.FOLDER], 'Asnet Kalai');

  // reordered columns still map by name
  wb = await sheetWith(ws => {
    ws.addRow(['Environment', 'Birthdate', 'Participant Name As Per ICF', 'Expressions', 'Skintone', 'Hat']);
    ws.addRow(['Indoor', '1990/05', 'MIA ROSS', 'Neutral', 'Deep Brown', 'x']);
  });
  const reordered = META.readMetadataWorkbook(wb, 'Mia Ross');
  check('reordered columns map by header name',
    [reordered.rows[0].values[META.OUT.NAME], reordered.rows[0].values[META.OUT.ENVIRONMENT],
     reordered.rows[0].values[META.OUT.HEAD_AND_HAIR]],
    ['MIA ROSS', 'Indoor', 'Hat']);

  // an unrelated workbook is rejected, and says what it saw
  wb = await sheetWith(ws => {
    ws.addRow(['Invoice number', 'Client', 'Amount due', 'Currency']);
    ws.addRow([1, 'Acme', 100, 'USD']);
  });
  const rejected = META.readMetadataWorkbook(wb, 'x');
  check('unrelated workbook rejected', rejected.ok, false);
  check('reason names the sheet', /MetaData_v1/.test(rejected.reason), true);
  check('reason lists the headers actually seen', /Invoice number/.test(rejected.reason), true);
  check('diagnostics attached', Array.isArray(rejected.sheets) && rejected.sheets.length > 0, true);

  // Head_Attributes is a reference sheet, never a data sheet
  wb = await sheetWith(ws => { ws.addRow(HEADERS); ws.addRow(['A', 'a@x', 'Neutral']); });
  const ha = wb.addWorksheet('Head_Attributes');
  ha.addRow(['Head and Hair', 'Facial Features ', 'Accessories and jewellery', 'Others']);
  check('reference sheet skipped', META.readMetadataWorkbook(wb, 'A').sheetName, 'MetaData_v1');

  console.log('\n' + (fail === 0 ? 'ALL ' + pass + ' CHECKS PASSED' : pass + ' passed, ' + fail + ' FAILED'));
  process.exit(fail === 0 ? 0 : 1);
})();

/* QA over the 30-column delivery layout. No deps: node tests/deliverymerge.test.js */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const sandbox = { window: {}, console, URL, URLSearchParams, TextEncoder, TextDecoder };
sandbox.global = sandbox;
vm.createContext(sandbox);
for (const f of ['js/core.js', 'js/metadata.js', 'js/linkfiller.js', 'js/deliverymerge.js']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), sandbox, { filename: f });
}
const { MetadataMerge: META, DeliveryMerge: DM, MergerCore: CORE } = sandbox.window;
const OUT = META.OUT;

let pass = 0, fail = 0;
const check = (label, actual, expected) => {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a === e) { pass++; console.log('  ok   ' + label); }
  else { fail++; console.log('  FAIL ' + label + '\n         got      ' + a + '\n         expected ' + e); }
};

const blank = () => new Array(META.OUTPUT_HEADERS.length).fill(null);
const mk = (patch) => { const v = blank(); Object.keys(patch).forEach(k => { v[k] = patch[k]; }); return { values: v, source: 'AgencyA.xlsx' }; };

console.log('\nQA on the delivery layout');

// whitespace
let rows = [mk({ [OUT.NAME]: '  ABRAHAM   OTIENO  ' })];
let issues = DM.runQa(rows);
check('whitespace collapsed in F', rows[0].values[OUT.NAME], 'ABRAHAM OTIENO');

// birthdate variants -> YYYY/MM in column M
[['1990/05', '1990/05'], ['1988-11', '1988/11'], ['07/1992', '1992/07'], ['1993/4', '1993/04']].forEach(([input, want]) => {
  const r = [mk({ [OUT.BIRTHDATE]: input })];
  DM.runQa(r);
  check('birthdate ' + input + ' -> ' + want, r[0].values[OUT.BIRTHDATE], want);
});
rows = [mk({ [OUT.BIRTHDATE]: 'not recorded' })];
issues = DM.runQa(rows);
check('unparseable birthdate preserved', rows[0].values[OUT.BIRTHDATE], 'not recorded');
check('unparseable birthdate flagged', issues.some(i => i.severity === 'review' && /Birthdate/.test(i.rule)), true);

// slash placeholder in Others (column U), not in the other attribute columns
rows = [mk({ [OUT.OTHERS]: ' / ' })];
DM.runQa(rows);
check('slash-only Others cleared', rows[0].values[OUT.OTHERS], null);
rows = [mk({ [OUT.FACIAL_FEATURES]: 'Face/neck tattoos' })];
DM.runQa(rows);
check('real value containing a slash survives', rows[0].values[OUT.FACIAL_FEATURES], 'Face/neck tattoos');

// vocabulary, per column
rows = [mk({ [OUT.HEAD_AND_HAIR]: 'Glases', [OUT.FACIAL_FEATURES]: 'Moustache',
             [OUT.ACCESSORIES]: 'Makeup; Earrings', [OUT.OTHERS]: 'Freckels' })];
issues = DM.runQa(rows);
check('R corrected', rows[0].values[OUT.HEAD_AND_HAIR], 'Glasses');
check('S corrected (British spelling)', rows[0].values[OUT.FACIAL_FEATURES], 'Mustache');
check('T separator normalised', rows[0].values[OUT.ACCESSORIES], 'Makeup, Earrings');
check('U corrected (transposition)', rows[0].values[OUT.OTHERS], 'Freckles');
check('client typo maps to canonical', (() => {
  const r = [mk({ [OUT.OTHERS]: 'Other tattons' })]; DM.runQa(r); return r[0].values[OUT.OTHERS];
})(), 'Other tattoos');

rows = [mk({ [OUT.HEAD_AND_HAIR]: 'Sombrero' })];
issues = DM.runQa(rows);
check('unknown value kept verbatim', rows[0].values[OUT.HEAD_AND_HAIR], 'Sombrero');
check('unknown value flagged', issues.some(i => i.severity === 'review' && /accepted list/.test(i.rule)), true);

// column V follows the attribute block
rows = [mk({})];
DM.runQa(rows);
check('no attributes -> V is N/A-none apply', rows[0].values[OUT.NONE], CORE.NONE_APPLY_TEXT);
rows = [mk({ [OUT.HEAD_AND_HAIR]: 'Hat', [OUT.NONE]: 'N/A-none apply' })];
DM.runQa(rows);
check('attribute present -> V cleared', rows[0].values[OUT.NONE], null);
// ordering: a cleared placeholder must count as empty for V
rows = [mk({ [OUT.OTHERS]: '/' })];
DM.runQa(rows);
check('placeholder cleared BEFORE V is decided', [rows[0].values[OUT.OTHERS], rows[0].values[OUT.NONE]],
  [null, CORE.NONE_APPLY_TEXT]);

// B and C are never written
rows = [mk({ [OUT.BIRTHDATE]: 'nonsense', [OUT.HEAD_AND_HAIR]: 'Sombrero' })];
DM.runQa(rows);
check('QA never writes column B', rows[0].values[OUT.QA_RESULT], null);
check('QA never writes column C', rows[0].values[OUT.COMMENTS], null);

console.log('\nWorkbook detection');
check('.xlsx is metadata', DM.isMetadataWorkbook({ name: 'Batch.xlsx' }), true);
check('.xlsm is metadata', DM.isMetadataWorkbook({ name: 'Batch.xlsm' }), true);
check('Excel lock file ignored', DM.isMetadataWorkbook({ name: '~$Batch.xlsx' }), false);
check('.mp4 is not metadata', DM.isMetadataWorkbook({ name: 'X_INDOOR_NEUTRAL.mp4' }), false);

console.log('\nParticipant identity from path or filename');
const ROOTS = ['/Agency/Powerling/Batch_5'];
check('per-participant folder', DM.participantKeysForFile(
  { name: 'A_INDOOR_NEUTRAL.mp4', path_display: '/Agency/Powerling/Batch_5/Abraham Otieno/A_INDOOR_NEUTRAL.mp4' },
  ROOTS, []), ['abraham otieno']);
check('flat media, name from filename', DM.participantKeysForFile(
  { name: 'ABRAHAM OTIENO_INDOOR_NEUTRAL.mp4', path_display: '/Agency/Powerling/Batch_5/ABRAHAM OTIENO_INDOOR_NEUTRAL.mp4' },
  ROOTS, ['abraham otieno']), ['abraham otieno']);
check('unrelated filename yields nothing', DM.participantKeysForFile(
  { name: 'readme.txt', path_display: '/Agency/Powerling/Batch_5/readme.txt' },
  ROOTS, ['abraham otieno']), []);

console.log('\n' + (fail === 0 ? 'ALL ' + pass + ' CHECKS PASSED' : pass + ' passed, ' + fail + ' FAILED'));
process.exit(fail === 0 ? 0 : 1);

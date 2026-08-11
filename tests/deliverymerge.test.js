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

console.log('\nColumns A, B, C are generated output, never sourced');

// participant_sequence: one number per participant, repeated across their video rows
let annotated = [
  mk({ [OUT.FOLDER]: 'Abraham Otieno' }), mk({ [OUT.FOLDER]: 'Abraham Otieno' }),
  mk({ [OUT.FOLDER]: 'Asnet Kalai' }),    mk({ [OUT.FOLDER]: 'ABRAHAM  OTIENO' }),
];
DM.annotate(annotated, [], null);
check('sequence numbers participants, not rows',
  annotated.map(r => r.values[OUT.SEQUENCE]), [1, 1, 2, 1]);
check('clean rows read OK',
  annotated.map(r => r.values[OUT.QA_RESULT]), ['OK', 'OK', 'OK', 'OK']);
check('clean rows have no comment', annotated[0].values[OUT.COMMENTS], null);

// verdicts escalate: auto-fixed < review < error
annotated = [mk({ [OUT.FOLDER]: 'A' }), mk({ [OUT.FOLDER]: 'B' }), mk({ [OUT.FOLDER]: 'C' })];
DM.annotate(annotated, [
  { row: 2, severity: 'fixed',  rule: 'Whitespace tidied', message: 'Spaces removed.' },
  { row: 3, severity: 'fixed',  rule: 'Whitespace tidied', message: 'Spaces removed.' },
  { row: 3, severity: 'review', rule: 'Birthdate needs a manual fix', message: 'Unreadable.' },
  { row: 4, severity: 'error',  rule: 'Workbook skipped', message: 'Bad file.' },
], null);
check('auto-fixed only', annotated[0].values[OUT.QA_RESULT], DM.QA_FIXED);
check('review beats auto-fixed', annotated[1].values[OUT.QA_RESULT], DM.QA_REVIEW);
check('error beats review', annotated[2].values[OUT.QA_RESULT], DM.QA_ERROR);
check('comments join every finding for that row',
  annotated[1].values[OUT.COMMENTS],
  'Whitespace tidied: Spaces removed. | Birthdate needs a manual fix: Unreadable.');

// link-mapping problems land in the same verdict
annotated = [mk({ [OUT.FOLDER]: 'A' })];
DM.annotate(annotated, [], [{ problems: ['No unambiguous assent file for this participant.'] }]);
check('link problems force review', annotated[0].values[OUT.QA_RESULT], DM.QA_REVIEW);
check('link problems appear in comments', annotated[0].values[OUT.COMMENTS],
  'No unambiguous assent file for this participant.');

// a row can carry both a QA finding and a link problem; both must survive
annotated = [mk({ [OUT.FOLDER]: 'Asnet Kalai' })];
DM.annotate(annotated,
  [{ row: 2, severity: 'review', rule: 'Birthdate needs a manual fix', message: 'Unreadable.' }],
  [{ problems: ['No unambiguous assent file for this participant.'] }]);
check('QA finding and link problem both appear', annotated[0].values[OUT.COMMENTS],
  'Birthdate needs a manual fix: Unreadable. | No unambiguous assent file for this participant.');
check('verdict stays Review', annotated[0].values[OUT.QA_RESULT], DM.QA_REVIEW);

// file-level issues have no row and must not be pinned to row 2
annotated = [mk({ [OUT.FOLDER]: 'A' })];
DM.annotate(annotated, [{ row: null, severity: 'error', rule: 'Workbook skipped', message: 'x' }], null);
check('file-level issue not attributed to a row', annotated[0].values[OUT.QA_RESULT], DM.QA_OK);

console.log('\n' + (fail === 0 ? 'ALL ' + pass + ' CHECKS PASSED' : pass + ' passed, ' + fail + ' FAILED'));
process.exit(fail === 0 ? 0 : 1);

/* Node harness: stub `window`, load both modules, exercise the spec's edge cases. */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = '/Users/davidk/Desktop/Projects/ExcelMerger';
// URL/URLSearchParams are ambient in a browser but not in a bare vm context.
const sandbox = { window: {}, console, URL, URLSearchParams };
sandbox.global = sandbox;
vm.createContext(sandbox);
for (const f of ['js/core.js', 'js/linkfiller.js']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), sandbox, { filename: f });
}
const LF = sandbox.window.LinkFiller;

let pass = 0, fail = 0;
const check = (label, actual, expected) => {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a === e) { pass++; console.log('  ok   ' + label); }
  else { fail++; console.log('  FAIL ' + label + '\n         got      ' + a + '\n         expected ' + e); }
};

const ROOT_PATH = '/Agency Collection/Powerling/Batch_5';

/* ---- unit: folder URL parsing ---- */
console.log('\nFolder URL parsing');
check('team path decodes',
  LF.parseFolderUrl('https://www.dropbox.com/home/Centific%20Team%20Folder/Agency%20Collection/Powerling/Batch_5_July%2029th%202026').path,
  '/Centific Team Folder/Agency Collection/Powerling/Batch_5_July 29th 2026');
check('scl/fo share link rejected', LF.parseFolderUrl('https://www.dropbox.com/scl/fo/9xk2/broken-share-link').ok, false);
check('bare path accepted', LF.parseFolderUrl('/Agency Collection/Aqlama/20260721').path, '/Agency Collection/Aqlama/20260721');
check('trailing slash trimmed', LF.parseFolderUrl('/a/b/').path, '/a/b');
check('non-dropbox host rejected', LF.parseFolderUrl('https://example.com/home/x').ok, false);

/* ---- unit: shared link normalisation ---- */
console.log('\nShared link normalisation');
check('st= dropped, dl=0 kept',
  LF.normalizeSharedLink('https://www.dropbox.com/scl/fi/abc123/CLIP.mp4?rlkey=k9x&st=ephemeral9&dl=0'),
  'https://www.dropbox.com/scl/fi/abc123/CLIP.mp4?rlkey=k9x&dl=0');
check('dl=1 forced to dl=0',
  LF.normalizeSharedLink('https://www.dropbox.com/scl/fi/abc/F.pdf?rlkey=zz&dl=1'),
  'https://www.dropbox.com/scl/fi/abc/F.pdf?rlkey=zz&dl=0');

/* ---- unit: name normalisation ---- */
console.log('\nName normalisation');
check('case + separators collapse',
  [LF.normalizeName('Abraham Otieno'), LF.normalizeName('ABRAHAM OTIENO'), LF.normalizeName('ABRAHAM-OTIENO')],
  ['abraham otieno', 'abraham otieno', 'abraham otieno']);
check('near-duplicate names stay distinct',
  LF.normalizeName('Ronald Okoth') === LF.normalizeName('Ronald Okothh'), false);
check('lowercase leading name normalises', LF.normalizeName('faith Kwenyu'), 'faith kwenyu');

/* ---- unit: classification + derivation ---- */
console.log('\nClassification and derivation');
const cfg = LF.AGENCY_PRESETS.Powerling;
check('.MOV uppercase is video', LF.classifyFile('X_INDOOR_NEUTRAL.MOV', cfg), 'video');
check('consent pdf is icf', LF.classifyFile('ABRAHAM-OTIENO-CONSENT-FORM.pdf', cfg), 'icf');
check('assent jpg is assent', LF.classifyFile('ABRAHAM_OTIENO_ASSENT-FORM.jpg', cfg), 'assent');
check('stray xlsx ignored', LF.classifyFile('batch_metadata.xlsx', cfg), 'other');
check('Aqlama ICF pdf is icf', LF.classifyFile('Asnet Kalai – ICF.pdf', LF.AGENCY_PRESETS.Aqlama), 'icf');
check('env from filename', [LF.deriveEnvironment('A_INDOOR_NEUTRAL.mp4'), LF.deriveEnvironment('A_OUTDOOR_ANGRY.mp4')], ['Indoor', 'Outdoor']);
check('neutral token', LF.deriveExpression('A_INDOOR_NEUTRAL.mp4', cfg).expression, 'Neutral');
const typo = LF.deriveExpression('FAITH_KWENYU_INDOOR_N3UTRAL.MOV', cfg);
check('typo N3UTRAL still neutral', [typo.expression, typo.fuzzy], ['Neutral', true]);
check('smiling is non-neutral', LF.deriveExpression('A_INDOOR_SMILING.mp4', cfg).expression, 'Non-Neutral');
check('non-neutral token extracted', LF.deriveExpression('A_INDOOR_SMILING.mp4', cfg).token, 'smiling');
check('Aqlama phrasing reads neutral',
  LF.deriveExpression('Asnet Kalai – Indoor Neutral Expression.mov', LF.AGENCY_PRESETS.Aqlama).expression, 'Neutral');

/* ---- real agency video filenames ---- */
console.log('\nReal video filenames');
const VIDEOS = [
  ['ABRAHAM-OTIENO_INDOOR_N3UTRAL.mp4',                                'Indoor',  'Neutral',     null,       'abraham otieno'],
  ['ALISHER-IMBULANI_OUTDOOR_NEUTRAL.mp4',                             'Outdoor', 'Neutral',     null,       'alisher imbulani'],
  ['Sophie Strnadelová – Outdoor Non Neutral Expression.mp4 Frowning', 'Outdoor', 'Non-Neutral', 'frowning', 'sophie strnadelova'],
  ['Najeeb – Indoor Non Neutral Expression Smiling.mp4',               'Indoor',  'Non-Neutral', 'smiling',  'najeeb'],
  ['Maria Vogiatzi – Indoor Non Neutral Expression Smiling,mp4',       'Indoor',  'Non-Neutral', 'smiling',  'maria vogiatzi'],
  ['Chizuru Watanabe – Indoor Neutral Expression.mov',                 'Indoor',  'Neutral',     null,       'chizuru watanabe'],
  ['R B Vaishali - Indoor NonNeutral angry.mp4',                       'Indoor',  'Non-Neutral', 'angry',    'r b vaishali'],
];
VIDEOS.forEach(function (row) {
  const name = row[0], env = row[1], expr = row[2], token = row[3];
  const short = name.length > 34 ? name.slice(0, 34) + '…' : name;
  check(short + ' is a video', LF.classifyFile(name, cfg), 'video');
  check(short + ' env', LF.deriveEnvironment(name), env);
  const derived = LF.deriveExpression(name, cfg);
  check(short + ' expression', derived.expression, expr);
  if (token) check(short + ' token', derived.token, token);
});

// the trap: "Non Neutral" contains "neutral"
check('"Non Neutral" is never read as Neutral',
  ['Non Neutral', 'NonNeutral', 'Non-Neutral', 'non neutral', 'Non N3UTRAL'].map(function (variant) {
    return LF.deriveExpression('X – Indoor ' + variant + ' Expression.mp4', cfg).expression;
  }), ['Non-Neutral', 'Non-Neutral', 'Non-Neutral', 'Non-Neutral', 'Non-Neutral']);
check('plain Neutral still reads Neutral',
  LF.deriveExpression('X – Indoor Neutral Expression.mp4', cfg).expression, 'Neutral');

// accents must fold, not vanish
check('accented name folds to the unaccented spelling',
  LF.normalizeName('Sophie Strnadelová'), LF.normalizeName('Sophie Strnadelova'));
check('consent files are still not videos',
  [LF.classifyFile('ABRAHAM-OTIENO-CONSENT.pdf', cfg), LF.classifyFile('X_ASSENT.jpg', cfg)],
  ['icf', 'assent']);

/* ---- integration ---- */
const row = (name, expr, env, k) => {
  const v = new Array(30).fill(null);
  v[LF.COL.NAME] = name; v[LF.COL.EXPRESSION] = expr;
  v[LF.COL.ENVIRONMENT] = env; v[LF.COL.NON_NEUTRAL] = k || null;
  v[10 + 0] = v[10 + 0]; // keep shape explicit
  for (let i = LF.UNTOUCHED_FIRST; i <= LF.UNTOUCHED_LAST; i++) if (v[i] == null) v[i] = 'preexisting';
  return v;
};
let rowNo = 1;
const R = (name, expr, env, k) => ({ excelRow: ++rowNo, values: row(name, expr, env, k) });

const sheet = { rows: [
  R('Abraham Otieno', 'Neutral', 'Indoor'),                       // 2
  R('Abraham Otieno', 'Non-Neutral', 'Indoor', 'Smiling'),        // 3
  R('Abraham Otieno', 'Neutral', 'Outdoor'),                      // 4
  R('Amani Granvill', 'Non-Neutral', 'Indoor'),                   // 5  ambiguous, K empty
  R('Amani Granvill', 'Non-Neutral', 'Indoor', 'Frowning'),       // 6  same slot, K resolves it
  R('Asnet Kalai', 'Non-Neutral', 'Outdoor', 'Angry'),            // 7  no assent
  R('Ronald Okoth', 'Neutral', 'Indoor'),                         // 8
  R('Ronald Okothh', 'Neutral', 'Indoor'),                        // 9  different person
  R('faith Kwenyu', 'Neutral', 'Indoor'),                         // 10 .MOV + typo
  R('Adam khashaba', 'Neutral', 'Indoor'),                        // 11 duplicate ICF
]};

const f = (folder, name, link) => ({ name, path_display: ROOT_PATH + '/' + folder + '/' + name, link: link || '' });
const LINKED = 'https://www.dropbox.com/scl/fi/exist1/F?rlkey=r1&st=tok&dl=0';

const files = [
  f('Abraham Otieno', 'ABRAHAM_OTIENO_INDOOR_NEUTRAL.mp4', LINKED),
  f('Abraham Otieno', 'ABRAHAM_OTIENO_INDOOR_SMILING.mp4'),
  f('Abraham Otieno', 'ABRAHAM_OTIENO_OUTDOOR_NEUTRAL.mp4'),
  f('Abraham Otieno', 'ABRAHAM-OTIENO-CONSENT-FORM.pdf', LINKED),
  f('Abraham Otieno', 'ABRAHAM_OTIENO_ASSENT-FORM.jpg', LINKED),
  f('Abraham Otieno', 'batch_metadata.xlsx'),                       // stray, must be ignored

  f('Amani Granvill', 'AMANI_GRANVILL_INDOOR_SMILING.mp4'),         // two non-neutral,
  f('Amani Granvill', 'AMANI_GRANVILL_INDOOR_FROWNING.mp4'),        // one slot
  f('Amani Granvill', 'AMANI-GRANVILL-CONSENT.pdf'),
  f('Amani Granvill', 'AMANI_GRANVILL_ASSENT.jpg'),

  f('Asnet Kalai', 'ASNET_KALAI_OUTDOOR_ANGRY.mp4'),
  f('Asnet Kalai', 'ASNET-KALAI-CONSENT.pdf'),                      // no assent

  f('Ronald Okoth', 'RONALD_OKOTH_INDOOR_NEUTRAL.mp4'),
  f('Ronald Okoth', 'RONALD-OKOTH-CONSENT.pdf'),
  f('Ronald Okoth', 'RONALD_OKOTH_ASSENT.jpg'),
  f('Ronald Okothh', 'RONALD_OKOTHH_INDOOR_NEUTRAL.mp4'),
  f('Ronald Okothh', 'RONALD-OKOTHH-CONSENT.pdf'),
  f('Ronald Okothh', 'RONALD_OKOTHH_ASSENT.jpg'),

  f('faith Kwenyu', 'FAITH_KWENYU_INDOOR_N3UTRAL.MOV'),             // typo + uppercase ext
  f('faith Kwenyu', 'FAITH-KWENYU-CONSENT.pdf'),
  f('faith Kwenyu', 'FAITH_KWENYU_ASSENT.png'),

  f('Adam khashaba', 'ADAM_KHASHABA_INDOOR_NEUTRAL.mp4'),
  f('Adam khashaba', 'ICF 1.pdf'),                                  // duplicate consent
  f('Adam khashaba', 'ICF 2.pdf'),
  f('Adam khashaba', 'ADAM_KHASHABA_ASSENT.pdf'),

  f('B. Idagiza', 'B_IDAGIZA_INDOOR_NEUTRAL.mp4'),                  // no sheet row
  f('B. Idagiza', 'B-IDAGIZA-CONSENT.pdf'),
];

const plan = LF.planFill(sheet, files, { config: cfg, roots: [ROOT_PATH] });
const byRow = {};
plan.rows.forEach(r => { byRow[r.excelRow] = r; });
const guardKinds = plan.guards.map(g => g.kind).sort();

console.log('\nIntegration — row outcomes');
check('r2 all three links exist -> writable', byRow[2].status, 'filled');
check('r3 sole indoor non-neutral matched', [byRow[3].status, byRow[3].videoReason], ['awaiting-link', 'unique']);
check('r3 video is the SMILING file', byRow[3].video.name, 'ABRAHAM_OTIENO_INDOOR_SMILING.mp4');
check('r4 matched but video unlinked', byRow[4].status, 'awaiting-link');
check('AC repeats and points at first row', [byRow[3].icfRepeatOf, byRow[4].icfRepeatOf], [2, 2]);
check('r2 has no repeat marker', byRow[2].icfRepeatOf, null);
check('r5 ambiguous slot held (K empty)', byRow[5].status, 'review');
check('r6 SAME slot resolved by column K', [byRow[6].status, byRow[6].videoReason], ['awaiting-link', 'column-k']);
check('r6 picked the FROWNING file', byRow[6].video.name, 'AMANI_GRANVILL_INDOOR_FROWNING.mp4');
check('r7 missing assent held', [byRow[7].status, byRow[7].assent], ['review', null]);
check('r7 assent NOT borrowed', byRow[7].assent, null);
check('r8 Ronald Okoth matched', byRow[8].status, 'awaiting-link');
check('r9 Ronald Okothh matched separately', byRow[9].status, 'awaiting-link');
check('near-duplicates got different videos',
  byRow[8].video.name !== byRow[9].video.name, true);
check('r10 typo + .MOV matched', byRow[10].status, 'awaiting-link');
check('r11 duplicate ICF held', [byRow[11].status, byRow[11].icf], ['review', null]);

console.log('\nIntegration — guards');
check('all five guard kinds raised', guardKinds,
  ['ambiguous-slot', 'duplicate-icf', 'missing-assent', 'unmatched-name'].sort());
check('unmatched names B. Idagiza',
  plan.guards.filter(g => g.kind === 'unmatched-name')[0].who, 'folder "B. Idagiza"');
check('ambiguous guard raised once', plan.guards.filter(g => g.kind === 'ambiguous-slot').length, 1);

console.log('\nIntegration — metrics and writes');
check('only fully-linked rows are writable', plan.metrics.rowsFilled, 1);
check('matched-but-unlinked counted separately', plan.metrics.rowsAwaitingLink, 6);
check('matched total unchanged', plan.metrics.rowsMatched, 7);
check('rows held for review', plan.metrics.rowsReview, 3);
check('reused links counted', plan.metrics.linksReused, 3);
check('stray xlsx never became a link',
  plan.rows.some(r => [r.video, r.icf, r.assent].some(x => x && /\.xlsx$/i.test(x.name))), false);
check('links to create > 0', plan.metrics.linksToCreate > 0, true);
check('filesNeedingLinks excludes already-linked',
  LF.filesNeedingLinks(plan).some(x => x.linkExisted), false);

/* applyPlan must only touch E/AC/AD */
const touched = new Set();
const fakeSheet = { getRow: (n) => ({ getCell: (c) => ({ set value(v) { touched.add(c); } }), commit() {} }) };
const written = LF.applyPlan(fakeSheet, plan);
check('unlinked rows are NOT written as blanks', written, 1);
check('only columns E(5) AC(29) AD(30) touched', Array.from(touched).sort((a, b) => a - b), [5, 29, 30]);

console.log('\n' + (fail === 0 ? 'ALL ' + pass + ' CHECKS PASSED' : pass + ' passed, ' + fail + ' FAILED'));
process.exit(fail === 0 ? 0 : 1);

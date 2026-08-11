/* Dropbox header encoding and scope accounting. node tests/dropbox.test.js */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const store = {};
const sandbox = {
  window: {}, console, URL, URLSearchParams, TextEncoder, Headers,
  localStorage: { getItem: k => (k in store ? store[k] : null), setItem: (k, v) => { store[k] = String(v); }, removeItem: k => { delete store[k]; } },
  sessionStorage: { getItem: () => null, setItem() {}, removeItem() {} },
  location: { origin: 'https://x.io', pathname: '/p/' },
  crypto: {}, btoa: () => '', fetch: () => {},
};
sandbox.global = sandbox;
sandbox.window = sandbox;
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(ROOT, 'js/dropbox.js'), 'utf8'), sandbox, { filename: 'dropbox.js' });
const DBX = sandbox.window.DropboxSource;

let pass = 0, fail = 0;
const check = (label, actual, expected) => {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a === e) { pass++; console.log('  ok   ' + label); }
  else { fail++; console.log('  FAIL ' + label + '\n         got      ' + a + '\n         expected ' + e); }
};

/* asciiArg is private; lift it out of the source so the real implementation is tested. */
const src = fs.readFileSync(path.join(ROOT, 'js/dropbox.js'), 'utf8');
const asciiArg = eval('(' + /function asciiArg\(value\) \{[\s\S]*?\n  \}/.exec(src)[0].replace('function asciiArg', 'function') + ')');

console.log('\nDropbox-API-Arg must be a valid ByteString');
const REAL_NAMES = [
  ['en dashes (Aqlama)', '/Agency/Aqlama – Centaurus – Metadata – 20260728.xlsx'],
  ['plain ascii',        '/Agency/Powerling/Batch_5/Abraham Otieno/A_INDOOR_NEUTRAL.mp4'],
  ['accents',            '/Agency/Zoë Müller/ICF.pdf'],
  ['cjk',                '/Agency/日本語/meta.xlsx'],
  ['emoji',              '/Agency/Batch 🎬/meta.xlsx'],
];
REAL_NAMES.forEach(([label, p]) => {
  const escaped = asciiArg({ path: p });
  let accepted = true;
  try { new Headers({ 'Dropbox-API-Arg': escaped }); } catch (e) { accepted = false; }
  check(label + ' — header accepted', accepted, true);
  check(label + ' — decodes to the original path', JSON.parse(escaped).path, p);
});

// the bug this guards against: raw JSON throws for anything above Latin-1
let rawThrew = false;
try { new Headers({ 'Dropbox-API-Arg': JSON.stringify({ path: REAL_NAMES[0][1] }) }); }
catch (e) { rawThrew = true; }
check('unescaped JSON would have thrown', rawThrew, true);

console.log('\nScope accounting');
check('nothing recorded yet -> no false alarm', DBX.missingScopes(), []);
store['merger.dropbox.scopes'] = 'account_info.read files.metadata.read sharing.read sharing.write';
check('sign-in predating files.content.read', DBX.missingScopes(), ['files.content.read']);
store['merger.dropbox.scopes'] = DBX.SCOPES.join(' ');
check('full grant is clean', DBX.missingScopes(), []);
store['merger.dropbox.scopes'] = 'account_info.read';
check('re-read is not memoised', DBX.missingScopes(),
  ['files.metadata.read', 'files.content.read', 'sharing.read', 'sharing.write']);

console.log('\nRequired scopes');
check('all five are requested', DBX.SCOPES,
  ['account_info.read', 'files.metadata.read', 'files.content.read', 'sharing.read', 'sharing.write']);

console.log('\n' + (fail === 0 ? 'ALL ' + pass + ' CHECKS PASSED' : pass + ' passed, ' + fail + ' FAILED'));
process.exit(fail === 0 ? 0 : 1);

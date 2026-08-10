/* Merge + QA engine. Pure browser JS, no network, no build step. */
(function (global) {
  'use strict';

  const CANONICAL_HEADERS = [
    'Participant Name As Per ICF',
    'Participant Email',
    'Name of Parents/Legal Guardian',
    'Email of Parents/Legal Guardian',
    'Expressions',
    'For Non Neutral, Please Select',
    'Age Group',
    'Birthdate',
    'Gender',
    'Main Ethnicity',
    'Secondary Ethnicity',
    'Environment',
    'Head and Hair',
    'Facial Features',
    'Accessories and jewellery',
    'Others',
    'None',
    'Skintone',
    'Device',
    'Country of Collection',
    'State Abbreviation',
  ];

  const COL = {
    PARTICIPANT_NAME: 0,  // A
    PARTICIPANT_EMAIL: 1, // B
    BIRTHDATE: 7,         // H
    HEAD_AND_HAIR: 12,    // M
    FACIAL_FEATURES: 13,  // N
    ACCESSORIES: 14,      // O
    OTHERS: 15,           // P
    NONE: 16,             // Q
  };

  const HEAD_ATTRIBUTE_START = COL.HEAD_AND_HAIR; // M
  const HEAD_ATTRIBUTE_END = COL.OTHERS;          // P
  const NONE_APPLY_TEXT = 'N/A-none apply';

  // Header spelling drifts between agencies; map normalized variants onto the canonical index.
  const HEADER_ALIASES = {
    dob: 'Birthdate',
    dateofbirth: 'Birthdate',
    birthday: 'Birthdate',
    accessoriesandjewelry: 'Accessories and jewellery',
    accessoriesjewellery: 'Accessories and jewellery',
    accessoriesjewelry: 'Accessories and jewellery',
    participantnameasperlcf: 'Participant Name As Per ICF',
    nameofparentlegalguardian: 'Name of Parents/Legal Guardian',
    emailofparentlegalguardian: 'Email of Parents/Legal Guardian',
    stateabbrev: 'State Abbreviation',
    stateabbreviations: 'State Abbreviation',
    countryofcollections: 'Country of Collection',
  };

  const MONTH_NAMES = {
    jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3, apr: 4, april: 4,
    may: 5, jun: 6, june: 6, jul: 7, july: 7, aug: 8, august: 8, sep: 9, sept: 9,
    september: 9, oct: 10, october: 10, nov: 11, november: 11, dec: 12, december: 12,
  };

  const MIN_BIRTH_YEAR = 1900;

  function columnLetter(index) {
    let n = index + 1;
    let s = '';
    while (n > 0) {
      const rem = (n - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function normalizeHeaderKey(value) {
    return String(value == null ? '' : value)
      .toLowerCase()
      .replace(/[\u00A0\u2007\u202F]/g, ' ')
      .replace(/[^a-z0-9]/g, '');
  }

  const CANONICAL_BY_KEY = (function () {
    const map = {};
    CANONICAL_HEADERS.forEach(function (h, i) { map[normalizeHeaderKey(h)] = i; });
    Object.keys(HEADER_ALIASES).forEach(function (alias) {
      const target = CANONICAL_HEADERS.indexOf(HEADER_ALIASES[alias]);
      if (target >= 0) map[alias] = target;
    });
    return map;
  })();

  function normalizeWhitespace(text) {
    return text
      .replace(/[\u00A0\u2007\u202F]/g, ' ')
      .replace(/[\u200B-\u200D\uFEFF]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  /* ExcelJS cell values arrive as plain scalars or as rich-text/hyperlink/formula wrappers. */
  function unwrapCellValue(value) {
    if (value == null) return null;
    if (value instanceof Date) return value;
    const type = typeof value;
    if (type === 'string' || type === 'number' || type === 'boolean') return value;
    if (Array.isArray(value.richText)) {
      return value.richText.map(function (part) { return part.text || ''; }).join('');
    }
    if (value.text != null) return value.text;
    if (value.result != null) return unwrapCellValue(value.result);
    if (value.error) return null;
    if (value.hyperlink) return value.hyperlink;
    return String(value);
  }

  function isBlank(value) {
    return value == null || (typeof value === 'string' && value.trim() === '');
  }

  function pad2(n) {
    return (n < 10 ? '0' : '') + n;
  }

  function formatYearMonth(year, month) {
    return String(year) + '/' + pad2(month);
  }

  function excelSerialToDate(serial) {
    // Excel's 1900 epoch, offset by its phantom leap day. 25569 = days from 1900 epoch to Unix epoch.
    return new Date(Math.round((serial - 25569) * 86400000));
  }

  function expandTwoDigitYear(yy, currentYear) {
    const currentTwoDigit = currentYear % 100;
    return yy > currentTwoDigit ? 1900 + yy : 2000 + yy;
  }

  /*
   * Coerce anything an agency might type into YYYY/MM.
   * Returns { value, status, message } where status is ok | fixed | review | unparseable | empty.
   */
  function parseBirthdate(raw, options) {
    const opts = options || {};
    const dayFirst = opts.dayFirst !== false;
    const currentYear = opts.currentYear || new Date().getUTCFullYear();

    function finish(year, month, status, message) {
      if (!(month >= 1 && month <= 12)) {
        return { value: null, status: 'unparseable', message: 'Month "' + month + '" is not between 1 and 12.' };
      }
      if (year < MIN_BIRTH_YEAR || year > currentYear) {
        return {
          value: formatYearMonth(year, month),
          status: 'review',
          message: 'Year ' + year + ' is outside the plausible range ' + MIN_BIRTH_YEAR + '-' + currentYear + '.',
        };
      }
      return { value: formatYearMonth(year, month), status: status, message: message || '' };
    }

    if (isBlank(raw)) return { value: null, status: 'empty', message: '' };

    if (raw instanceof Date) {
      if (isNaN(raw.getTime())) return { value: null, status: 'unparseable', message: 'Invalid date value.' };
      return finish(raw.getUTCFullYear(), raw.getUTCMonth() + 1, 'fixed', 'Converted from a date cell.');
    }

    if (typeof raw === 'number' && isFinite(raw)) {
      const n = raw;
      if (Number.isInteger(n) && n >= MIN_BIRTH_YEAR && n <= currentYear) {
        return { value: null, status: 'review', message: 'Only a year (' + n + ') was supplied; the month is missing.' };
      }
      if (Number.isInteger(n) && n >= 190001 && n <= 999912) {
        const y = Math.floor(n / 100);
        const m = n % 100;
        return finish(y, m, 'fixed', 'Interpreted numeric value as YYYYMM.');
      }
      if (Number.isInteger(n) && n >= 19000101 && n <= 99991231) {
        const y8 = Math.floor(n / 10000);
        const m8 = Math.floor(n / 100) % 100;
        return finish(y8, m8, 'fixed', 'Interpreted numeric value as YYYYMMDD.');
      }
      if (n > 0 && n < 80000) {
        const d = excelSerialToDate(n);
        if (!isNaN(d.getTime())) {
          return finish(d.getUTCFullYear(), d.getUTCMonth() + 1, 'fixed', 'Converted from an Excel date serial number.');
        }
      }
      return { value: null, status: 'unparseable', message: 'Numeric value ' + n + ' is not a recognisable date.' };
    }

    const text = normalizeWhitespace(String(raw));
    if (!text) return { value: null, status: 'empty', message: '' };

    // Month spelled out, in either order: "May 1990", "1990 May", "12 May 1990".
    const nameMatch = text.match(/^([A-Za-z]{3,9})[\s.,\-\/]+(\d{2,4})$/) || text.match(/^(\d{2,4})[\s.,\-\/]+([A-Za-z]{3,9})$/);
    if (nameMatch) {
      const aIsName = /[A-Za-z]/.test(nameMatch[1]);
      const nameToken = (aIsName ? nameMatch[1] : nameMatch[2]).toLowerCase();
      const numToken = aIsName ? nameMatch[2] : nameMatch[1];
      const month = MONTH_NAMES[nameToken];
      if (month) {
        let year = parseInt(numToken, 10);
        if (numToken.length <= 2) {
          year = expandTwoDigitYear(year, currentYear);
          return finish(year, month, 'review', 'Expanded 2-digit year "' + numToken + '" to ' + year + '. Please confirm.');
        }
        return finish(year, month, 'fixed', 'Converted month name to numeric format.');
      }
    }

    const longNameMatch = text.match(/^(\d{1,2})[\s.,\-\/]+([A-Za-z]{3,9})[\s.,\-\/]+(\d{2,4})$/)
      || text.match(/^([A-Za-z]{3,9})[\s.,\-\/]+(\d{1,2})[\s.,\-\/]+(\d{2,4})$/);
    if (longNameMatch) {
      const monthToken = /[A-Za-z]/.test(longNameMatch[1]) ? longNameMatch[1] : longNameMatch[2];
      const month = MONTH_NAMES[monthToken.toLowerCase()];
      if (month) {
        const yearToken = longNameMatch[3];
        let year = parseInt(yearToken, 10);
        if (yearToken.length <= 2) {
          year = expandTwoDigitYear(year, currentYear);
          return finish(year, month, 'review', 'Expanded 2-digit year "' + yearToken + '" to ' + year + '. Please confirm.');
        }
        return finish(year, month, 'fixed', 'Converted month name to numeric format.');
      }
    }

    const digitsOnly = text.replace(/\D/g, '');
    if (/^\d+$/.test(text)) {
      if (text.length === 4) {
        const y = parseInt(text, 10);
        if (y >= MIN_BIRTH_YEAR && y <= currentYear) {
          return { value: null, status: 'review', message: 'Only a year (' + y + ') was supplied; the month is missing.' };
        }
      }
      if (digitsOnly.length === 6) {
        return finish(parseInt(digitsOnly.slice(0, 4), 10), parseInt(digitsOnly.slice(4, 6), 10), 'fixed', 'Interpreted as YYYYMM.');
      }
      if (digitsOnly.length === 8) {
        return finish(parseInt(digitsOnly.slice(0, 4), 10), parseInt(digitsOnly.slice(4, 6), 10), 'fixed', 'Interpreted as YYYYMMDD.');
      }
    }

    const parts = text.split(/[^0-9]+/).filter(function (p) { return p !== ''; });

    if (parts.length === 2) {
      const a = parts[0];
      const b = parts[1];
      if (a.length === 4) {
        return finish(parseInt(a, 10), parseInt(b, 10), 'fixed', 'Reformatted to YYYY/MM.');
      }
      if (b.length === 4) {
        return finish(parseInt(b, 10), parseInt(a, 10), 'fixed', 'Reordered MM/YYYY to YYYY/MM.');
      }
      // Both 2-digit: assume MM/YY, which is a guess worth a human glance.
      const mm = parseInt(a, 10);
      const yy = expandTwoDigitYear(parseInt(b, 10), currentYear);
      return finish(yy, mm, 'review', 'Read "' + text + '" as MM/YY and expanded the year to ' + yy + '. Please confirm.');
    }

    if (parts.length === 3) {
      const p0 = parseInt(parts[0], 10);
      const p1 = parseInt(parts[1], 10);
      const p2 = parseInt(parts[2], 10);

      if (parts[0].length === 4) {
        return finish(p0, p1, 'fixed', 'Reformatted from YYYY/MM/DD to YYYY/MM.');
      }
      if (parts[2].length === 4) {
        if (p0 > 12 && p1 <= 12) {
          return finish(p2, p1, 'fixed', 'Reformatted from DD/MM/YYYY to YYYY/MM.');
        }
        if (p1 > 12 && p0 <= 12) {
          return finish(p2, p0, 'fixed', 'Reformatted from MM/DD/YYYY to YYYY/MM.');
        }
        if (p0 <= 12 && p1 <= 12) {
          const month = dayFirst ? p1 : p0;
          const order = dayFirst ? 'DD/MM/YYYY' : 'MM/DD/YYYY';
          return finish(p2, month, 'review',
            'Day and month are both <= 12, so "' + text + '" is ambiguous. Assumed ' + order + '. Please confirm.');
        }
      }
      if (parts[2].length <= 2 && parts[0].length <= 2) {
        const year = expandTwoDigitYear(p2, currentYear);
        const month = dayFirst ? p1 : p0;
        return finish(year, month, 'review', 'Read "' + text + '" using a 2-digit year expanded to ' + year + '. Please confirm.');
      }
    }

    return { value: null, status: 'unparseable', message: 'Could not recognise "' + text + '" as a date.' };
  }

  /*
   * Controlled vocabulary for the head attribute block. These are the only values the
   * client accepts in M-P; anything else is either a typo we can fix or a real problem.
   */
  const ATTRIBUTE_VOCABULARY = {};
  ATTRIBUTE_VOCABULARY[COL.HEAD_AND_HAIR] = ['Glasses', 'Religious headwear', 'Hat', 'Scarf'];
  ATTRIBUTE_VOCABULARY[COL.FACIAL_FEATURES] = ['Mustache', 'Beard', 'Dimples', 'Facial scars', 'Facial moles', 'Acne', 'Face/neck tattoos'];
  ATTRIBUTE_VOCABULARY[COL.ACCESSORIES] = ['Makeup', 'Necklace', 'Earrings', 'Nose piercing', 'Lip piercing', 'Eyebrow piercing'];
  ATTRIBUTE_VOCABULARY[COL.OTHERS] = ['Freckles', 'Wrinkles', 'Bindi', 'Other tattoos', 'Other piercings', 'Other - not specified'];

  const MULTI_VALUE_SPLIT = /[;,]/;
  const MULTI_VALUE_JOIN = ', ';

  /* Fold away the differences that are never meaningful: case, punctuation, spacing. */
  function vocabKey(text) {
    return String(text)
      .toLowerCase()
      .replace(/[\/\\\-_.,;:]+/g, ' ')
      .replace(/[^a-z0-9 ]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  /* Damerau-Levenshtein (optimal string alignment) so a swapped pair costs 1, not 2. */
  function editDistance(a, b) {
    const m = a.length;
    const n = b.length;
    if (!m) return n;
    if (!n) return m;

    const d = [];
    for (let i = 0; i <= m; i++) d.push([i]);
    for (let j = 1; j <= n; j++) d[0][j] = j;

    for (let i = 1; i <= m; i++) {
      for (let j = 1; j <= n; j++) {
        const cost = a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1;
        d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + cost);
        if (i > 1 && j > 1 && a.charAt(i - 1) === b.charAt(j - 2) && a.charAt(i - 2) === b.charAt(j - 1)) {
          d[i][j] = Math.min(d[i][j], d[i - 2][j - 2] + cost);
        }
      }
    }
    return d[m][n];
  }

  /* Short words tolerate fewer edits, otherwise "Hat" would absorb half the alphabet. */
  function maxEdits(length) {
    if (length < 5) return 1;
    if (length < 9) return 2;
    return 3;
  }

  /*
   * Resolve one submitted value against a column's vocabulary.
   * status: exact | normalized | corrected | ambiguous | unknown
   */
  function matchVocabulary(rawValue, allowed) {
    const key = vocabKey(rawValue);
    if (!key) return { status: 'unknown', value: null };

    for (let i = 0; i < allowed.length; i++) {
      if (vocabKey(allowed[i]) === key) {
        return {
          status: String(rawValue).trim() === allowed[i] ? 'exact' : 'normalized',
          value: allowed[i],
        };
      }
    }

    let best = null;
    let bestDistance = Infinity;
    let tied = false;

    for (let i = 0; i < allowed.length; i++) {
      const distance = editDistance(key, vocabKey(allowed[i]));
      if (distance < bestDistance) {
        bestDistance = distance;
        best = allowed[i];
        tied = false;
      } else if (distance === bestDistance) {
        tied = true;
      }
    }

    if (best && bestDistance <= maxEdits(Math.max(key.length, vocabKey(best).length))) {
      // A tie means two vocabulary entries are equally plausible, so guessing would be a coin flip.
      return tied ? { status: 'ambiguous', value: null } : { status: 'corrected', value: best, distance: bestDistance };
    }
    return { status: 'unknown', value: null };
  }

  /*
   * Validate a whole cell, which may carry several comma or semicolon separated values.
   * Confident fixes are applied; anything unrecognised is preserved verbatim and reported.
   */
  function validateAttributeCell(rawValue, allowed) {
    const parts = String(rawValue).split(MULTI_VALUE_SPLIT)
      .map(function (part) { return normalizeWhitespace(part); })
      .filter(function (part) { return part !== ''; });

    if (!parts.length) return { value: null, corrected: [], unresolved: [] };

    const resolved = [];
    const corrected = [];
    const unresolved = [];

    parts.forEach(function (part) {
      const match = matchVocabulary(part, allowed);
      if (match.status === 'exact') {
        resolved.push(match.value);
      } else if (match.status === 'normalized' || match.status === 'corrected') {
        resolved.push(match.value);
        corrected.push(part + '" to "' + match.value);
      } else {
        resolved.push(part);
        unresolved.push(part);
      }
    });

    return { value: resolved.join(MULTI_VALUE_JOIN), corrected: corrected, unresolved: unresolved };
  }

  /* A cell holding only slashes and spaces is a placeholder for "nothing", not data. */
  function isSlashPlaceholder(value) {
    if (typeof value !== 'string') return false;
    const trimmed = value.trim();
    if (trimmed === '') return false;
    return /^[\/\s\\|-]+$/.test(trimmed) && trimmed.indexOf('/') !== -1;
  }

  function detectHeaderRow(worksheet, maxScanRows) {
    const limit = maxScanRows || 12;
    let best = null;

    for (let rowNumber = 1; rowNumber <= Math.min(limit, worksheet.rowCount || limit); rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      if (!row) continue;

      const mapping = {};   // canonical index -> source column number
      const unmatched = [];
      let matched = 0;
      const cellCount = Math.max(row.cellCount || 0, row.actualCellCount || 0);
      const scanWidth = Math.max(cellCount, CANONICAL_HEADERS.length);

      for (let c = 1; c <= scanWidth; c++) {
        const raw = unwrapCellValue(row.getCell(c).value);
        if (isBlank(raw)) continue;
        const key = normalizeHeaderKey(raw);
        if (!key) continue;
        const canonicalIndex = CANONICAL_BY_KEY[key];
        if (canonicalIndex !== undefined) {
          if (mapping[canonicalIndex] === undefined) {
            mapping[canonicalIndex] = c;
            matched++;
          }
        } else {
          unmatched.push(normalizeWhitespace(String(raw)));
        }
      }

      if (!best || matched > best.matched) {
        best = { rowNumber: rowNumber, mapping: mapping, matched: matched, unmatched: unmatched };
      }
      if (matched === CANONICAL_HEADERS.length) break;
    }

    return best;
  }

  function findHeaderInWorkbook(workbook) {
    let fallback = null;
    let result = null;

    workbook.eachSheet(function (worksheet) {
      if (result) return;
      if (worksheet.state === 'veryHidden') return;
      const detected = detectHeaderRow(worksheet);
      if (!detected) return;
      if (detected.matched >= Math.ceil(CANONICAL_HEADERS.length * 0.6)) {
        result = { worksheet: worksheet, header: detected };
      } else if (!fallback || detected.matched > fallback.header.matched) {
        fallback = { worksheet: worksheet, header: detected };
      }
    });

    return result || fallback;
  }

  function makeIssue(fields) {
    return {
      severity: fields.severity,
      file: fields.file,
      sheet: fields.sheet || '',
      sourceRow: fields.sourceRow == null ? '' : fields.sourceRow,
      mergedRow: fields.mergedRow == null ? '' : fields.mergedRow,
      column: fields.column || '',
      header: fields.header || '',
      rule: fields.rule,
      original: fields.original == null ? '' : String(fields.original),
      corrected: fields.corrected == null ? '' : String(fields.corrected),
      message: fields.message || '',
    };
  }

  /*
   * Read one workbook into canonical rows, applying every correction rule.
   * `startMergedRow` is the 1-based data position this file's first row will occupy.
   */
  async function processWorkbook(arrayBuffer, fileName, startMergedRow, options) {
    const opts = options || {};
    const issues = [];
    const rows = [];

    const workbook = new ExcelJS.Workbook();
    try {
      await workbook.xlsx.load(arrayBuffer);
    } catch (err) {
      issues.push(makeIssue({
        severity: 'error', file: fileName, rule: 'File could not be read',
        message: 'ExcelJS could not open this file (' + (err && err.message ? err.message : 'unknown error') +
          '). Only .xlsx and .xlsm are supported - re-save .xls or .csv files as .xlsx.',
      }));
      return { rows: rows, issues: issues, ok: false, sheetName: '', headerRow: 0 };
    }

    const found = findHeaderInWorkbook(workbook);
    if (!found || found.header.matched === 0) {
      issues.push(makeIssue({
        severity: 'error', file: fileName, rule: 'Header row not found',
        message: 'No sheet in this file contains the expected column headers, so it was skipped.',
      }));
      return { rows: rows, issues: issues, ok: false, sheetName: '', headerRow: 0 };
    }

    const worksheet = found.worksheet;
    const mapping = found.header.mapping;
    const headerRowNumber = found.header.rowNumber;

    const missing = [];
    for (let i = 0; i < CANONICAL_HEADERS.length; i++) {
      if (mapping[i] === undefined) missing.push(CANONICAL_HEADERS[i]);
    }
    if (missing.length) {
      issues.push(makeIssue({
        severity: 'error', file: fileName, sheet: worksheet.name, sourceRow: headerRowNumber,
        rule: 'Missing columns',
        message: 'These expected columns are not present and were left blank in the merged output: ' + missing.join(', ') + '.',
      }));
    }
    if (found.header.unmatched.length) {
      issues.push(makeIssue({
        severity: 'review', file: fileName, sheet: worksheet.name, sourceRow: headerRowNumber,
        rule: 'Unrecognised columns',
        message: 'These extra columns were found and were NOT carried into the merged output: ' +
          found.header.unmatched.join(', ') + '.',
      }));
    }

    let mergedRow = startMergedRow;

    for (let rowNumber = headerRowNumber + 1; rowNumber <= worksheet.rowCount; rowNumber++) {
      const sourceRow = worksheet.getRow(rowNumber);
      if (!sourceRow) continue;

      const values = new Array(CANONICAL_HEADERS.length).fill(null);
      let hasAnyValue = false;

      for (let i = 0; i < CANONICAL_HEADERS.length; i++) {
        const sourceCol = mapping[i];
        if (sourceCol === undefined) continue;
        const raw = unwrapCellValue(sourceRow.getCell(sourceCol).value);
        values[i] = raw;
        if (!isBlank(raw)) hasAnyValue = true;
      }

      if (!hasAnyValue) continue;

      // Row 1 of the merged sheet is the header, so data row N lands on Excel row N+1.
      const excelRow = mergedRow + 1;

      // Rule 1: whitespace hygiene on every text cell.
      for (let i = 0; i < values.length; i++) {
        if (typeof values[i] !== 'string') continue;
        const cleaned = normalizeWhitespace(values[i]);
        if (cleaned !== values[i]) {
          issues.push(makeIssue({
            severity: 'fixed', file: fileName, sheet: worksheet.name, sourceRow: rowNumber, mergedRow: excelRow,
            column: columnLetter(i), header: CANONICAL_HEADERS[i], rule: 'Whitespace cleaned',
            original: values[i], corrected: cleaned,
            message: 'Trimmed surrounding spaces and collapsed repeated spaces.',
          }));
        }
        values[i] = cleaned === '' ? null : cleaned;
      }

      // Rule 2: Birthdate (column H) is always YYYY/MM.
      const birthRaw = values[COL.BIRTHDATE];
      if (!isBlank(birthRaw)) {
        const parsed = parseBirthdate(birthRaw, opts);
        const originalText = birthRaw instanceof Date ? birthRaw.toISOString().slice(0, 10) : String(birthRaw);

        if (parsed.value != null) {
          values[COL.BIRTHDATE] = parsed.value;
          if (parsed.value !== originalText) {
            issues.push(makeIssue({
              severity: parsed.status === 'review' ? 'review' : 'fixed',
              file: fileName, sheet: worksheet.name, sourceRow: rowNumber, mergedRow: excelRow,
              column: 'H', header: 'Birthdate', rule: 'Birthdate reformatted to YYYY/MM',
              original: originalText, corrected: parsed.value, message: parsed.message,
            }));
          }
        } else {
          // Nothing safe to write, so the original is preserved rather than guessed at.
          values[COL.BIRTHDATE] = originalText;
          issues.push(makeIssue({
            severity: 'review', file: fileName, sheet: worksheet.name, sourceRow: rowNumber, mergedRow: excelRow,
            column: 'H', header: 'Birthdate', rule: 'Birthdate needs a manual fix',
            original: originalText, corrected: originalText,
            message: parsed.message + ' The original value was kept so nothing is lost - please correct it at source.',
          }));
        }
      }

      // Rule 3: column P holds a slash placeholder -> clear it. Real content is left untouched.
      const othersValue = values[COL.OTHERS];
      if (isSlashPlaceholder(othersValue)) {
        issues.push(makeIssue({
          severity: 'fixed', file: fileName, sheet: worksheet.name, sourceRow: rowNumber, mergedRow: excelRow,
          column: 'P', header: 'Others', rule: 'Slash placeholder cleared',
          original: othersValue, corrected: '',
          message: 'The cell contained only a slash placeholder, so it was emptied.',
        }));
        values[COL.OTHERS] = null;
      }

      // Rule 4: M-P accept only the client's vocabulary. Runs after rule 3 so a cleared P is skipped.
      Object.keys(ATTRIBUTE_VOCABULARY).forEach(function (key) {
        const index = Number(key);
        if (isBlank(values[index])) return;

        const allowed = ATTRIBUTE_VOCABULARY[index];
        const original = String(values[index]);
        const outcome = validateAttributeCell(original, allowed);
        values[index] = outcome.value;

        if (outcome.unresolved.length) {
          const fixedNote = outcome.corrected.length
            ? ' Also corrected "' + outcome.corrected.join('", "') + '".'
            : '';
          issues.push(makeIssue({
            severity: 'review', file: fileName, sheet: worksheet.name, sourceRow: rowNumber, mergedRow: excelRow,
            column: columnLetter(index), header: CANONICAL_HEADERS[index], rule: 'Value not on the accepted list',
            original: original, corrected: outcome.value,
            message: 'Not recognised: "' + outcome.unresolved.join('", "') + '". Kept as submitted for a human to fix.' +
              fixedNote + ' Accepted values are: ' + allowed.join(', ') + '.',
          }));
        } else if (outcome.corrected.length) {
          issues.push(makeIssue({
            severity: 'fixed', file: fileName, sheet: worksheet.name, sourceRow: rowNumber, mergedRow: excelRow,
            column: columnLetter(index), header: CANONICAL_HEADERS[index], rule: 'Value matched to the accepted list',
            original: original, corrected: outcome.value,
            message: 'Corrected "' + outcome.corrected.join('", "') + '".',
          }));
        }
      });

      // Rule 5: Head attribute block M-P drives column Q. Runs last so cleared cells count as empty.
      let headAttributeFilled = false;
      const filledHeadColumns = [];
      for (let i = HEAD_ATTRIBUTE_START; i <= HEAD_ATTRIBUTE_END; i++) {
        if (!isBlank(values[i])) {
          headAttributeFilled = true;
          filledHeadColumns.push(columnLetter(i));
        }
      }

      const noneValue = values[COL.NONE];
      if (headAttributeFilled) {
        if (!isBlank(noneValue)) {
          issues.push(makeIssue({
            severity: 'fixed', file: fileName, sheet: worksheet.name, sourceRow: rowNumber, mergedRow: excelRow,
            column: 'Q', header: 'None', rule: 'Head attribute conflict resolved',
            original: noneValue, corrected: '',
            message: 'Column ' + filledHeadColumns.join(', ') + ' already has head attribute data, so column Q was emptied.',
          }));
          values[COL.NONE] = null;
        }
      } else if (noneValue !== NONE_APPLY_TEXT) {
        issues.push(makeIssue({
          severity: 'fixed', file: fileName, sheet: worksheet.name, sourceRow: rowNumber, mergedRow: excelRow,
          column: 'Q', header: 'None', rule: 'Head attribute placeholder added',
          original: noneValue, corrected: NONE_APPLY_TEXT,
          message: 'Columns M-P are all empty, so column Q was set to "' + NONE_APPLY_TEXT + '".',
        }));
        values[COL.NONE] = NONE_APPLY_TEXT;
      }

      rows.push({ values: values, file: fileName, sheet: worksheet.name, sourceRow: rowNumber, mergedRow: excelRow });
      mergedRow++;
    }

    return { rows: rows, issues: issues, ok: true, sheetName: worksheet.name, headerRow: headerRowNumber };
  }

  function findDuplicates(rows) {
    const issues = [];
    const seen = {};

    rows.forEach(function (row) {
      const email = row.values[COL.PARTICIPANT_EMAIL];
      const name = row.values[COL.PARTICIPANT_NAME];
      const key = isBlank(email)
        ? (isBlank(name) ? null : 'name:' + String(name).toLowerCase())
        : 'email:' + String(email).toLowerCase();
      if (!key) return;

      if (seen[key]) {
        const first = seen[key];
        issues.push(makeIssue({
          severity: 'review', file: row.file, sheet: row.sheet, sourceRow: row.sourceRow, mergedRow: row.mergedRow,
          column: isBlank(email) ? 'A' : 'B',
          header: isBlank(email) ? 'Participant Name As Per ICF' : 'Participant Email',
          rule: 'Possible duplicate participant',
          original: String(email || name),
          corrected: String(email || name),
          message: 'Also appears in ' + first.file + ' at source row ' + first.sourceRow +
            ' (merged row ' + first.mergedRow + '). Both rows were kept - remove one if this is a genuine duplicate.',
        }));
      } else {
        seen[key] = row;
      }
    });

    return issues;
  }

  const SEVERITY_ORDER = { error: 0, review: 1, fixed: 2 };

  function summarize(issues) {
    const counts = { error: 0, review: 0, fixed: 0 };
    issues.forEach(function (issue) {
      if (counts[issue.severity] !== undefined) counts[issue.severity]++;
    });
    return counts;
  }

  /* Merge every supplied file. `files` is [{ name, buffer, source }]. */
  async function mergeFiles(files, options) {
    const opts = options || {};
    const allRows = [];
    let allIssues = [];
    const fileReports = [];

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      if (opts.onProgress) opts.onProgress(i, files.length, file.name);

      const startMergedRow = allRows.length + 1;
      const result = await processWorkbook(file.buffer, file.name, startMergedRow, opts);

      allRows.push.apply(allRows, result.rows);
      allIssues = allIssues.concat(result.issues);

      fileReports.push({
        name: file.name,
        source: file.source || 'Upload',
        ok: result.ok,
        sheetName: result.sheetName,
        headerRow: result.headerRow,
        rowCount: result.rows.length,
        firstMergedRow: result.rows.length ? startMergedRow + 1 : null,
        lastMergedRow: result.rows.length ? startMergedRow + result.rows.length : null,
        issueCount: result.issues.length,
      });
    }

    allIssues = allIssues.concat(findDuplicates(allRows));
    allIssues.sort(function (a, b) {
      const bySeverity = SEVERITY_ORDER[a.severity] - SEVERITY_ORDER[b.severity];
      if (bySeverity !== 0) return bySeverity;
      return (a.mergedRow || 0) - (b.mergedRow || 0);
    });

    return { rows: allRows, issues: allIssues, fileReports: fileReports, counts: summarize(allIssues) };
  }

  const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } };
  const HEADER_FONT = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };

  function styleHeaderRow(row) {
    row.font = HEADER_FONT;
    row.fill = HEADER_FILL;
    row.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    row.height = 30;
  }

  function autoWidth(worksheet, maxWidth) {
    const cap = maxWidth || 40;
    worksheet.columns.forEach(function (column) {
      let widest = 10;
      column.eachCell({ includeEmpty: false }, function (cell) {
        const length = cell.value == null ? 0 : String(cell.value).length;
        if (length > widest) widest = length;
      });
      column.width = Math.min(cap, widest + 2);
    });
  }

  const REVIEW_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE699' } };

  /*
   * Amber a cell whenever a human still needs to look at it, and attach the reason as a
   * comment. A clean run leaves no highlights at all, so any colour means "not ready to send".
   */
  function highlightReviewCells(sheet, issues) {
    issues.forEach(function (issue) {
      if (issue.severity !== 'review') return;
      if (!issue.mergedRow || !issue.header) return;

      const columnIndex = CANONICAL_HEADERS.indexOf(issue.header);
      if (columnIndex < 0) return;

      const cell = sheet.getCell(issue.mergedRow, columnIndex + 1);
      cell.fill = REVIEW_FILL;
      cell.note = issue.rule + ': ' + issue.message;
    });
  }

  function buildMergedWorkbook(mergeResult, options) {
    const opts = options || {};
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Agency Excel Merger';
    workbook.created = opts.now || new Date();

    const sheet = workbook.addWorksheet('Merged Data', {
      views: [{ state: 'frozen', ySplit: 1 }],
    });

    sheet.addRow(CANONICAL_HEADERS);
    styleHeaderRow(sheet.getRow(1));

    mergeResult.rows.forEach(function (row) {
      sheet.addRow(row.values.map(function (v) { return v == null ? null : v; }));
    });

    // Keep Birthdate as literal text so Excel cannot re-interpret "1990/05" on open.
    sheet.getColumn(COL.BIRTHDATE + 1).numFmt = '@';

    highlightReviewCells(sheet, mergeResult.issues);

    sheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: CANONICAL_HEADERS.length },
    };
    autoWidth(sheet);

    if (opts.includeQaSheet) addQaSheet(workbook, mergeResult);

    return workbook;
  }

  const QA_HEADERS = ['Severity', 'Source File', 'Sheet', 'Source Row', 'Merged Row', 'Column', 'Header', 'Check', 'Original Value', 'Corrected Value', 'Details'];

  const SEVERITY_LABEL = { error: 'ERROR', review: 'NEEDS REVIEW', fixed: 'AUTO-FIXED' };
  const SEVERITY_COLOR = { error: 'FFF8CBAD', review: 'FFFFE699', fixed: 'FFD9EAD3' };

  function addQaSheet(workbook, mergeResult) {
    const sheet = workbook.addWorksheet('QA Report', { views: [{ state: 'frozen', ySplit: 1 }] });

    sheet.addRow(QA_HEADERS);
    styleHeaderRow(sheet.getRow(1));

    mergeResult.issues.forEach(function (issue) {
      const row = sheet.addRow([
        SEVERITY_LABEL[issue.severity] || issue.severity,
        issue.file, issue.sheet, issue.sourceRow, issue.mergedRow,
        issue.column, issue.header, issue.rule, issue.original, issue.corrected, issue.message,
      ]);
      const color = SEVERITY_COLOR[issue.severity];
      if (color) {
        row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
      }
      row.alignment = { vertical: 'top', wrapText: true };
    });

    sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: QA_HEADERS.length } };
    autoWidth(sheet, 60);

    const sources = workbook.addWorksheet('Sources');
    sources.addRow(['Source File', 'Origin', 'Sheet Used', 'Header Row', 'Rows Merged', 'Merged Rows From', 'Merged Rows To', 'Status']);
    styleHeaderRow(sources.getRow(1));
    mergeResult.fileReports.forEach(function (report) {
      sources.addRow([
        report.name, report.source, report.sheetName, report.headerRow || '', report.rowCount,
        report.firstMergedRow == null ? '' : report.firstMergedRow,
        report.lastMergedRow == null ? '' : report.lastMergedRow,
        report.ok ? 'Merged' : 'Skipped',
      ]);
    });
    autoWidth(sources);
  }

  function buildQaWorkbook(mergeResult) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Agency Excel Merger';
    addQaSheet(workbook, mergeResult);
    return workbook;
  }

  global.MergerCore = {
    CANONICAL_HEADERS: CANONICAL_HEADERS,
    COL: COL,
    NONE_APPLY_TEXT: NONE_APPLY_TEXT,
    columnLetter: columnLetter,
    normalizeWhitespace: normalizeWhitespace,
    editDistance: editDistance,
    ATTRIBUTE_VOCABULARY: ATTRIBUTE_VOCABULARY,
    vocabKey: vocabKey,
    matchVocabulary: matchVocabulary,
    parseBirthdate: parseBirthdate,
    isSlashPlaceholder: isSlashPlaceholder,
    processWorkbook: processWorkbook,
    mergeFiles: mergeFiles,
    buildMergedWorkbook: buildMergedWorkbook,
    buildQaWorkbook: buildQaWorkbook,
    SEVERITY_LABEL: SEVERITY_LABEL,
  };
})(window);

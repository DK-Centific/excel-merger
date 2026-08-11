/*
 * The delivery pipeline.
 *
 * Folders -> every workbook inside them -> one merged sheet in the client's 30-column
 * layout, QA'd, with the Dropbox links mapped in.
 *
 * QA findings never touch columns B and C: the delivered sheet keeps the client's exact
 * layout, and the analysis rides along as extra sheets at the end of the workbook.
 */
(function (global) {
  'use strict';

  const CORE = global.MergerCore;
  const META = global.MetadataMerge;
  const LF = global.LinkFiller;

  const OUT = META.OUT;

  /* Vocabulary group -> its column in the delivery layout. */
  const ATTRIBUTE_COLUMNS = [
    { vocabKey: 12, column: OUT.HEAD_AND_HAIR },
    { vocabKey: 13, column: OUT.FACIAL_FEATURES },
    { vocabKey: 14, column: OUT.ACCESSORIES },
    { vocabKey: 15, column: OUT.OTHERS },
  ];

  const EXCEL_NAME = /\.xls[xm]$/i;
  const TEMP_NAME = /^~\$/;

  function isMetadataWorkbook(file) {
    const name = file.name || '';
    return EXCEL_NAME.test(name) && !TEMP_NAME.test(name);
  }

  function columnLetterOf(index) {
    return CORE.columnLetter(index);
  }

  function issue(fields) {
    return {
      severity: fields.severity,
      source: fields.source || '',
      row: fields.row || null,
      column: fields.column || '',
      header: fields.header || '',
      rule: fields.rule,
      original: fields.original == null ? '' : String(fields.original),
      corrected: fields.corrected == null ? '' : String(fields.corrected),
      message: fields.message || '',
    };
  }

  /* ---------- QA over the merged rows ---------- */

  /*
   * Applied in order, because later rules depend on earlier ones: a placeholder cleared in
   * Others has to count as empty when deciding column V.
   */
  function runQa(rows, options) {
    const opts = options || {};
    const issues = [];

    rows.forEach(function (row, index) {
      const values = row.values;
      const excelRow = index + 2; // header occupies row 1
      const where = row.source || '';

      // 1. Whitespace: invisible, and it breaks every comparison downstream.
      values.forEach(function (value, i) {
        if (typeof value !== 'string') return;
        const tidy = CORE.normalizeWhitespace(value);
        if (tidy !== value) {
          values[i] = tidy === '' ? null : tidy;
          issues.push(issue({
            severity: 'fixed', source: where, row: excelRow,
            column: columnLetterOf(i), header: META.OUTPUT_HEADERS[i],
            rule: 'Whitespace tidied', original: value, corrected: values[i],
            message: 'Leading, trailing or repeated spaces removed.',
          }));
        }
      });

      // 2. Birthdate -> YYYY/MM, kept as text so Excel cannot reinterpret it.
      const rawBirth = values[OUT.BIRTHDATE];
      if (rawBirth != null && String(rawBirth).trim() !== '') {
        const parsed = CORE.parseBirthdate(rawBirth, { dayFirst: opts.dayFirst !== false });
        if (parsed.status === 'unparseable') {
          issues.push(issue({
            severity: 'review', source: where, row: excelRow,
            column: columnLetterOf(OUT.BIRTHDATE), header: 'Birthdate',
            rule: 'Birthdate needs a manual fix', original: rawBirth, corrected: rawBirth,
            message: parsed.message + ' Left exactly as submitted.',
          }));
        } else if (parsed.value != null) {
          const changed = String(rawBirth) !== parsed.value;
          values[OUT.BIRTHDATE] = parsed.value;
          if (changed || parsed.status === 'review') {
            issues.push(issue({
              severity: parsed.status === 'review' ? 'review' : 'fixed',
              source: where, row: excelRow,
              column: columnLetterOf(OUT.BIRTHDATE), header: 'Birthdate',
              rule: parsed.status === 'review' ? 'Birthdate reformatted, please confirm' : 'Birthdate reformatted to YYYY/MM',
              original: rawBirth, corrected: parsed.value, message: parsed.message,
            }));
          }
        }
      }

      // 3. A cell holding only slashes is a placeholder for nothing, not data.
      const others = values[OUT.OTHERS];
      if (others != null && CORE.isSlashPlaceholder(others)) {
        values[OUT.OTHERS] = null;
        issues.push(issue({
          severity: 'fixed', source: where, row: excelRow,
          column: columnLetterOf(OUT.OTHERS), header: 'Others',
          rule: 'Slash placeholder cleared', original: others, corrected: '',
          message: 'The cell held only a slash, so it was emptied.',
        }));
      }

      // 4. The attribute columns accept only the client's vocabulary.
      ATTRIBUTE_COLUMNS.forEach(function (group) {
        const raw = values[group.column];
        if (raw == null || String(raw).trim() === '') return;

        const allowed = CORE.ATTRIBUTE_VOCABULARY[group.vocabKey];
        const parts = String(raw).split(/[;,]/)
          .map(function (part) { return CORE.normalizeWhitespace(part); })
          .filter(function (part) { return part !== ''; });

        const resolved = [];
        const corrected = [];
        const unresolved = [];

        parts.forEach(function (part) {
          const match = CORE.matchVocabulary(part, allowed);
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

        values[group.column] = resolved.join(', ');
        if (unresolved.length) {
          issues.push(issue({
            severity: 'review', source: where, row: excelRow,
            column: columnLetterOf(group.column), header: META.OUTPUT_HEADERS[group.column],
            rule: 'Value not on the accepted list', original: raw, corrected: values[group.column],
            message: 'Not recognised: "' + unresolved.join('", "') + '". Kept as submitted. ' +
              'Accepted values are: ' + allowed.join(', ') + '.',
          }));
        } else if (corrected.length) {
          issues.push(issue({
            severity: 'fixed', source: where, row: excelRow,
            column: columnLetterOf(group.column), header: META.OUTPUT_HEADERS[group.column],
            rule: 'Value matched to the accepted list', original: raw, corrected: values[group.column],
            message: 'Corrected "' + corrected.join('", "') + '".',
          }));
        }
      });

      // 5. Column V follows the attribute block. Runs last so a cleared cell counts as empty.
      const anyAttribute = ATTRIBUTE_COLUMNS.some(function (group) {
        const value = values[group.column];
        return value != null && String(value).trim() !== '';
      });
      const none = values[OUT.NONE];
      if (anyAttribute && none != null && String(none).trim() !== '') {
        values[OUT.NONE] = null;
        issues.push(issue({
          severity: 'fixed', source: where, row: excelRow,
          column: columnLetterOf(OUT.NONE), header: 'None',
          rule: 'None cleared', original: none, corrected: '',
          message: 'An attribute is present, so None must be empty.',
        }));
      } else if (!anyAttribute && String(none == null ? '' : none).trim() !== CORE.NONE_APPLY_TEXT) {
        const before = none;
        values[OUT.NONE] = CORE.NONE_APPLY_TEXT;
        issues.push(issue({
          severity: 'fixed', source: where, row: excelRow,
          column: columnLetterOf(OUT.NONE), header: 'None',
          rule: 'None set', original: before, corrected: CORE.NONE_APPLY_TEXT,
          message: 'No attributes are present, so None is "' + CORE.NONE_APPLY_TEXT + '".',
        }));
      }
    });

    return issues;
  }

  /* ---------- participant identity ---------- */

  /*
   * A participant is identified by their folder when the media sits in per-participant
   * folders, and otherwise by the participant name appearing in the filename. Agencies use
   * both layouts, so both are supported.
   */
  function participantKeysForFile(file, roots, knownKeys) {
    const path = file.path_display || file.path_lower || file.name;
    const keys = [];

    for (let i = 0; i < roots.length; i++) {
      const root = String(roots[i] || '').replace(/\/+$/, '');
      let rest = path;
      if (root && path.toLowerCase().indexOf(root.toLowerCase() + '/') === 0) {
        rest = path.slice(root.length + 1);
      } else {
        continue;
      }
      const segments = rest.split('/').filter(Boolean);
      // Every folder level between the root and the file is a candidate identity.
      for (let s = 0; s < segments.length - 1; s++) {
        const key = LF.normalizeName(segments[s]);
        if (key && keys.indexOf(key) === -1) keys.push(key);
      }
      break;
    }

    // Fall back to the filename when the media is flat.
    const base = LF.normalizeName(String(file.name || '').replace(/\.[^.]+$/, ''));
    if (base) {
      knownKeys.forEach(function (known) {
        if (known && base.indexOf(known) !== -1 && keys.indexOf(known) === -1) keys.push(known);
      });
    }
    return keys;
  }

  /*
   * Fill columns A, B and C. These never come from the agency files — they are the tool's
   * own output, written after QA and link mapping so each row carries its own verdict.
   */
  const QA_OK = 'OK';
  const QA_FIXED = 'Auto-fixed';
  const QA_REVIEW = 'Review';
  const QA_ERROR = 'Error';

  function annotate(rows, issues, planRows) {
    // A: one number per participant, repeated across that participant's video rows.
    const sequence = {};
    let next = 0;
    rows.forEach(function (row) {
      const key = LF.normalizeName(row.values[OUT.FOLDER]) || LF.normalizeName(row.values[OUT.NAME]);
      if (!key) { row.values[OUT.SEQUENCE] = null; return; }
      if (sequence[key] == null) { next += 1; sequence[key] = next; }
      row.values[OUT.SEQUENCE] = sequence[key];
    });

    const byRow = {};
    (issues || []).forEach(function (item) {
      if (!item.row) return; // file-level, not attributable to a row
      if (!byRow[item.row]) byRow[item.row] = [];
      byRow[item.row].push(item);
    });

    rows.forEach(function (row, index) {
      const excelRow = index + 2;
      const found = byRow[excelRow] || [];
      const plan = (planRows && planRows[index]) || null;
      const problems = (plan && plan.problems) || [];

      let verdict = QA_OK;
      if (found.some(function (i) { return i.severity === 'error'; })) verdict = QA_ERROR;
      else if (found.some(function (i) { return i.severity === 'review'; }) || problems.length) verdict = QA_REVIEW;
      else if (found.length) verdict = QA_FIXED;

      const comments = found.map(function (item) {
        return item.rule + (item.message ? ': ' + item.message : '');
      }).concat(problems);

      row.values[OUT.QA_RESULT] = verdict;
      row.values[OUT.COMMENTS] = comments.length ? comments.join(' | ') : null;
    });
  }

  global.DeliveryMerge = {
    QA_OK: QA_OK,
    QA_FIXED: QA_FIXED,
    QA_REVIEW: QA_REVIEW,
    QA_ERROR: QA_ERROR,
    annotate: annotate,
    ATTRIBUTE_COLUMNS: ATTRIBUTE_COLUMNS,
    isMetadataWorkbook: isMetadataWorkbook,
    runQa: runQa,
    participantKeysForFile: participantKeysForFile,
  };
})(window);

/*
 * Agency metadata -> delivery sheet.
 *
 * Reads whatever metadata workbooks turn up in the Dropbox tree and produces rows in the
 * 30-column delivery shape (the "Merged Example" layout).
 *
 * Everything is matched by HEADER NAME, never by column position, because the agency
 * layouts vary — sometimes one workbook per participant folder, sometimes one workbook
 * listing every participant.
 *
 * The attribute block is recomputed from the "x" marks rather than read out of the
 * consolidation columns. That is deliberate: those cells hold formulas
 * (=IFERROR(LEFT(IF($M2="x",$M$1&", ",...)) whose cached values may be absent or stale, and
 * a formula copied into another workbook would break. Recomputing is the paste-special.
 */
(function (global) {
  'use strict';

  const CORE = global.MergerCore;

  /* The delivery layout, A..AD. Column S carries a leading space in the client's file. */
  const OUTPUT_HEADERS = [
    'participant_sequence',            // A
    'QA Result',                       // B
    'Comments',                        // C
    'Dropbox Folder Name',             // D
    'Dropbox URL',                     // E  <- link filler
    'Participant Name As Per ICF',     // F
    'Participant Email',               // G
    'Name of Parents/Legal Guardian',  // H
    'Email of Parents/Legal Guardian', // I
    'Expressions',                     // J
    'For Non Neutral, Please Select',  // K
    'Age Group',                       // L
    'Birthdate',                       // M
    'Gender',                          // N
    'Main Ethnicity',                  // O
    'Secondary Ethnicity',             // P
    'Environment',                     // Q
    'Head and Hair',                   // R
    ' Facial Features',                // S
    'Accessories and jewellery',       // T
    'Others',                          // U
    'None',                            // V
    'Skintone',                        // W
    'Device',                          // X
    'Specific Device Model',           // Y
    'Date of Recording',               // Z
    'Country of Collection',           // AA
    'State Abbreviation',              // AB
    'ICF URL',                         // AC
    'Assent URL',                      // AD
  ];

  const OUT = {
    SEQUENCE: 0, QA_RESULT: 1, COMMENTS: 2, FOLDER: 3, VIDEO_URL: 4,
    NAME: 5, EMAIL: 6, GUARDIAN: 7, GUARDIAN_EMAIL: 8,
    EXPRESSION: 9, NON_NEUTRAL: 10, AGE: 11, BIRTHDATE: 12, GENDER: 13,
    ETHNICITY1: 14, ETHNICITY2: 15, ENVIRONMENT: 16,
    HEAD_AND_HAIR: 17, FACIAL_FEATURES: 18, ACCESSORIES: 19, OTHERS: 20, NONE: 21,
    SKINTONE: 22, DEVICE: 23, DEVICE_MODEL: 24, RECORDED: 25,
    COUNTRY: 26, STATE: 27, ICF_URL: 28, ASSENT_URL: 29,
  };

  /* The four consolidation groups, in delivery-column order. */
  const GROUP_COLUMNS = [OUT.HEAD_AND_HAIR, OUT.FACIAL_FEATURES, OUT.ACCESSORIES, OUT.OTHERS];
  const GROUP_KEYS = [12, 13, 14, 15]; // matches MergerCore.ATTRIBUTE_VOCABULARY keys (M,N,O,P)

  const NONE_APPLY = 'N/A-none apply';

  /*
   * Source headers we know how to carry across. Header text drifts between agency
   * templates, so each output column lists every spelling seen.
   */
  const SOURCE_ALIASES = {};
  SOURCE_ALIASES[OUT.NAME] = ['participant name as per icf', 'participant name', 'name as per icf'];
  SOURCE_ALIASES[OUT.EMAIL] = ['participant email'];
  SOURCE_ALIASES[OUT.GUARDIAN] = ['name of parents legal guardian', 'name of parents/legal guardian', 'guardian name'];
  SOURCE_ALIASES[OUT.GUARDIAN_EMAIL] = ['email of parents legal guardian', 'email of parents/legal guardian', 'guardian email'];
  SOURCE_ALIASES[OUT.EXPRESSION] = ['expressions', 'expression'];
  SOURCE_ALIASES[OUT.NON_NEUTRAL] = ['for non neutral please select', 'for non neutral, please select', 'non neutral'];
  SOURCE_ALIASES[OUT.AGE] = ['age group'];
  SOURCE_ALIASES[OUT.BIRTHDATE] = ['birthdate', 'birth date', 'date of birth'];
  SOURCE_ALIASES[OUT.GENDER] = ['gender'];
  SOURCE_ALIASES[OUT.ETHNICITY1] = ['main ethnicity'];
  SOURCE_ALIASES[OUT.ETHNICITY2] = ['secondary ethnicity'];
  SOURCE_ALIASES[OUT.ENVIRONMENT] = ['environment'];
  SOURCE_ALIASES[OUT.SKINTONE] = ['skintone', 'skin tone'];
  SOURCE_ALIASES[OUT.DEVICE] = ['device'];
  SOURCE_ALIASES[OUT.DEVICE_MODEL] = ['specific device model', 'device model'];
  SOURCE_ALIASES[OUT.RECORDED] = ['date of recording', 'recording date'];
  SOURCE_ALIASES[OUT.COUNTRY] = ['country of collection', 'country'];
  SOURCE_ALIASES[OUT.STATE] = ['state abbreviation', 'state'];

  /* Consolidation columns are deliberately absent: they are recomputed, never read. */
  const IGNORED_HEADERS = ['head and hair', 'facial features', 'accessories and jewellery', 'others', 'none'];

  function norm(text) {
    return String(text == null ? '' : text)
      .toLowerCase()
      .replace(/[\/\\\-_.,;:()]+/g, ' ')
      .replace(/[^a-z0-9 ]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function cellText(value) {
    if (value == null) return null;
    if (value instanceof Date) return value;
    if (typeof value === 'object') {
      // ExcelJS wraps rich text and formulas; take the computed side, never the formula.
      if (value.result !== undefined) return value.result;
      if (value.text !== undefined) return value.text;
      if (Array.isArray(value.richText)) return value.richText.map(function (r) { return r.text; }).join('');
      return null;
    }
    return value;
  }

  /* Every vocabulary term, mapped to the group it belongs to and its canonical spelling. */
  function buildAttributeIndex() {
    const index = {};
    GROUP_KEYS.forEach(function (key, groupIndex) {
      (CORE.ATTRIBUTE_VOCABULARY[key] || []).forEach(function (term) {
        index[CORE.vocabKey(term)] = { group: groupIndex, canonical: term };
      });
    });
    return index;
  }

  /*
   * Map a worksheet's header row. Returns which source column feeds each output column,
   * plus the attribute columns keyed by their canonical term.
   *
   * Attribute headers are matched through the merger's vocabulary matcher, so the client's
   * "Other tattons" resolves to the canonical "Other tattoos" without a special case.
   */
  /*
   * Headers come from a controlled template, so matching them is deliberately stricter than
   * matching cell values. The only typo worth tolerating is the client's own "Other
   * tattons"; a general fuzzy match reads a data cell "Acme" as the attribute "Acne" and
   * makes a row of invoice data look like a header row.
   */
  const MIN_FUZZY_HEADER_LENGTH = 6;

  function matchAttributeHeader(raw, attributeIndex) {
    const key = CORE.vocabKey(raw);
    if (!key) return null;

    const exact = attributeIndex[key];
    if (exact) return exact;
    if (key.length < MIN_FUZZY_HEADER_LENGTH) return null;

    for (let g = 0; g < GROUP_KEYS.length; g++) {
      const terms = CORE.ATTRIBUTE_VOCABULARY[GROUP_KEYS[g]] || [];
      for (let t = 0; t < terms.length; t++) {
        const termKey = CORE.vocabKey(terms[t]);
        if (termKey.length < MIN_FUZZY_HEADER_LENGTH) continue;
        if (CORE.editDistance(key, termKey) <= 1) return { group: g, canonical: terms[t] };
      }
    }
    return null;
  }

  function mapHeaders(worksheet, headerRowNumber) {
    const attributeIndex = buildAttributeIndex();
    const fields = {};
    const attributes = [];
    let matched = 0;

    const headerRow = worksheet.getRow(headerRowNumber);

    /*
     * Collect the populated cells directly rather than sweeping 1..columnCount.
     * ExcelJS reports columnCount as 0 for sheets that carry no explicit dimension
     * metadata — which is exactly what happens when the data is not inside an Excel
     * table — and the old sweep then examined nothing and matched no headers at all.
     */
    const cells = [];
    headerRow.eachCell({ includeEmpty: false }, function (cell, col) {
      cells.push({ col: col, raw: cellText(cell.value) });
    });
    if (!cells.length) {
      // Last resort for rows eachCell declines to walk.
      const width = Math.max(worksheet.columnCount || 0, worksheet.actualColumnCount || 0, 60);
      for (let col = 1; col <= width; col++) {
        cells.push({ col: col, raw: cellText(headerRow.getCell(col).value) });
      }
    }

    const seen = [];

    for (let c = 0; c < cells.length; c++) {
      const col = cells[c].col;
      const raw = cells[c].raw;
      const key = norm(raw);
      if (!key) continue;
      seen.push(String(raw).trim());

      if (IGNORED_HEADERS.indexOf(key) !== -1) continue;

      const attribute = matchAttributeHeader(raw, attributeIndex);
      if (attribute) {
        attributes.push({ col: col, group: attribute.group, canonical: attribute.canonical });
        matched++;
        continue;
      }

      Object.keys(SOURCE_ALIASES).forEach(function (outIndex) {
        if (fields[outIndex] != null) return;
        if (SOURCE_ALIASES[outIndex].map(norm).indexOf(key) !== -1) {
          fields[outIndex] = col;
          matched++;
        }
      });
    }

    return { fields: fields, attributes: attributes, matched: matched, seen: seen };
  }

  /* A metadata sheet is one whose header row we can actually read. */
  /*
   * The header row is not always row 1 — templates carry title and instruction rows above
   * it. Scan a generous window and take the best-matching row. rowCount is unreliable for
   * the same reason columnCount is, so it only widens the window, never narrows it.
   */
  const HEADER_SCAN_ROWS = 25;
  const MIN_HEADER_MATCHES = 3;

  function findHeaderRow(worksheet) {
    let best = null;
    const limit = Math.max(Math.min(worksheet.rowCount || 0, HEADER_SCAN_ROWS), HEADER_SCAN_ROWS);

    for (let rowNumber = 1; rowNumber <= limit; rowNumber++) {
      const map = mapHeaders(worksheet, rowNumber);
      if (!best || map.matched > best.map.matched) best = { rowNumber: rowNumber, map: map };
    }
    if (best) best.accepted = best.map.matched >= MIN_HEADER_MATCHES;
    return best && best.accepted ? best : null;
  }

  /*
   * What a workbook looks like from the outside, for when nothing matched and the user
   * needs to see why rather than being told "no recognisable headers".
   */
  function describeWorkbook(workbook) {
    return workbook.worksheets.map(function (worksheet) {
      let best = null;
      const limit = Math.max(Math.min(worksheet.rowCount || 0, HEADER_SCAN_ROWS), HEADER_SCAN_ROWS);
      for (let rowNumber = 1; rowNumber <= limit; rowNumber++) {
        const map = mapHeaders(worksheet, rowNumber);
        if (!best || map.matched > best.map.matched) best = { rowNumber: rowNumber, map: map };
      }
      return {
        sheet: worksheet.name,
        rows: worksheet.rowCount || 0,
        bestRow: best ? best.rowNumber : null,
        matched: best ? best.map.matched : 0,
        headersSeen: best ? best.map.seen.slice(0, 12) : [],
      };
    });
  }

  /*
   * Recompute the four consolidation columns from the "x" marks, then column V, exactly as
   * the template's formulas would have: join the ticked headers with ", ", and set
   * "N/A-none apply" only when all four groups are empty.
   */
  function consolidate(values, row, attributes) {
    const groups = [[], [], [], []];

    attributes.forEach(function (attribute) {
      const raw = cellText(row.getCell(attribute.col).value);
      if (raw == null) return;
      if (String(raw).trim().toLowerCase() !== 'x') return;
      groups[attribute.group].push(attribute.canonical);
    });

    let anyFilled = false;
    groups.forEach(function (terms, i) {
      values[GROUP_COLUMNS[i]] = terms.length ? terms.join(', ') : null;
      if (terms.length) anyFilled = true;
    });
    values[OUT.NONE] = anyFilled ? null : NONE_APPLY;
  }

  /*
   * Read one metadata workbook into delivery-shaped rows.
   * `folderName` is the Dropbox folder the workbook came from; it only becomes column D
   * when the workbook does not name participants itself.
   */
  function readMetadataWorkbook(workbook, folderName) {
    const out = { rows: [], sheetName: '', ok: false, reason: '' };

    for (let i = 0; i < workbook.worksheets.length; i++) {
      const worksheet = workbook.worksheets[i];
      if (/^head_attributes$/i.test(worksheet.name)) continue;

      const found = findHeaderRow(worksheet);
      if (!found) continue;

      const map = found.map;
      worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber <= found.rowNumber) return;

        const values = new Array(OUTPUT_HEADERS.length).fill(null);
        Object.keys(map.fields).forEach(function (outIndex) {
          const value = cellText(row.getCell(map.fields[outIndex]).value);
          values[Number(outIndex)] = value === '' ? null : value;
        });
        consolidate(values, row, map.attributes);

        const name = values[OUT.NAME];
        if (name == null || String(name).trim() === '') return; // blank template row

        values[OUT.FOLDER] = folderName || name;
        out.rows.push({ values: values, sourceRow: rowNumber });
      });

      out.sheetName = worksheet.name;
      out.ok = true;
      return out;
    }

    const seen = describeWorkbook(workbook)
      .map(function (sheet) {
        return sheet.sheet + ' (best row ' + sheet.bestRow + ', ' + sheet.matched + ' matched: ' +
          (sheet.headersSeen.join(' | ') || 'no text in that row') + ')';
      })
      .join('; ');
    out.reason = 'No sheet had at least ' + MIN_HEADER_MATCHES + ' recognisable headers. Saw — ' + seen;
    out.sheets = describeWorkbook(workbook);
    return out;
  }

  /* ---------- output workbook ---------- */

  const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } };
  const REVIEW_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE699' } };

  function buildDeliveryWorkbook(rows, options) {
    const opts = options || {};
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Centific Merge';
    workbook.created = opts.now || new Date();

    const sheet = workbook.addWorksheet('Merged', { views: [{ state: 'frozen', ySplit: 1 }] });
    sheet.addRow(OUTPUT_HEADERS);

    const header = sheet.getRow(1);
    header.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    header.fill = HEADER_FILL;
    header.alignment = { vertical: 'middle' };

    rows.forEach(function (row) {
      sheet.addRow(row.values);
    });

    sheet.columns.forEach(function (column, i) {
      const label = OUTPUT_HEADERS[i] || '';
      column.width = Math.min(Math.max(label.length + 4, 12), 42);
    });
    // Birthdate is text like 2012/6; let Excel leave it alone.
    sheet.getColumn(OUT.BIRTHDATE + 1).numFmt = '@';
    sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: OUTPUT_HEADERS.length } };

    if (opts.highlight) {
      opts.highlight.forEach(function (mark) {
        const cell = sheet.getCell(mark.row, mark.column);
        cell.fill = REVIEW_FILL;
        if (mark.note) cell.note = mark.note;
      });
    }

    /*
     * QA lives here, at the end of the workbook — never in columns B and C, which stay
     * exactly as the client's layout expects them.
     */
    if (opts.qa && opts.qa.length) {
      const qa = workbook.addWorksheet('QA Report');
      qa.addRow(['Severity', 'Source file', 'Row', 'Col', 'Header', 'Check', 'Original', 'Corrected', 'Details']);
      qa.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      qa.getRow(1).fill = HEADER_FILL;
      opts.qa.forEach(function (item) {
        qa.addRow([
          (item.severity || '').toUpperCase(), item.source || '', item.row || '',
          item.column || '', item.header || '', item.rule || '',
          item.original || '', item.corrected || '', item.message || '',
        ]);
      });
      [12, 30, 7, 6, 22, 30, 24, 24, 60].forEach(function (width, i) {
        qa.getColumn(i + 1).width = width;
      });
      qa.views = [{ state: 'frozen', ySplit: 1 }];
      qa.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 9 } };
    }

    /* The client's file carries this reference sheet; keep it so the output matches. */
    const reference = workbook.addWorksheet('Head_Attributes');
    const groups = GROUP_KEYS.map(function (key) { return CORE.ATTRIBUTE_VOCABULARY[key] || []; });
    reference.addRow(['Head and Hair', 'Facial Features ', 'Accessories and jewellery', 'Others']);
    reference.getRow(1).font = { bold: true };
    const depth = Math.max.apply(null, groups.map(function (g) { return g.length; }));
    for (let i = 0; i < depth; i++) {
      reference.addRow(groups.map(function (g) { return g[i] || null; }));
    }
    reference.columns.forEach(function (column) { column.width = 26; });

    return workbook;
  }

  global.MetadataMerge = {
    OUTPUT_HEADERS: OUTPUT_HEADERS,
    OUT: OUT,
    NONE_APPLY: NONE_APPLY,
    norm: norm,
    mapHeaders: mapHeaders,
    findHeaderRow: findHeaderRow,
    describeWorkbook: describeWorkbook,
    readMetadataWorkbook: readMetadataWorkbook,
    buildDeliveryWorkbook: buildDeliveryWorkbook,
  };
})(window);

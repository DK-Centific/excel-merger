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
  function mapHeaders(worksheet, headerRowNumber) {
    const attributeIndex = buildAttributeIndex();
    const fields = {};
    const attributes = [];
    let matched = 0;

    const headerRow = worksheet.getRow(headerRowNumber);
    const width = Math.max(worksheet.columnCount || 0, headerRow.cellCount || 0);

    for (let col = 1; col <= width; col++) {
      const raw = cellText(headerRow.getCell(col).value);
      const key = norm(raw);
      if (!key) continue;

      if (IGNORED_HEADERS.indexOf(key) !== -1) continue;

      const vkey = CORE.vocabKey(raw);
      let attribute = attributeIndex[vkey];
      if (!attribute) {
        // Tolerate the template's own typos, e.g. "Other tattons".
        for (let g = 0; g < GROUP_KEYS.length; g++) {
          const hit = CORE.matchVocabulary(raw, CORE.ATTRIBUTE_VOCABULARY[GROUP_KEYS[g]]);
          if (hit.status === 'corrected' || hit.status === 'normalized' || hit.status === 'exact') {
            attribute = { group: g, canonical: hit.value };
            break;
          }
        }
      }
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

    return { fields: fields, attributes: attributes, matched: matched };
  }

  /* A metadata sheet is one whose header row we can actually read. */
  function findHeaderRow(worksheet) {
    let best = null;
    const limit = Math.min(worksheet.rowCount || 1, 10);
    for (let rowNumber = 1; rowNumber <= limit; rowNumber++) {
      const map = mapHeaders(worksheet, rowNumber);
      if (!best || map.matched > best.map.matched) best = { rowNumber: rowNumber, map: map };
    }
    return best && best.map.matched >= 4 ? best : null;
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

    out.reason = 'No sheet in this workbook has recognisable metadata headers.';
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
    readMetadataWorkbook: readMetadataWorkbook,
    buildDeliveryWorkbook: buildDeliveryWorkbook,
  };
})(window);

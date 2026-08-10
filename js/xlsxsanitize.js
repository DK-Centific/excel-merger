/*
 * Make real agency workbooks readable by ExcelJS.
 *
 * The client's Centaurus template defines an Excel Table (a ListObject) over its data
 * range. ExcelJS 4.x throws parsing those table parts —
 *   "Cannot set properties of undefined (setting 'filterButton')"
 * — and the whole workbook fails to load. Every filled-in copy of that template carries the
 * same table, so without this the feature cannot read a single real file.
 *
 * Skipping the table part alone is not enough: the worksheet still references it and
 * reconciliation then fails on a missing model. So the table has to be removed at the zip
 * level, along with every reference to it, before ExcelJS ever sees the file.
 *
 * Nothing of value is lost — a table is a formatting and filtering construct. The cells,
 * formulas, styles and sheets are untouched.
 */
(function (global) {
  'use strict';

  const TABLE_PART = /^xl\/tables\//;
  const WORKSHEET_XML = /^xl\/worksheets\/[^/]+\.xml$/;
  const WORKSHEET_RELS = /^xl\/worksheets\/_rels\/[^/]+\.rels$/;
  const CONTENT_TYPES = '[Content_Types].xml';

  function decode(bytes) {
    return new TextDecoder('utf-8').decode(bytes);
  }

  function encode(text) {
    return new TextEncoder().encode(text);
  }

  /*
   * Returns the original buffer untouched when there is no table, so the common case pays
   * nothing and cannot be broken by this code path.
   */
  function stripTables(arrayBuffer) {
    if (!global.fflate) return { buffer: arrayBuffer, stripped: 0 };

    let entries;
    try {
      entries = global.fflate.unzipSync(new Uint8Array(arrayBuffer));
    } catch (err) {
      // Not a zip we can read; hand it back and let ExcelJS report the real problem.
      return { buffer: arrayBuffer, stripped: 0, error: err.message };
    }

    const names = Object.keys(entries);
    const tableParts = names.filter(function (name) { return TABLE_PART.test(name); });
    if (!tableParts.length) return { buffer: arrayBuffer, stripped: 0 };

    const rebuilt = {};
    names.forEach(function (name) {
      if (TABLE_PART.test(name)) return;

      if (WORKSHEET_XML.test(name)) {
        let xml = decode(entries[name]);
        xml = xml.replace(/<tableParts[\s\S]*?<\/tableParts>/g, '').replace(/<tableParts[^>]*\/>/g, '');
        rebuilt[name] = encode(xml);
        return;
      }
      if (WORKSHEET_RELS.test(name)) {
        let xml = decode(entries[name]);
        xml = xml.replace(/<Relationship\b[^>]*\/tables\/[^>]*\/>/g, '');
        rebuilt[name] = encode(xml);
        return;
      }
      if (name === CONTENT_TYPES) {
        let xml = decode(entries[name]);
        xml = xml.replace(/<Override\b[^>]*spreadsheetml\.table\+xml[^>]*\/>/g, '');
        rebuilt[name] = encode(xml);
        return;
      }
      rebuilt[name] = entries[name];
    });

    const zipped = global.fflate.zipSync(rebuilt, { level: 6 });
    // Copy out of fflate's view so the ArrayBuffer is exactly the zip and nothing more.
    return { buffer: zipped.buffer.slice(zipped.byteOffset, zipped.byteOffset + zipped.byteLength),
             stripped: tableParts.length };
  }

  /* Load a workbook from any agency file, table or no table. */
  async function loadWorkbook(arrayBuffer) {
    const workbook = new ExcelJS.Workbook();
    try {
      await workbook.xlsx.load(arrayBuffer.slice(0));
      return { workbook: workbook, sanitized: false };
    } catch (err) {
      const result = stripTables(arrayBuffer);
      if (!result.stripped) throw err;
      const retry = new ExcelJS.Workbook();
      await retry.xlsx.load(result.buffer);
      return { workbook: retry, sanitized: true, strippedTables: result.stripped };
    }
  }

  global.XlsxSanitize = {
    stripTables: stripTables,
    loadWorkbook: loadWorkbook,
  };
})(window);

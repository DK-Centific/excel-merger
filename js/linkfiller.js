/*
 * Metadata Link Filler — matching engine.
 *
 * Fills the three Dropbox link columns on an already-merged Centaurus sheet. Everything
 * here is pure: it takes a sheet and a flat file listing and returns a plan. No network,
 * no DOM, so the whole correctness surface is testable without Dropbox credentials.
 *
 * The link-creation step is mechanical; the file-to-row matching is where a messy batch
 * goes wrong, which is why the guards below flag rather than guess.
 */
(function (global) {
  'use strict';

  /*
   * Fixed column positions from the spec. This sheet is NOT the 21-column output of the
   * merger — it is the wider Centaurus sheet where F-AB are already populated upstream.
   * Zero-based: D=3, E=4, J=9, K=10, Q=16, AC=28, AD=29.
   */
  const COL = {
    NAME: 3,          // D — folder / participant name
    VIDEO_URL: 4,     // E — written, one per row
    EXPRESSION: 9,    // J — Neutral | Non-Neutral
    NON_NEUTRAL: 10,  // K — which non-neutral expression, disambiguates a contested slot
    ENVIRONMENT: 16,  // Q — Indoor | Outdoor
    ICF_URL: 28,      // AC — written, repeats per participant
    ASSENT_URL: 29,   // AD — written, repeats per participant
  };

  const WRITE_COLUMNS = [COL.VIDEO_URL, COL.ICF_URL, COL.ASSENT_URL];

  /* F-AB inclusive (5..27) is read for the name join and never written. */
  const UNTOUCHED_FIRST = 5;
  const UNTOUCHED_LAST = 27;

  /* Filename conventions differ per agency, so type rules are config rather than hardcoded. */
  const AGENCY_PRESETS = {
    Powerling: {
      agency: 'Powerling',
      videoExt: ['.mp4', '.mov'],
      icfMatch: ['consent', 'icf'],
      assentMatch: ['assent'],
      assentExt: ['.pdf', '.jpg', '.jpeg', '.png'],
      expressionNeutralToken: 'neutral',
      fuzzyExpression: true,
    },
    Aqlama: {
      agency: 'Aqlama',
      videoExt: ['.mp4', '.mov'],
      icfMatch: ['icf', 'consent'],
      assentMatch: ['assent'],
      assentExt: ['.pdf', '.jpg', '.jpeg', '.png'],
      expressionNeutralToken: 'neutral',
      fuzzyExpression: true,
    },
  };

  const NEUTRAL = 'Neutral';
  const NON_NEUTRAL = 'Non-Neutral';
  const INDOOR = 'Indoor';
  const OUTDOOR = 'Outdoor';

  /* ---------- small helpers ---------- */

  function lower(text) {
    return String(text == null ? '' : text).toLowerCase();
  }

  function extensionOf(name) {
    const match = /\.[^.]+$/.exec(String(name));
    return match ? match[0].toLowerCase() : '';
  }

  function baseName(path) {
    const parts = String(path).split('/');
    return parts[parts.length - 1] || '';
  }

  /*
   * Collapse case and separators so "Abraham Otieno", "ABRAHAM OTIENO" and
   * "ABRAHAM-OTIENO" reach the same key. Deliberately exact after normalizing:
   * "Ronald Okoth" and "Ronald Okothh" are different people and must not merge.
   */
  function normalizeName(text) {
    return lower(text)
      .replace(/[._\-–—]+/g, ' ')
      .replace(/[^a-z0-9 ]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  /* ---------- Dropbox folder URLs ---------- */

  /*
   * A browsable folder path can be enumerated; an scl/fo share link cannot, and telling
   * the two apart up front is cheaper than failing halfway through a run.
   */
  function parseFolderUrl(raw) {
    const input = String(raw == null ? '' : raw).trim();
    if (!input) return null;

    if (/\/scl\/fo\//i.test(input) || /\/sh\//i.test(input)) {
      return {
        input: input,
        ok: false,
        path: null,
        reason: "Share links can't be listed. Use the folder's path, e.g. /Team Folder/Agency Collection/…",
      };
    }

    let path = input;

    if (/^https?:\/\//i.test(input)) {
      let parsed;
      try {
        parsed = new URL(input);
      } catch (err) {
        return { input: input, ok: false, path: null, reason: 'Not a valid URL or folder path.' };
      }
      if (!/(^|\.)dropbox\.com$/i.test(parsed.hostname)) {
        return { input: input, ok: false, path: null, reason: 'Not a Dropbox URL.' };
      }
      path = parsed.pathname;
      path = path.replace(/^\/home/i, '');
    }

    try {
      path = decodeURIComponent(path);
    } catch (err) {
      /* Leave the raw form if it carries a stray percent sign. */
    }

    path = path.replace(/\/+$/, '');
    if (path && path.charAt(0) !== '/') path = '/' + path;

    if (!path || path === '/') {
      return { input: input, ok: false, path: null, reason: 'No folder path found in that link.' };
    }
    return { input: input, ok: true, path: path, reason: '' };
  }

  function parseFolderUrls(text) {
    return String(text == null ? '' : text)
      .split(/\r?\n/)
      .map(function (line) { return parseFolderUrl(line); })
      .filter(Boolean);
  }

  /* ---------- shared links ---------- */

  /*
   * Store the durable form. The st= token expires and is not reproducible, so keeping it
   * would make a re-run produce a different-looking link for the same file.
   */
  function normalizeSharedLink(url) {
    const input = String(url == null ? '' : url).trim();
    if (!input) return '';

    let parsed;
    try {
      parsed = new URL(input);
    } catch (err) {
      return input;
    }

    const rlkey = parsed.searchParams.get('rlkey');
    parsed.search = '';
    if (rlkey) parsed.searchParams.set('rlkey', rlkey);
    parsed.searchParams.set('dl', '0');
    return parsed.toString();
  }

  /* ---------- filename classification ---------- */

  function classifyFile(name, config) {
    const base = lower(baseName(name));
    const ext = extensionOf(base);

    if (config.videoExt.indexOf(ext) !== -1) return 'video';

    const isAssent = config.assentMatch.some(function (token) { return base.indexOf(token) !== -1; });
    if (isAssent && config.assentExt.indexOf(ext) !== -1) return 'assent';

    const isIcf = config.icfMatch.some(function (token) { return base.indexOf(token) !== -1; });
    if (isIcf) return 'icf';

    // A stray metadata .xlsx can live beside the media; it is neither.
    return 'other';
  }

  function deriveEnvironment(name) {
    const base = lower(baseName(name));
    if (/\bindoor\b|indoor/.test(base)) return INDOOR;
    if (/\boutdoor\b|outdoor/.test(base)) return OUTDOOR;
    return null;
  }

  /* Split a filename into word-ish tokens so a typo can be compared token by token. */
  function tokenize(name) {
    return lower(baseName(name))
      .replace(/\.[^.]+$/, '')
      .split(/[^a-z0-9]+/)
      .filter(Boolean);
  }

  /*
   * NEUTRAL maps to Neutral; every other expression token maps to Non-Neutral. Fuzzy so a
   * filename typo like N3UTRAL still reads as neutral rather than silently flipping the row
   * into the wrong slot.
   */
  function deriveExpression(name, config) {
    const target = config.expressionNeutralToken;
    const tokens = tokenize(name);

    for (let i = 0; i < tokens.length; i++) {
      if (tokens[i] === target) return { expression: NEUTRAL, token: tokens[i], fuzzy: false };
    }

    if (config.fuzzyExpression && global.MergerCore && global.MergerCore.editDistance) {
      for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];
        if (Math.abs(token.length - target.length) > 2) continue;
        const distance = global.MergerCore.editDistance(token, target);
        if (distance > 0 && distance <= 2) {
          return { expression: NEUTRAL, token: token, fuzzy: true };
        }
      }
    }

    return { expression: NON_NEUTRAL, token: nonNeutralToken(tokens, name), fuzzy: false };
  }

  /* The descriptive word (SMILING, FROWNING…) is what column K disambiguates against. */
  const STRUCTURAL_TOKENS = ['indoor', 'outdoor', 'mp4', 'mov', 'video', 'expression'];

  function nonNeutralToken(tokens, name) {
    const env = lower(deriveEnvironment(name) || '');
    for (let i = tokens.length - 1; i >= 0; i--) {
      const token = tokens[i];
      if (token.length < 3) continue;
      if (token === env) continue;
      if (STRUCTURAL_TOKENS.indexOf(token) !== -1) continue;
      if (/^\d+$/.test(token)) continue;
      return token;
    }
    return '';
  }

  /* ---------- grouping ---------- */

  /*
   * The participant is the first path segment below the scanned batch folder — the folder
   * name is the identity, which is why an unmatched folder name is reported rather than
   * guessed at from the filename.
   */
  function participantFolderOf(filePath, rootPath) {
    const root = String(rootPath || '').replace(/\/+$/, '');
    let rest = String(filePath);
    if (root && lower(rest).indexOf(lower(root) + '/') === 0) {
      rest = rest.slice(root.length + 1);
    } else {
      rest = rest.replace(/^\//, '');
    }
    const parts = rest.split('/').filter(Boolean);
    return parts.length > 1 ? parts[0] : '';
  }

  function groupByParticipant(files, roots, config) {
    const groups = {};

    files.forEach(function (file) {
      const path = file.path_display || file.path_lower || file.name;

      /*
       * A caller that already knows whose file this is wins — that is how flat media, where
       * the participant is only identifiable from the filename, gets attributed.
       */
      let folder = file.participantFolder || '';
      if (!folder) {
        for (let i = 0; i < roots.length; i++) {
          const candidate = participantFolderOf(path, roots[i]);
          if (candidate) { folder = candidate; break; }
        }
      }
      if (!folder) folder = participantFolderOf(path, '');
      if (!folder) return;

      const key = normalizeName(folder);
      if (!groups[key]) {
        groups[key] = { key: key, folder: folder, videos: [], icf: [], assent: [], other: [] };
      }

      const group = groups[key];
      const name = file.name || baseName(path);
      const kind = classifyFile(name, config);
      const entry = {
        name: name,
        path: path,
        link: file.link ? normalizeSharedLink(file.link) : '',
        linkExisted: !!file.link,
        // Carried through so the caller can date a row from the media it matched.
        recordedAt: file.recordedAt || file.server_modified || null,
      };

      if (kind === 'video') {
        entry.environment = deriveEnvironment(name);
        const expression = deriveExpression(name, config);
        entry.expression = expression.expression;
        entry.expressionToken = expression.token;
        entry.expressionFuzzy = expression.fuzzy;
        group.videos.push(entry);
      } else if (kind === 'icf') {
        group.icf.push(entry);
      } else if (kind === 'assent') {
        group.assent.push(entry);
      } else {
        group.other.push(entry);
      }
    });

    return groups;
  }

  /* ---------- planning ---------- */

  function makeGuard(kind, severity, title, who, detail) {
    return { kind: kind, severity: severity, title: title, who: who, detail: detail };
  }

  /*
   * Choose the video for a row. Several videos can legitimately share (Environment,
   * Non-Neutral) — column K names which one this row wants, so try that before giving up.
   */
  function pickVideo(candidates, row, config) {
    if (candidates.length === 0) return { video: null, reason: 'none' };
    if (candidates.length === 1) return { video: candidates[0], reason: 'unique' };

    const wanted = normalizeName(row.values[COL.NON_NEUTRAL]);
    if (wanted) {
      const byK = candidates.filter(function (video) {
        const token = normalizeName(video.expressionToken);
        if (!token) return false;
        if (token === wanted || wanted.indexOf(token) !== -1 || token.indexOf(wanted) !== -1) return true;
        if (config.fuzzyExpression && global.MergerCore && global.MergerCore.editDistance) {
          return global.MergerCore.editDistance(token, wanted) <= 2;
        }
        return false;
      });
      if (byK.length === 1) return { video: byK[0], reason: 'column-k' };
    }
    return { video: null, reason: 'ambiguous' };
  }

  /*
   * files: [{ name, path_display, link }] — link is the existing shared link, if any.
   * sheet: { rows: [{ excelRow, values: [] }] }
   */
  function planFill(sheet, files, options) {
    const opts = options || {};
    const config = opts.config || AGENCY_PRESETS.Powerling;
    const roots = opts.roots || [];

    const groups = groupByParticipant(files || [], roots, config);
    const rows = [];
    const guards = [];
    const sheetKeys = {};

    (sheet.rows || []).forEach(function (row) {
      const key = normalizeName(row.values[COL.NAME]);
      if (key) sheetKeys[key] = true;
    });

    /* A folder that matches no sheet row is reported, never forced onto a near-miss row. */
    Object.keys(groups).forEach(function (key) {
      if (!sheetKeys[key]) {
        guards.push(makeGuard(
          'unmatched-name', 'error', "Folder name didn't match a sheet row",
          'folder "' + groups[key].folder + '"',
          'No participant name in column D normalizes to it. Confirm the mapping or fix the folder name.'
        ));
      }
    });

    /* Per-participant guards, raised once rather than repeated on every row. */
    Object.keys(groups).forEach(function (key) {
      const group = groups[key];
      if (!sheetKeys[key]) return;

      if (group.icf.length === 0) {
        guards.push(makeGuard('missing-icf', 'error', 'Consent file not found', group.folder,
          'No ICF/consent file in this folder. AC left blank — never borrowed from another participant.'));
      } else if (group.icf.length > 1) {
        guards.push(makeGuard('duplicate-icf', 'warn', 'More than one consent file', group.folder,
          'Found ' + group.icf.map(function (f) { return f.name; }).join(', ') + '. Ambiguous, so AC is held.'));
      }

      if (group.assent.length === 0) {
        guards.push(makeGuard('missing-assent', 'error', 'Assent file not found', group.folder,
          'No assent file in this folder. AD left blank — never borrowed from another participant.'));
      } else if (group.assent.length > 1) {
        guards.push(makeGuard('duplicate-assent', 'warn', 'More than one assent file', group.folder,
          'Found ' + group.assent.map(function (f) { return f.name; }).join(', ') + '. Ambiguous, so AD is held.'));
      }
    });

    const icfFor = {};
    const assentFor = {};
    Object.keys(groups).forEach(function (key) {
      icfFor[key] = groups[key].icf.length === 1 ? groups[key].icf[0] : null;
      assentFor[key] = groups[key].assent.length === 1 ? groups[key].assent[0] : null;
    });

    const ambiguousSeen = {};
    const firstRowForParticipant = {};

    (sheet.rows || []).forEach(function (row) {
      const key = normalizeName(row.values[COL.NAME]);
      const group = groups[key];

      const plan = {
        excelRow: row.excelRow,
        name: row.values[COL.NAME],
        environment: row.values[COL.ENVIRONMENT],
        expression: row.values[COL.EXPRESSION],
        nonNeutral: row.values[COL.NON_NEUTRAL],
        video: null, icf: null, assent: null,
        icfRepeatOf: null, assentRepeatOf: null,
        status: 'skipped',
        problems: [],
      };

      if (!group) {
        plan.status = 'skipped';
        plan.problems.push('No scanned folder matches this participant.');
        rows.push(plan);
        return;
      }

      const wantEnv = String(row.values[COL.ENVIRONMENT] || '').trim();
      const wantExpr = String(row.values[COL.EXPRESSION] || '').trim();
      const candidates = group.videos.filter(function (video) {
        const envOk = !wantEnv || !video.environment ||
          lower(video.environment) === lower(wantEnv);
        const exprOk = !wantExpr || lower(video.expression) === lower(wantExpr);
        return envOk && exprOk;
      });

      const picked = pickVideo(candidates, row, config);
      if (picked.video) {
        plan.video = picked.video;
        plan.videoReason = picked.reason;
      } else if (picked.reason === 'ambiguous') {
        plan.problems.push(candidates.length + ' videos match ' + wantEnv + ' / ' + wantExpr +
          ' and column K does not single one out.');
        const guardKey = key + '|' + wantEnv + '|' + wantExpr;
        if (!ambiguousSeen[guardKey]) {
          ambiguousSeen[guardKey] = true;
          guards.push(makeGuard('ambiguous-slot', 'warn', 'Two videos, one slot',
            group.folder + ' · ' + wantEnv,
            candidates.map(function (v) { return v.name; }).join(' and ') +
            ' both map to ' + wantExpr + '. Needs column K to pick the row.'));
        }
      } else {
        plan.problems.push('No video file matches ' + (wantEnv || '?') + ' / ' + (wantExpr || '?') + '.');
      }

      plan.icf = icfFor[key] || null;
      plan.assent = assentFor[key] || null;
      if (!plan.icf) plan.problems.push('No unambiguous consent file for this participant.');
      if (!plan.assent) plan.problems.push('No unambiguous assent file for this participant.');

      /* AC/AD repeat across a participant's rows; show later rows as pointing at the first. */
      if (firstRowForParticipant[key] == null) {
        firstRowForParticipant[key] = row.excelRow;
      } else {
        if (plan.icf) plan.icfRepeatOf = firstRowForParticipant[key];
        if (plan.assent) plan.assentRepeatOf = firstRowForParticipant[key];
      }

      /*
       * Matching cleanly is not the same as being writable. A file with no shared link yet
       * would write an empty cell, so those rows wait for the create step rather than
       * counting as filled.
       */
      const matched = !!(plan.video && plan.icf && plan.assent);
      const linked = matched && !!plan.video.link && !!plan.icf.link && !!plan.assent.link;

      if (linked) plan.status = 'filled';
      else if (matched) plan.status = 'awaiting-link';
      else if (plan.video || plan.icf || plan.assent) plan.status = 'review';
      else plan.status = 'missing';

      rows.push(plan);
    });

    return {
      rows: rows,
      guards: guards,
      groups: groups,
      config: config,
      metrics: summarize(rows, guards),
    };
  }

  function summarize(rows, guards) {
    let filled = 0;
    let awaiting = 0;
    let review = 0;
    let missing = 0;
    const links = {};
    let reused = 0;
    let created = 0;

    rows.forEach(function (plan) {
      if (plan.status === 'filled') filled++;
      else if (plan.status === 'awaiting-link') awaiting++;
      else if (plan.status === 'review') review++;
      else if (plan.status === 'missing') missing++;

      [plan.video, plan.icf, plan.assent].forEach(function (file) {
        if (!file || links[file.path]) return;
        links[file.path] = true;
        if (file.linkExisted) reused++; else created++;
      });
    });

    return {
      rowsFilled: filled,
      rowsAwaitingLink: awaiting,
      rowsMatched: filled + awaiting,
      rowsReview: review,
      rowsMissing: missing,
      linksResolved: reused + created,
      linksReused: reused,
      linksToCreate: created,
      filesMatched: Object.keys(links).length,
      guardCount: guards.length,
    };
  }

  /* Only clean rows are written, and only ever into E, AC and AD. */
  function applyPlan(worksheet, plan) {
    let written = 0;

    plan.rows.forEach(function (row) {
      if (row.status !== 'filled') return;
      // Belt and braces: never blank out a cell by writing an unresolved link.
      if (!row.video.link || !row.icf.link || !row.assent.link) return;
      const sheetRow = worksheet.getRow(row.excelRow);
      sheetRow.getCell(COL.VIDEO_URL + 1).value = row.video.link;
      sheetRow.getCell(COL.ICF_URL + 1).value = row.icf.link;
      sheetRow.getCell(COL.ASSENT_URL + 1).value = row.assent.link;
      sheetRow.commit && sheetRow.commit();
      written++;
    });

    return written;
  }

  /* Every file whose link must be created before a write can happen. */
  function filesNeedingLinks(plan) {
    const seen = {};
    const out = [];
    plan.rows.forEach(function (row) {
      [row.video, row.icf, row.assent].forEach(function (file) {
        if (!file || file.linkExisted || seen[file.path]) return;
        seen[file.path] = true;
        out.push(file);
      });
    });
    return out;
  }

  global.LinkFiller = {
    COL: COL,
    WRITE_COLUMNS: WRITE_COLUMNS,
    UNTOUCHED_FIRST: UNTOUCHED_FIRST,
    UNTOUCHED_LAST: UNTOUCHED_LAST,
    AGENCY_PRESETS: AGENCY_PRESETS,
    NEUTRAL: NEUTRAL,
    NON_NEUTRAL: NON_NEUTRAL,
    normalizeName: normalizeName,
    parseFolderUrl: parseFolderUrl,
    parseFolderUrls: parseFolderUrls,
    normalizeSharedLink: normalizeSharedLink,
    classifyFile: classifyFile,
    deriveEnvironment: deriveEnvironment,
    deriveExpression: deriveExpression,
    groupByParticipant: groupByParticipant,
    planFill: planFill,
    applyPlan: applyPlan,
    filesNeedingLinks: filesNeedingLinks,
  };
})(window);

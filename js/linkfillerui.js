/*
 * Link filler UI wiring.
 *
 * Kept separate from app.js so the merger keeps working untouched if this half fails to
 * load. Shares the design system and the existing file-grabbing, nothing else.
 */
(function (global) {
  'use strict';

  const LF = global.LinkFiller;
  const DBX = global.DropboxSource;
  const MAX_TABLE_ROWS = 300;

  const state = {
    workbook: null,
    fileName: '',
    sheetName: '',
    plan: null,
    scannedFiles: [],
    browsePath: null,
    browseChildren: [],
    browseSelection: [],
  };

  function $(id) { return document.getElementById(id); }

  function icon(pathData, className) {
    const NS = 'http://www.w3.org/2000/svg';
    const svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('viewBox', '0 0 24 24');
    svg.setAttribute('class', 'ico' + (className ? ' ' + className : ''));
    svg.setAttribute('aria-hidden', 'true');
    const path = document.createElementNS(NS, 'path');
    path.setAttribute('d', pathData);
    svg.appendChild(path);
    return svg;
  }

  const ICON_TICK = 'M5 13l4 4L19 7';
  const ICON_FOLDER = 'M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7z';

  function showError(message) {
    const box = $('lf-error');
    box.textContent = message;
    box.classList.remove('hidden');
  }

  function clearError() {
    $('lf-error').classList.add('hidden');
  }

  function setBusy(text) {
    $('lf-progress-text').textContent = text;
    $('lf-progress').classList.toggle('hidden', !text);
  }

  /* ---------- mode switching ---------- */

  function setMode(mode) {
    document.querySelectorAll('.modes .chip').forEach(function (chip) {
      const on = chip.dataset.mode === mode;
      chip.classList.toggle('is-active', on);
      chip.setAttribute('aria-selected', on ? 'true' : 'false');
    });
    $('view-merge').classList.toggle('hidden', mode !== 'merge');
    $('view-linkfiller').classList.toggle('hidden', mode !== 'linkfiller');
  }

  function bindTabs(tabAttr, panelAttr, onChange) {
    document.querySelectorAll('[data-' + tabAttr + ']').forEach(function (tab) {
      tab.addEventListener('click', function () {
        const name = tab.dataset[tabAttr];
        document.querySelectorAll('[data-' + tabAttr + ']').forEach(function (other) {
          const on = other === tab;
          other.classList.toggle('is-active', on);
          other.setAttribute('aria-selected', on ? 'true' : 'false');
        });
        document.querySelectorAll('[data-' + panelAttr + ']').forEach(function (panel) {
          panel.classList.toggle('is-active', panel.dataset[panelAttr] === name);
        });
        if (onChange) onChange(name);
      });
    });
  }

  /* ---------- the merged sheet ---------- */

  async function loadSheet(buffer, name) {
    clearError();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    state.workbook = workbook;
    state.fileName = name;

    const select = $('lf-sheet-select');
    select.textContent = '';
    workbook.worksheets.forEach(function (sheet) {
      const option = document.createElement('option');
      option.value = sheet.name;
      option.textContent = sheet.name + ' · ' + sheet.rowCount + ' rows';
      select.appendChild(option);
    });

    // The spec's default sheet name, when it happens to be there.
    const preferred = workbook.worksheets.filter(function (s) { return /^merged$/i.test(s.name); })[0];
    state.sheetName = preferred ? preferred.name : (workbook.worksheets[0] || {}).name || '';
    select.value = state.sheetName;

    $('lf-sheet-row').textContent = '';
    $('lf-sheet-row').appendChild(icon('M8 3h5l5 5v13H8z', 'file-icon'));
    const label = document.createElement('span');
    label.className = 'file-name';
    label.textContent = name;
    $('lf-sheet-row').appendChild(label);

    $('lf-sheet-info').classList.remove('hidden');
    refresh();
  }

  function clearSheet() {
    state.workbook = null;
    state.fileName = '';
    state.sheetName = '';
    $('lf-sheet-info').classList.add('hidden');
    refresh();
  }

  /* Read the chosen worksheet into the plain shape planFill expects. */
  function readSheetRows() {
    const worksheet = state.workbook.getWorksheet(state.sheetName);
    if (!worksheet) return { worksheet: null, rows: [] };

    const rows = [];
    worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
      if (rowNumber === 1) return; // header
      const values = [];
      for (let i = 1; i <= 30; i++) {
        const cell = row.getCell(i);
        const value = cell && cell.value;
        values.push(value == null ? null : (value.text != null ? value.text : value));
      }
      if (values.every(function (v) { return v == null || String(v).trim() === ''; })) return;
      rows.push({ excelRow: rowNumber, values: values });
    });
    return { worksheet: worksheet, rows: rows };
  }

  /* ---------- source folders ---------- */

  function renderChips() {
    const box = $('lf-chips');
    box.textContent = '';
    const parsed = LF.parseFolderUrls($('lf-urls').value);

    parsed.forEach(function (entry) {
      const chip = document.createElement('div');
      chip.className = 'urlchip' + (entry.ok ? '' : ' is-bad');

      const status = document.createElement('span');
      status.className = 'st';
      if (entry.ok) status.appendChild(icon(ICON_TICK));
      else status.textContent = '!';

      const path = document.createElement('span');
      path.className = 'path';
      path.textContent = entry.ok ? entry.path : entry.input;

      const why = document.createElement('span');
      why.className = 'why';
      why.textContent = entry.ok ? 'folder path' : entry.reason;

      chip.appendChild(status);
      chip.appendChild(path);
      chip.appendChild(why);
      box.appendChild(chip);
    });

    return parsed;
  }

  function activeSourceTab() {
    const active = document.querySelector('[data-srctab].is-active');
    return active ? active.dataset.srctab : 'paste';
  }

  /* Whichever source tab is showing is the one that feeds the scan. */
  function validRoots() {
    if (activeSourceTab() === 'browse') return state.browseSelection.slice();
    return LF.parseFolderUrls($('lf-urls').value)
      .filter(function (entry) { return entry.ok; })
      .map(function (entry) { return entry.path; });
  }

  /* ---------- folder browser ---------- */

  function browseRoot() {
    const configured = (global.MERGER_CONFIG && global.MERGER_CONFIG.dropboxBrowseRoot) || '';
    return String(configured).replace(/\/+$/, '');
  }

  function segmentsOf(path) {
    return String(path || '').split('/').filter(Boolean);
  }

  function renderCrumbs() {
    const box = $('lf-crumbs');
    box.textContent = '';

    const root = browseRoot();
    const rootSegs = segmentsOf(root);
    const curSegs = segmentsOf(state.browsePath);
    // Never offer a way above the configured root.
    const extra = curSegs.slice(rootSegs.length);

    const crumbs = [{ label: rootSegs[rootSegs.length - 1] || 'Dropbox', path: root }];
    extra.forEach(function (segment, i) {
      crumbs.push({ label: segment, path: root + '/' + extra.slice(0, i + 1).join('/') });
    });

    crumbs.forEach(function (crumb, i) {
      if (i > 0) {
        const sep = document.createElement('span');
        sep.className = 'crumb-sep';
        sep.textContent = '›';
        box.appendChild(sep);
      }
      const last = i === crumbs.length - 1;
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'crumb' + (last ? ' is-current' : '');
      button.textContent = crumb.label;
      if (last) {
        button.setAttribute('aria-current', 'true');
      } else {
        button.addEventListener('click', function () { browseTo(crumb.path); });
      }
      box.appendChild(button);
    });
  }

  function isSelected(path) {
    return state.browseSelection.indexOf(path) !== -1;
  }

  function toggleSelection(path) {
    const at = state.browseSelection.indexOf(path);
    if (at === -1) state.browseSelection.push(path);
    else state.browseSelection.splice(at, 1);
    renderNodes();
    renderSelectionBar();
    refresh();
  }

  function renderSelectionBar() {
    const count = state.browseSelection.length;
    const sum = $('lf-selcount');
    sum.textContent = '';
    if (!count) {
      sum.textContent = 'No folders selected';
    } else {
      const strong = document.createElement('b');
      strong.textContent = String(count);
      sum.appendChild(strong);
      sum.appendChild(document.createTextNode(
        ' folder' + (count === 1 ? '' : 's') + ' selected — scanned recursively'));
    }
    $('lf-clear-sel').classList.toggle('hidden', count === 0);
    $('lf-use-selected').disabled = count === 0;
  }

  function renderNodes() {
    const box = $('lf-nodes');
    box.textContent = '';

    if (!state.browseChildren.length) {
      const empty = document.createElement('div');
      empty.className = 'node-empty';
      empty.textContent = 'No subfolders here.';
      box.appendChild(empty);
      return;
    }

    state.browseChildren.forEach(function (folder) {
      const selected = isSelected(folder.path);

      const node = document.createElement('div');
      node.className = 'node' + (selected ? ' is-selected' : '');
      node.tabIndex = 0;
      node.setAttribute('role', 'checkbox');
      node.setAttribute('aria-checked', selected ? 'true' : 'false');
      node.setAttribute('aria-label', folder.name);

      const check = document.createElement('span');
      check.className = 'node-check';
      check.appendChild(icon(ICON_TICK));

      const name = document.createElement('span');
      name.className = 'node-name';
      name.textContent = folder.name;

      const open = document.createElement('button');
      open.type = 'button';
      open.className = 'node-open';
      open.textContent = 'Open ›';
      open.setAttribute('aria-label', 'Open ' + folder.name);
      open.addEventListener('click', function (event) {
        // Opening a folder should not also tick it.
        event.stopPropagation();
        browseTo(folder.path);
      });

      node.appendChild(check);
      node.appendChild(icon(ICON_FOLDER, 'node-folder'));
      node.appendChild(name);
      node.appendChild(open);

      node.addEventListener('click', function () { toggleSelection(folder.path); });
      node.addEventListener('keydown', function (event) {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          toggleSelection(folder.path);
        }
      });

      box.appendChild(node);
    });
  }

  async function browseTo(path) {
    clearError();
    state.browsePath = path || browseRoot();
    renderCrumbs();

    const box = $('lf-nodes');
    box.textContent = '';
    const loading = document.createElement('div');
    loading.className = 'node-empty';
    loading.textContent = 'Loading…';
    box.appendChild(loading);

    try {
      const result = await DBX.listFolderChildren(state.browsePath);
      state.browseChildren = result.folders;
      renderNodes();
    } catch (err) {
      state.browseChildren = [];
      box.textContent = '';
      const failed = document.createElement('div');
      failed.className = 'node-empty';
      failed.textContent = err && err.message ? err.message : String(err);
      box.appendChild(failed);
    }
    renderSelectionBar();
  }

  /* ---------- state-driven enabling ---------- */

  function refresh() {
    const configured = DBX.isConfigured();
    const account = DBX.getAccount();
    const connected = !!account;

    $('lf-unconfigured').classList.toggle('hidden', configured);
    $('lf-account-row').classList.toggle('hidden', !configured);

    const status = $('lf-status');
    if (connected) {
      const name = (account.name && account.name.display_name) || account.email || 'Dropbox';
      status.textContent = 'Connected as ' + name;
    } else {
      status.textContent = 'Not connected';
    }
    status.classList.toggle('is-on', connected);
    $('lf-connect').classList.toggle('hidden', connected);
    $('lf-disconnect').classList.toggle('hidden', !connected);
    $('lf-marker-connect').classList.toggle('is-done', connected);

    $('lf-stage-folders').classList.toggle('is-locked', !connected);
    $('lf-urls').disabled = !connected;

    // Connecting while Browse is open should populate the tree without another click.
    if (connected && activeSourceTab() === 'browse' && state.browsePath === null) {
      browseTo(browseRoot());
    }
    renderSelectionBar();

    $('cfg-sheet').textContent = state.workbook
      ? state.fileName + ' › ' + state.sheetName
      : '—';

    $('lf-scan').disabled = !(state.workbook && connected && validRoots().length > 0);
  }

  /* ---------- scan ---------- */

  async function scan() {
    clearError();
    if (!state.workbook) { showError('Load the merged sheet first.'); return; }

    const roots = validRoots();
    const config = LF.AGENCY_PRESETS[$('lf-agency').value] || LF.AGENCY_PRESETS.Powerling;

    try {
      setBusy('Listing files…');
      let files = [];
      for (let i = 0; i < roots.length; i++) {
        setBusy('Listing ' + roots[i] + '…');
        const found = await DBX.listFolderRecursive(roots[i], function (count) {
          setBusy('Listing ' + roots[i] + ' — ' + count + ' files…');
        });
        files = files.concat(found);
      }

      // Read-only pass. Nothing is created here; that needs an explicit confirmation.
      setBusy('Checking existing shared links…');
      await DBX.resolveExistingLinks(files, function (done, total) {
        setBusy('Checking existing shared links — ' + done + ' of ' + total + '…');
      });

      const sheet = readSheetRows();
      state.scannedFiles = files;
      state.plan = LF.planFill({ rows: sheet.rows }, files, { config: config, roots: roots });

      setBusy('');
      renderResults();
    } catch (err) {
      setBusy('');
      showError(err && err.message ? err.message : String(err));
    }
  }

  /* ---------- results ---------- */

  function addStat(container, value, label, className) {
    const box = document.createElement('div');
    box.className = 'stat' + (className ? ' ' + className : '');
    const v = document.createElement('div');
    v.className = 'stat-value';
    v.textContent = String(value);
    const l = document.createElement('div');
    l.className = 'stat-label';
    l.textContent = label;
    box.appendChild(v);
    box.appendChild(l);
    container.appendChild(box);
  }

  function linkCell(file, repeatOf, missingText) {
    const wrap = document.createElement('div');
    wrap.className = 'lf-link';

    const url = document.createElement('span');
    url.className = 'url';

    if (repeatOf) {
      wrap.classList.add('is-repeat');
      url.textContent = '↳ same as row ' + repeatOf;
      wrap.appendChild(url);
      return wrap;
    }
    if (!file) {
      wrap.classList.add('is-empty');
      url.textContent = missingText;
      wrap.appendChild(url);
      return wrap;
    }

    url.textContent = file.link || file.name;
    url.title = file.path;
    wrap.appendChild(url);

    const pill = document.createElement('span');
    pill.className = 'pill ' + (file.linkExisted ? 'pill-reuse' : 'pill-new');
    pill.textContent = file.linkExisted ? 'reuse' : 'new';
    wrap.appendChild(pill);
    return wrap;
  }

  const STATUS_META = {
    filled: { cls: 'ok', label: 'Ready' },
    'awaiting-link': { cls: 'pending', label: 'Needs link' },
    review: { cls: 'warn', label: 'Review' },
    missing: { cls: 'err', label: 'Missing' },
    skipped: { cls: 'skip', label: 'Not scanned' },
  };

  function renderResults() {
    const plan = state.plan;
    const metrics = plan.metrics;

    const stats = $('lf-stats');
    stats.textContent = '';
    addStat(stats, metrics.rowsFilled, 'Rows ready to write', 'fixed');
    addStat(stats, metrics.rowsAwaitingLink, 'Rows waiting on a link');
    addStat(stats, metrics.linksResolved, metrics.linksReused + ' reused · ' + metrics.linksToCreate + ' to create');
    addStat(stats, metrics.rowsReview + metrics.rowsMissing, 'Rows need review', 'review');

    const body = $('lf-body');
    body.textContent = '';
    const shown = plan.rows.slice(0, MAX_TABLE_ROWS);

    shown.forEach(function (row, index) {
      const tr = document.createElement('tr');
      if (row.status === 'review') tr.className = 'is-review';
      if (row.status === 'missing') tr.className = 'is-missing';
      if (row.status === 'awaiting-link') tr.className = 'is-pending';

      const cells = [
        String(index + 1).padStart(2, '0'),
        row.name || '—',
        row.environment || '—',
        row.expression || '—',
      ];
      cells.forEach(function (text, i) {
        const td = document.createElement('td');
        if (i === 0) td.className = 'num';
        td.textContent = text;
        tr.appendChild(td);
      });

      const videoTd = document.createElement('td');
      videoTd.appendChild(linkCell(row.video, null, 'no video matched'));
      tr.appendChild(videoTd);

      const icfTd = document.createElement('td');
      icfTd.appendChild(linkCell(row.icf, row.icfRepeatOf, 'no consent file'));
      tr.appendChild(icfTd);

      const assentTd = document.createElement('td');
      assentTd.appendChild(linkCell(row.assent, row.assentRepeatOf, 'no assent file'));
      tr.appendChild(assentTd);

      const statusTd = document.createElement('td');
      const meta = STATUS_META[row.status] || STATUS_META.skipped;
      const badge = document.createElement('span');
      badge.className = 'rowstat ' + meta.cls;
      const dot = document.createElement('span');
      dot.className = 'd';
      badge.appendChild(dot);
      badge.appendChild(document.createTextNode(meta.label));
      if (row.problems.length) badge.title = row.problems.join(' ');
      statusTd.appendChild(badge);
      tr.appendChild(statusTd);

      body.appendChild(tr);
    });

    const truncated = $('lf-truncated');
    if (plan.rows.length > MAX_TABLE_ROWS) {
      truncated.textContent = 'Showing the first ' + MAX_TABLE_ROWS + ' of ' + plan.rows.length +
        ' rows. The downloaded sheet contains all of them.';
      truncated.classList.remove('hidden');
    } else {
      truncated.classList.add('hidden');
    }

    renderGuards(plan.guards);
    renderRules(plan.config);

    const pending = LF.filesNeedingLinks(plan);
    const createBtn = $('lf-create-links');
    createBtn.classList.toggle('hidden', pending.length === 0);
    createBtn.textContent = 'Create ' + pending.length + ' missing link' + (pending.length === 1 ? '' : 's');
    createBtn.disabled = false;

    $('lf-results').classList.remove('hidden');
    $('lf-results').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function renderGuards(guards) {
    const box = $('lf-guards');
    box.textContent = '';
    $('lf-guard-count').textContent = String(guards.length);

    if (!guards.length) {
      const empty = document.createElement('div');
      empty.className = 'guard';
      empty.innerHTML = '<div class="guard-body"><b>Nothing held back.</b>' +
        '<p>Every scanned row matched a file and a link.</p></div>';
      box.appendChild(empty);
      return;
    }

    guards.forEach(function (guard) {
      const item = document.createElement('div');
      item.className = 'guard';

      const ico = document.createElement('div');
      ico.className = 'guard-ico ' + (guard.severity === 'error' ? 'error' : 'warn');
      ico.textContent = guard.severity === 'error' ? '!' : '?';

      const body = document.createElement('div');
      body.className = 'guard-body';
      const title = document.createElement('b');
      title.textContent = guard.title;
      const detail = document.createElement('p');
      const who = document.createElement('span');
      who.className = 'guard-who';
      who.textContent = guard.who;
      detail.appendChild(who);
      detail.appendChild(document.createTextNode(' — ' + guard.detail));

      body.appendChild(title);
      body.appendChild(detail);
      item.appendChild(ico);
      item.appendChild(body);
      box.appendChild(item);
    });
  }

  function renderRules(config) {
    const box = $('lf-rules');
    box.textContent = '';
    const rules = [
      [config.videoExt.join(' / '), 'video → column E'],
      [config.icfMatch.join(' · '), '→ column AC'],
      [config.assentMatch.join(' · '), '→ column AD (' + config.assentExt.join(' ') + ')'],
      ['"' + config.expressionNeutralToken + '"', 'J = Neutral' + (config.fuzzyExpression ? ' · fuzzy (N3UTRAL ok)' : '')],
      ['name join', 'case + separators normalized, exact after that'],
    ];
    rules.forEach(function (pair) {
      const rule = document.createElement('div');
      rule.className = 'rule';
      const code = document.createElement('code');
      code.textContent = pair[0];
      const arrow = document.createElement('span');
      arrow.className = 'arrow';
      arrow.textContent = '→';
      const text = document.createElement('span');
      text.textContent = pair[1];
      rule.appendChild(code);
      rule.appendChild(arrow);
      rule.appendChild(text);
      box.appendChild(rule);
    });
  }

  /* ---------- create links (mutating, confirmed) ---------- */

  async function createMissingLinks() {
    const pending = LF.filesNeedingLinks(state.plan);
    if (!pending.length) return;

    const ok = global.confirm(
      'This creates ' + pending.length + ' new Dropbox shared link' + (pending.length === 1 ? '' : 's') +
      '.\n\nEach one is an anyone-with-the-link URL to a participant file, including consent and assent forms.\n\n' +
      'Continue?'
    );
    if (!ok) return;

    clearError();
    $('lf-create-links').disabled = true;

    try {
      for (let i = 0; i < pending.length; i++) {
        const file = pending[i];
        setBusy('Creating link ' + (i + 1) + ' of ' + pending.length + '…');
        const url = await DBX.createSharedLink(file.path);
        file.link = LF.normalizeSharedLink(url);
        file.linkExisted = true;
      }
      setBusy('');

      // Re-plan so statuses and counts reflect the links that now exist.
      const sheet = readSheetRows();
      const config = LF.AGENCY_PRESETS[$('lf-agency').value] || LF.AGENCY_PRESETS.Powerling;
      state.plan = LF.planFill({ rows: sheet.rows }, state.scannedFiles, { config: config, roots: validRoots() });
      renderResults();
    } catch (err) {
      setBusy('');
      $('lf-create-links').disabled = false;
      showError(err && err.message ? err.message : String(err));
    }
  }

  /* ---------- download ---------- */

  async function download() {
    clearError();
    if (!state.plan) return;

    try {
      setBusy('Writing links…');
      const worksheet = state.workbook.getWorksheet(state.sheetName);
      const written = LF.applyPlan(worksheet, state.plan);

      const buffer = await state.workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = state.fileName.replace(/\.xlsx?$/i, '') + '_linked.xlsx';
      document.body.appendChild(anchor);
      anchor.click();
      document.body.removeChild(anchor);
      URL.revokeObjectURL(url);

      setBusy('');
      const metrics = state.plan.metrics;
      const notes = [];
      if (metrics.rowsAwaitingLink) {
        notes.push(metrics.rowsAwaitingLink + ' matched cleanly but have no shared link yet — use ' +
          '"Create missing links" first, then download again.');
      }
      if (metrics.rowsReview + metrics.rowsMissing) {
        notes.push((metrics.rowsReview + metrics.rowsMissing) + ' are held for review — fix those and re-run.');
      }
      if (notes.length) {
        showError(written + ' row' + (written === 1 ? '' : 's') + ' written. ' + notes.join(' '));
      }
    } catch (err) {
      setBusy('');
      showError(err && err.message ? err.message : String(err));
    }
  }

  /* ---------- init ---------- */

  function init() {
    if (!LF || !DBX || !$('view-linkfiller')) return;

    document.querySelectorAll('.modes .chip').forEach(function (chip) {
      chip.addEventListener('click', function () { setMode(chip.dataset.mode); });
    });
    bindTabs('lftab', 'lfpanel');
    bindTabs('srctab', 'srcpanel', function (name) {
      // Load the tree the first time Browse is opened, not on every switch.
      if (name === 'browse' && DBX.getAccount() && state.browsePath === null) browseTo(browseRoot());
      refresh();
    });

    const dropzone = $('lf-dropzone');
    const input = $('lf-file-input');
    dropzone.addEventListener('click', function () { input.click(); });
    dropzone.addEventListener('keydown', function (event) {
      if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); input.click(); }
    });
    ['dragenter', 'dragover'].forEach(function (name) {
      dropzone.addEventListener(name, function (event) {
        event.preventDefault();
        dropzone.classList.add('is-over');
      });
    });
    ['dragleave', 'drop'].forEach(function (name) {
      dropzone.addEventListener(name, function () { dropzone.classList.remove('is-over'); });
    });
    dropzone.addEventListener('drop', async function (event) {
      event.preventDefault();
      const file = event.dataTransfer.files[0];
      if (file) await loadSheet(await file.arrayBuffer(), file.name);
    });
    input.addEventListener('change', async function () {
      const file = input.files[0];
      if (file) await loadSheet(await file.arrayBuffer(), file.name);
      input.value = '';
    });

    $('lf-clear-sheet').addEventListener('click', clearSheet);
    $('lf-sheet-select').addEventListener('change', function () {
      state.sheetName = $('lf-sheet-select').value;
      refresh();
    });

    $('lf-urls').addEventListener('input', function () { renderChips(); refresh(); });

    $('lf-use-selected').addEventListener('click', function () { scan(); });
    $('lf-clear-sel').addEventListener('click', function () {
      state.browseSelection = [];
      renderNodes();
      renderSelectionBar();
      refresh();
    });
    $('lf-agency').addEventListener('change', refresh);
    $('lf-scan').addEventListener('click', scan);
    $('lf-create-links').addEventListener('click', createMissingLinks);
    $('lf-download').addEventListener('click', download);

    $('lf-open-settings').addEventListener('click', function () {
      const btn = $('settings-btn');
      if (btn) btn.click();
    });

    $('lf-connect').addEventListener('click', async function () {
      clearError();
      try {
        // The redirect loses the page, so hand the typed URLs forward.
        await DBX.signIn({ urls: $('lf-urls').value, mode: 'linkfiller' });
      } catch (err) {
        showError(err && err.message ? err.message : String(err));
      }
    });
    $('lf-disconnect').addEventListener('click', function () { DBX.signOut(); refresh(); });

    $('lf-sp-fetch').addEventListener('click', async function () {
      clearError();
      const url = $('lf-sp-url').value.trim();
      if (!url) return;
      try {
        setBusy('Fetching from SharePoint…');
        const files = await global.SharePointSource.fetchFiles([url]);
        setBusy('');
        if (!files.length) { showError('That link did not return a workbook.'); return; }
        await loadSheet(files[0].buffer, files[0].name);
      } catch (err) {
        setBusy('');
        showError(err && err.message ? err.message : String(err));
      }
    });

    const dbxRedirect = $('dbx-redirect-uri');
    if (dbxRedirect) dbxRedirect.textContent = DBX.redirectUri();

    document.addEventListener('merger:settings-saved', refresh);

    refresh();

    // Completes a Dropbox redirect, if we came back from one.
    DBX.restore().then(function () {
      const resumed = DBX.takeResumeState();
      if (resumed) {
        if (resumed.urls) $('lf-urls').value = resumed.urls;
        if (resumed.mode === 'linkfiller') setMode('linkfiller');
      }
      renderChips();
      refresh();
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})(window);

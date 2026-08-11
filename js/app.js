/* UI wiring. All rendering uses textContent so agency data can never inject markup. */
(function () {
  'use strict';

  const MAX_QA_ROWS_RENDERED = 500;

  const state = {
    files: [],
    result: null,
    severityFilter: 'all',
    search: '',
  };

  const $ = function (id) { return document.getElementById(id); };
  const SVG_NS = 'http://www.w3.org/2000/svg';

  const ICON_SHEET = 'M4 4h16v16H4z M4 9h16 M4 14h16 M9 4v16';
  const ICON_CLOSE = 'M6 6l12 12M18 6L6 18';

  /* ---------- helpers ---------- */

  function icon(pathData, className) {
    const svg = document.createElementNS(SVG_NS, 'svg');
    svg.setAttribute('viewBox', '0 0 24 24');
    svg.setAttribute('aria-hidden', 'true');
    svg.setAttribute('class', 'ico' + (className ? ' ' + className : ''));
    pathData.split(' M').forEach(function (segment, index) {
      const path = document.createElementNS(SVG_NS, 'path');
      path.setAttribute('d', index === 0 ? segment : 'M' + segment);
      svg.appendChild(path);
    });
    return svg;
  }

  function formatBytes(bytes) {
    if (!bytes && bytes !== 0) return '';
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(0) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  }

  function todayStamp() {
    const d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }

  function showError(message) {
    const box = $('global-error');
    box.textContent = message;
    box.classList.remove('hidden');
  }

  function clearError() {
    $('global-error').classList.add('hidden');
  }

  function setProgress(message) {
    const box = $('progress');
    if (message == null) {
      box.classList.add('hidden');
      return;
    }
    $('progress-text').textContent = message;
    box.classList.remove('hidden');
  }

  async function downloadWorkbook(workbook, filename) {
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  /* ---------- file queue ---------- */

  function addFiles(entries) {
    entries.forEach(function (entry) {
      const duplicate = state.files.some(function (f) {
        return f.name === entry.name && f.buffer.byteLength === entry.buffer.byteLength;
      });
      if (!duplicate) state.files.push(entry);
    });
    renderFileList();
  }

  function removeFile(index) {
    state.files.splice(index, 1);
    renderFileList();
  }

  function renderFileList() {
    const list = $('file-list');
    const wrap = $('file-list-wrap');
    list.textContent = '';

    state.files.forEach(function (file, index) {
      const li = document.createElement('li');

      const name = document.createElement('span');
      name.className = 'file-name';
      name.textContent = file.name;

      const meta = document.createElement('span');
      meta.className = 'file-meta';
      meta.textContent = file.source + ' · ' + formatBytes(file.buffer.byteLength);

      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'file-remove';
      remove.setAttribute('aria-label', 'Remove ' + file.name);
      remove.appendChild(icon(ICON_CLOSE));
      remove.addEventListener('click', function () { removeFile(index); });

      li.appendChild(icon(ICON_SHEET, 'file-icon'));
      li.appendChild(name);
      li.appendChild(meta);
      li.appendChild(remove);
      list.appendChild(li);
    });

    $('file-count').textContent = String(state.files.length);
    wrap.classList.toggle('hidden', state.files.length === 0);
    $('merge-btn').disabled = state.files.length === 0;
  }

  /* Takes a plain array: a live FileList empties mid-loop when the input is reset. */
  async function readLocalFiles(files) {
    const accepted = [];
    const rejected = [];

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
        rejected.push(file.name);
        continue;
      }
      accepted.push({ name: file.name, buffer: await file.arrayBuffer(), source: 'Upload' });
    }

    if (rejected.length) {
      showError('Skipped ' + rejected.length + ' file(s) that are not .xlsx or .xlsm: ' + rejected.join(', ') +
        '. Re-save them as .xlsx and try again.');
    } else {
      clearError();
    }
    addFiles(accepted);
  }

  /* ---------- merge ---------- */

  async function runMerge() {
    clearError();
    $('merge-btn').disabled = true;
    setProgress('Reading files…');

    try {
      const options = {
        dayFirst: $('date-order').value === 'day',
        onProgress: function (index, total, name) {
          setProgress('Checking ' + name + ' (' + (index + 1) + ' of ' + total + ')…');
        },
      };

      const result = await MergerCore.mergeFiles(state.files, options);
      state.result = result;
      renderResults();
      setProgress(null);
    } catch (err) {
      setProgress(null);
      showError('Merge failed: ' + (err && err.message ? err.message : String(err)));
    } finally {
      $('merge-btn').disabled = state.files.length === 0;
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

  function renderStats() {
    const stats = $('stats');
    stats.textContent = '';
    const result = state.result;
    const mergedFiles = result.fileReports.filter(function (r) { return r.ok; }).length;

    addStat(stats, mergedFiles + ' / ' + result.fileReports.length, 'Files merged');
    addStat(stats, result.rows.length, 'Rows merged');
    addStat(stats, result.counts.fixed, 'Auto-fixed', 'fixed');
    addStat(stats, result.counts.review, 'Needs review', 'review');
    addStat(stats, result.counts.error, 'Errors', 'error');
  }

  function renderSources() {
    const tbody = $('sources-body');
    tbody.textContent = '';

    state.result.fileReports.forEach(function (report) {
      const tr = document.createElement('tr');
      const range = report.firstMergedRow == null
        ? '—'
        : report.firstMergedRow + '–' + report.lastMergedRow;

      [report.name, report.source, String(report.rowCount), range].forEach(function (text) {
        const td = document.createElement('td');
        td.textContent = text;
        tr.appendChild(td);
      });

      const status = document.createElement('td');
      status.textContent = report.ok ? 'Merged' : 'Skipped';
      if (!report.ok) status.className = 'skipped';
      tr.appendChild(status);

      tbody.appendChild(tr);
    });
  }

  function filteredIssues() {
    const term = state.search.trim().toLowerCase();
    return state.result.issues.filter(function (issue) {
      if (state.severityFilter !== 'all' && issue.severity !== state.severityFilter) return false;
      if (!term) return true;
      return (issue.file + ' ' + issue.rule + ' ' + issue.header + ' ' + issue.original + ' ' +
        issue.corrected + ' ' + issue.message).toLowerCase().indexOf(term) !== -1;
    });
  }

  function renderQaTable() {
    const body = $('qa-body');
    body.textContent = '';

    const issues = filteredIssues();
    const visible = issues.slice(0, MAX_QA_ROWS_RENDERED);

    visible.forEach(function (issue) {
      const tr = document.createElement('tr');

      const sevCell = document.createElement('td');
      const badge = document.createElement('span');
      badge.className = 'sev sev-' + issue.severity;
      badge.textContent = MergerCore.SEVERITY_LABEL[issue.severity] || issue.severity;
      sevCell.appendChild(badge);
      tr.appendChild(sevCell);

      const columns = [
        { text: issue.file },
        { text: String(issue.sourceRow), cls: 'num' },
        { text: String(issue.mergedRow), cls: 'num' },
        { text: issue.column, cls: 'num' },
        { text: issue.rule, cls: 'col-check' },
        { text: issue.original, cls: 'mono' },
        { text: issue.corrected, cls: 'mono' },
        { text: issue.message, cls: 'col-detail' },
      ];

      columns.forEach(function (col) {
        const td = document.createElement('td');
        if (col.cls) td.className = col.cls;
        td.textContent = col.text;
        tr.appendChild(td);
      });

      body.appendChild(tr);
    });

    $('qa-empty').classList.toggle('hidden', issues.length !== 0);

    const truncated = $('qa-truncated');
    if (issues.length > MAX_QA_ROWS_RENDERED) {
      truncated.textContent = 'Showing the first ' + MAX_QA_ROWS_RENDERED + ' of ' + issues.length +
        ' issues. Download the QA report to see them all.';
      truncated.classList.remove('hidden');
    } else {
      truncated.classList.add('hidden');
    }
  }

  function renderResults() {
    renderStats();
    renderSources();
    renderQaTable();
    $('results').classList.remove('hidden');
    $('results').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  /* ---------- sharepoint panel ---------- */

  function refreshSharePointPanel() {
    if (!$('sp-unconfigured')) return; // merger view not present
    const configured = SharePointSource.isConfigured();
    $('sp-unconfigured').classList.toggle('hidden', configured);
    $('sp-configured').classList.toggle('hidden', !configured);

    const account = SharePointSource.getAccount();
    const signedIn = !!account;
    const status = $('sp-status');

    if (signedIn) {
      status.textContent = 'Signed in as ' + (account.username || account.name);
    } else {
      status.textContent = 'Not signed in';
    }
    status.classList.toggle('is-on', signedIn);

    $('sp-signin').classList.toggle('hidden', signedIn);
    $('sp-switch').classList.toggle('hidden', !signedIn);
    $('sp-signout').classList.toggle('hidden', !signedIn);
    $('sp-marker-signin').classList.toggle('is-done', signedIn);

    // Picking a location is meaningless until we know whose permissions apply.
    $('sp-stage-url').classList.toggle('is-locked', !signedIn);
    $('sp-urls').disabled = !signedIn;
    $('sp-fetch').disabled = !signedIn;
  }

  async function handleSignIn(forceAccountPicker) {
    clearError();
    try {
      await SharePointSource.signIn(forceAccountPicker);
      refreshSharePointPanel();
      $('sp-urls').focus();
    } catch (err) {
      const message = err && err.message ? err.message : String(err);
      if (/user_cancelled|popup_window_error|interaction_in_progress/i.test(message)) {
        showError('Sign-in was cancelled or the popup was blocked. Allow popups for this site and try again.');
      } else {
        showError('Sign-in failed: ' + message);
      }
    }
  }

  async function fetchFromSharePoint() {
    clearError();
    const urls = $('sp-urls').value.split('\n')
      .map(function (line) { return line.trim(); })
      .filter(function (line) { return line !== ''; });

    if (!urls.length) {
      showError('Paste at least one SharePoint link first.');
      return;
    }

    $('sp-fetch').disabled = true;
    setProgress('Connecting to SharePoint…');

    try {
      const outcome = await SharePointSource.fetchFromUrls(urls, function (index, total, label) {
        setProgress('Downloading ' + label + ' (' + (index + 1) + ' of ' + total + ')…');
      });

      addFiles(outcome.files);
      refreshSharePointPanel();

      if (outcome.errors.length) {
        showError('Some links could not be used:\n' + outcome.errors.map(function (e) {
          return '• ' + e.url + ' — ' + e.message;
        }).join('\n'));
      }
    } catch (err) {
      showError('SharePoint fetch failed: ' + (err && err.message ? err.message : String(err)));
    } finally {
      setProgress(null);
      $('sp-fetch').disabled = false;
    }
  }

  /* ---------- settings ---------- */

  function openSettings() {
    const config = SharePointSource.getConfig();
    $('cfg-client').value = config.clientId;
    $('cfg-tenant').value = config.tenantId;
    $('redirect-uri').textContent = window.location.origin + window.location.pathname;

    if (window.DropboxSource) {
      $('cfg-dropbox').value = window.DropboxSource.getConfig().appKey;
      $('dbx-redirect-uri').textContent = window.DropboxSource.redirectUri();
    }
    $('settings-modal').classList.remove('hidden');
  }

  function closeSettings() {
    $('settings-modal').classList.add('hidden');
  }

  /* Shared chrome: works with or without the standalone merger view present. */
  function initSettings() {
    const openers = [$('settings-btn'), $('open-settings-link')];
    openers.forEach(function (button) {
      if (button) button.addEventListener('click', openSettings);
    });
    $('cfg-cancel').addEventListener('click', closeSettings);
    $('cfg-save').addEventListener('click', saveSettings);
    $('settings-modal').addEventListener('click', function (event) {
      if (event.target === $('settings-modal')) closeSettings();
    });
    document.addEventListener('keydown', function (event) {
      if (event.key === 'Escape') closeSettings();
    });
  }

  function saveSettings() {
    SharePointSource.saveConfig($('cfg-client').value.trim(), $('cfg-tenant').value.trim() || 'common');
    if (window.DropboxSource) window.DropboxSource.saveConfig($('cfg-dropbox').value.trim());
    closeSettings();
    refreshSharePointPanel();
    document.dispatchEvent(new CustomEvent('merger:settings-saved'));
  }

  /* ---------- wiring ---------- */

  /* Scoped to this view: the link filler has its own tab groups using the same classes. */
  function initTabs() {
    const scope = $('view-merge') || document;
    scope.querySelectorAll('.tab[data-tab]').forEach(function (tab) {
      tab.addEventListener('click', function () {
        scope.querySelectorAll('.tab[data-tab]').forEach(function (t) {
          const active = t === tab;
          t.classList.toggle('is-active', active);
          t.setAttribute('aria-selected', String(active));
        });
        scope.querySelectorAll('.tab-panel[data-panel]').forEach(function (panel) {
          panel.classList.toggle('is-active', panel.dataset.panel === tab.dataset.tab);
        });
      });
    });
  }

  function initDropzone() {
    const zone = $('dropzone');
    const input = $('file-input');

    zone.addEventListener('click', function () { input.click(); });
    zone.addEventListener('keydown', function (event) {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        input.click();
      }
    });

    input.addEventListener('change', function () {
      const picked = Array.from(input.files);
      input.value = '';
      readLocalFiles(picked);
    });

    ['dragenter', 'dragover'].forEach(function (name) {
      zone.addEventListener(name, function (event) {
        event.preventDefault();
        zone.classList.add('is-over');
      });
    });
    ['dragleave', 'drop'].forEach(function (name) {
      zone.addEventListener(name, function (event) {
        event.preventDefault();
        zone.classList.remove('is-over');
      });
    });
    zone.addEventListener('drop', function (event) {
      if (event.dataTransfer && event.dataTransfer.files.length) readLocalFiles(Array.from(event.dataTransfer.files));
    });
  }

  function initFilters() {
    document.querySelectorAll('#severity-filters .chip').forEach(function (chip) {
      chip.addEventListener('click', function () {
        state.severityFilter = chip.dataset.severity;
        document.querySelectorAll('#severity-filters .chip').forEach(function (c) {
          c.classList.toggle('is-active', c === chip);
        });
        renderQaTable();
      });
    });

    $('qa-search').addEventListener('input', function (event) {
      state.search = event.target.value;
      renderQaTable();
    });
  }

  function initDownloads() {
    $('download-merged').addEventListener('click', async function () {
      if (!state.result) return;
      const includeQa = $('include-qa').getAttribute('aria-checked') === 'true';
      const workbook = MergerCore.buildMergedWorkbook(state.result, { includeQaSheet: includeQa });
      await downloadWorkbook(workbook, 'Merged_Delivery_' + todayStamp() + '.xlsx');
    });

    $('download-qa').addEventListener('click', async function () {
      if (!state.result) return;
      const workbook = MergerCore.buildQaWorkbook(state.result);
      await downloadWorkbook(workbook, 'QA_Report_' + todayStamp() + '.xlsx');
    });
  }

  function initTheme() {
    let saved = null;
    try { saved = localStorage.getItem('merger.theme'); } catch (err) { /* private browsing */ }
    document.documentElement.setAttribute('data-theme', saved || 'dark');

    $('theme-toggle').addEventListener('click', function () {
      const next = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', next);
      try { localStorage.setItem('merger.theme', next); } catch (err) { /* private browsing */ }
    });
  }

  function initSwitches() {
    document.querySelectorAll('.switch[role="switch"]').forEach(function (toggle) {
      toggle.addEventListener('click', function () {
        toggle.setAttribute('aria-checked', toggle.getAttribute('aria-checked') === 'true' ? 'false' : 'true');
      });
    });
  }

  function init() {
    // Theme and the settings dialog are shared chrome, so they are wired first and
    // unconditionally — the standalone merger view is optional.
    initTheme();
    initSettings();

    if (!$('view-merge')) return;

    initTabs();
    initDropzone();
    initFilters();
    initSwitches();
    initDownloads();

    $('clear-files').addEventListener('click', function () {
      state.files = [];
      renderFileList();
    });

    $('merge-btn').addEventListener('click', runMerge);
    $('sp-fetch').addEventListener('click', fetchFromSharePoint);

    $('sp-signin').addEventListener('click', function () { handleSignIn(false); });
    $('sp-switch').addEventListener('click', function () { handleSignIn(true); });

    $('sp-signout').addEventListener('click', async function () {
      try {
        await SharePointSource.signOut();
      } finally {
        refreshSharePointPanel();
      }
    });

    refreshSharePointPanel();
    renderFileList();

    // A reload should not look like a sign-out if MSAL still holds the session.
    SharePointSource.restore().then(refreshSharePointPanel);
  }

  document.addEventListener('DOMContentLoaded', init);
})();

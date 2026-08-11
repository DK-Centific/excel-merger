/*
 * Dropbox access for the link filler.
 *
 * Browser-only, so this uses the PKCE authorization-code flow — a public client with no
 * secret, same posture as the SharePoint module. Dormant until an app key is configured;
 * every entry point checks isConfigured() first so the rest of the app works without it.
 *
 * NOTE: this module has not been exercised against the live Dropbox API — there is no app
 * key registered yet. The shapes follow the v2 HTTP docs; expect to verify the team-folder
 * path root and pagination on the first real run.
 */
(function (global) {
  'use strict';

  /*
   * Requested explicitly rather than relying on "whatever the app has selected". Being
   * explicit means a token can never come back quietly missing a scope, and it documents
   * exactly what this tool needs:
   *   account_info.read   users/get_current_account — the connection check and, on a
   *                       Business team, root_info for addressing team folders
   *   files.metadata.read files/list_folder — browsing and enumerating
   *   sharing.read        sharing/list_shared_links — reusing existing links
   *   sharing.write       sharing/create_shared_link_with_settings — only on confirm
   */
  const SCOPES = [
    'account_info.read',
    'files.metadata.read',
    // Needed to read the agency metadata workbooks themselves. Only .xlsx files are ever
    // downloaded; videos and consent PDFs are linked, never fetched.
    'files.content.read',
    'sharing.read',
    'sharing.write',
  ];

  /* Which scope each endpoint needs, so a 401 can name it even on an empty body. */
  const SCOPE_FOR = {
    '/users/get_current_account': 'account_info.read',
    '/files/list_folder': 'files.metadata.read',
    '/files/list_folder/continue': 'files.metadata.read',
    '/files/download': 'files.content.read',
    '/sharing/list_shared_links': 'sharing.read',
    '/sharing/create_shared_link_with_settings': 'sharing.write',
  };

  const AUTH_URL = 'https://www.dropbox.com/oauth2/authorize';
  const TOKEN_URL = 'https://api.dropboxapi.com/oauth2/token';
  const API = 'https://api.dropboxapi.com/2';
  const CONTENT = 'https://content.dropboxapi.com/2';

  const STORE_KEY = 'merger.dropbox.appKey';
  const VERIFIER_KEY = 'merger.dropbox.verifier';
  const TOKEN_KEY = 'merger.dropbox.token';
  const REFRESH_KEY = 'merger.dropbox.refresh';
  const RESUME_KEY = 'merger.dropbox.resume';

  let token = null;
  let account = null;
  let pathRoot = null;
  let granted = null; // null = not told yet this session

  const GRANTED_KEY = 'merger.dropbox.scopes';

  /* The token response tells us what was actually granted, so a gap is visible
     immediately rather than as a 401 halfway through a run. */
  function rememberScopes(payload) {
    if (!payload || !payload.scope) return;
    granted = String(payload.scope).split(/\s+/).filter(Boolean);
    try {
      global.localStorage.setItem(GRANTED_KEY, granted.join(' '));
    } catch (err) { /* private mode */ }
  }

  /*
   * Never memoise a storage read: reconnecting with a wider grant must be seen straight
   * away, otherwise the warning outlives the problem it describes.
   */
  function grantedScopes() {
    if (granted && granted.length) return granted;
    const stored = global.localStorage.getItem(GRANTED_KEY);
    return stored ? stored.split(/\s+/).filter(Boolean) : null;
  }

  function missingScopes() {
    const have = grantedScopes();
    if (!have || !have.length) return []; // unknown, not necessarily missing
    return SCOPES.filter(function (scope) { return have.indexOf(scope) === -1; });
  }

  function redirectUri() {
    return global.location.origin + global.location.pathname;
  }

  function getConfig() {
    const stored = global.localStorage.getItem(STORE_KEY);
    const fromFile = (global.MERGER_CONFIG && global.MERGER_CONFIG.dropboxAppKey) || '';
    return { appKey: (stored || fromFile || '').trim() };
  }

  function saveConfig(appKey) {
    const value = String(appKey || '').trim();
    if (value) global.localStorage.setItem(STORE_KEY, value);
    else global.localStorage.removeItem(STORE_KEY);
  }

  function isConfigured() {
    return !!getConfig().appKey;
  }

  function getAccount() {
    return account;
  }

  /* ---------- PKCE ---------- */

  function base64Url(bytes) {
    let binary = '';
    const view = new Uint8Array(bytes);
    for (let i = 0; i < view.length; i++) binary += String.fromCharCode(view[i]);
    return global.btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }

  function randomVerifier() {
    const bytes = new Uint8Array(64);
    global.crypto.getRandomValues(bytes);
    return base64Url(bytes);
  }

  async function challengeFor(verifier) {
    const digest = await global.crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier));
    return base64Url(digest);
  }

  /*
   * Redirect rather than popup: Dropbox has no MSAL equivalent, and a redirect avoids the
   * popup-blocker failure mode. Anything the user typed is stashed so the round trip is
   * invisible to them.
   */
  async function signIn(resumeState) {
    if (!isConfigured()) throw new Error('No Dropbox app key configured.');

    const verifier = randomVerifier();
    global.sessionStorage.setItem(VERIFIER_KEY, verifier);
    if (resumeState != null) {
      global.sessionStorage.setItem(RESUME_KEY, JSON.stringify(resumeState));
    }

    const params = new URLSearchParams({
      client_id: getConfig().appKey,
      response_type: 'code',
      code_challenge: await challengeFor(verifier),
      code_challenge_method: 'S256',
      redirect_uri: redirectUri(),
      scope: SCOPES.join(' '),
      // offline returns a refresh token, so each person signs in once per device
      // rather than once per session.
      token_access_type: 'offline',
    });
    global.location.assign(AUTH_URL + '?' + params.toString());
  }

  /* Anything the user had typed before the redirect, handed back once. */
  function takeResumeState() {
    const raw = global.sessionStorage.getItem(RESUME_KEY);
    global.sessionStorage.removeItem(RESUME_KEY);
    if (!raw) return null;
    try {
      return JSON.parse(raw);
    } catch (err) {
      return null;
    }
  }

  async function exchangeCode(code) {
    const verifier = global.sessionStorage.getItem(VERIFIER_KEY);
    global.sessionStorage.removeItem(VERIFIER_KEY);
    if (!verifier) throw new Error('Sign-in could not be completed: the session expired.');

    const body = new URLSearchParams({
      code: code,
      grant_type: 'authorization_code',
      client_id: getConfig().appKey,
      code_verifier: verifier,
      redirect_uri: redirectUri(),
    });

    const response = await fetch(TOKEN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body.toString(),
    });
    if (!response.ok) {
      throw new Error('Dropbox rejected the sign-in (' + response.status + ').');
    }
    const payload = await response.json();
    token = payload.access_token;
    global.sessionStorage.setItem(TOKEN_KEY, token);
    rememberScopes(payload);
    /*
     * The refresh token persists so a teammate signs in once on their machine and never
     * again. It is that person's own delegated access, not a shared credential — Dropbox
     * still attributes every link to them. "Disconnect" clears it.
     */
    if (payload.refresh_token) {
      global.localStorage.setItem(REFRESH_KEY, payload.refresh_token);
    }
    return token;
  }

  async function refreshAccessToken() {
    const refresh = global.localStorage.getItem(REFRESH_KEY);
    if (!refresh) return null;

    const body = new URLSearchParams({
      grant_type: 'refresh_token',
      refresh_token: refresh,
      client_id: getConfig().appKey,
    });

    const response = await fetch(TOKEN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body.toString(),
    });
    if (!response.ok) {
      // Revoked or expired: drop it so the UI falls back to a normal sign-in.
      global.localStorage.removeItem(REFRESH_KEY);
      return null;
    }
    const payload = await response.json();
    token = payload.access_token;
    global.sessionStorage.setItem(TOKEN_KEY, token);
    rememberScopes(payload);
    return token;
  }

  /*
   * Call on load. Completes a redirect if we came back with a code, otherwise picks up a
   * token still held for this session. Returns the account, or null.
   */
  async function restore() {
    if (!isConfigured()) return null;

    const url = new URL(global.location.href);
    const code = url.searchParams.get('code');

    if (code) {
      // Strip the code from the address bar so a refresh cannot replay it.
      url.searchParams.delete('code');
      url.searchParams.delete('state');
      global.history.replaceState({}, '', url.toString());
      try {
        await exchangeCode(code);
      } catch (err) {
        return null;
      }
    } else {
      token = global.sessionStorage.getItem(TOKEN_KEY);
      if (!token) await refreshAccessToken();
    }

    if (!token) return null;
    try {
      account = await rpc('/users/get_current_account', null);
      adoptPathRoot(account);
      return account;
    } catch (err) {
      // An expired access token is recoverable if the refresh token still holds.
      token = null;
      global.sessionStorage.removeItem(TOKEN_KEY);
      if (await refreshAccessToken()) {
        try {
          account = await rpc('/users/get_current_account', null);
          adoptPathRoot(account);
          return account;
        } catch (retryErr) {
          token = null;
          global.sessionStorage.removeItem(TOKEN_KEY);
        }
      }
      return null;
    }
  }

  /*
   * Only worth setting when the account actually has a team space — for a personal Dropbox
   * root and home are the same namespace and the header is unnecessary.
   */
  function adoptPathRoot(info) {
    pathRoot = null;
    const root = info && info.root_info;
    if (!root || !root.root_namespace_id) return;
    if (root.root_namespace_id === root.home_namespace_id) return;
    pathRoot = JSON.stringify({ '.tag': 'root', root: root.root_namespace_id });
  }

  function signOut() {
    token = null;
    account = null;
    pathRoot = null;
    granted = null;
    global.sessionStorage.removeItem(TOKEN_KEY);
    global.localStorage.removeItem(REFRESH_KEY);
    global.localStorage.removeItem(GRANTED_KEY);
  }

  /* ---------- API ---------- */

  /*
   * A 401 from Dropbox always means the token cannot do this, and a token carries the
   * scopes it was issued with — refreshing never upgrades them. So there is nothing to
   * retry: the permission has to be granted and the user has to sign in again. Clearing
   * the stored token here is what makes the UI offer that instead of looping.
   *
   * `needed` names the scope this endpoint requires, so the message is actionable even
   * when Dropbox returns an empty body.
   */
  function authFailure(response, text, needed) {
    const declared = /"required_scope":\s*"([^"]+)"/.exec(text || '');
    const scope = declared ? declared[1] : needed;
    signOut();

    const error = new Error(
      'Dropbox refused this call (401). ' +
      (scope ? 'It needs the "' + scope + '" permission. ' : 'The sign-in is no longer valid. ') +
      'Open your Dropbox app\'s Permissions tab, tick it, click Submit, then connect again — ' +
      'a permission added after you connected does not apply to an existing sign-in.'
    );
    error.missingScope = scope || 'unknown';
    return error;
  }

  async function rpc(endpoint, body) {
    if (!token) throw new Error('Not signed in to Dropbox.');

    const headers = {
      Authorization: 'Bearer ' + token,
      'Content-Type': 'application/json',
    };
    /*
     * On a Dropbox Business team, a member's default namespace is their personal space, so
     * a team folder path like /Centific Team Folder/... resolves to not_found without this.
     * Addressing the team root makes those paths work as typed.
     */
    if (pathRoot) headers['Dropbox-API-Path-Root'] = pathRoot;

    const response = await fetch(API + endpoint, {
      method: 'POST',
      headers: headers,
      body: body == null ? 'null' : JSON.stringify(body),
    });

    if (response.status === 429) {
      const retry = Number(response.headers.get('Retry-After') || '5');
      const error = new Error('Dropbox rate limit hit. Retry in ' + retry + 's.');
      error.retryAfter = retry;
      throw error;
    }
    if (!response.ok) {
      const text = await response.text();
      if (response.status === 401) throw authFailure(response, text, SCOPE_FOR[endpoint]);
      throw new Error('Dropbox ' + endpoint + ' failed (' + response.status + '): ' + text.slice(0, 200));
    }
    return response.json();
  }

  /*
   * Recursive listing: one call plus cursor continuations beats walking level by level,
   * and cannot miss a subfolder. Paginates until has_more is false.
   */
  async function listFolderRecursive(path, onProgress) {
    const files = [];
    let page = await rpc('/files/list_folder', {
      path: path === '/' ? '' : path,
      recursive: true,
      include_non_downloadable_files: false,
    });

    for (;;) {
      page.entries.forEach(function (entry) {
        if (entry['.tag'] !== 'file') return;
        files.push({
          name: entry.name,
          path_display: entry.path_display,
          path_lower: entry.path_lower,
          id: entry.id,
          // Column Z (Date of Recording) is taken from this.
          server_modified: entry.server_modified || null,
          size: entry.size == null ? null : entry.size,
          link: '',
        });
      });
      if (onProgress) onProgress(files.length);
      if (!page.has_more) break;
      page = await rpc('/files/list_folder/continue', { cursor: page.cursor });
    }

    return files;
  }

  /*
   * One level only, for the folder browser. Deliberately does not count each child's
   * subfolders: that would be an extra API call per row and burn the rate limit on a
   * listing the user is only skimming.
   */
  async function listFolderChildren(path) {
    const folders = [];
    const files = [];
    let page = await rpc('/files/list_folder', {
      path: !path || path === '/' ? '' : path,
      recursive: false,
      include_non_downloadable_files: false,
    });

    for (;;) {
      page.entries.forEach(function (entry) {
        const item = {
          name: entry.name,
          path: entry.path_display || entry.path_lower,
        };
        if (entry['.tag'] === 'folder') folders.push(item);
        else if (entry['.tag'] === 'file') files.push(item);
      });
      if (!page.has_more) break;
      page = await rpc('/files/list_folder/continue', { cursor: page.cursor });
    }

    folders.sort(function (a, b) { return a.name.localeCompare(b.name, undefined, { numeric: true }); });
    return { folders: folders, files: files };
  }

  /*
   * Fetch a file's bytes. Used only for the metadata workbooks — the media and consent
   * files are linked, never downloaded, so participant video never touches this browser.
   */
  async function downloadFile(path) {
    if (!token) throw new Error('Not signed in to Dropbox.');

    const headers = {
      Authorization: 'Bearer ' + token,
      'Dropbox-API-Arg': JSON.stringify({ path: path }),
    };
    if (pathRoot) headers['Dropbox-API-Path-Root'] = pathRoot;

    const response = await fetch(CONTENT + '/files/download', { method: 'POST', headers: headers });

    if (!response.ok) {
      const text = await response.text();
      if (response.status === 401) throw authFailure(response, text, SCOPE_FOR['/files/download']);
      throw new Error('Could not download ' + path + ' (' + response.status + '): ' + text.slice(0, 160));
    }
    return response.arrayBuffer();
  }

  /* Read-only. Returns the existing link for a path, or '' when there is none. */
  async function existingSharedLink(path) {
    const result = await rpc('/sharing/list_shared_links', { path: path, direct_only: true });
    const links = result.links || [];
    return links.length ? links[0].url : '';
  }

  /*
   * Mutating: this publishes an anyone-with-the-link URL. Only ever called after the user
   * has seen the count and confirmed.
   */
  async function createSharedLink(path) {
    try {
      const result = await rpc('/sharing/create_shared_link_with_settings', { path: path });
      return result.url;
    } catch (err) {
      // A link created between our check and now comes back as a conflict; reuse it.
      if (/shared_link_already_exists/.test(err.message || '')) {
        return existingSharedLink(path);
      }
      throw err;
    }
  }

  /* Resolve existing links for many files, in small batches to stay under rate limits. */
  async function resolveExistingLinks(files, onProgress) {
    const BATCH = 4;
    for (let i = 0; i < files.length; i += BATCH) {
      const slice = files.slice(i, i + BATCH);
      await Promise.all(slice.map(async function (file) {
        try {
          const url = await existingSharedLink(file.path_display || file.path_lower);
          if (url) file.link = url;
        } catch (err) {
          file.linkError = err.message;
        }
      }));
      if (onProgress) onProgress(Math.min(i + BATCH, files.length), files.length);
    }
    return files;
  }

  global.DropboxSource = {
    isConfigured: isConfigured,
    getConfig: getConfig,
    saveConfig: saveConfig,
    redirectUri: redirectUri,
    signIn: signIn,
    signOut: signOut,
    restore: restore,
    getAccount: getAccount,
    missingScopes: missingScopes,
    SCOPES: SCOPES,
    takeResumeState: takeResumeState,
    listFolderRecursive: listFolderRecursive,
    listFolderChildren: listFolderChildren,
    downloadFile: downloadFile,
    existingSharedLink: existingSharedLink,
    createSharedLink: createSharedLink,
    resolveExistingLinks: resolveExistingLinks,
  };
})(window);

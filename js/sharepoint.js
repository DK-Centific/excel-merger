/* SharePoint / OneDrive fetch via Microsoft Graph, using MSAL browser PKCE (no client secret). */
(function (global) {
  'use strict';

  const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';
  const SCOPES = ['Files.Read.All', 'Sites.Read.All'];
  const EXCEL_PATTERN = /\.(xlsx|xlsm)$/i;

  let msalInstance = null;
  let activeAccount = null;

  function getConfig() {
    const stored = {};
    try {
      stored.clientId = localStorage.getItem('merger.clientId') || '';
      stored.tenantId = localStorage.getItem('merger.tenantId') || '';
    } catch (err) {
      // Private browsing can block localStorage; fall back to the bundled config.
    }
    const base = global.MERGER_CONFIG || {};
    return {
      clientId: stored.clientId || base.clientId || '',
      tenantId: stored.tenantId || base.tenantId || 'common',
    };
  }

  function saveConfig(clientId, tenantId) {
    localStorage.setItem('merger.clientId', clientId || '');
    localStorage.setItem('merger.tenantId', tenantId || '');
    msalInstance = null;
    activeAccount = null;
  }

  function isConfigured() {
    return !!getConfig().clientId;
  }

  async function getMsal() {
    if (msalInstance) return msalInstance;

    const config = getConfig();
    if (!config.clientId) {
      throw new Error('No Azure app client ID is configured. Open Settings and add one, or use file upload instead.');
    }
    if (!global.msal || !global.msal.PublicClientApplication) {
      throw new Error('The MSAL library failed to load, so SharePoint sign-in is unavailable.');
    }

    msalInstance = new global.msal.PublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: 'https://login.microsoftonline.com/' + (config.tenantId || 'common'),
        redirectUri: global.location.origin + global.location.pathname,
      },
      cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false },
    });

    await msalInstance.initialize();
    await msalInstance.handleRedirectPromise();

    const accounts = msalInstance.getAllAccounts();
    if (accounts.length) activeAccount = accounts[0];

    return msalInstance;
  }

  /*
   * Pick up an account MSAL already holds for this session, so a page reload does not
   * look like a sign-out. Returns null when not configured or not signed in.
   */
  async function restore() {
    if (!isConfigured()) return null;
    try {
      const instance = await getMsal();
      const accounts = instance.getAllAccounts();
      if (!accounts.length) return null;
      activeAccount = accounts[0];
      instance.setActiveAccount(activeAccount);
      return activeAccount;
    } catch (err) {
      return null;
    }
  }

  async function signIn(forceAccountPicker) {
    const instance = await getMsal();
    const request = { scopes: SCOPES };
    if (forceAccountPicker) request.prompt = 'select_account';
    const result = await instance.loginPopup(request);
    activeAccount = result.account;
    instance.setActiveAccount(activeAccount);
    return activeAccount;
  }

  async function signOut() {
    if (!msalInstance || !activeAccount) return;
    await msalInstance.logoutPopup({ account: activeAccount });
    activeAccount = null;
  }

  function getAccount() {
    return activeAccount;
  }

  async function getToken() {
    const instance = await getMsal();
    if (!activeAccount) {
      const accounts = instance.getAllAccounts();
      if (accounts.length) activeAccount = accounts[0];
    }
    if (!activeAccount) await signIn();

    try {
      const result = await instance.acquireTokenSilent({ scopes: SCOPES, account: activeAccount });
      return result.accessToken;
    } catch (err) {
      const result = await instance.acquireTokenPopup({ scopes: SCOPES, account: activeAccount });
      return result.accessToken;
    }
  }

  function toBase64(text) {
    const bytes = new TextEncoder().encode(text);
    let binary = '';
    bytes.forEach(function (b) { binary += String.fromCharCode(b); });
    return btoa(binary);
  }

  /* Graph accepts any SharePoint or OneDrive URL encoded as a sharing token. */
  function encodeSharingUrl(url) {
    return 'u!' + toBase64(url.trim()).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  }

  async function graphGet(url, token) {
    const response = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    if (!response.ok) {
      let detail = '';
      try {
        const body = await response.json();
        detail = body && body.error && body.error.message ? body.error.message : '';
      } catch (err) {
        detail = response.statusText;
      }
      const error = new Error('Graph request failed (' + response.status + ')' + (detail ? ': ' + detail : ''));
      error.status = response.status;
      throw error;
    }
    return response.json();
  }

  async function resolveDriveItem(url, token) {
    const shareId = encodeSharingUrl(url);
    return graphGet(GRAPH_ROOT + '/shares/' + shareId + '/driveItem?$select=id,name,size,folder,file,parentReference,@microsoft.graph.downloadUrl', token);
  }

  async function listFolderChildren(driveId, itemId, token) {
    const results = [];
    let next = GRAPH_ROOT + '/drives/' + driveId + '/items/' + itemId +
      '/children?$select=id,name,size,folder,file,parentReference,@microsoft.graph.downloadUrl&$top=200';

    while (next) {
      const page = await graphGet(next, token);
      (page.value || []).forEach(function (child) { results.push(child); });
      next = page['@odata.nextLink'] || null;
    }
    return results;
  }

  async function downloadItem(item, token) {
    const preAuthUrl = item['@microsoft.graph.downloadUrl'];
    if (preAuthUrl) {
      const response = await fetch(preAuthUrl);
      if (response.ok) return response.arrayBuffer();
    }

    const driveId = item.parentReference && item.parentReference.driveId;
    if (!driveId) throw new Error('Could not determine which drive "' + item.name + '" lives in.');

    const response = await fetch(GRAPH_ROOT + '/drives/' + driveId + '/items/' + item.id + '/content', {
      headers: { Authorization: 'Bearer ' + token },
    });
    if (!response.ok) throw new Error('Download failed for "' + item.name + '" (' + response.status + ').');
    return response.arrayBuffer();
  }

  /*
   * Turn a list of pasted URLs into downloaded workbooks.
   * Folder links expand to every .xlsx/.xlsm inside them.
   * Returns { files: [{name, buffer, source}], errors: [{url, message}] }.
   */
  async function fetchFromUrls(urls, onProgress) {
    const token = await getToken();
    const files = [];
    const errors = [];

    for (let i = 0; i < urls.length; i++) {
      const url = urls[i];
      if (onProgress) onProgress(i, urls.length, url);

      try {
        const item = await resolveDriveItem(url, token);

        if (item.folder) {
          const children = await listFolderChildren(item.parentReference.driveId, item.id, token);
          const workbooks = children.filter(function (child) { return child.file && EXCEL_PATTERN.test(child.name); });

          if (!workbooks.length) {
            errors.push({ url: url, message: 'Folder "' + item.name + '" contains no .xlsx or .xlsm files.' });
            continue;
          }
          for (let j = 0; j < workbooks.length; j++) {
            const child = workbooks[j];
            if (onProgress) onProgress(i, urls.length, child.name);
            files.push({
              name: child.name,
              buffer: await downloadItem(child, token),
              source: 'SharePoint: ' + item.name,
            });
          }
        } else if (EXCEL_PATTERN.test(item.name)) {
          files.push({ name: item.name, buffer: await downloadItem(item, token), source: 'SharePoint' });
        } else {
          errors.push({ url: url, message: '"' + item.name + '" is not an .xlsx or .xlsm file.' });
        }
      } catch (err) {
        errors.push({ url: url, message: err && err.message ? err.message : String(err) });
      }
    }

    return { files: files, errors: errors };
  }

  global.SharePointSource = {
    isConfigured: isConfigured,
    getConfig: getConfig,
    saveConfig: saveConfig,
    restore: restore,
    signIn: signIn,
    signOut: signOut,
    getAccount: getAccount,
    fetchFromUrls: fetchFromUrls,
  };
})(window);

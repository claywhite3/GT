/* =========================================================================
   Vusion TopStock — Ground Truth Collection
   Shared application module: data definitions + storage layer
   =========================================================================

   STORAGE LAYER
   -------------
   All record persistence goes through saveRecord(). It appends rows to a
   SharePoint-hosted Excel workbook through Microsoft Graph. Nothing else in
   the app touches the backend directly.

   Each row in the sheet = ONE out-of-shelf scan, with these columns:
     timestamp | email | storeType | storeAddress | cameraBarcode |
     outOfShelfBarcode | rootCause
   ========================================================================= */

/* ---- CONFIG ------------------------------------------------------------ */
// Microsoft Graph / SharePoint Excel backend. Register this browser app as an
// Entra ID SPA; do not add a client secret in a browser app.
const SHAREPOINT_EXCEL_CONFIG = {
  tenantId: 'YOUR_TENANT_ID_OR_DOMAIN',
  clientId: 'YOUR_ENTRA_APP_CLIENT_ID',
  redirectUri: '', // Blank = current page URL. Register the exact URI(s) in Entra ID.

  // Ask your Microsoft 365 admin whether Files.ReadWrite.All plus user/site
  // access is sufficient, or whether a narrower site-scoped setup is available.
  scopes: ['Sites.ReadWrite.All'],

  graphBaseUrl: 'https://graph.microsoft.com/v1.0',
  msalCdnUrl: 'https://alcdn.msauth.net/browser/3.30.0/js/msal-browser.min.js',

  // Site lookup: either paste siteId directly, or set hostname + path.
  siteId: '',
  siteHostname: 'YOUR_TENANT.sharepoint.com',
  sitePath: '/sites/YOUR_SITE',

  // Workbook lookup: paste driveItemId directly, resolve by fileWebUrl, or
  // resolve by filePath relative to the default SharePoint document library.
  driveId: '', // Optional; set for a non-default document library.
  driveItemId: '',
  fileWebUrl: 'https://vusion365-my.sharepoint.com/:x:/g/personal/clay_white_vusion_com/IQAsBgq8iD3bS7vbrb8vxcp6AVdLxK7x7X4zuI-qSvq7v-I?e=LpaxYq',
  filePath: '/GroundTruth.xlsx',

  tableName: 'GroundTruth',
};

/* ---- STORE DATA -------------------------------------------------------- */
const STORE_TYPES = [
  {
    id: 'sainsburys',
    name: "Sainsbury's",
    color: '#F06C00',
    initial: 'S',
  },
  {
    id: 'lidl',
    name: 'Lidl',
    color: '#0050AA',
    initial: 'L',
  },
];

// Individual store locations, keyed by store type id.
// Add more entries here as needed — the UI updates automatically.
const STORES_BY_TYPE = {
  sainsburys: [
    { id: 'sby-witney', name: 'Sainsbury\u2019s Witney', addr: 'Witan Way\nWitney OX28 4FF' },
  ],
  lidl: [
    // Add Lidl stores here, e.g.
    // { id: 'lidl-xxxx', name: 'Lidl ...', addr: '...\n...' },
  ],
};

/* ---- ROOT CAUSES ------------------------------------------------------- */
// From the supplied list (percentages were sample data and are omitted).
const ROOT_CAUSES = [
  'Neighbouring product drifted into slot',
  'Detection picking up adjacent slot product',
  'Planogram changes after Blink ID',
  'Model can\u2019t recognise pallet / down arrows',
  'Camera moved, missing, or blind spot',
  'Slot not photographed (no camera coverage)',
  'Obstruction in front of shelf',
  'Outdated snapshot',
  'Captana failed to detect (underdetection)',
  'ESL not on correct shelf edge',
  'Detection picking up packaging, not SKU',
];

/* ---- SESSION HELPERS --------------------------------------------------- */
const Session = {
  get: (k, fallback = '') => sessionStorage.getItem(k) || fallback,
  set: (k, v) => sessionStorage.setItem(k, v),
  del: (k) => sessionStorage.removeItem(k),
  // Guard: redirect to login if no email recorded yet.
  requireLogin() {
    if (!sessionStorage.getItem('email')) {
      window.location.href = 'index.html';
    }
  },
  initials() {
    const email = sessionStorage.getItem('email') || '';
    const namePart = email.split('@')[0] || '';
    const letters = namePart.replace(/[^a-zA-Z]/g, '');
    return (letters.slice(0, 2) || 'GT').toUpperCase();
  },
};

/* ---- RECORD QUEUE + STORAGE -------------------------------------------- */
/*
   Records are appended to a local queue immediately (so nothing is lost on a
   flaky connection) and flushed to the backend. Each record represents one
   out-of-shelf scan with its single root cause.
*/
const Records = {
  QUEUE_KEY: 'gt_pending_records',

  _readQueue() {
    try { return JSON.parse(localStorage.getItem(this.QUEUE_KEY) || '[]'); }
    catch (e) { return []; }
  },
  _writeQueue(q) {
    localStorage.setItem(this.QUEUE_KEY, JSON.stringify(q));
  },

  // Build a record from current session + the issue specifics.
  build({ outOfShelfBarcode, rootCause }) {
    return {
      timestamp: new Date().toISOString(),
      email: Session.get('email'),
      storeType: Session.get('storeTypeName'),
      storeAddress: (Session.get('storeName') + ' \u2014 ' + Session.get('storeAddr')).trim(),
      cameraBarcode: Session.get('cameraBarcode'),
      outOfShelfBarcode: outOfShelfBarcode || '',
      rootCause: rootCause || '',
    };
  },
};

/* ---- MICROSOFT GRAPH AUTH + EXCEL HELPERS ------------------------------ */
const GRAPH_ROW_FIELDS = [
  'timestamp',
  'email',
  'storeType',
  'storeAddress',
  'cameraBarcode',
  'outOfShelfBarcode',
  'rootCause',
];

let msalScriptPromise = null;
let msalClientPromise = null;
let sharePointTargetPromise = null;
let postRedirectFlushScheduled = false;

function hasConfigValue(value) {
  if (typeof value !== 'string') return Boolean(value);
  const trimmed = value.trim();
  return Boolean(trimmed) && !trimmed.includes('YOUR_') && !/^\{.+\}$/.test(trimmed);
}

function isSharePointExcelConfigured() {
  const cfg = SHAREPOINT_EXCEL_CONFIG;
  const hasSiteId = hasConfigValue(cfg.siteId);
  const hasSiteLookup = hasConfigValue(cfg.siteHostname) && hasConfigValue(cfg.sitePath);
  const hasWorkbookId = hasConfigValue(cfg.driveItemId) && (hasSiteId || hasConfigValue(cfg.driveId));
  const hasWorkbookPath = hasConfigValue(cfg.filePath) && (hasSiteId || hasSiteLookup || hasConfigValue(cfg.driveId));
  const hasWorkbookUrl = hasConfigValue(cfg.fileWebUrl);

  return (
    hasConfigValue(cfg.tenantId) &&
    hasConfigValue(cfg.clientId) &&
    Array.isArray(cfg.scopes) &&
    cfg.scopes.length > 0 &&
    hasConfigValue(cfg.tableName) &&
    (hasWorkbookId || hasWorkbookPath || hasWorkbookUrl)
  );
}

function normalizeGraphPath(path) {
  const trimmed = String(path || '').trim();
  if (!trimmed) return '';
  return trimmed.startsWith('/') ? trimmed : `/${trimmed}`;
}

function encodeGraphPath(path) {
  return normalizeGraphPath(path)
    .split('/')
    .map((part) => (part ? encodeURIComponent(part) : ''))
    .join('/');
}

function getRedirectUri() {
  if (hasConfigValue(SHAREPOINT_EXCEL_CONFIG.redirectUri)) {
    return SHAREPOINT_EXCEL_CONFIG.redirectUri.trim();
  }
  return window.location.href.split('#')[0].split('?')[0];
}

function loadMsalBrowser() {
  if (window.msal && window.msal.PublicClientApplication) {
    return Promise.resolve(window.msal);
  }
  if (msalScriptPromise) return msalScriptPromise;

  msalScriptPromise = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = SHAREPOINT_EXCEL_CONFIG.msalCdnUrl;
    script.async = true;
    script.onload = () => resolve(window.msal);
    script.onerror = () => reject(new Error('Unable to load MSAL.js from CDN.'));
    document.head.appendChild(script);
  });

  return msalScriptPromise;
}

async function getMsalClient() {
  if (!isSharePointExcelConfigured()) return null;
  if (msalClientPromise) return msalClientPromise;

  msalClientPromise = (async () => {
    const msalBrowser = await loadMsalBrowser();
    const client = new msalBrowser.PublicClientApplication({
      auth: {
        clientId: SHAREPOINT_EXCEL_CONFIG.clientId.trim(),
        authority: `https://login.microsoftonline.com/${SHAREPOINT_EXCEL_CONFIG.tenantId.trim()}`,
        redirectUri: getRedirectUri(),
      },
      cache: {
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: false,
      },
    });

    if (typeof client.initialize === 'function') {
      await client.initialize();
    }

    const redirectResult = await client.handleRedirectPromise();
    if (redirectResult && redirectResult.account) {
      client.setActiveAccount(redirectResult.account);
      schedulePostRedirectFlush();
    } else {
      const [account] = client.getAllAccounts();
      if (account && !client.getActiveAccount()) client.setActiveAccount(account);
    }

    return client;
  })();

  return msalClientPromise;
}

function schedulePostRedirectFlush() {
  if (postRedirectFlushScheduled) return;
  postRedirectFlushScheduled = true;
  window.setTimeout(() => { flushRecords(); }, 0);
}

function shouldUseInteractiveAuth(error) {
  const msalBrowser = window.msal || {};
  const code = String(error && (error.errorCode || error.error || error.message || ''));
  return (
    (msalBrowser.InteractionRequiredAuthError && error instanceof msalBrowser.InteractionRequiredAuthError) ||
    /interaction_required|consent_required|login_required|no_account/i.test(code)
  );
}

async function acquireGraphAccessToken() {
  const client = await getMsalClient();
  if (!client) return null;

  const account = client.getActiveAccount() || client.getAllAccounts()[0];
  const tokenRequest = {
    scopes: SHAREPOINT_EXCEL_CONFIG.scopes,
  };
  if (account) tokenRequest.account = account;

  if (account) {
    try {
      const response = await client.acquireTokenSilent(tokenRequest);
      if (response.account) client.setActiveAccount(response.account);
      return response.accessToken;
    } catch (error) {
      if (!shouldUseInteractiveAuth(error)) {
        console.warn('[auth] Silent token acquisition failed:', error);
        return null;
      }
    }
  }

  try {
    const response = await client.acquireTokenPopup(tokenRequest);
    if (response.account) client.setActiveAccount(response.account);
    return response.accessToken;
  } catch (popupError) {
    console.warn('[auth] Popup token acquisition failed; falling back to redirect:', popupError);
    try {
      await client.acquireTokenRedirect({
        ...tokenRequest,
        redirectStartPage: window.location.href,
      });
    } catch (redirectError) {
      console.warn('[auth] Redirect token acquisition failed:', redirectError);
    }
    return null;
  }
}

async function graphFetch(path, accessToken, options = {}) {
  const url = path.startsWith('https://') ? path : `${SHAREPOINT_EXCEL_CONFIG.graphBaseUrl}${path}`;
  const headers = new Headers(options.headers || {});
  headers.set('Authorization', `Bearer ${accessToken}`);
  if (options.body && !headers.has('Content-Type')) {
    headers.set('Content-Type', 'application/json');
  }

  return fetch(url, {
    ...options,
    headers,
  });
}

async function graphJson(path, accessToken, options = {}) {
  const response = await graphFetch(path, accessToken, options);
  if (!response.ok) {
    const details = await response.text().catch(() => '');
    throw new Error(`Graph request failed (${response.status}): ${details}`);
  }
  return response.json();
}

function toBase64Url(value) {
  const bytes = new TextEncoder().encode(value);
  let binary = '';
  bytes.forEach((byte) => { binary += String.fromCharCode(byte); });
  return btoa(binary).replace(/=+$/g, '').replace(/\+/g, '-').replace(/\//g, '_');
}

async function resolveSiteId(accessToken) {
  const cfg = SHAREPOINT_EXCEL_CONFIG;
  const site = await graphJson(
    `/sites/${encodeURIComponent(cfg.siteHostname.trim())}:${encodeGraphPath(cfg.sitePath)}?$select=id`,
    accessToken
  );
  return site.id;
}

async function resolveDriveItemFromWebUrl(accessToken) {
  const shareId = `u!${toBase64Url(SHAREPOINT_EXCEL_CONFIG.fileWebUrl.trim())}`;
  return graphJson(
    `/shares/${shareId}/driveItem?$select=id,name,parentReference`,
    accessToken
  );
}

async function resolveDriveItemId(accessToken, siteId) {
  const cfg = SHAREPOINT_EXCEL_CONFIG;
  const filePath = encodeGraphPath(cfg.filePath);
  const path = hasConfigValue(cfg.driveId)
    ? `/drives/${encodeURIComponent(cfg.driveId.trim())}/root:${filePath}?$select=id,name`
    : `/sites/${encodeURIComponent(siteId)}/drive/root:${filePath}?$select=id,name`;
  const item = await graphJson(path, accessToken);
  return item.id;
}

async function resolveSharePointExcelTarget(accessToken) {
  if (sharePointTargetPromise) return sharePointTargetPromise;

  sharePointTargetPromise = (async () => {
    const cfg = SHAREPOINT_EXCEL_CONFIG;

    if (hasConfigValue(cfg.fileWebUrl)) {
      const item = await resolveDriveItemFromWebUrl(accessToken);
      return {
        siteId: item.parentReference && item.parentReference.siteId,
        driveId: item.parentReference && item.parentReference.driveId,
        itemId: item.id,
      };
    }

    const siteId = hasConfigValue(cfg.siteId) ? cfg.siteId.trim() : await resolveSiteId(accessToken);
    const itemId = hasConfigValue(cfg.driveItemId) ? cfg.driveItemId.trim() : await resolveDriveItemId(accessToken, siteId);

    return {
      siteId,
      driveId: hasConfigValue(cfg.driveId) ? cfg.driveId.trim() : '',
      itemId,
    };
  })();

  return sharePointTargetPromise;
}

function buildRowsAddPath(target) {
  const tableName = encodeURIComponent(SHAREPOINT_EXCEL_CONFIG.tableName.trim());
  const itemId = encodeURIComponent(target.itemId);

  if ((hasConfigValue(SHAREPOINT_EXCEL_CONFIG.driveId) || !target.siteId) && target.driveId) {
    return `/drives/${encodeURIComponent(target.driveId)}/items/${itemId}/workbook/tables/${tableName}/rows/add`;
  }

  return `/sites/${encodeURIComponent(target.siteId)}/drive/items/${itemId}/workbook/tables/${tableName}/rows/add`;
}

/*
   saveRecord(record) -> Promise<{ok:boolean, queued:boolean}>
   Public entry point. Always queues locally, then tries to flush.
*/
async function saveRecord(record) {
  const q = Records._readQueue();
  q.push(record);
  Records._writeQueue(q);
  const result = await flushRecords();
  return { ok: result.flushedAll, queued: !result.flushedAll };
}

/*
   flushRecords() -> attempts to send every queued record to the backend.
   Removes successfully-sent records from the queue.
*/
async function flushRecords() {
  let q = Records._readQueue();
  if (q.length === 0) return { flushedAll: true, remaining: 0 };

  if (!isSharePointExcelConfigured()) {
    // No backend configured yet — keep everything queued locally.
    console.warn('[storage] SharePoint Excel backend is not configured. Records held in local queue:', q.length);
    return { flushedAll: false, remaining: q.length };
  }

  const stillPending = [];
  for (const rec of q) {
    try {
      const ok = await pushToBackend(rec);
      if (!ok) stillPending.push(rec);
    } catch (e) {
      console.error('[storage] push failed, re-queueing:', e);
      stillPending.push(rec);
    }
  }
  Records._writeQueue(stillPending);
  return { flushedAll: stillPending.length === 0, remaining: stillPending.length };
}

/* -------------------------------------------------------------------------
   pushToBackend(record) -> Promise<boolean>

   Appends exactly one row to the configured Excel table. The record shape stays
   identical, and a false return leaves the record in the existing local queue.
   ------------------------------------------------------------------------- */
async function pushToBackend(record) {
  if (!isSharePointExcelConfigured()) return false;

  try {
    const accessToken = await acquireGraphAccessToken();
    if (!accessToken) return false;

    const target = await resolveSharePointExcelTarget(accessToken);
    const values = [GRAPH_ROW_FIELDS.map((field) => record[field] || '')];
    const response = await graphFetch(buildRowsAddPath(target), accessToken, {
      method: 'POST',
      body: JSON.stringify({ values }),
    });

    if (!response.ok) {
      const details = await response.text().catch(() => '');
      console.warn('[storage] Graph row append failed:', response.status, details);
      if (response.status === 400 || response.status === 404) sharePointTargetPromise = null;
      return false;
    }

    return true;
  } catch (error) {
    console.error('[storage] Graph push failed:', error);
    sharePointTargetPromise = null;
    return false;
  }
}

/* Initialize auth on load; try to flush queued records whenever network returns. */
if (typeof window !== 'undefined') {
  window.addEventListener('load', () => {
    getMsalClient().catch((error) => {
      console.warn('[auth] MSAL initialization failed:', error);
    });
  });
  window.addEventListener('online', () => { flushRecords(); });
}

/* ---- SERVICE WORKER REGISTRATION --------------------------------------- */
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('service-worker.js').catch((e) => {
      console.warn('[pwa] SW registration failed:', e);
    });
  });
}

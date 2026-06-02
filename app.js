/* =========================================================================
   Vusion TopStock — Ground Truth Collection
   Shared application module: data definitions + storage layer
   =========================================================================

   STORAGE LAYER
   -------------
   All record persistence goes through saveRecord(). Right now it posts to a
   Google Apps Script web app (Google Sheet backend). To move to Microsoft
   Excel on OneDrive later, you only need to replace the body of
   `pushToBackend()` with a Microsoft Graph call — nothing else in the app
   touches the backend directly.

   Each row in the sheet = ONE out-of-shelf scan, with these columns:
     timestamp | email | storeType | storeAddress | cameraBarcode |
     outOfShelfBarcode | rootCause
   ========================================================================= */

/* ---- CONFIG ------------------------------------------------------------ */
// Paste your deployed Google Apps Script Web App URL here.
// (Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone)
const GOOGLE_SHEET_ENDPOINT = 'https://script.google.com/macros/s/AKfycbzksD9-gUojfwkC6Ksu3gTdNR6Lt5cbGFJKNp9Oz-w-EnFgjEGyoT2mQGePtNA1xOoQ/exec';

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

  if (!GOOGLE_SHEET_ENDPOINT) {
    // No backend configured yet — keep everything queued locally.
    console.warn('[storage] No GOOGLE_SHEET_ENDPOINT set. Records held in local queue:', q.length);
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

   *** THIS IS THE ONLY FUNCTION YOU SWAP FOR MICROSOFT GRAPH / EXCEL ***

   Current implementation: POST JSON to a Google Apps Script web app, which
   appends a row to the Google Sheet.

   To move to Excel on OneDrive, replace the body with a Microsoft Graph
   call to:
     POST /me/drive/items/{file-id}/workbook/tables/{table}/rows/add
   using an access token acquired via MSAL. The record shape stays identical.
   ------------------------------------------------------------------------- */
async function pushToBackend(record) {
  // Google Apps Script web apps don't return CORS headers, so we send as a
  // simple request and treat a resolved fetch as success. The Apps Script
  // should parse e.postData.contents as JSON and append a row.
  await fetch(GOOGLE_SHEET_ENDPOINT, {
    method: 'POST',
    mode: 'no-cors',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(record),
  });
  // With no-cors we can't read the response; assume success if no throw.
  return true;
}

/* Try to flush any leftover records whenever the app loads + regains network. */
if (typeof window !== 'undefined') {
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

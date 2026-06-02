# Vusion TopStock — Ground Truth PWA

A mobile PWA for collecting out-of-shelf ground truth: log in, pick a customer
and store, scan a Captana camera, then loop through scanning out-of-shelf
locations and tagging each with a single root cause. Records are written one
row per out-of-shelf scan.

## Pages / flow

```
index.html         Login  (any valid email + 6+ char password; no user DB)
   ↓
store_type.html    Pick customer: Sainsbury's / Lidl
   ↓
store_select.html  Pick the specific store (searchable list)
   ↓
scan_camera.html   Scan the Captana camera barcode  ← loop returns here per camera
   ↓
scan_oos.html      Scan an out-of-shelf barcode       ← loop returns here per issue
   ↓
root_cause.html    Pick ONE root cause, then:
                     • "Next Issue"  → save + back to scan_oos (same camera)
                     • "All Done"    → confirm → save + back to scan_camera
```

Every scan screen has: rear camera by default, a **flip** button to the
front camera (for reaching under shelves), a torch toggle (Android, where
supported), an **Enter manually** link beneath the window, and a **Retry scan**
link once a code is captured. 1D barcodes are prioritised; QR is allowed as a
fallback.

## Data model

One spreadsheet row per out-of-shelf scan:

| Timestamp | Email | Store Type | Store Address | Camera Barcode | Out-of-Shelf Barcode | Root Cause |
|-----------|-------|------------|---------------|----------------|----------------------|------------|

Records are queued in `localStorage` first, then flushed to the backend, so a
flaky connection never loses data. Leftover records re-flush on load and when
the network returns.

## Setup (Google Sheet backend, current)

1. Create a Google Sheet → **Extensions → Apps Script**.
2. Paste `google-apps-script.gs`, run `setupHeaders` once.
3. **Deploy → New deployment → Web app** (Execute as: *Me*, Access: *Anyone*).
   Copy the Web app URL.
4. Open `app.js`, set `GOOGLE_SHEET_ENDPOINT` to that URL.
5. Serve the folder over **HTTPS** (camera access requires a secure context).
   Any static host works — GitHub Pages, Netlify, etc. `localhost` is also
   treated as secure for testing.

## Adding stores / root causes

Everything data-related lives at the top of `app.js`:
- `STORE_TYPES` — the customer cards.
- `STORES_BY_TYPE` — individual stores per customer. Currently one Sainsbury's
  (Witney). Add objects `{ id, name, addr }` here; the UI updates automatically.
- `ROOT_CAUSES` — the 11 cards on the root-cause grid.

## Moving to Excel on OneDrive (later, on Codex)

The backend is isolated to ONE function: `pushToBackend(record)` in `app.js`.
Nothing else touches the network. To switch to Microsoft Excel:

1. Register an app in Azure AD (Entra ID); add MSAL.js for interactive sign-in
   and request the `Files.ReadWrite` scope.
2. Replace the body of `pushToBackend` with a Microsoft Graph call:
   `POST /me/drive/items/{file-id}/workbook/tables/{table}/rows/add`
   passing the same record fields as a row.
3. The record shape, queue, retry, and all UI stay exactly as they are.

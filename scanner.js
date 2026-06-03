/* =========================================================================
   Shared scanner controller
   Used by scan_camera.html and scan_oos.html.

   ENGINE SELECTION
   ----------------
   1. Scandit Data Capture SDK v8  (primary) — used when window.SCANDIT_LICENSE_KEY
      is set (see config.js) AND the SDK loads successfully.
   2. html5-qrcode                 (fallback) — used automatically if Scandit has
      no key or fails to load.
   ========================================================================= */
const ScanController = (function () {
  const SCANNER_VERSION = 'scanner v0.21.0';
  let opts = {};
  let scanCompleted = false;
  let lastInvalidShake = 0;
  let engine = null;            // 'scandit' | 'html5'
  let facingMode = 'environment';
  let torchOn = false;

  let html5QrCode = null;
  let currentTrack = null;

  let sdcContext = null;
  let sdcView = null;
  let sdcBarcodeCapture = null;
  let sdcCamera = null;
  let SDC = null;
  let SDCBarcode = null;

  const successSound = new Audio('https://assets.mixkit.co/active_storage/sfx/2013/2013-preview.mp3');
  function $(id) { return document.getElementById(id); }

  function log(msg) {
    const line = '[scanner] ' + msg;
    console.log(line);
    const panel = $('scanLog');
    if (panel) {
      panel.style.display = 'block';
      const t = new Date().toISOString().slice(11, 19);
      panel.textContent += t + '  ' + msg + '\n';
      panel.scrollTop = panel.scrollHeight;
    }
  }

  function H5Formats() {
    return [
      Html5QrcodeSupportedFormats.CODE_128,
      Html5QrcodeSupportedFormats.CODE_39,
      Html5QrcodeSupportedFormats.CODE_93,
      Html5QrcodeSupportedFormats.EAN_13,
      Html5QrcodeSupportedFormats.EAN_8,
      Html5QrcodeSupportedFormats.UPC_A,
      Html5QrcodeSupportedFormats.UPC_E,
      Html5QrcodeSupportedFormats.ITF,
      Html5QrcodeSupportedFormats.CODABAR,
      Html5QrcodeSupportedFormats.QR_CODE,
      Html5QrcodeSupportedFormats.DATA_MATRIX,
    ];
  }

  async function init(options) {
    options = options || {};
    log(SCANNER_VERSION + ' init; onComplete=' + (typeof options.onComplete));
    if (!options.onComplete && opts.onComplete) {
      log('init called without onComplete; keeping existing one');
      options.onComplete = opts.onComplete;
    }
    opts = options;
    if (opts.manualTitle) $('manualTitle').textContent = opts.manualTitle;
    if (opts.manualSub) $('manualSub').textContent = opts.manualSub;
    if (opts.manualPlaceholder) $('manualInput').placeholder = opts.manualPlaceholder;
    await start();
  }

  function resetUiForStart() {
    scanCompleted = false;
    $('completeScreen').style.display = 'none';
    $('rescanBtn').classList.add('hidden');
    $('manualEntryBtn').classList.remove('hidden');
    $('cameraState').classList.remove('hidden');
    $('cameraRetryBtn').classList.add('hidden');
    $('cameraStateText').textContent = 'Starting camera\u2026';
  }

  let starting = false;
  async function start() {
    if (starting) { log('start() ignored \u2014 already starting'); return; }
    starting = true;
    try {
      resetUiForStart();
      await stop();

      const key = (typeof window !== 'undefined' && window.SCANDIT_LICENSE_KEY) ? window.SCANDIT_LICENSE_KEY : '';
      if (key) {
        log('Scandit key present (length ' + key.length + '). Attempting Scandit\u2026');
        try {
          await startScandit(key);
          engine = 'scandit';
          log('Scandit started OK \u2713');
          $('cameraState').classList.add('hidden');
          return;
        } catch (e) {
          log('Scandit FAILED: ' + (e && e.message ? e.message : String(e)));
          console.warn('[scanner] Scandit failed, falling back to html5-qrcode:', e);
          await stopScandit();
        }
      } else {
        log('No Scandit key set \u2014 using html5-qrcode fallback.');
      }
      await startHtml5();
      engine = 'html5';
      log('Using html5-qrcode engine.');
    } finally {
      starting = false;
    }
  }

  function handleSuccess(decodedText) {
    if (scanCompleted) return;
    const value = String(decodedText || '').trim();
    if (!value) { shake(); return; }
    scanCompleted = true;
    log('handleSuccess: ' + value);
    complete(value);
  }

  function shake() {
    const now = Date.now();
    if (now - lastInvalidShake < 600) return;
    lastInvalidShake = now;
    const card = $('scannerBox');
    card.classList.add('shake');
    setTimeout(() => card.classList.remove('shake'), 400);
  }

  async function complete(value) {
    successSound.play().catch(() => {});
    if (navigator.vibrate) navigator.vibrate([100, 50, 100]);

    $('completeScreen').style.display = 'flex';
    $('rescanBtn').classList.remove('hidden');
    $('manualEntryBtn').classList.add('hidden');

    $('idCard').classList.add('found');
    $('idValue').textContent = value;
    $('idBadge').textContent = 'Scanned \u2713';

    const checkEl = $('completeCheck');
    if (checkEl) { checkEl.style.animation = 'none'; void checkEl.offsetWidth; checkEl.style.animation = ''; }

    if (typeof opts.onComplete === 'function') {
      log('firing onComplete for: ' + value);
      try {
        opts.onComplete(value);
        const nb = document.getElementById('nextBtn');
        log('nextBtn disabled now? ' + (nb ? nb.classList.contains('disabled') : 'no nextBtn'));
      } catch (e) { log('onComplete error: ' + e); }
    } else {
      log('NO onComplete callback registered');
    }

    try { await stop(); } catch (e) { log('stop after complete failed: ' + e); }
  }

  /* ---- SCANDIT engine --------------------------------------------------- */
  async function startScandit(licenseKey) {
    log('importing core module\u2026');
    if (!SDC) SDC = await import('@scandit/web-datacapture-core');
    log('core OK; importing barcode module\u2026');
    if (!SDCBarcode) SDCBarcode = await import('@scandit/web-datacapture-barcode');
    log('barcode module OK');

    const { DataCaptureView, Camera, DataCaptureContext, FrameSourceState } = SDC;
    const { barcodeCaptureLoader, BarcodeCaptureSettings, BarcodeCapture, Symbology } = SDCBarcode;

    const mount = $('qrReader');
    mount.innerHTML = '';
    sdcView = new DataCaptureView();
    sdcView.connectToElement(mount);
    log('view mounted; creating context (loading WASM)\u2026');
    try {
      const st = $('cameraStateText');
      if (st) st.textContent = 'Loading scanner\u2026 (first launch can take a few seconds)';
    } catch (e) {}

    sdcContext = await DataCaptureContext.forLicenseKey(licenseKey, {
      libraryLocation: 'https://cdn.jsdelivr.net/npm/@scandit/web-datacapture-barcode@8/sdc-lib/',
      moduleLoaders: [barcodeCaptureLoader()],
    });
    log('context created; setting view context\u2026');
    await sdcView.setContext(sdcContext);

    log('selecting camera\u2026');
    sdcCamera = null;
    try {
      // Scandit Web v8: prefer the rear (world-facing) camera explicitly.
      if (Camera.pickBestGuessForPosition && SDC.CameraPosition) {
        sdcCamera = Camera.pickBestGuessForPosition(SDC.CameraPosition.WorldFacing);
      } else if (Camera.atPosition && SDC.CameraPosition) {
        sdcCamera = Camera.atPosition(SDC.CameraPosition.WorldFacing);
      }
    } catch (e) { log('rear camera pick unavailable: ' + (e && e.message ? e.message : e)); }
    if (!sdcCamera) {
      sdcCamera = (Camera.default !== undefined) ? Camera.default
                : (Camera.pickBestGuess ? Camera.pickBestGuess() : null);
      log('using default camera');
    }
    if (!sdcCamera) throw new Error('no camera available');
    facingMode = 'environment';
    const camSettings = BarcodeCapture.recommendedCameraSettings;
    if (sdcCamera && camSettings) await sdcCamera.applySettings(camSettings);
    await sdcContext.setFrameSource(sdcCamera);
    log('turning camera on\u2026');
    await sdcContext.frameSource.switchToDesiredState(FrameSourceState.On);
    setTimeout(async () => {
      try {
        if (sdcContext && sdcContext.frameSource &&
            sdcContext.frameSource.currentState !== FrameSourceState.On) {
          await sdcContext.frameSource.switchToDesiredState(FrameSourceState.On);
          log('re-asserted camera On');
        }
      } catch (e) {}
    }, 1200);

    const settings = new BarcodeCaptureSettings();
    const wanted = [
      'Code128', 'Code39', 'Code93',
      'EAN13UPCA', 'EAN8', 'UPCE',
      'InterleavedTwoOfFive', 'Codabar',
      'QR', 'DataMatrix',
    ];
    const valid = [];
    const skipped = [];
    wanted.forEach(name => {
      if (Symbology[name] !== undefined) valid.push(Symbology[name]);
      else skipped.push(name);
    });
    if (skipped.length) log('skipped unknown symbologies: ' + skipped.join(', '));
    settings.enableSymbologies(valid);
    log('symbologies enabled (' + valid.length + '); creating capture mode\u2026');

    sdcBarcodeCapture = await BarcodeCapture.forContext(sdcContext, settings);
    sdcBarcodeCapture.addListener({
      didScan: (mode, session) => {
        log('didScan fired');
        let barcode = null;
        try {
          barcode = session.newlyRecognizedBarcode
            || (session.newlyRecognizedBarcodes && session.newlyRecognizedBarcodes[0]);
        } catch (e) { log('barcode read error: ' + e); }
        if (!barcode) { log('didScan: no barcode in session'); return; }
        const data = (barcode.data != null ? barcode.data : (barcode.rawData || ''));
        log('scanned: ' + data);
        handleSuccess(String(data));
      },
    });
    await sdcBarcodeCapture.setEnabled(true);
    log('capture enabled');

    try {
      const torchBtn = $('torchBtn');
      let torchOk = false;
      try { if (sdcCamera) torchOk = await sdcCamera.isTorchAvailable(); } catch (e) {}
      if (torchOk) {
        torchBtn.classList.remove('hidden');
        log('torch available');
      } else {
        torchBtn.classList.add('hidden');
        log('torch not available on this camera');
      }
    } catch (e) { log('torch check failed: ' + e); }
  }

  async function stopScandit() {
    try { if (sdcBarcodeCapture) await sdcBarcodeCapture.setEnabled(false); } catch (e) {}
    try { if (sdcContext && SDC) await sdcContext.frameSource.switchToDesiredState(SDC.FrameSourceState.Off); } catch (e) {}
    try { if (sdcContext) await sdcContext.dispose(); } catch (e) {}
    sdcBarcodeCapture = null; sdcContext = null; sdcCamera = null;
    if (sdcView) { try { sdcView.detachFromElement && sdcView.detachFromElement(); } catch (e) {} sdcView = null; }
  }

  /* ---- HTML5-QRCODE engine (fallback) ----------------------------------- */
  async function startHtml5() {
    const stateText = $('cameraStateText');
    const retryBtn = $('cameraRetryBtn');

    if (typeof Html5Qrcode === 'undefined') {
      stateText.textContent = 'Scanner failed to load. Use Manual Entry below.';
      retryBtn.classList.remove('hidden');
      return;
    }
    $('qrReader').innerHTML = '';

    html5QrCode = new Html5Qrcode('qrReader', { verbose: false });
    const config = {
      fps: 12,
      formatsToSupport: H5Formats(),
      experimentalFeatures: { useBarCodeDetectorIfSupported: true },
    };
    try {
      await html5QrCode.start({ facingMode }, config, handleSuccess, () => {});
      $('cameraState').classList.add('hidden');
      setTimeout(setupTorchHtml5, 500);
    } catch (err) {
      console.error('Camera init failed:', err);
      if (err && /permission|notallowed|denied/i.test(String(err))) {
        stateText.textContent = 'Camera access denied. Grant permission, or use Manual Entry below.';
      } else if (err && /notfound|nocamera/i.test(String(err))) {
        stateText.textContent = 'No camera found. Use Manual Entry below.';
      } else {
        stateText.textContent = 'Could not start camera. Tap Try Again, or use Manual Entry below.';
      }
      retryBtn.classList.remove('hidden');
    }
  }

  async function stopHtml5() {
    if (html5QrCode) {
      try { await html5QrCode.stop(); } catch (e) {}
      try { html5QrCode.clear(); } catch (e) {}
      html5QrCode = null;
    }
    currentTrack = null;
  }

  function setupTorchHtml5() {
    const btn = $('torchBtn');
    try {
      const stream = document.querySelector('#qrReader video')?.srcObject;
      if (!stream) return;
      const track = stream.getVideoTracks()[0];
      if (!track) return;
      currentTrack = track;
      const caps = track.getCapabilities ? track.getCapabilities() : {};
      if (caps.torch) btn.classList.remove('hidden'); else btn.classList.add('hidden');
    } catch (e) { btn.classList.add('hidden'); }
  }

  /* ---- shared controls -------------------------------------------------- */
  async function flipCamera() {
    if (engine === 'scandit') {
      const wantUser = (facingMode === 'environment');
      let cam = null;
      try {
        const pos = wantUser ? SDC.CameraPosition.UserFacing : SDC.CameraPosition.WorldFacing;
        // Scandit Web v8: pickBestGuessForPosition is the correct API (atPosition is gone).
        if (SDC.Camera.pickBestGuessForPosition) {
          cam = SDC.Camera.pickBestGuessForPosition(pos);
        } else if (SDC.Camera.atPosition) {
          cam = SDC.Camera.atPosition(pos);   // legacy fallback
        }
      } catch (e) { log('flip camera pick failed: ' + (e && e.message ? e.message : e)); }
      if (!cam) {
        log('no ' + (wantUser ? 'front' : 'rear') + ' camera available');
        return;
      }
      try {
        torchOn = false;
        $('torchBtn').classList.remove('on');
        sdcCamera = cam;
        facingMode = wantUser ? 'user' : 'environment';
        const camSettings = SDCBarcode.BarcodeCapture.recommendedCameraSettings;
        if (camSettings) await sdcCamera.applySettings(camSettings);
        await sdcContext.setFrameSource(sdcCamera);
        await sdcContext.frameSource.switchToDesiredState(SDC.FrameSourceState.On);
        const torchBtn = $('torchBtn');
        let torchOk = false;
        try { torchOk = await sdcCamera.isTorchAvailable(); } catch (e) {}
        if (torchOk) torchBtn.classList.remove('hidden');
        else torchBtn.classList.add('hidden');
        log('switched to ' + (wantUser ? 'front' : 'rear') + ' camera');
      } catch (e) { log('flip failed: ' + (e && e.message ? e.message : e)); }
    } else {
      facingMode = (facingMode === 'environment') ? 'user' : 'environment';
      torchOn = false;
      $('torchBtn').classList.add('hidden');
      await start();
    }
  }

  async function toggleTorch() {
    torchOn = !torchOn;
    if (engine === 'scandit') {
      try {
        if (sdcCamera) {
          // Scandit Web v8: setDesiredTorchState is the correct async API.
          if (sdcCamera.setDesiredTorchState) {
            await sdcCamera.setDesiredTorchState(torchOn ? SDC.TorchState.On : SDC.TorchState.Off);
          } else {
            sdcCamera.desiredTorchState = torchOn ? SDC.TorchState.On : SDC.TorchState.Off; // legacy
          }
          $('torchBtn').classList.toggle('on', torchOn);
        }
      } catch (e) { torchOn = !torchOn; log('torch toggle failed: ' + e); }
    } else {
      if (!currentTrack) { torchOn = !torchOn; return; }
      try {
        await currentTrack.applyConstraints({ advanced: [{ torch: torchOn }] });
        $('torchBtn').classList.toggle('on', torchOn);
      } catch (e) { torchOn = !torchOn; }
    }
  }

  function openManual() {
    const input = $('manualInput');
    input.value = '';
    $('manualError').textContent = '';
    $('manualOverlay').classList.add('open');
    setTimeout(() => input.focus(), 50);
  }
  function closeManual() { $('manualOverlay').classList.remove('open'); }
  function submitManual() {
    const value = $('manualInput').value.trim();
    if (!value) { $('manualError').textContent = 'Please enter a value.'; return; }
    closeManual();
    scanCompleted = true;
    complete(value);
  }

  function rescan() { start(); }

  async function stop() {
    await stopScandit();
    await stopHtml5();
  }

  return {
    init, flipCamera, toggleTorch, openManual, closeManual,
    submitManual, rescan, retry: start, stop,
  };
})();

window.ScanController = ScanController;

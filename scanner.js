/* =========================================================================
   Shared scanner controller
   Used by scan_camera.html and scan_oos.html.

   ENGINE SELECTION
   ----------------
   1. Scandit Data Capture SDK v8  (primary) — used when window.SCANDIT_LICENSE_KEY
      is set (see config.js) AND the SDK loads successfully.
   2. html5-qrcode                 (fallback) — used automatically if Scandit has
      no key or fails to load.

   The public interface is identical for both engines, so the pages don't care
   which one is running:
     ScanController.init({ onComplete, manualTitle, manualSub, manualPlaceholder })
     ScanController.flipCamera()  toggleTorch()  openManual()  closeManual()
     ScanController.submitManual()  rescan()  retry()  stop()
   ========================================================================= */
const ScanController = (function () {
  let opts = {};
  let scanCompleted = false;
  let lastInvalidShake = 0;
  let engine = null;            // 'scandit' | 'html5'
  let facingMode = 'environment';
  let torchOn = false;

  // html5-qrcode state
  let html5QrCode = null;
  let currentTrack = null;

  // Scandit state
  let sdcContext = null;
  let sdcView = null;
  let sdcBarcodeCapture = null;
  let sdcCamera = null;
  let SDC = null;               // core module namespace
  let SDCBarcode = null;        // barcode module namespace

  const successSound = new Audio('https://assets.mixkit.co/active_storage/sfx/2013/2013-preview.mp3');
  function $(id) { return document.getElementById(id); }

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

  /* ---- shared lifecycle ------------------------------------------------- */
  async function init(options) {
    opts = options || {};
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

  async function start() {
    resetUiForStart();
    await stop();   // tear down whatever was running

    const key = (typeof window !== 'undefined' && window.SCANDIT_LICENSE_KEY) ? window.SCANDIT_LICENSE_KEY : '';
    if (key) {
      try {
        await startScandit(key);
        engine = 'scandit';
        $('cameraState').classList.add('hidden');
        return;
      } catch (e) {
        console.warn('[scanner] Scandit failed, falling back to html5-qrcode:', e);
        await stopScandit();
      }
    }
    await startHtml5();
    engine = 'html5';
  }

  function handleSuccess(decodedText) {
    if (scanCompleted) return;
    const value = String(decodedText || '').trim();
    if (!value) { shake(); return; }
    scanCompleted = true;
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
    await stop();

    $('completeScreen').style.display = 'flex';
    $('rescanBtn').classList.remove('hidden');
    $('manualEntryBtn').classList.add('hidden');

    $('idCard').classList.add('found');
    $('idValue').textContent = value;
    $('idBadge').textContent = 'Scanned \u2713';

    const checkEl = $('completeCheck');
    if (checkEl) { checkEl.style.animation = 'none'; void checkEl.offsetWidth; checkEl.style.animation = ''; }

    if (typeof opts.onComplete === 'function') opts.onComplete(value);
  }

  /* ---- SCANDIT engine --------------------------------------------------- */
  async function startScandit(licenseKey) {
    if (!SDC) SDC = await import('@scandit/web-datacapture-core');
    if (!SDCBarcode) SDCBarcode = await import('@scandit/web-datacapture-barcode');

    const { DataCaptureView, Camera, DataCaptureContext, FrameSourceState } = SDC;
    const { barcodeCaptureLoader, BarcodeCaptureSettings, BarcodeCapture, Symbology } = SDCBarcode;

    const mount = $('qrReader');
    mount.innerHTML = '';
    sdcView = new DataCaptureView();
    sdcView.connectToElement(mount);

    sdcContext = await DataCaptureContext.forLicenseKey(licenseKey, {
      libraryLocation: 'https://cdn.jsdelivr.net/npm/@scandit/web-datacapture-barcode@8/sdc-lib/',
      moduleLoaders: [barcodeCaptureLoader()],
    });
    await sdcView.setContext(sdcContext);

    sdcCamera = Camera.pickBestGuess ? Camera.pickBestGuess() : Camera.default;
    const camSettings = BarcodeCapture.recommendedCameraSettings;
    if (sdcCamera && camSettings) await sdcCamera.applySettings(camSettings);
    await sdcContext.setFrameSource(sdcCamera);
    await sdcContext.frameSource.switchToDesiredState(FrameSourceState.On);

    const settings = new BarcodeCaptureSettings();
    settings.enableSymbologies([
      Symbology.Code128, Symbology.Code39, Symbology.Code93,
      Symbology.Ean13Upca, Symbology.Ean8, Symbology.Upce,
      Symbology.InterleavedTwoOfFive, Symbology.Codabar,
      Symbology.QR, Symbology.DataMatrix,
    ]);

    sdcBarcodeCapture = await BarcodeCapture.forContext(sdcContext, settings);
    sdcBarcodeCapture.addListener({
      didScan: (mode, session) => {
        const barcode = session.newlyRecognizedBarcode ||
          (session.newlyRecognizedBarcodes && session.newlyRecognizedBarcodes[0]);
        if (!barcode) return;
        handleSuccess(barcode.data || '');
      },
    });
    await sdcBarcodeCapture.setEnabled(true);
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
    facingMode = (facingMode === 'environment') ? 'user' : 'environment';
    torchOn = false;
    $('torchBtn').classList.add('hidden');
    if (engine === 'scandit') {
      try {
        const pos = (facingMode === 'user')
          ? SDC.CameraPosition.UserFacing : SDC.CameraPosition.WorldFacing;
        const cam = SDC.Camera.atPosition(pos) || sdcCamera;
        if (cam) {
          sdcCamera = cam;
          const camSettings = SDCBarcode.BarcodeCapture.recommendedCameraSettings;
          if (camSettings) await sdcCamera.applySettings(camSettings);
          await sdcContext.setFrameSource(sdcCamera);
          await sdcContext.frameSource.switchToDesiredState(SDC.FrameSourceState.On);
        }
      } catch (e) { console.warn('[scanner] flip failed:', e); }
    } else {
      await start();
    }
  }

  async function toggleTorch() {
    torchOn = !torchOn;
    if (engine === 'scandit') {
      try {
        const desired = torchOn ? SDC.TorchState.On : SDC.TorchState.Off;
        if (sdcCamera && sdcCamera.setDesiredTorchState) {
          await sdcCamera.setDesiredTorchState(desired);
          $('torchBtn').classList.toggle('on', torchOn);
        }
      } catch (e) { torchOn = !torchOn; }
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

// scanner.js is loaded as a module, so its top-level binding is module-scoped.
// Expose it globally so inline onclick handlers and page scripts can reach it.
window.ScanController = ScanController;

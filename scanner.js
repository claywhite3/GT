/* =========================================================================
   Shared scanner controller (built on html5-qrcode)
   Used by scan_camera.html and scan_oos.html.

   Usage:
     ScanController.init({
       onComplete: (value) => { ... },   // called with the scanned/typed value
       manualTitle: 'Enter Camera ID',
       manualSub:   'Type the code printed on the Captana camera.',
       manualPlaceholder: 'e.g. CAP-00123'
     });
   ========================================================================= */
const ScanController = (function () {
  let html5QrCode = null;
  let scanCompleted = false;
  let lastInvalidShake = 0;
  let currentTrack = null;
  let torchOn = false;
  let facingMode = 'environment';   // rear default
  let opts = {};

  const FORMATS = () => [
    Html5QrcodeSupportedFormats.CODE_128,
    Html5QrcodeSupportedFormats.CODE_39,
    Html5QrcodeSupportedFormats.CODE_93,
    Html5QrcodeSupportedFormats.EAN_13,
    Html5QrcodeSupportedFormats.EAN_8,
    Html5QrcodeSupportedFormats.UPC_A,
    Html5QrcodeSupportedFormats.UPC_E,
    Html5QrcodeSupportedFormats.ITF,
    Html5QrcodeSupportedFormats.CODABAR,
    // QR fallback
    Html5QrcodeSupportedFormats.QR_CODE,
    Html5QrcodeSupportedFormats.DATA_MATRIX,
  ];

  const successSound = new Audio('https://assets.mixkit.co/active_storage/sfx/2013/2013-preview.mp3');

  function $(id) { return document.getElementById(id); }

  async function init(options) {
    opts = options || {};
    // wire manual modal text
    if (opts.manualTitle) $('manualTitle').textContent = opts.manualTitle;
    if (opts.manualSub) $('manualSub').textContent = opts.manualSub;
    if (opts.manualPlaceholder) $('manualInput').placeholder = opts.manualPlaceholder;
    await start();
  }

  async function start() {
    scanCompleted = false;
    const stateEl = $('cameraState');
    const stateText = $('cameraStateText');
    const retryBtn = $('cameraRetryBtn');

    // reset UI
    $('completeScreen').style.display = 'none';
    $('rescanBtn').classList.add('hidden');
    $('manualEntryBtn').classList.remove('hidden');
    stateEl.classList.remove('hidden');
    retryBtn.classList.add('hidden');
    stateText.textContent = 'Starting camera\u2026';

    if (html5QrCode) {
      try { await html5QrCode.stop(); } catch (e) {}
      try { html5QrCode.clear(); } catch (e) {}
      html5QrCode = null;
    }

    if (typeof Html5Qrcode === 'undefined') {
      stateText.textContent = 'Scanner library failed to load. Use Manual Entry below.';
      retryBtn.classList.remove('hidden');
      return;
    }

    html5QrCode = new Html5Qrcode('qrReader', { verbose: false });
    const config = {
      fps: 12,
      qrbox: undefined,
      aspectRatio: undefined,
      disableFlip: false,
      formatsToSupport: FORMATS(),
      experimentalFeatures: { useBarCodeDetectorIfSupported: true },
    };

    try {
      await html5QrCode.start({ facingMode }, config, handleSuccess, () => {});
      stateEl.classList.add('hidden');
      setTimeout(setupTorch, 500);
    } catch (err) {
      console.error('Camera init failed:', err);
      if (err && /permission|notallowed|denied/i.test(String(err))) {
        stateText.textContent = 'Camera access denied. Grant permission in settings, or use Manual Entry below.';
      } else if (err && /notfound|nocamera/i.test(String(err))) {
        stateText.textContent = 'No camera found. Use Manual Entry below.';
      } else {
        stateText.textContent = 'Could not start camera. Tap Try Again, or use Manual Entry below.';
      }
      retryBtn.classList.remove('hidden');
    }
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

    if (html5QrCode) { try { await html5QrCode.stop(); } catch (e) {} }

    $('completeScreen').style.display = 'flex';
    $('rescanBtn').classList.remove('hidden');
    $('manualEntryBtn').classList.add('hidden');

    const card = $('idCard');
    card.classList.add('found');
    $('idValue').textContent = value;
    $('idBadge').textContent = 'Scanned \u2713';

    const checkEl = $('completeCheck');
    if (checkEl) { checkEl.style.animation = 'none'; void checkEl.offsetWidth; checkEl.style.animation = ''; }

    if (typeof opts.onComplete === 'function') opts.onComplete(value);
  }

  // ── Camera flip ──
  async function flipCamera() {
    facingMode = (facingMode === 'environment') ? 'user' : 'environment';
    torchOn = false;
    $('torchBtn').classList.add('hidden');
    await start();
  }

  // ── Torch ──
  function setupTorch() {
    const btn = $('torchBtn');
    try {
      const stream = document.querySelector('#qrReader video')?.srcObject;
      if (!stream) return;
      const track = stream.getVideoTracks()[0];
      if (!track) return;
      currentTrack = track;
      const caps = track.getCapabilities ? track.getCapabilities() : {};
      if (caps.torch) btn.classList.remove('hidden');
      else btn.classList.add('hidden');
    } catch (e) { btn.classList.add('hidden'); }
  }

  async function toggleTorch() {
    if (!currentTrack) return;
    torchOn = !torchOn;
    try {
      await currentTrack.applyConstraints({ advanced: [{ torch: torchOn }] });
      $('torchBtn').classList.toggle('on', torchOn);
    } catch (e) { torchOn = !torchOn; }
  }

  // ── Manual entry ──
  function openManual() {
    const overlay = $('manualOverlay');
    const input = $('manualInput');
    input.value = '';
    $('manualError').textContent = '';
    overlay.classList.add('open');
    setTimeout(() => input.focus(), 50);
  }
  function closeManual() { $('manualOverlay').classList.remove('open'); }
  function submitManual() {
    const value = $('manualInput').value.trim();
    if (!value) {
      $('manualError').textContent = 'Please enter a value.';
      return;
    }
    closeManual();
    scanCompleted = true;
    complete(value);
  }

  function rescan() { start(); }

  // expose
  return { init, flipCamera, toggleTorch, openManual, closeManual, submitManual, rescan, retry: start };
})();

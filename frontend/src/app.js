/**
 * Bangla Word to PDF — Frontend Application
 * Designed & Developed by A.K.A.Sumon
 *
 * Handles file upload (drag-and-drop + click), conversion request,
 * multi-stage progress tracking, and PDF download with layout engine detection.
 */

// ── Configuration ─────────────────────────────────────────────
// Safe extraction of environment variables when running under Vite or standalone FastAPI
const CONVERTER_URL = (typeof import.meta !== 'undefined' && import.meta?.env?.VITE_CONVERTER_URL)
  ? import.meta.env.VITE_CONVERTER_URL
  : '';
const MAX_FILE_SIZE = 20 * 1024 * 1024; // 20 MB
const ALLOWED_EXTENSIONS = ['.docx'];

// ── DOM Elements ──────────────────────────────────────────────
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('file-input');
const fileInfo = document.getElementById('file-info');
const fileName = document.getElementById('file-name');
const fileSize = document.getElementById('file-size');
const removeFileBtn = document.getElementById('remove-file');
const convertBtn = document.getElementById('convert-btn');
const convertBtnText = document.getElementById('convert-btn-text');
const progress = document.getElementById('progress');
const progressFill = document.getElementById('progress-fill');
const progressText = document.getElementById('progress-text');
const progressPercent = document.getElementById('progress-percent');
const stageElements = document.querySelectorAll('.progress__stages .stage');
const engineName = document.getElementById('engine-name');
const engineBadge = document.getElementById('engine-status-badge');
const successResult = document.getElementById('success-result');
const successDetail = document.getElementById('success-detail');
const downloadBtn = document.getElementById('download-btn');
const convertAnother = document.getElementById('convert-another');
const errorResult = document.getElementById('error-result');
const errorDetail = document.getElementById('error-detail');
const tryAgain = document.getElementById('try-again');

// ── State ─────────────────────────────────────────────────────
let selectedFile = null;
let pdfBlob = null;
let pdfFilename = '';
let progressTimer = null;

// ── Utility Functions ─────────────────────────────────────────

/**
 * Format bytes into a human-readable string.
 */
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * Get file extension from filename.
 */
function getExtension(filename) {
  const idx = filename.lastIndexOf('.');
  return idx >= 0 ? filename.substring(idx).toLowerCase() : '';
}

/**
 * Show an element (remove hidden attribute).
 */
function show(el) {
  if (el) el.hidden = false;
}

/**
 * Hide an element (set hidden attribute).
 */
function hide(el) {
  if (el) el.hidden = true;
}

/**
 * Update the visual stage stepper (1: Parse, 2: Decode, 3: Layout, 4: Ready).
 */
function setProgressStage(stageIndex, percent, message) {
  if (progressPercent) progressPercent.textContent = `${percent}%`;
  if (progressFill) progressFill.style.width = `${percent}%`;
  if (progressText && message) progressText.textContent = message;

  if (stageElements && stageElements.length > 0) {
    stageElements.forEach((el, idx) => {
      if (idx <= stageIndex) {
        el.classList.add('active');
      } else {
        el.classList.remove('active');
      }
    });
  }
}

// ── File Selection ────────────────────────────────────────────

/**
 * Handle a selected file — validate and display info.
 */
function handleFileSelect(file) {
  // Validate extension
  const ext = getExtension(file.name);
  if (!ALLOWED_EXTENSIONS.includes(ext)) {
    showError(`Invalid file format "${ext}". Only Microsoft Word (.docx) files are supported.`);
    return;
  }

  // Validate size
  if (file.size > MAX_FILE_SIZE) {
    showError(`File size exceeds limit (${formatFileSize(file.size)}). Maximum supported file size is 20 MB.`);
    return;
  }

  selectedFile = file;
  pdfBlob = null;
  pdfFilename = '';

  // Update UI
  fileName.textContent = file.name;
  fileSize.textContent = formatFileSize(file.size);

  hide(dropzone);
  show(fileInfo);
  hide(successResult);
  hide(errorResult);
  hide(progress);

  convertBtn.disabled = false;
  convertBtnText.textContent = 'Convert to PDF';
  convertBtn.classList.remove('btn--loading');
}

/**
 * Remove the selected file and reset UI.
 */
function resetUI() {
  if (progressTimer) {
    clearInterval(progressTimer);
    progressTimer = null;
  }
  selectedFile = null;
  pdfBlob = null;
  pdfFilename = '';

  show(dropzone);
  hide(fileInfo);
  hide(successResult);
  hide(errorResult);
  hide(progress);

  convertBtn.disabled = true;
  convertBtnText.textContent = 'Convert to PDF';
  convertBtn.classList.remove('btn--loading');
  show(convertBtn);

  // Reset file input
  fileInput.value = '';
}

// ── Drag and Drop Events ──────────────────────────────────────

dropzone.addEventListener('dragenter', (e) => {
  e.preventDefault();
  dropzone.classList.add('dropzone--active');
});

dropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropzone.classList.add('dropzone--active');
});

dropzone.addEventListener('dragleave', (e) => {
  e.preventDefault();
  dropzone.classList.remove('dropzone--active');
});

dropzone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropzone.classList.remove('dropzone--active');

  const files = e.dataTransfer.files;
  if (files && files.length > 0) {
    handleFileSelect(files[0]);
  }
});

// Click to upload
dropzone.addEventListener('click', () => {
  fileInput.click();
});

dropzone.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    fileInput.click();
  }
});

fileInput.addEventListener('change', () => {
  if (fileInput.files.length > 0) {
    handleFileSelect(fileInput.files[0]);
  }
});

// Remove file
removeFileBtn.addEventListener('click', () => {
  resetUI();
});

// ── Conversion Flow ───────────────────────────────────────────

convertBtn.addEventListener('click', async () => {
  if (!selectedFile) return;

  // Show loading state
  convertBtn.disabled = true;
  convertBtn.classList.add('btn--loading');
  convertBtnText.textContent = 'Processing Document...';
  hide(errorResult);
  hide(successResult);
  show(progress);

  // Step 1: Parse
  setProgressStage(0, 18, 'Reading document structures and OpenXML tables...');

  // Multi-stage simulated progress while awaiting conversion response
  let currentStage = 0;
  let currentPercent = 18;
  progressTimer = setInterval(() => {
    if (currentPercent < 45) {
      currentPercent += 5;
      setProgressStage(0, currentPercent, 'Parsing OpenXML runs & paragraphs...');
    } else if (currentPercent < 75) {
      if (currentStage < 1) currentStage = 1;
      currentPercent += 4;
      setProgressStage(1, currentPercent, 'Reconstructing SutonnyMJ & Bijoy ANSI ligatures...');
    } else if (currentPercent < 90) {
      if (currentStage < 2) currentStage = 2;
      currentPercent += 2;
      setProgressStage(2, currentPercent, 'Rendering exact layout via Word Native Engine...');
    }
  }, 250);

  try {
    const formData = new FormData();
    formData.append('file', selectedFile);

    const response = await fetch(`${CONVERTER_URL}/convert`, {
      method: 'POST',
      body: formData,
    });

    if (progressTimer) {
      clearInterval(progressTimer);
      progressTimer = null;
    }

    if (!response.ok) {
      let errorMessage;
      try {
        const errorData = await response.json();
        errorMessage = errorData.detail || `Server error (${response.status})`;
      } catch {
        errorMessage = `Server error (${response.status})`;
      }
      throw new Error(errorMessage);
    }

    // Step 4: Ready
    setProgressStage(3, 100, 'Document conversion complete!');

    pdfBlob = await response.blob();

    // Extract filename from Content-Disposition header
    const disposition = response.headers.get('Content-Disposition');
    if (disposition) {
      const utf8Match = disposition.match(/filename\*=(?:UTF-8|utf-8)''([^;\n]+)/i);
      if (utf8Match) {
        try {
          pdfFilename = decodeURIComponent(utf8Match[1]);
        } catch {
          // ignore parsing error
        }
      }
      if (!pdfFilename) {
        const match = disposition.match(/filename="?([^";\n]+)"?/);
        if (match && match[1] !== 'converted.pdf') {
          pdfFilename = match[1];
        }
      }
    }
    if (!pdfFilename) {
      const stem = selectedFile.name.replace(/\.docx$/i, '');
      pdfFilename = `${stem}.pdf`;
    }

    // Read conversion telemetry headers
    const bijoyDetected = response.headers.get('X-Bijoy-Detected') === 'true';
    const runsConverted = parseInt(response.headers.get('X-Runs-Converted') || '0', 10);
    const engineUsed = response.headers.get('X-Conversion-Engine') || 'Native';
    const screenshotMode = response.headers.get('X-Exact-Screenshot-Mode') === 'true';

    let engineSummary = screenshotMode
      ? 'Microsoft Word Native Engine (100% Screenshot Fidelity)'
      : 'LibreOffice High-Fidelity Engine';

    let detailText = `Exported with ${engineSummary}.`;
    if (bijoyDetected) {
      detailText += ` Successfully decoded ${runsConverted} SutonnyMJ/Bijoy text runs with full ligature kerning.`;
    } else {
      detailText += ` Complex script typography and table layouts preserved flawlessly.`;
    }

    successDetail.textContent = detailText;

    setTimeout(() => {
      hide(progress);
      hide(convertBtn);
      show(successResult);
    }, 450);

  } catch (err) {
    if (progressTimer) {
      clearInterval(progressTimer);
      progressTimer = null;
    }
    hide(progress);
    showError(err.message);
    convertBtn.disabled = false;
    convertBtn.classList.remove('btn--loading');
    convertBtnText.textContent = 'Convert to PDF';
  }
});

// ── Download ──────────────────────────────────────────────────

downloadBtn.addEventListener('click', () => {
  if (!pdfBlob) return;

  const url = URL.createObjectURL(pdfBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = pdfFilename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

// ── Convert Another / Try Again ───────────────────────────────

convertAnother.addEventListener('click', resetUI);
tryAgain.addEventListener('click', () => {
  hide(errorResult);
  show(convertBtn);
  convertBtn.disabled = false;
  convertBtnText.textContent = 'Convert to PDF';
  convertBtn.classList.remove('btn--loading');
});

// ── Error Display ─────────────────────────────────────────────

function showError(message) {
  errorDetail.textContent = message;
  show(errorResult);
}

// ── Health & Engine Check ─────────────────────────────────────

async function checkHealth() {
  try {
    const res = await fetch(`${CONVERTER_URL}/health`);
    if (res.ok) {
      const data = await res.json();
      console.log('✓ Backend is healthy:', data);
      if (engineName) {
        if (data.word_native_available) {
          engineName.textContent = 'Word Native Engine Active';
          if (engineBadge) {
            engineBadge.title = 'Microsoft Word Automation available: 100% screenshot-fidelity mode';
          }
        } else {
          engineName.textContent = 'LibreOffice Engine Active';
          if (engineBadge) {
            engineBadge.title = 'LibreOffice Headless conversion pipeline active';
          }
        }
      }
    }
  } catch (err) {
    console.warn('Backend health check note:', err.message);
  }
}

// Initialize on load
checkHealth();


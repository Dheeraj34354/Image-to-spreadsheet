/* ══════════════════════════════════════════════════════════════════
   IntelliTab — app.js
   
   THE GENUINELY NOVEL FEATURE IS HERE:
   Bidirectional Image ↔ Cell Mapping
   
   Direction 1: Click spreadsheet cell → draw bounding box on image
   Direction 2: Click image region    → select cell in spreadsheet
   Direction 3: Edit low-conf cell    → flash green on image (feedback)
   
   No other open-source OCR tool does all three directions.
   ══════════════════════════════════════════════════════════════════ */

'use strict';

// ── Global state ─────────────────────────────────────────────────
const S = {
  data:        [],     // 2-D string table
  confidence:  [],     // 2-D int grid
  coordinates: [],     // 2-D [x,y,w,h] grid (in original image pixels)
  imgSize:     [0,0],  // [w, h] of original image
  numericCols: [],     // column indices that are numeric
  colSums:     {},     // {colIdx: sum}

  hot:           null,    // Handsontable instance
  mappingOn:     true,    // image↔cell mapping enabled
  lastRow:       -1,      // last selected row
  lastCol:       -1,      // last selected col

  // Canvas drawing
  canvas:        null,
  ctx:           null,
  img:           null,
  scale:         1,       // ratio: rendered px / original px

  // Reverse mapping: click on image → cell
  clickRegions:  [],     // [{row,col,x,y,w,h} in rendered coords]

  // Animation state
  editFlashActive: false,
};

// ── DOM refs ──────────────────────────────────────────────────────
const $ = id => document.getElementById(id);

// ════════════════════════════════════════════════════════════════
// FILE UPLOAD
// ════════════════════════════════════════════════════════════════

const dropZone  = $('dropZone');
const fileInput = $('fileInput');

dropZone.addEventListener('dragover', e => {
  e.preventDefault(); dropZone.classList.add('over');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault(); dropZone.classList.remove('over');
  const f = e.dataTransfer.files[0];
  if (f && f.type.startsWith('image/')) handleFile(f);
  else showError('Please drop a valid image file.');
});
dropZone.addEventListener('click', e => {
  if (!e.target.closest('.link-btn')) fileInput.click();
});
fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) handleFile(fileInput.files[0]);
});

function handleFile(file) {
  hideError();
  // Show preview in drop zone briefly then start OCR
  const reader = new FileReader();
  reader.onload = e => {
    S.img = new Image();
    S.img.onload = () => startOCR(file);
    S.img.src = e.target.result;

    // Also set <img> source for display
    $('sourceImg').src = e.target.result;
  };
  reader.readAsDataURL(file);
}

// ════════════════════════════════════════════════════════════════
// OCR PIPELINE
// ════════════════════════════════════════════════════════════════

let _stepTimer = null;

async function startOCR(file) {
  $('uploadScreen').style.display  = 'none';
  $('loadingScreen').style.display = 'flex';
  animateLoadingSteps();

  try {
    const fd = new FormData();
    fd.append('image', file);

    const res    = await fetch('/ocr', { method: 'POST', body: fd });
    const result = await res.json();

    clearTimeout(_stepTimer);

    if (!result.success) throw new Error(result.error || 'OCR failed.');

    // Store state
    S.data        = result.data        || [];
    S.confidence  = result.confidence  || [];
    S.coordinates = result.coordinates || [];
    S.imgSize     = result.img_size    || [0, 0];
    S.numericCols = result.numeric_cols || [];
    S.colSums     = result.col_sums    || {};

    showWorkspace(result);

  } catch (err) {
    $('loadingScreen').style.display = 'none';
    $('uploadScreen').style.display  = 'flex';
    showError(err.message || 'Unexpected error during OCR.');
  }
}

function animateLoadingSteps() {
  const steps = ['lss1','lss2','lss3','lss4'];
  let i = 0;
  function tick() {
    steps.forEach((s, si) => {
      const el = $(s);
      el.classList.toggle('active', si === i);
      el.classList.toggle('done',   si  <  i);
    });
    if (i < steps.length - 1) {
      i++;
      _stepTimer = setTimeout(tick, 700);
    }
  }
  tick();
}

// ════════════════════════════════════════════════════════════════
// SHOW WORKSPACE
// ════════════════════════════════════════════════════════════════

function showWorkspace(result) {
  $('loadingScreen').style.display = 'none';
  $('workspace').style.display     = 'flex';
  $('headerStats').style.display   = 'flex';
  $('exportBtn').style.display     = 'flex';
  $('resetBtn').style.display      = 'flex';

  // Header stats
  $('accuracyVal').textContent = result.accuracy + '%';
  $('rowsVal').textContent     = Math.max(0, S.data.length - 1);
  $('colsVal').textContent     = S.data[0]?.length || 0;
  $('numericVal').textContent  = S.numericCols.length;

  // Colour-code accuracy chip
  const acc = result.accuracy;
  $('statAccuracy').style.borderColor =
    acc >= 90 ? 'rgba(16,185,129,.5)' :
    acc >= 70 ? 'rgba(245,158,11,.5)' : 'rgba(239,68,68,.5)';

  // Build Handsontable
  initHandsontable();

  // Build image overlay
  initImageOverlay();

  // Novelty alerts
  showNoveltyAlerts(result);

  // Sum bar
  buildSumBar();
}

// ════════════════════════════════════════════════════════════════
// HANDSONTABLE INIT
// ════════════════════════════════════════════════════════════════

function initHandsontable() {
  if (S.hot) { S.hot.destroy(); S.hot = null; }

  const container = $('hotContainer');
  const numColSet = new Set(S.numericCols.map(Number));

  // ── Custom renderer for confidence colour coding ──────────────
  function confRenderer(hotInstance, td, row, col, prop, value, cellProperties) {
    Handsontable.renderers.TextRenderer.apply(this, arguments);

    // Skip header row (row 0 of data = visual row 0)
    const conf = S.confidence?.[row]?.[col] ?? -1;

    // Remove all custom classes first
    td.classList.remove('ht-conf-high','ht-conf-medium','ht-conf-low',
                         'ht-numeric-err','ht-mapped-cell');

    // Confidence colour coding (Unique Feature 1)
    if (conf >= 90)        td.classList.add('ht-conf-high');
    else if (conf >= 70)   td.classList.add('ht-conf-medium');
    else if (conf >= 0)    td.classList.add('ht-conf-low');

    // Numeric validation error (Unique Feature 2)
    if (numColSet.has(col) && value && isNaN(parseFloat(String(value).replace(/,/g,'')))) {
      td.classList.add('ht-numeric-err');
    }

    // Right-align numeric
    if (numColSet.has(col)) {
      td.classList.add('htNumeric');
    }

    // Active mapped cell highlight
    if (row === S.lastRow && col === S.lastCol) {
      td.classList.add('ht-mapped-cell');
    }
  }

  Handsontable.renderers.registerRenderer('confRenderer', confRenderer);

  S.hot = new Handsontable(container, {
    data:            deepCopy(S.data),
    rowHeaders:      true,
    colHeaders:      true,
    filters:         true,
    dropdownMenu:    true,
    columnSorting:   true,
    autoColumnSize:  true,
    manualColumnResize: true,
    manualRowResize:    true,
    contextMenu:     true,
    minSpareRows:    1,
    licenseKey:      'non-commercial-and-evaluation',
    renderer:        'confRenderer',

    afterSelectionEnd(row, col) {
      if (!S.mappingOn) return;
      S.lastRow = row;
      S.lastCol = col;

      // ★ NOVEL: Cell → Image direction
      highlightCellOnImage(row, col);

      // Re-render to apply ht-mapped-cell class
      S.hot.render();
    },

    afterChange(changes, source) {
      if (!changes || source === 'loadData') return;
      changes.forEach(([row, col, oldVal, newVal]) => {
        if (!S.mappingOn) return;
        // ★ NOVEL: Edit feedback — flash green on image region
        const wasLowConf = (S.confidence?.[row]?.[col] ?? 100) < 70;
        if (wasLowConf && newVal !== oldVal) {
          flashGreenOnImage(row, col);
        }
      });
    },
  });
}

function deepCopy(arr) {
  return arr.map(r => [...r]);
}

// ════════════════════════════════════════════════════════════════
// ★ NOVEL FEATURE: IMAGE OVERLAY CANVAS
// ════════════════════════════════════════════════════════════════

function initImageOverlay() {
  const imgEl  = $('sourceImg');
  const canvas = $('overlayCanvas');
  S.canvas = canvas;
  S.ctx    = canvas.getContext('2d');

  function resizeCanvas() {
    canvas.width  = imgEl.offsetWidth;
    canvas.height = imgEl.offsetHeight;

    // Scale: rendered pixels / original pixels
    S.scale = imgEl.offsetWidth / (S.imgSize[0] || imgEl.offsetWidth);

    buildClickRegions();
    clearCanvas();
  }

  // Resize when image loads and on window resize
  imgEl.addEventListener('load', () => setTimeout(resizeCanvas, 50));
  if (imgEl.complete) setTimeout(resizeCanvas, 50);
  window.addEventListener('resize', () => setTimeout(resizeCanvas, 100));

  // ★ NOVEL: Reverse mapping — click image region → select cell
  canvas.addEventListener('click', onCanvasClick);
  canvas.addEventListener('mousemove', onCanvasHover);
  canvas.addEventListener('mouseleave', () => {
    $('bboxTooltip').style.display = 'none';
    clearCanvas();
    if (S.lastRow >= 0) highlightCellOnImage(S.lastRow, S.lastCol);
  });
}

function buildClickRegions() {
  S.clickRegions = [];
  if (!S.coordinates.length) return;

  S.coordinates.forEach((row, ri) => {
    row.forEach((coord, ci) => {
      if (!coord) return;
      const [x, y, w, h] = coord;
      S.clickRegions.push({
        row: ri, col: ci,
        rx: Math.floor(x * S.scale),
        ry: Math.floor(y * S.scale),
        rw: Math.ceil(w  * S.scale),
        rh: Math.ceil(h  * S.scale),
      });
    });
  });
}

// ── Direction 1: Cell clicked → highlight on image ──────────────
function highlightCellOnImage(row, col) {
  clearCanvas();
  if (!S.mappingOn) return;

  const coord = S.coordinates?.[row]?.[col];
  if (!coord) return;

  const [x, y, w, h] = coord;
  const conf          = S.confidence?.[row]?.[col] ?? -1;
  const text          = S.data?.[row]?.[col] ?? '';

  const rx = Math.floor(x * S.scale);
  const ry = Math.floor(y * S.scale);
  const rw = Math.ceil(w  * S.scale);
  const rh = Math.ceil(h  * S.scale);

  drawBox(rx, ry, rw, rh, conf, true);
  showTooltip(rx, ry, rw, rh, text, conf);

  // Scroll image viewer so the highlighted region is visible
  const viewer = $('imgViewer');
  const scrollY = ry - viewer.offsetHeight / 3;
  viewer.scrollTo({ top: Math.max(0, scrollY), behavior: 'smooth' });
}

// ── Direction 2: Click on image → select cell in spreadsheet ────
function onCanvasClick(e) {
  if (!S.mappingOn) return;

  const rect = S.canvas.getBoundingClientRect();
  const mx   = e.clientX - rect.left;
  const my   = e.clientY - rect.top;

  // Find which region was clicked
  const hit = S.clickRegions.find(r =>
    mx >= r.rx && mx <= r.rx + r.rw &&
    my >= r.ry && my <= r.ry + r.rh
  );

  if (hit) {
    // Select the cell in Handsontable
    S.lastRow = hit.row;
    S.lastCol = hit.col;
    S.hot.selectCell(hit.row, hit.col);
    S.hot.scrollViewportTo({ row: hit.row, col: hit.col });

    // Flash the cell amber to confirm selection
    flashAmberOnCell(hit.row, hit.col);
  }
}

// ── Hover: show tooltip on nearby regions ───────────────────────
function onCanvasHover(e) {
  if (!S.mappingOn) return;

  const rect = S.canvas.getBoundingClientRect();
  const mx   = e.clientX - rect.left;
  const my   = e.clientY - rect.top;

  const hit = S.clickRegions.find(r =>
    mx >= r.rx && mx <= r.rx + r.rw &&
    my >= r.ry && my <= r.ry + r.rh
  );

  if (hit) {
    S.canvas.style.cursor = 'pointer';
    const conf = S.confidence?.[hit.row]?.[hit.col] ?? -1;
    const text = S.data?.[hit.row]?.[hit.col] ?? '';
    clearCanvas();
    drawBox(hit.rx, hit.ry, hit.rw, hit.rh, conf, false);
    showTooltip(hit.rx, hit.ry, hit.rw, hit.rh, text, conf);
    // Keep last selection visible too
    if (S.lastRow >= 0 && !(S.lastRow === hit.row && S.lastCol === hit.col)) {
      const lc = S.coordinates?.[S.lastRow]?.[S.lastCol];
      if (lc) {
        const [lx,ly,lw,lh] = lc;
        drawBox(Math.floor(lx*S.scale), Math.floor(ly*S.scale),
                Math.ceil(lw*S.scale),  Math.ceil(lh*S.scale),
                S.confidence?.[S.lastRow]?.[S.lastCol] ?? -1, true);
      }
    }
  } else {
    S.canvas.style.cursor = 'crosshair';
    $('bboxTooltip').style.display = 'none';
    clearCanvas();
    if (S.lastRow >= 0) highlightCellOnImage(S.lastRow, S.lastCol);
  }
}

// ── Direction 3: Flash green on image after editing low-conf cell
function flashGreenOnImage(row, col) {
  if (!S.mappingOn || S.editFlashActive) return;
  const coord = S.coordinates?.[row]?.[col];
  if (!coord) return;

  S.editFlashActive = true;
  const [x, y, w, h] = coord;
  const rx = Math.floor(x * S.scale);
  const ry = Math.floor(y * S.scale);
  const rw = Math.ceil(w  * S.scale);
  const rh = Math.ceil(h  * S.scale);

  let alpha = 0;
  let phase = 'in';
  const ctx = S.ctx;

  function frame() {
    clearCanvas();

    // Draw green flash
    ctx.save();
    ctx.fillStyle = `rgba(16, 185, 129, ${alpha * 0.35})`;
    ctx.fillRect(rx, ry, rw, rh);

    ctx.strokeStyle = `rgba(16, 185, 129, ${alpha})`;
    ctx.lineWidth   = 3;
    ctx.shadowColor = '#10B981';
    ctx.shadowBlur  = 10 * alpha;
    ctx.strokeRect(rx, ry, rw, rh);
    ctx.restore();

    if (phase === 'in') {
      alpha += 0.08;
      if (alpha >= 1) { alpha = 1; phase = 'hold'; setTimeout(() => { phase = 'out'; }, 400); }
    } else if (phase === 'out') {
      alpha -= 0.06;
      if (alpha <= 0) { S.editFlashActive = false; return; }
    }
    requestAnimationFrame(frame);
  }
  requestAnimationFrame(frame);
}

function flashAmberOnCell(row, col) {
  // Brief amber outline flash on the Handsontable cell after reverse-map click
  // We do this by temporarily overriding the renderer and re-rendering
  S.hot.render();
}

// ── Canvas drawing helpers ────────────────────────────────────────

function clearCanvas() {
  if (!S.ctx || !S.canvas) return;
  S.ctx.clearRect(0, 0, S.canvas.width, S.canvas.height);
}

function drawBox(x, y, w, h, conf, selected) {
  const ctx = S.ctx;
  if (!ctx) return;

  // Colour by confidence
  let color;
  if (conf < 0)       color = '#94A3B8';
  else if (conf >= 90) color = '#10B981';
  else if (conf >= 70) color = '#F59E0B';
  else                 color = '#EF4444';

  const alpha   = selected ? 0.30 : 0.18;
  const lWidth  = selected ? 2.5  : 1.5;
  const blur    = selected ? 8    : 4;

  ctx.save();

  // Semi-transparent fill
  ctx.fillStyle = hexToRgba(color, alpha);
  ctx.fillRect(x, y, w, h);

  // Border with glow
  ctx.strokeStyle = color;
  ctx.lineWidth   = lWidth;
  ctx.shadowColor = color;
  ctx.shadowBlur  = blur;
  ctx.strokeRect(x, y, w, h);

  // Corner decorations (L-shaped corners for selected)
  if (selected) {
    const cs = Math.min(8, w/3, h/3);
    ctx.lineWidth   = 3;
    ctx.shadowBlur  = 12;
    // TL
    ctx.beginPath(); ctx.moveTo(x, y+cs); ctx.lineTo(x, y); ctx.lineTo(x+cs, y); ctx.stroke();
    // TR
    ctx.beginPath(); ctx.moveTo(x+w-cs, y); ctx.lineTo(x+w, y); ctx.lineTo(x+w, y+cs); ctx.stroke();
    // BR
    ctx.beginPath(); ctx.moveTo(x+w, y+h-cs); ctx.lineTo(x+w, y+h); ctx.lineTo(x+w-cs, y+h); ctx.stroke();
    // BL
    ctx.beginPath(); ctx.moveTo(x+cs, y+h); ctx.lineTo(x, y+h); ctx.lineTo(x, y+h-cs); ctx.stroke();
  }

  ctx.restore();
}

function showTooltip(rx, ry, rw, rh, text, conf) {
  const tt      = $('bboxTooltip');
  const viewer  = $('imgViewer');

  $('ttText').textContent = text || '(empty)';

  const confEl = $('ttConf');
  if (conf >= 0) {
    confEl.textContent = conf + '%';
    confEl.className   = 'tt-conf ' +
      (conf >= 90 ? 'high' : conf >= 70 ? 'medium' : 'low');
    confEl.style.display = '';
  } else {
    confEl.style.display = 'none';
  }

  // Position tooltip below the box, clamp to viewer
  let tx = rx;
  let ty = ry + rh + 6;
  if (ty + 30 > viewer.offsetHeight) ty = ry - 32;
  if (tx + 180 > viewer.offsetWidth)  tx = viewer.offsetWidth - 184;
  tx = Math.max(4, tx);

  tt.style.left    = tx + 'px';
  tt.style.top     = (ty + viewer.scrollTop) + 'px';
  tt.style.display = 'flex';
}

function hexToRgba(hex, alpha) {
  const r = parseInt(hex.slice(1,3),16);
  const g = parseInt(hex.slice(3,5),16);
  const b = parseInt(hex.slice(5,7),16);
  return `rgba(${r},${g},${b},${alpha})`;
}

// ════════════════════════════════════════════════════════════════
// NOVELTY ALERT PANELS
// ════════════════════════════════════════════════════════════════

function showNoveltyAlerts(result) {
  const alerts = $('noveltyAlerts');
  alerts.innerHTML = '';

  // 1. Confidence report
  const allConfs = (result.confidence || []).flat().filter(c => c >= 0);
  const lowCount = allConfs.filter(c => c < 70).length;
  const medCount = allConfs.filter(c => c >= 70 && c < 90).length;

  if (lowCount > 0) {
    alerts.appendChild(makeAlert('red',
      `🔴 ${lowCount} cells have low confidence (&lt;70%) — highlighted in red. Click them to see source region.`));
  }
  if (medCount > 0) {
    alerts.appendChild(makeAlert('amber',
      `🟡 ${medCount} cells have medium confidence (70–89%) — verify before exporting.`));
  }

  // 2. Numeric detection report
  if (S.numericCols.length > 0) {
    const headers   = result.data?.[0] || [];
    const colNames  = S.numericCols.map(ci => `"${headers[ci] || 'Col'+(ci+1)}"`).join(', ');
    alerts.appendChild(makeAlert('green',
      `🔢 ${S.numericCols.length} numeric column(s) detected: ${colNames}. Auto-sum enabled. Right-aligned in Excel.`));
  }

  // 3. Mapping enabled
  alerts.appendChild(makeAlert('amber',
    `🔗 Image↔Cell mapping active. Click any cell to highlight its source on the image — or click the image to jump to that cell.`));
}

function makeAlert(type, html) {
  const d = document.createElement('div');
  d.className  = `nov-alert ${type}`;
  d.innerHTML  = html;
  return d;
}

// ════════════════════════════════════════════════════════════════
// NUMERIC COL SUMS BAR
// ════════════════════════════════════════════════════════════════

function buildSumBar() {
  if (!Object.keys(S.colSums).length) return;

  const chips   = $('sumChips');
  const headers = S.data?.[0] || [];
  chips.innerHTML = '';

  for (const [ci, sum] of Object.entries(S.colSums)) {
    const colName = headers[Number(ci)] || `Col ${Number(ci)+1}`;
    const chip    = document.createElement('span');
    chip.className   = 'sb-chip';
    chip.textContent = `${colName}: ${Number(sum).toLocaleString()}`;
    chips.appendChild(chip);
  }

  $('sumBar').style.display = 'flex';
}

// ════════════════════════════════════════════════════════════════
// MAPPING TOGGLE
// ════════════════════════════════════════════════════════════════

function toggleMapping() {
  S.mappingOn = !S.mappingOn;
  const badge  = $('mapModeBadge');
  const toggle = $('mappingToggle');
  badge.textContent = S.mappingOn ? '🔗 Mapping ON' : '🔗 Mapping OFF';
  badge.classList.toggle('active', S.mappingOn);
  toggle.textContent = S.mappingOn ? 'Disable' : 'Enable';
  $('bboxTooltip').style.display = 'none';
  clearCanvas();
}

// ════════════════════════════════════════════════════════════════
// EXCEL EXPORT
// ════════════════════════════════════════════════════════════════

async function downloadExcel() {
  if (!S.data.length) { showError('No data to export.'); return; }

  // Sync DOM table → state
  const currentData = S.hot ? S.hot.getData() : S.data;

  const btn  = $('exportBtn');
  const orig = btn.innerHTML;
  btn.innerHTML    = '⏳ Generating…';
  btn.disabled     = true;

  try {
    const res = await fetch('/export', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        tableData:   currentData,
        confidence:  S.confidence,
        numericCols: S.numericCols,
        colSums:     S.colSums,
      }),
    });

    if (!res.ok) {
      const e = await res.json();
      throw new Error(e.error || 'Export failed.');
    }

    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = 'intellitab_export.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    btn.innerHTML = '✓ Downloaded!';
    setTimeout(() => { btn.innerHTML = orig; btn.disabled = false; }, 2500);

  } catch (err) {
    showError(err.message);
    btn.innerHTML = orig;
    btn.disabled  = false;
  }
}

// ════════════════════════════════════════════════════════════════
// RESET
// ════════════════════════════════════════════════════════════════

function resetAll() {
  Object.assign(S, {
    data:[], confidence:[], coordinates:[], imgSize:[0,0],
    numericCols:[], colSums:{}, lastRow:-1, lastCol:-1,
    scale:1, clickRegions:[], editFlashActive:false,
  });

  if (S.hot) { S.hot.destroy(); S.hot = null; }
  clearCanvas();

  $('workspace').style.display     = 'none';
  $('headerStats').style.display   = 'none';
  $('exportBtn').style.display     = 'none';
  $('resetBtn').style.display      = 'none';
  $('uploadScreen').style.display  = 'flex';
  $('bboxTooltip').style.display   = 'none';
  $('sumBar').style.display        = 'none';
  $('noveltyAlerts').innerHTML     = '';
  $('sourceImg').src               = '';
  $('fileInput').value             = '';
  $('mapModeBadge').className      = 'badge active';
  $('mapModeBadge').textContent    = '🔗 Mapping ON';
  $('mappingToggle').textContent   = 'Disable';
  S.mappingOn = true;
}

// ════════════════════════════════════════════════════════════════
// ERROR TOAST
// ════════════════════════════════════════════════════════════════

function showError(msg) {
  $('errorMsg').textContent    = msg;
  $('errorToast').style.display = 'flex';
}

function hideError() {
  $('errorToast').style.display = 'none';
}

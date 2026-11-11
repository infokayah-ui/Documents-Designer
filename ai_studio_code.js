/**
 * DOCUMENT DESIGNER PRO
 * Portable Office Web App & PWA Core Engine
 */

// Application State
const State = {
  currentMenu: 'background',
  paperSize: 'A4',
  orientation: 'portrait',
  margins: { top: 0.75, bottom: 0.75, left: 0.75, right: 0.75 }, // in inches
  objects: [], // Array of { id, type, x, y, width, height, zIndex, locked, el, data }
  selectedObjectId: null,
  history: [],
  historyIndex: -1,
  cleanerOriginalImageData: null
};

const DPI = 96;
const paperPage = document.getElementById('paperPage');
const layersContainer = document.getElementById('layersContainer');
const textEditorRoot = document.getElementById('textEditorRoot');
const contextToolbar = document.getElementById('contextToolbar');
const rulerWrapper = document.getElementById('rulerWrapper');

// ==========================================
// 1. INITIALIZATION & MENU DISPATCH
// ==========================================
window.addEventListener('DOMContentLoaded', () => {
  setupNavigation();
  setupPageSetupModal();
  setupHistory();
  setupGlobalShortcuts();
  switchMenu('background');
  updatePageDimensions();
});

function switchMenu(menuKey) {
  State.currentMenu = menuKey;
  
  // Highlight Bottom Nav
  document.querySelectorAll('.bottom-nav .nav-item').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.menu === menuKey);
  });

  // Ruler visible only for Text Editor & Word
  if (menuKey === 'texteditor' || menuKey === 'word') {
    rulerWrapper.style.display = 'block';
    textEditorRoot.style.display = 'block';
  } else {
    rulerWrapper.style.display = 'none';
    if (menuKey !== 'texteditor' && menuKey !== 'word') {
      textEditorRoot.style.display = 'none';
    }
  }

  // Clear or re-evaluate selections: Strictly scoped to current menu
  deselectAllObjects();
  autoSelectTopObjectForMenu(menuKey);

  // Render Secondary Toolbar for the chosen Menu
  renderContextToolbar(menuKey);
}

function setupNavigation() {
  document.querySelectorAll('.bottom-nav .nav-item').forEach(btn => {
    btn.addEventListener('click', () => switchMenu(btn.dataset.menu));
  });
}

// ==========================================
// 2. CONTEXT TOOLBARS (The 7 Specific Menus)
// ==========================================
function renderContextToolbar(menu) {
  contextToolbar.innerHTML = '';
  const obj = getSelectedObject();

  switch(menu) {
    case 'background':
      contextToolbar.innerHTML = `
        <label class="btn btn-secondary"><input type="file" id="bgFileInput" accept="image/*" style="display:none">🖼️ Add BG Image</label>
        <button class="btn btn-secondary" id="btnFitBg">Auto Fit Page</button>
        <button class="btn btn-secondary" id="btnDelBg">🗑️ Remove</button>
      `;
      document.getElementById('bgFileInput').onchange = (e) => handleImageUpload(e, 'background');
      document.getElementById('btnFitBg').onclick = fitBackgroundToPage;
      document.getElementById('btnDelBg').onclick = removeCurrentObject;
      break;

    case 'docimage':
      contextToolbar.innerHTML = `
        <label class="btn btn-secondary"><input type="file" id="docImgInput" accept="image/*" style="display:none">📷 Add Doc Image</label>
        <button class="btn btn-secondary" id="btnRemoveWhiteDoc">⚪ Remove White BG</button>
        <label>Bright: <input type="range" id="docBrightness" min="-100" max="100" value="0"></label>
        <label>Contrast: <input type="range" id="docContrast" min="-100" max="100" value="0"></label>
        <button class="btn btn-secondary" id="btnDelObj">🗑️</button>
      `;
      document.getElementById('docImgInput').onchange = (e) => handleImageUpload(e, 'docimage');
      document.getElementById('btnRemoveWhiteDoc').onclick = () => removeWhiteBackground(obj, 220);
      document.getElementById('docBrightness').oninput = (e) => applyCSSFilters(obj, e.target.value, null);
      document.getElementById('docContrast').oninput = (e) => applyCSSFilters(obj, null, e.target.value);
      document.getElementById('btnDelObj').onclick = removeCurrentObject;
      break;

    case 'signature':
      contextToolbar.innerHTML = `
        <label class="btn btn-secondary"><input type="file" id="sigFileInput" accept="image/*" style="display:none">✍️ Add Signature</label>
        <button class="btn btn-secondary" id="btnRemoveSigBg">⚪ Cut White BG</button>
        <select id="sigColorPicker">
          <option value="original">Original</option>
          <option value="#000000">Black</option>
          <option value="#0038ff">Blue</option>
          <option value="#e11d48">Red</option>
          <option value="#16a34a">Green</option>
        </select>
        <label>Opacity: <input type="range" id="sigOpacity" min="10" max="100" value="100"></label>
        <button class="btn btn-secondary" id="btnDelSig">🗑️</button>
      `;
      document.getElementById('sigFileInput').onchange = (e) => handleImageUpload(e, 'signature');
      document.getElementById('btnRemoveSigBg').onclick = () => removeWhiteBackground(obj, 200);
      document.getElementById('sigColorPicker').onchange = (e) => tintSignature(obj, e.target.value);
      document.getElementById('sigOpacity').oninput = (e) => {
        if(obj) { obj.el.style.opacity = e.target.value / 100; recordState(); }
      };
      document.getElementById('btnDelSig').onclick = removeCurrentObject;
      break;

    case 'word':
      contextToolbar.innerHTML = `
        <label class="btn btn-secondary"><input type="file" id="docxInput" accept=".docx" style="display:none">📘 Import DOCX</label>
        <button class="icon-btn" onclick="document.execCommand('bold')"><b>B</b></button>
        <button class="icon-btn" onclick="document.execCommand('italic')"><i>I</i></button>
        <button class="icon-btn" onclick="document.execCommand('underline')"><u>U</u></button>
        <button class="icon-btn" onclick="document.execCommand('justifyLeft')">Left</button>
        <button class="icon-btn" onclick="document.execCommand('justifyCenter')">Center</button>
        <button class="icon-btn" onclick="document.execCommand('justifyRight')">Right</button>
        <select onchange="document.execCommand('fontSize', false, this.value)">
          <option value="3">12pt</option>
          <option value="4">14pt</option>
          <option value="5">18pt</option>
        </select>
      `;
      document.getElementById('docxInput').onchange = handleDocxImport;
      break;

    case 'texteditor':
      contextToolbar.innerHTML = `
        <button class="icon-btn" onclick="document.execCommand('bold')"><b>B</b></button>
        <button class="icon-btn" onclick="document.execCommand('italic')"><i>I</i></button>
        <button class="icon-btn" onclick="document.execCommand('underline')"><u>U</u></button>
        <button class="icon-btn" onclick="document.execCommand('justifyLeft')">Left</button>
        <button class="icon-btn" onclick="document.execCommand('justifyCenter')">Center</button>
        <button class="icon-btn" onclick="document.execCommand('justifyRight')">Right</button>
        <button class="icon-btn" onclick="document.execCommand('justifyFull')">Justify</button>
        <select id="fontFamilySelector" onchange="document.execCommand('fontName', false, this.value)">
          <option value="Pyidaungsu, sans-serif">Pyidaungsu (Myanmar)</option>
          <option value="Arial, sans-serif">Arial</option>
          <option value="'Times New Roman', serif">Times New Roman</option>
        </select>
        <input type="color" onchange="document.execCommand('foreColor', false, this.value)" title="Text Color">
      `;
      break;

    case 'pdf':
      contextToolbar.innerHTML = `
        <label class="btn btn-secondary"><input type="file" id="pdfFileInput" accept="application/pdf" style="display:none">📕 Import PDF</label>
        <button class="btn btn-secondary" id="btnPrevPdf">◀</button>
        <span id="pdfPageIndicator" style="font-size:12px;">Page 1</span>
        <button class="btn btn-secondary" id="btnNextPdf">▶</button>
      `;
      document.getElementById('pdfFileInput').onchange = handlePdfImport;
      break;

    case 'cleaner':
      contextToolbar.innerHTML = `
        <button class="btn btn-primary" id="btnAutoClean">⚡ Auto Clean</button>
        <label>White: <input type="range" id="cleanWhiteness" min="150" max="255" value="215"></label>
        <label>Clarity: <input type="range" id="cleanClarity" min="1" max="10" value="5"></label>
        <button class="btn btn-secondary" id="btnTogglePreview">Compare</button>
        <button class="btn btn-secondary" id="btnResetClean">Reset</button>
      `;
      document.getElementById('btnAutoClean').onclick = () => runPaperCleaner(true);
      document.getElementById('cleanWhiteness').oninput = () => runPaperCleaner(false);
      document.getElementById('cleanClarity').oninput = () => runPaperCleaner(false);
      document.getElementById('btnResetClean').onclick = resetPaperCleaner;
      break;
  }
}

// ==========================================
// 3. OBJECT ADJUSTMENT & BOUNDING BOX ENGINE
// ==========================================
function createObjectWrapper(type, contentElement, initialWidth, initialHeight) {
  const id = 'obj_' + Date.now();
  const wrapper = document.createElement('div');
  wrapper.className = `doc-object type-${type}`;
  wrapper.id = id;
  wrapper.style.left = '50px';
  wrapper.style.top = '50px';
  wrapper.style.width = initialWidth + 'px';
  wrapper.style.height = initialHeight + 'px';
  wrapper.style.zIndex = State.objects.length + 1;

  contentElement.style.width = '100%';
  contentElement.style.height = '100%';
  contentElement.style.pointerEvents = 'none';
  wrapper.appendChild(contentElement);

  // Resize Handles
  ['nw', 'ne', 'sw', 'se'].forEach(dir => {
    const handle = document.createElement('div');
    handle.className = `resize-handle handle-${dir}`;
    handle.dataset.handle = dir;
    wrapper.appendChild(handle);
  });

  layersContainer.appendChild(wrapper);

  const objRecord = {
    id,
    type,
    x: 50,
    y: 50,
    width: initialWidth,
    height: initialHeight,
    zIndex: State.objects.length + 1,
    locked: false,
    el: wrapper,
    imgElement: contentElement.tagName === 'IMG' ? contentElement : null
  };

  State.objects.push(objRecord);
  attachTransformGestures(objRecord);
  selectObject(objRecord);
  recordState();
  return objRecord;
}

function selectObject(obj) {
  // Enforce rule: Menu strictly controls selection
  if (obj.type !== State.currentMenu && !(State.currentMenu === 'pdf' && obj.type === 'signature')) {
    return;
  }
  deselectAllObjects();
  State.selectedObjectId = obj.id;
  obj.el.classList.add('selected');
  updateLockStateUI();
}

function deselectAllObjects() {
  State.selectedObjectId = null;
  State.objects.forEach(o => o.el.classList.remove('selected'));
}

function getSelectedObject() {
  return State.objects.find(o => o.id === State.selectedObjectId);
}

function autoSelectTopObjectForMenu(menu) {
  const match = State.objects.slice().reverse().find(o => o.type === menu);
  if (match) selectObject(match);
}

// Touch & Mouse Drag / Scale Gestures
function attachTransformGestures(obj) {
  const el = obj.el;
  let isDragging = false;
  let isResizing = false;
  let activeHandle = null;
  let startX, startY, startW, startH, startLeft, startTop;

  const onPointerDown = (e) => {
    // Only interact if matches active menu
    if (obj.type !== State.currentMenu && !(State.currentMenu === 'pdf' && obj.type === 'signature')) return;
    if (obj.locked) return;

    selectObject(obj);
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    const clientY = e.touches ? e.touches[0].clientY : e.clientY;

    if (e.target.classList.contains('resize-handle')) {
      isResizing = true;
      activeHandle = e.target.dataset.handle;
      startX = clientX;
      startY = clientY;
      startW = obj.width;
      startH = obj.height;
      startLeft = obj.x;
      startTop = obj.y;
    } else {
      isDragging = true;
      startX = clientX - obj.x;
      startY = clientY - obj.y;
    }
    e.stopPropagation();
  };

  const onPointerMove = (e) => {
    if (!isDragging && !isResizing) return;
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    const clientY = e.touches ? e.touches[0].clientY : e.clientY;

    if (isDragging) {
      obj.x = clientX - startX;
      obj.y = clientY - startY;
      el.style.left = obj.x + 'px';
      el.style.top = obj.y + 'px';
    } else if (isResizing) {
      const dx = clientX - startX;
      const ratio = startH / startW;

      if (activeHandle === 'se') {
        obj.width = Math.max(30, startW + dx);
        obj.height = obj.width * ratio;
      } else if (activeHandle === 'sw') {
        const newW = Math.max(30, startW - dx);
        obj.x = startLeft + (startW - newW);
        obj.width = newW;
        obj.height = obj.width * ratio;
        el.style.left = obj.x + 'px';
      }
      el.style.width = obj.width + 'px';
      el.style.height = obj.height + 'px';
    }
  };

  const onPointerUp = () => {
    if (isDragging || isResizing) {
      isDragging = false;
      isResizing = false;
      recordState();
    }
  };

  el.addEventListener('mousedown', onPointerDown);
  el.addEventListener('touchstart', onPointerDown, { passive: false });
  window.addEventListener('mousemove', onPointerMove);
  window.addEventListener('touchmove', onPointerMove, { passive: false });
  window.addEventListener('mouseup', onPointerUp);
  window.addEventListener('touchend', onPointerUp);
}

// ==========================================
// 4. IMAGE PROCESSING (Signature Tint, Clean)
// ==========================================
function handleImageUpload(e, type) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    const img = new Image();
    img.src = evt.target.result;
    img.onload = () => {
      let w = img.naturalWidth;
      let h = img.naturalHeight;
      // Normalise size
      if (w > 500) { h = (500 / w) * h; w = 500; }
      createObjectWrapper(type, img, w, h);
    };
  };
  reader.readAsDataURL(file);
}

function removeWhiteBackground(obj, threshold = 210) {
  if (!obj || !obj.imgElement) return;
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');
  const img = obj.imgElement;

  canvas.width = img.naturalWidth;
  canvas.height = img.naturalHeight;
  ctx.drawImage(img, 0, 0);

  const imgData = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const data = imgData.data;

  for (let i = 0; i < data.length; i += 4) {
    const r = data[i], g = data[i+1], b = data[i+2];
    const brightness = (r * 0.299 + g * 0.587 + b * 0.114);
    if (brightness > threshold) {
      data[i + 3] = 0; // Alpha to transparent
    }
  }
  ctx.putImageData(imgData, 0, 0);
  img.src = canvas.toDataURL();
  recordState();
}

function tintSignature(obj, hexColor) {
  if (!obj || !obj.imgElement || hexColor === 'original') return;
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');
  const img = obj.imgElement;

  canvas.width = img.naturalWidth;
  canvas.height = img.naturalHeight;
  ctx.drawImage(img, 0, 0);

  ctx.globalCompositeOperation = 'source-in';
  ctx.fillStyle = hexColor;
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  img.src = canvas.toDataURL();
  recordState();
}

function applyCSSFilters(obj, bright, contrast) {
  if (!obj) return;
  const bVal = bright !== null ? bright : 0;
  const cVal = contrast !== null ? contrast : 0;
  obj.el.style.filter = `brightness(${100 + parseInt(bVal)}%) contrast(${100 + parseInt(cVal)}%)`;
}

// Paper Cleaner Filter
function runPaperCleaner(isAuto) {
  const obj = getSelectedObject() || State.objects.find(o => o.type === 'docimage');
  if (!obj || !obj.imgElement) return;

  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');
  const img = obj.imgElement;
  canvas.width = img.naturalWidth;
  canvas.height = img.naturalHeight;
  ctx.drawImage(img, 0, 0);

  if (!State.cleanerOriginalImageData) {
    State.cleanerOriginalImageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
  }

  const imgData = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const data = imgData.data;
  const whitenessThreshold = isAuto ? 190 : parseInt(document.getElementById('cleanWhiteness').value);
  const contrastFactor = 1.35;

  for (let i = 0; i < data.length; i += 4) {
    let lum = data[i] * 0.299 + data[i+1] * 0.587 + data[i+2] * 0.114;
    if (lum > whitenessThreshold) {
      data[i] = 255;
      data[i+1] = 255;
      data[i+2] = 255;
    } else {
      // Enhance text sharpness & depth
      let c = ((lum - 128) * contrastFactor) + 128;
      c = Math.max(0, Math.min(255, c));
      data[i] = c; data[i+1] = c; data[i+2] = c;
    }
  }

  ctx.putImageData(imgData, 0, 0);
  img.src = canvas.toDataURL();
}

function resetPaperCleaner() {
  const obj = getSelectedObject();
  if (obj && obj.imgElement && State.cleanerOriginalImageData) {
    const canvas = document.createElement('canvas');
    canvas.width = State.cleanerOriginalImageData.width;
    canvas.height = State.cleanerOriginalImageData.height;
    canvas.getContext('2d').putImageData(State.cleanerOriginalImageData, 0, 0);
    obj.imgElement.src = canvas.toDataURL();
    State.cleanerOriginalImageData = null;
  }
}

function fitBackgroundToPage() {
  const obj = getSelectedObject();
  if (obj && obj.type === 'background') {
    obj.x = 0;
    obj.y = 0;
    obj.width = paperPage.clientWidth;
    obj.height = paperPage.clientHeight;
    obj.el.style.left = '0px';
    obj.el.style.top = '0px';
    obj.el.style.width = obj.width + 'px';
    obj.el.style.height = obj.height + 'px';
    recordState();
  }
}

// ==========================================
// 5. WORD & PDF IMPORTERS
// ==========================================
function handleDocxImport(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(event) {
    mammoth.convertToHtml({ arrayBuffer: event.target.result })
      .then(res => {
        textEditorRoot.innerHTML = res.value;
        switchMenu('texteditor');
      })
      .catch(console.error);
  };
  reader.readAsArrayBuffer(file);
}

let activePdfDoc = null;
let currentPdfPage = 1;

function handlePdfImport(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function() {
    const typedarray = new Uint8Array(this.result);
    pdfjsLib.getDocument(typedarray).promise.then(pdf => {
      activePdfDoc = pdf;
      currentPdfPage = 1;
      renderPdfPage(1);
    });
  };
  reader.readAsArrayBuffer(file);
}

function renderPdfPage(pageNumber) {
  if (!activePdfDoc) return;
  activePdfDoc.getPage(pageNumber).then(page => {
    const scale = 1.5;
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    canvas.height = viewport.height;
    canvas.width = viewport.width;

    page.render({ canvasContext: context, viewport }).promise.then(() => {
      const img = new Image();
      img.src = canvas.toDataURL();
      img.onload = () => {
        // Clear prior background and set this page as base
        State.objects = State.objects.filter(o => o.type !== 'background');
        createObjectWrapper('background', img, paperPage.clientWidth, paperPage.clientHeight);
      };
      document.getElementById('pdfPageIndicator').innerText = `Page ${pageNumber} of ${activePdfDoc.numPages}`;
    });
  });
}

// ==========================================
// 6. PAGE SETUP & MARGINS
// ==========================================
function setupPageSetupModal() {
  document.getElementById('btnPageSetup').onclick = () => {
    document.getElementById('pageSetupModal').style.display = 'flex';
  };
  document.getElementById('btnApplyPageSetup').onclick = () => {
    State.paperSize = document.getElementById('selPaperSize').value;
    State.orientation = document.getElementById('selOrientation').value;
    State.margins.top = parseFloat(document.getElementById('marginTop').value) || 0;
    State.margins.bottom = parseFloat(document.getElementById('marginBottom').value) || 0;
    State.margins.left = parseFloat(document.getElementById('marginLeft').value) || 0;
    State.margins.right = parseFloat(document.getElementById('marginRight').value) || 0;

    updatePageDimensions();
    document.getElementById('pageSetupModal').style.display = 'none';
  };
}

function updatePageDimensions() {
  paperPage.className = `paper-page ${State.paperSize.toLowerCase()} ${State.orientation}`;
  
  // Apply Margins in inches to display overlay and editor padding
  const t = State.margins.top * DPI;
  const b = State.margins.bottom * DPI;
  const l = State.margins.left * DPI;
  const r = State.margins.right * DPI;

  const overlay = document.getElementById('marginOverlay');
  overlay.style.top = `${t}px`;
  overlay.style.bottom = `${b}px`;
  overlay.style.left = `${l}px`;
  overlay.style.right = `${r}px`;

  textEditorRoot.style.padding = `${t}px ${r}px ${b}px ${l}px`;
}

// ==========================================
// 7. UNDO / REDO / LOCK / LAYER CONTROLS
// ==========================================
function setupHistory() {
  document.getElementById('btnUndo').onclick = undo;
  document.getElementById('btnRedo').onclick = redo;
  document.getElementById('btnLock').onclick = toggleLock;
  document.getElementById('btnLayerUp').onclick = () => changeLayer(1);
  document.getElementById('btnLayerDown').onclick = () => changeLayer(-1);
}

function recordState() {
  const snapshot = {
    objects: State.objects.map(o => ({
      id: o.id,
      type: o.type,
      x: o.x,
      y: o.y,
      width: o.width,
      height: o.height,
      zIndex: o.zIndex,
      locked: o.locked,
      src: o.imgElement ? o.imgElement.src : null,
      filter: o.el.style.filter,
      opacity: o.el.style.opacity
    })),
    textHTML: textEditorRoot.innerHTML
  };
  State.history = State.history.slice(0, State.historyIndex + 1);
  State.history.push(JSON.stringify(snapshot));
  State.historyIndex++;
}

function undo() {
  if (State.historyIndex > 0) {
    State.historyIndex--;
    restoreSnapshot(JSON.parse(State.history[State.historyIndex]));
  }
}

function redo() {
  if (State.historyIndex < State.history.length - 1) {
    State.historyIndex++;
    restoreSnapshot(JSON.parse(State.history[State.historyIndex]));
  }
}

function restoreSnapshot(data) {
  layersContainer.innerHTML = '';
  layersContainer.appendChild(textEditorRoot);
  textEditorRoot.innerHTML = data.textHTML;
  State.objects = [];

  data.objects.forEach(item => {
    const img = new Image();
    img.src = item.src;
    const obj = createObjectWrapper(item.type, img, item.width, item.height);
    obj.x = item.x;
    obj.y = item.y;
    obj.locked = item.locked;
    obj.zIndex = item.zIndex;
    obj.el.style.left = item.x + 'px';
    obj.el.style.top = item.y + 'px';
    obj.el.style.zIndex = item.zIndex;
    obj.el.style.filter = item.filter;
    obj.el.style.opacity = item.opacity;
    if (obj.locked) obj.el.classList.add('locked');
  });
}

function toggleLock() {
  const obj = getSelectedObject();
  if (!obj) return;
  obj.locked = !obj.locked;
  obj.el.classList.toggle('locked', obj.locked);
  updateLockStateUI();
  recordState();
}

function updateLockStateUI() {
  const obj = getSelectedObject();
  document.getElementById('btnLock').innerText = obj && obj.locked ? '🔒' : '🔓';
}

function changeLayer(delta) {
  const obj = getSelectedObject();
  if (!obj) return;
  obj.zIndex = Math.max(1, obj.zIndex + delta);
  obj.el.style.zIndex = obj.zIndex;
  recordState();
}

function removeCurrentObject() {
  const obj = getSelectedObject();
  if (!obj) return;
  obj.el.remove();
  State.objects = State.objects.filter(o => o.id !== obj.id);
  deselectAllObjects();
  recordState();
}

// ==========================================
// 8. PROJECT SAVE / OPEN & EXPORT / SHARE
// ==========================================
document.getElementById('btnSaveProject').onclick = () => {
  const projectData = {
    version: '1.0',
    paperSize: State.paperSize,
    orientation: State.orientation,
    margins: State.margins,
    textHTML: textEditorRoot.innerHTML,
    objects: State.objects.map(o => ({
      type: o.type, x: o.x, y: o.y, width: o.width, height: o.height,
      zIndex: o.zIndex, locked: o.locked, src: o.imgElement ? o.imgElement.src : null
    }))
  };
  const blob = new Blob([JSON.stringify(projectData)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `DocProject_${Date.now()}.ddpro`;
  a.click();
};

document.getElementById('btnOpenProject').onclick = () => {
  document.getElementById('projectFileInput').click();
};

document.getElementById('projectFileInput').onchange = (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    const data = JSON.parse(evt.target.result);
    State.paperSize = data.paperSize;
    State.orientation = data.orientation;
    State.margins = data.margins;
    updatePageDimensions();
    restoreSnapshot(data);
  };
  reader.readAsText(file);
};

document.getElementById('btnExport').onclick = () => {
  document.getElementById('exportModal').style.display = 'flex';
};

document.getElementById('btnExportPNG').onclick = () => exportAsRaster('png');
document.getElementById('btnExportJPEG').onclick = () => exportAsRaster('jpeg');

function exportAsRaster(format) {
  deselectAllObjects();
  document.getElementById('marginOverlay').style.display = 'none';

  html2canvas(paperPage, { scale: 2 }).then(canvas => {
    document.getElementById('marginOverlay').style.display = 'block';
    const a = document.createElement('a');
    a.href = canvas.toDataURL(`image/${format}`, 0.95);
    a.download = `Document_${Date.now()}.${format}`;
    a.click();
  });
}

document.getElementById('btnExportPDF').onclick = () => {
  deselectAllObjects();
  document.getElementById('marginOverlay').style.display = 'none';

  html2canvas(paperPage, { scale: 2 }).then(canvas => {
    document.getElementById('marginOverlay').style.display = 'block';
    const imgData = canvas.toDataURL('image/jpeg', 0.95);
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({
      orientation: State.orientation,
      unit: 'mm',
      format: State.paperSize.toLowerCase()
    });

    const w = State.orientation === 'portrait' ? 210 : 297;
    const h = State.orientation === 'portrait' ? 297 : 210;
    pdf.addImage(imgData, 'JPEG', 0, 0, w, h);
    pdf.save(`Document_${Date.now()}.pdf`);
  });
};

document.getElementById('btnSystemShare').onclick = async () => {
  deselectAllObjects();
  document.getElementById('marginOverlay').style.display = 'none';

  const canvas = await html2canvas(paperPage, { scale: 2 });
  document.getElementById('marginOverlay').style.display = 'block';

  canvas.toBlob(async (blob) => {
    const file = new File([blob], 'document.png', { type: 'image/png' });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try {
        await navigator.share({
          files: [file],
          title: 'Document Designer Export',
          text: 'Here is my designed document.'
        });
      } catch (err) {
        console.log('Share canceled or failed:', err);
      }
    } else {
      alert('System Web Share is not supported in this browser. Downloading directly.');
      exportAsRaster('png');
    }
  });
};

function setupGlobalShortcuts() {
  window.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'z') { e.preventDefault(); undo(); }
    if ((e.ctrlKey || e.metaKey) && e.key === 'y') { e.preventDefault(); redo(); }
    if (e.key === 'Delete') removeCurrentObject();
  });
}
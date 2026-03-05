import { initializeApp }       from 'https://www.gstatic.com/firebasejs/12.10.0/firebase-app.js';
import { getAuth, createUserWithEmailAndPassword,
         signInWithEmailAndPassword, onAuthStateChanged,
         signOut, sendPasswordResetEmail } from 'https://www.gstatic.com/firebasejs/12.10.0/firebase-auth.js';
import { getFirestore, doc,
         getDoc, setDoc }        from 'https://www.gstatic.com/firebasejs/12.10.0/firebase-firestore.js';

// ─── Firebase init ────────────────────────────────────────────────────────────
const firebaseConfig = {
  apiKey:            'AIzaSyBxjbIOV8cxfVDLt_gAefRuWn21ULjDRD4',
  authDomain:        'simple-spreadsheets.firebaseapp.com',
  projectId:         'simple-spreadsheets',
  storageBucket:     'simple-spreadsheets.firebasestorage.app',
  messagingSenderId: '295138315683',
  appId:             '1:295138315683:web:018699551e44f32df3de16',
};

const app  = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db   = getFirestore(app);

const INIT_ROWS = 100; const INIT_COLS = 26; const EXPAND_BUFFER = 5; const SAVE_DELAY_MS = 1500;
const MIN_EMPTY_ROWS = 25; const ROWS_TO_ADD = 25;

// ─── DOM refs ─────────────────────────────────────────────────────────────────
const authOverlay = document.getElementById('auth-overlay');
const authTitle = document.getElementById('auth-title');
const authEmail = document.getElementById('auth-email');
const authPassword = document.getElementById('auth-password');
const authSubmit = document.getElementById('auth-submit');
const authToggle = document.getElementById('auth-toggle');
const authError = document.getElementById('auth-error');
const togglePassword = document.getElementById('toggle-password');
const iconEye = document.getElementById('icon-eye');
const iconEyeOff = document.getElementById('icon-eye-off');

const statusBar = document.getElementById('status-bar');
const sheetNameInput = document.getElementById('sheet-name');
const btnBold = document.getElementById('btn-bold');
const btnItalic = document.getElementById('btn-italic');
const btnSettings = document.getElementById('btn-settings');
// Status message: shows saving progress to the left of the Settings icon
const statusMsg = document.createElement('span');
statusMsg.id = 'status-message';
statusMsg.textContent = '';
statusMsg.style.fontSize = '12px';
statusMsg.style.color = 'var(--text-muted)';
statusMsg.style.opacity = '0';
statusMsg.style.transition = 'opacity 0.3s ease';
btnSettings.parentElement.insertBefore(statusMsg, btnSettings);
let statusHideTimer = null;
const signOutBtn = document.getElementById('sign-out-btn');
const menuToggle = document.getElementById('menu-toggle');

const settingsOverlay = document.getElementById('settings-overlay');
const checkSticky = document.getElementById('check-sticky');
const closeSettings = document.getElementById('close-settings');

const sidebar = document.getElementById('sidebar');
const sidebarOverlay = document.getElementById('sidebar-overlay');
const sheetList = document.getElementById('sheet-list');
const newSheetBtn = document.getElementById('new-sheet-btn');

const sheet = document.getElementById('sheet');
const wrapper = document.getElementById('sheet-wrapper');

// ─── State ────────────────────────────────────────────────────────────────────
let data = {}; let styles = {}; let config = { stickyTop: false };
let columnWidths = {}; // { colIndex: widthInPx }
let numRows = INIT_ROWS; let numCols = INIT_COLS;
let focusedCell = null; let currentUid = null; let currentSheetId = 'default';
let sheetsIndex = []; let trashIndex = []; let isSignUp = false; let saveTimer = null;
let resizingCol = null; let resizeStartX = 0; let resizeStartWidth = 0;
let fillHandleActive = false; let fillStartCell = null; let fillDragCells = [];
let selectedCells = []; // Array of {r, c} for multi-selection
let isSelecting = false; let selectionStartCell = null;
let isMultiSelectDragging = false; let multiSelectStart = null;
let undoStack = []; let redoStack = [];

// ─── Auth ─────────────────────────────────────────────────────────────────────
authToggle.onclick = () => {
  isSignUp = !isSignUp;
  authTitle.textContent = isSignUp ? 'Sign up' : 'Sign in';
  authSubmit.textContent = isSignUp ? 'Sign up' : 'Sign in';
  authToggle.textContent = isSignUp ? 'Already have an account? Sign in' : "Don't have an account? Sign up";
  forgotPassword.parentElement.style.visibility = isSignUp ? 'hidden' : 'visible';
  authError.textContent = '';
};

togglePassword.onclick = () => {
  const visible = authPassword.type === 'text';
  authPassword.type = visible ? 'password' : 'text';
  iconEye.style.display = visible ? '' : 'none';
  iconEyeOff.style.display = visible ? 'none' : '';
};

authSubmit.onclick = async () => {
  const email = authEmail.value.trim(); const pass = authPassword.value;
  if (!email || !pass) return;
  try {
    if (isSignUp) await createUserWithEmailAndPassword(auth, email, pass);
    else await signInWithEmailAndPassword(auth, email, pass);
  } catch (e) { authError.textContent = e.message; }
};

const logoutOverlay = document.getElementById('logout-overlay');
const confirmLogout = document.getElementById('confirm-logout');
const cancelLogout  = document.getElementById('cancel-logout');

signOutBtn.onclick = () => {
  logoutOverlay.classList.add('active');
};

confirmLogout.onclick = () => {
  logoutOverlay.classList.remove('active');
  signOut(auth);
};

cancelLogout.onclick = () => {
  logoutOverlay.classList.remove('active');
};

const forgotPassword = document.getElementById('forgot-password');

// ... (Auth Logic)

forgotPassword.onclick = async () => {
  const email = authEmail.value.trim();
  if (!email) { authError.textContent = 'Enter your email first.'; return; }
  try {
    await sendPasswordResetEmail(auth, email);
    authError.style.color = 'var(--accent)';
    authError.textContent = 'Reset email sent! Check your inbox.';
  } catch (e) { authError.style.color = 'var(--danger)'; authError.textContent = e.message; }
};



onAuthStateChanged(auth, async (user) => {
  if (user) {
    currentUid = user.uid;
    await loadIndex();
    await switchSheet(sheetsIndex[0]?.id || 'default');
    authOverlay.classList.remove('active');
    authOverlay.classList.add('hidden');
    statusBar.classList.add('visible');
  } else {
    currentUid = null;
    authOverlay.classList.add('active');
    authOverlay.classList.remove('hidden');
    statusBar.classList.remove('visible');
    sheet.innerHTML = ''; data = {}; styles = {}; config = { stickyTop: false };
    sheetsIndex = []; trashIndex = [];
  }
});

// ─── Sheet Management ────────────────────────────────────────────────────────

async function loadIndex() {
  const ref = doc(db, 'users', currentUid, 'meta', 'index');
  const snap = await getDoc(ref);
  if (snap.exists()) {
    const d = snap.data();
    sheetsIndex = d.sheets || [];
    trashIndex = d.trash || [];
    
    // Auto-cleanup trash older than 30 days
    const thirtyDaysAgo = Date.now() - (30 * 24 * 60 * 60 * 1000);
    const initialTrashCount = trashIndex.length;
    trashIndex = trashIndex.filter(item => item.deletedAt > thirtyDaysAgo);
    if (trashIndex.length !== initialTrashCount) await saveIndex();

  } else {
    sheetsIndex = [{ id: 'default', name: 'New Sheet 1' }];
    trashIndex = [];
    await setDoc(ref, { sheets: sheetsIndex, trash: trashIndex });
  }
  renderSheetList();
  renderTrash();
}

async function saveIndex() {
  await setDoc(doc(db, 'users', currentUid, 'meta', 'index'), { sheets: sheetsIndex, trash: trashIndex });
}

async function switchSheet(id) {
  if (saveTimer) await saveToCloud();
  currentSheetId = id;
  const meta = sheetsIndex.find(s => s.id === id);
  sheetNameInput.value = meta ? meta.name : 'Untitled Spreadsheet';
  await loadFromCloud();
  buildGrid();
  applyConfig();
  renderSheetList();
  sidebar.classList.remove('active'); sidebarOverlay.classList.remove('active');
  restoreScroll();
}

async function loadFromCloud() {
  const ref = doc(db, 'users', currentUid, 'sheets', currentSheetId);
  const snap = await getDoc(ref);
  if (snap.exists()) {
    const d = snap.data(); data = d.cells || {}; styles = d.styles || {};
    config = d.config || { stickyTop: false };
    columnWidths = d.columnWidths || {};
    numRows = Math.max(INIT_ROWS, d.numRows || INIT_ROWS);
    numCols = Math.max(INIT_COLS, d.numCols || INIT_COLS);
  } else {
    data = {}; styles = {}; config = { stickyTop: false };
    columnWidths = {};
    numRows = INIT_ROWS; numCols = INIT_COLS;
  }
}

async function saveToCloud() {
  // If not authenticated, consider as saved to avoid blocking UI
  if (!currentUid) {
    return;
  }
  
  // Clear any pending hide timer
  if (statusHideTimer) clearTimeout(statusHideTimer);
  
  // Show the status message
  statusMsg.style.opacity = '1';
  statusMsg.textContent = 'Saving…';
  
  const ref = doc(db, 'users', currentUid, 'sheets', currentSheetId);
  await setDoc(ref, {
    cells: data, styles, config, columnWidths, numRows, numCols, updatedAt: new Date().toISOString()
  });
  
  statusMsg.textContent = 'Saved!';
  
  // Hide the message after 2 seconds
  statusHideTimer = setTimeout(() => {
    statusMsg.style.opacity = '0';
    statusMsg.textContent = '';
  }, 2000);
}

function scheduleSave() {
  if (saveTimer) clearTimeout(saveTimer);
  saveTimer = setTimeout(saveToCloud, SAVE_DELAY_MS);
}

function renderSheetList() {
  sheetList.innerHTML = '';
  sheetsIndex.forEach(s => {
    const div = document.createElement('div');
    div.className = `sheet-item ${s.id === currentSheetId ? 'active' : ''}`;
    
    const nameSpan = document.createElement('span');
    nameSpan.className = 'sheet-item-name';
    nameSpan.textContent = s.name;
    nameSpan.onclick = () => switchSheet(s.id);
    
    const menuContainer = document.createElement('div');
    menuContainer.className = 'sheet-menu-container';

    const dotsBtn = document.createElement('button');
    dotsBtn.className = 'btn-item-action';
    dotsBtn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="1.5"></circle><circle cx="12" cy="5" r="1.5"></circle><circle cx="12" cy="19" r="1.5"></circle></svg>';
    
    const menu = document.createElement('div');
    menu.className = 'sheet-options-menu';
    
    const renameOpt = document.createElement('button');
    renameOpt.textContent = 'Rename';
    renameOpt.onclick = (e) => {
      e.stopPropagation(); menu.classList.remove('active');
      const input = document.createElement('input');
      input.value = s.name;
      div.innerHTML = ''; div.appendChild(input);
      input.focus();
      input.onblur = () => { s.name = input.value || 'Untitled'; saveIndex(); renderSheetList(); if(s.id === currentSheetId) sheetNameInput.value = s.name; };
      input.onkeydown = (ev) => { if(ev.key === 'Enter') input.blur(); };
    };

    const deleteOpt = document.createElement('button');
    deleteOpt.textContent = 'Delete';
    deleteOpt.className = 'btn-delete';
    deleteOpt.onclick = (e) => {
      e.stopPropagation(); menu.classList.remove('active');
      if (sheetsIndex.length <= 1) { alert("You can't delete your last spreadsheet."); return; }
      
      const overlay = document.getElementById('delete-overlay');
      const msg = document.getElementById('delete-confirm-msg');
      const confirmBtn = document.getElementById('confirm-delete');
      const cancelBtn = document.getElementById('cancel-delete');
      
      msg.textContent = `Move "${s.name}" to archive? It will be permanently deleted after 30 days.`;
      overlay.classList.add('active');
      
      confirmBtn.onclick = async () => {
        // MOVE TO ARCHIVE
        const itemToTrash = sheetsIndex.find(item => item.id === s.id);
        if (itemToTrash) {
          trashIndex.push({ ...itemToTrash, deletedAt: Date.now() });
          sheetsIndex = sheetsIndex.filter(item => item.id !== s.id);
          await saveIndex();
          overlay.classList.remove('active');
          if (s.id === currentSheetId) {
            await switchSheet(sheetsIndex[0].id);
          } else {
            renderSheetList();
            renderTrash();
          }
        }
      };
      cancelBtn.onclick = () => overlay.classList.remove('active');
    };

    dotsBtn.onclick = (e) => {
      e.stopPropagation();
      document.querySelectorAll('.sheet-options-menu').forEach(m => { if (m !== menu) m.classList.remove('active'); });
      menu.classList.toggle('active');
    };

    menu.appendChild(renameOpt); menu.appendChild(deleteOpt);
    menuContainer.appendChild(dotsBtn); menuContainer.appendChild(menu);
    div.appendChild(nameSpan); div.appendChild(menuContainer);
    sheetList.appendChild(div);
  });
}

function renderTrash() {
  const list = document.getElementById('trash-list');
  list.innerHTML = '';
  if (trashIndex.length === 0) {
    list.innerHTML = '<div style="font-size: 11px; color: var(--text-muted); padding: 10px; text-align: center;">Archive is empty</div>';
    return;
  }
  
  trashIndex.forEach(s => {
    const item = document.createElement('div');
    item.className = 'trash-item';
    
    const nameSpan = document.createElement('span');
    nameSpan.className = 'trash-name';
    nameSpan.textContent = s.name;
    
    const rightSide = document.createElement('div');
    rightSide.style.display = 'flex';
    rightSide.style.alignItems = 'center';
    
    const daysLeft = Math.ceil((30 * 24 * 60 * 60 * 1000 - (Date.now() - s.deletedAt)) / (24 * 60 * 60 * 1000));
    const daysSpan = document.createElement('span');
    daysSpan.className = 'trash-days';
    daysSpan.textContent = `${daysLeft}d`;
    
    const restoreBtn = document.createElement('button');
    restoreBtn.className = 'btn-restore';
    restoreBtn.textContent = 'Restore';
    restoreBtn.onclick = async () => {
      sheetsIndex.push({ id: s.id, name: s.name });
      trashIndex = trashIndex.filter(item => item.id !== s.id);
      await saveIndex();
      renderSheetList();
      renderTrash();
      await switchSheet(s.id);
    };
    
    rightSide.appendChild(daysSpan);
    rightSide.appendChild(restoreBtn);
    item.appendChild(nameSpan);
    item.appendChild(rightSide);
    list.appendChild(item);
  });
}

const trashHeader = document.getElementById('trash-header');
trashHeader.onclick = () => {
  const section = document.getElementById('trash-section');
  const list = document.getElementById('trash-list');
  section.classList.toggle('expanded');
  list.classList.toggle('hidden');
};


// Close menus when clicking anywhere else
window.addEventListener('click', () => {
  document.querySelectorAll('.sheet-options-menu').forEach(m => m.classList.remove('active'));
});


newSheetBtn.onclick = async () => {
  const id = 'sheet_' + Date.now();
  const num = sheetsIndex.length + trashIndex.length + 1;
  sheetsIndex.push({ id, name: `New Sheet ${num}` });
  await saveIndex(); await switchSheet(id);
};

// ─── Format & Config ─────────────────────────────────────────────────────────

function applyFormatting(r, c, inputEl) {
  const el = inputEl || getCell(r, c)?.querySelector('input');
  if (!el) return;
  const style = styles[`${r},${c}`] || {};
  el.style.fontWeight = style.bold ? 'bold' : 'normal';
  el.style.fontStyle  = style.italic ? 'italic' : 'normal';
}

function toggleStyle(prop) {
  if (!focusedCell) return;
  const key = `${focusedCell.r},${focusedCell.c}`;
  if (!styles[key]) styles[key] = {};
  styles[key][prop] = !styles[key][prop];
  applyFormatting(focusedCell.r, focusedCell.c);
  updateFormatButtons();
  scheduleSave();
}

function updateFormatButtons() {
  if (!focusedCell) { btnBold.classList.remove('active'); btnItalic.classList.remove('active'); return; }
  const style = styles[`${focusedCell.r},${focusedCell.c}`] || {};
  btnBold.classList.toggle('active', !!style.bold);
  btnItalic.classList.toggle('active', !!style.italic);
}


btnBold.onmousedown = (e) => { e.preventDefault(); toggleStyle('bold'); };
btnItalic.onmousedown = (e) => { e.preventDefault(); toggleStyle('italic'); };

// Deselect cleanly when the user switches to another tab/window
window.onblur = () => {
  if (focusedCell) {
    const td = getCell(focusedCell.r, focusedCell.c);
    td?.classList.remove('focused', 'editing');
    td?.querySelector('input')?.blur();
    focusedCell = null;
    updateFormatButtons();
  }
};


btnSettings.onclick = () => {
  checkSticky.checked = config.stickyTop;
  settingsOverlay.classList.add('active');
};

closeSettings.onclick = () => {
  config.stickyTop = checkSticky.checked;
  applyConfig();
  scheduleSave();
  settingsOverlay.classList.remove('active');
};

function applyConfig() {
  sheet.classList.toggle('sticky-top', config.stickyTop);
}

// ─── UI Interactions ─────────────────────────────────────────────────────────

sheetNameInput.oninput = () => {
  const meta = sheetsIndex.find(s => s.id === currentSheetId);
  if (meta) { meta.name = sheetNameInput.value; renderSheetList(); if (window.indexTimer) clearTimeout(window.indexTimer); window.indexTimer = setTimeout(saveIndex, 1000); }
};

menuToggle.onclick = () => { sidebar.classList.toggle('active'); sidebarOverlay.classList.toggle('active'); };
sidebarOverlay.onclick = () => { sidebar.classList.remove('active'); sidebarOverlay.classList.remove('active'); };

statusBar.onclick = (e) => {
  if (e.target.id === 'sheet-name' || e.target.closest('button') || e.target.closest('#format-bar')) return;
  deselectAll();
};

function deselectAll() {
  // Clear all selected cells
  selectedCells.forEach(cellPos => {
    const td = getCell(cellPos.r, cellPos.c);
    if (td) td.classList.remove('selected');
  });
  selectedCells = [];
  
  if (focusedCell) {
    const td = getCell(focusedCell.r, focusedCell.c);
    // Explicitly remove highlight classes first — sticky cells can
    // prevent the browser's native blur from cleaning these up reliably.
    td?.classList.remove('focused', 'editing');
    
    // Clear fill handle and preview
    const fillHandle = td?.querySelector('.fill-handle');
    if (fillHandle) fillHandle.remove();
    
    td?.querySelector('input')?.blur();
    focusedCell = null;
    updateFormatButtons();
  }
}

// Clicking anywhere on the sheet wrapper outside of a cell also deselects
wrapper.addEventListener('mousedown', (e) => {
  if (!e.target.closest('td')) deselectAll();
});

// ─── Multi-Cell Selection ────────────────────────────────────────────────────
function toggleCellSelection(r, c, td) {
  const cellPos = { r, c };
  const existing = selectedCells.findIndex(cell => cell.r === r && cell.c === c);
  
  if (existing !== -1) {
    selectedCells.splice(existing, 1);
    td.classList.remove('selected');
  } else {
    selectedCells.push(cellPos);
    td.classList.add('selected');
  }
}

function selectRange(r1, c1, r2, c2) {
  // Clear previous selection
  selectedCells.forEach(cellPos => {
    const td = getCell(cellPos.r, cellPos.c);
    if (td) td.classList.remove('selected');
  });
  selectedCells = [];
  
  const minR = Math.min(r1, r2);
  const maxR = Math.max(r1, r2);
  const minC = Math.min(c1, c2);
  const maxC = Math.max(c1, c2);
  
  for (let r = minR; r <= maxR; r++) {
    for (let c = minC; c <= maxC; c++) {
      const td = getCell(r, c);
      if (td) {
        td.classList.add('selected');
        selectedCells.push({ r, c });
      }
    }
  }
}

function deleteSelectedCells() {
  selectedCells.forEach(cellPos => {
    const key = `${cellPos.r},${cellPos.c}`;
    delete data[key];
    delete styles[key];
    const input = getCell(cellPos.r, cellPos.c)?.querySelector('input');
    if (input) {
      input.value = '';
      applyFormatting(cellPos.r, cellPos.c);
    }
  });
  scheduleSave();
}

function onMultiSelectDrag(e) {
  if (!isMultiSelectDragging || !multiSelectStart) return;
  
  // Find which cell the mouse is currently over
  const target = document.elementFromPoint(e.clientX, e.clientY);
  const cell = target?.closest('td');
  
  if (!cell) return;
  
  const r = parseInt(cell.dataset.r);
  const c = parseInt(cell.dataset.c);
  
  if (isNaN(r) || isNaN(c)) return;
  
  // Update selection range from start to current position
  selectRange(multiSelectStart.r, multiSelectStart.c, r, c);
}

function onMultiSelectEnd() {
  isMultiSelectDragging = false;
  multiSelectStart = null;
  document.removeEventListener('mousemove', onMultiSelectDrag);
  document.removeEventListener('mouseup', onMultiSelectEnd);
}

// Handle Delete key for selected cells
document.addEventListener('keydown', (e) => {
  if (e.key === 'Delete' && selectedCells.length > 0) {
    e.preventDefault();
    deleteSelectedCells();
  }
});

// Handle Ctrl+Z for undo
document.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
    e.preventDefault();
    performUndo();
  }
});

// Handle Ctrl+Y for redo
document.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && (e.key === 'y' || (e.shiftKey && e.key === 'z'))) {
    e.preventDefault();
    performRedo();
  }
});

// ─── Grid Logic ──────────────────────────────────────────────────────────────

function buildGrid() {
  sheet.innerHTML = '';
  for (let r = 0; r < numRows; r++) addRow(r, false);
}

function addRow(r, animate = true) {
  const tr = document.createElement('tr'); if (animate) tr.classList.add('new-row');
  for (let c = 0; c < numCols; c++) tr.appendChild(makeCell(r, c));
  sheet.appendChild(tr);
}

function addColToAllRows() {
  const rows = sheet.rows; const c = numCols - 1;
  for (let r = 0; r < rows.length; r++) rows[r].appendChild(makeCell(r, c));
}

function makeCell(r, c) {
  const td = document.createElement('td'); const input = document.createElement('input');
  const key = `${r},${c}`; input.value = data[key] || '';
  td.dataset.r = r; td.dataset.c = c; td.appendChild(input);
  
  // Apply custom column width if set
  if (columnWidths[c]) {
    td.style.width = columnWidths[c] + 'px';
  }
  
  // Pass input directly — td isn't in the DOM yet so getCell() would return null
  applyFormatting(r, c, input);
  
  // Add resize handle on the right edge of cells in the first row
  if (r === 0) {
    const resizeHandle = document.createElement('div');
    resizeHandle.className = 'col-resize-handle';
    resizeHandle.dataset.col = c;
    td.appendChild(resizeHandle);
    resizeHandle.onmousedown = (e) => startResizing(e, c);
    resizeHandle.ondblclick = (e) => resetColumnWidth(e, c);
  }
  
  input.onfocus = () => {
    deselectAll();
    focusedCell = { r, c };
    td.classList.add('focused');
    selectedCells = [{ r, c }];
    
    // Add fill handle to the focused cell
    addFillHandle(td, r, c);
    
    expandIfNeeded(r, c);
    updateFormatButtons();
  };
  
  td.onmousedown = (e) => {
    if (e.target === input) return; // Don't interfere with input focus
    
    if (e.ctrlKey || e.metaKey) {
      // Ctrl/Cmd+Drag to select range or Ctrl/Cmd+Click to toggle
      e.preventDefault();
      isMultiSelectDragging = true;
      multiSelectStart = { r, c };
      // Start with the clicked cell
      selectRange(r, c, r, c);
      document.addEventListener('mousemove', onMultiSelectDrag);
      document.addEventListener('mouseup', onMultiSelectEnd);
    } else if (e.shiftKey) {
      // Shift+Click to select range
      e.preventDefault();
      if (focusedCell) {
        selectRange(focusedCell.r, focusedCell.c, r, c);
      }
    }
  };
  input.oninput = () => {
    data[key] = input.value;
    if (!input.value) {
      // RESET STYLE ON EMPTY
      delete data[key];
      delete styles[key];
      applyFormatting(r, c);
      updateFormatButtons();
    }
    td.classList.add('editing');
    ensureEmptyRows();
    scheduleSave();
  };
  input.onblur = () => {
    td.classList.remove('focused', 'editing');
    if (!input.value) {
      delete styles[key];
      applyFormatting(r, c);
      updateFormatButtons();
      scheduleSave();
    }
  };
  input.onkeydown = (e) => handleKeyNav(e, r, c);
  return td;
}

function handleKeyNav(e, r, c) {
  if (e.ctrlKey && e.key === 'b') { e.preventDefault(); toggleStyle('bold'); return; }
  if (e.ctrlKey && e.key === 'i') { e.preventDefault(); toggleStyle('italic'); return; }

  const map = { Tab: [0, 1], Enter: [1, 0], ArrowUp: [-1, 0], ArrowDown: [1, 0] };
  if (map[e.key]) {
    e.preventDefault(); let [dr, dc] = map[e.key];
    if (e.key === 'Tab' && e.shiftKey) dc = -1;
    moveFocus(r + dr, c + dc);
  }
}

function moveFocus(r, c) {
  r = Math.max(0, r); c = Math.max(0, c);
  while (r >= numRows) { numRows++; addRow(numRows - 1); }
  while (c >= numCols) { numCols++; addColToAllRows(); }
  getCell(r, c)?.querySelector('input')?.focus();
}

function expandIfNeeded(r, c) {
  if (r >= numRows - EXPAND_BUFFER) { numRows++; addRow(numRows - 1); }
  if (c >= numCols - EXPAND_BUFFER) { numCols++; addColToAllRows(); }
}

function ensureEmptyRows() {
  // Find the last row with any content
  let lastFilledRow = -1;
  
  for (let r = numRows - 1; r >= 0; r--) {
    for (let c = 0; c < numCols; c++) {
      if (data[`${r},${c}`]) {
        lastFilledRow = r;
        break;
      }
    }
    if (lastFilledRow !== -1) break;
  }
  
  // Calculate how many empty rows we have after the last filled row
  const emptyRowsAtEnd = numRows - 1 - lastFilledRow;
  
  // If we don't have enough empty rows, add more
  if (emptyRowsAtEnd < MIN_EMPTY_ROWS) {
    const rowsNeeded = ROWS_TO_ADD;
    for (let i = 0; i < rowsNeeded; i++) {
      numRows++;
      addRow(numRows - 1, false);
    }
  }
}

function getCell(r, c) { return sheet.rows[r]?.cells[c]; }

// ─── Column Resize ──────────────────────────────────────────────────────────
function startResizing(e, colIndex) {
  e.preventDefault();
  e.stopPropagation();
  resizingCol = colIndex;
  resizeStartX = e.clientX;
  
  // Get the current width of the column
  const cell = getCell(0, colIndex);
  if (cell) {
    resizeStartWidth = cell.offsetWidth;
  } else {
    resizeStartWidth = columnWidths[colIndex] || 130; // Default to --cell-w which is 130px
  }
  
  document.addEventListener('mousemove', onResizing);
  document.addEventListener('mouseup', stopResizing);
}

function onResizing(e) {
  if (resizingCol === null) return;
  const deltaX = e.clientX - resizeStartX;
  const newWidth = Math.max(50, resizeStartWidth + deltaX); // Minimum 50px width
  
  columnWidths[resizingCol] = newWidth;
  
  // Apply width to all cells in this column
  for (let r = 0; r < numRows; r++) {
    const cell = getCell(r, resizingCol);
    if (cell) cell.style.width = newWidth + 'px';
  }
}

function stopResizing() {
  if (resizingCol !== null) {
    resizingCol = null;
    scheduleSave();
  }
  document.removeEventListener('mousemove', onResizing);
  document.removeEventListener('mouseup', stopResizing);
}

function resetColumnWidth(e, colIndex) {
  e.preventDefault();
  e.stopPropagation();
  
  // Delete the custom width (this will make it use default)
  delete columnWidths[colIndex];
  
  // Reset all cells in this column to default width
  for (let r = 0; r < numRows; r++) {
    const cell = getCell(r, colIndex);
    if (cell) cell.style.width = '';
  }
  
  scheduleSave();
}

// ─── Fill Handle (Auto-Fill) ────────────────────────────────────────────────────
function addFillHandle(td, r, c) {
  // Remove any existing fill handle
  const existing = td.querySelector('.fill-handle');
  if (existing) existing.remove();
  
  // Create the fill handle
  const handle = document.createElement('div');
  handle.className = 'fill-handle';
  td.appendChild(handle);
  
  handle.onmousedown = (e) => startFill(e, r, c);
}

function startFill(e, startRow, startCol) {
  e.preventDefault();
  e.stopPropagation();
  
  fillHandleActive = true;
  fillStartCell = { r: startRow, c: startCol };
  fillDragCells = [fillStartCell];
  
  // Get the source value
  const sourceKey = `${startRow},${startCol}`;
  const sourceValue = data[sourceKey] || '';
  const sourceStyle = styles[sourceKey] || {};
  
  document.addEventListener('mousemove', onFillDrag);
  document.addEventListener('mouseup', onFillEnd);
}

function onFillDrag(e) {
  if (!fillHandleActive || !fillStartCell) return;
  
  // Auto-scroll when dragging near edges
  const wrapperRect = wrapper.getBoundingClientRect();
  const scrollThreshold = 80; // pixels from edge to trigger scroll
  const scrollSpeed = 15; // pixels to scroll per iteration
  
  if (e.clientY > wrapperRect.bottom - scrollThreshold) {
    // Near bottom - scroll down
    wrapper.scrollTop += scrollSpeed;
  } else if (e.clientY < wrapperRect.top + scrollThreshold) {
    // Near top - scroll up
    wrapper.scrollTop -= scrollSpeed;
  }
  
  // Find which cell the mouse is over
  const target = document.elementFromPoint(e.clientX, e.clientY);
  const cell = target?.closest('td');
  
  if (!cell) return;
  
  const r = parseInt(cell.dataset.r);
  const c = parseInt(cell.dataset.c);
  
  if (isNaN(r) || isNaN(c)) return;
  
  // Determine the fill range
  const minR = Math.min(fillStartCell.r, r);
  const maxR = Math.max(fillStartCell.r, r);
  const minC = Math.min(fillStartCell.c, c);
  const maxC = Math.max(fillStartCell.c, c);
  
  // Clear previous highlighting
  fillDragCells.forEach(cellPos => {
    const td = getCell(cellPos.r, cellPos.c);
    if (td) td.classList.remove('fill-preview');
  });
  
  // Build the new fill range
  fillDragCells = [];
  
  // Handle both vertical and horizontal fill
  const isBothDims = minR < fillStartCell.r || maxR > fillStartCell.r && minC < fillStartCell.c || maxC > fillStartCell.c;
  
  if (fillStartCell.r === minR && fillStartCell.r === maxR) {
    // Horizontal fill
    for (let c = minC; c <= maxC; c++) {
      fillDragCells.push({ r: fillStartCell.r, c });
    }
  } else {
    // Vertical fill
    for (let r = minR; r <= maxR; r++) {
      fillDragCells.push({ r, c: fillStartCell.c });
    }
  }
  
  // Highlight the cells
  fillDragCells.forEach(cellPos => {
    const td = getCell(cellPos.r, cellPos.c);
    if (td) td.classList.add('fill-preview');
  });
}

function onFillEnd() {
  if (!fillHandleActive || !fillStartCell || fillDragCells.length === 0) {
    fillHandleActive = false;
    fillStartCell = null;
    fillDragCells = [];
    document.removeEventListener('mousemove', onFillDrag);
    document.removeEventListener('mouseup', onFillEnd);
    return;
  }
  
  // Save state before fill for undo
  const beforeState = {};
  const affectedCells = [];
  fillDragCells.forEach(cellPos => {
    if (cellPos.r === fillStartCell.r && cellPos.c === fillStartCell.c) return; // Skip source
    const key = `${cellPos.r},${cellPos.c}`;
    beforeState[key] = {
      data: data[key] || '',
      style: styles[key] ? { ...styles[key] } : {}
    };
    affectedCells.push(cellPos);
  });
  
  const sourceKey = `${fillStartCell.r},${fillStartCell.c}`;
  const sourceValue = data[sourceKey] || '';
  const sourceStyle = styles[sourceKey] || {};
  
  // Check if it's a date
  const dateMatch = tryParseDate(sourceValue);
  
  // Apply the fill
  fillDragCells.forEach((cellPos, idx) => {
    if (cellPos.r === fillStartCell.r && cellPos.c === fillStartCell.c) return; // Skip source
    
    const key = `${cellPos.r},${cellPos.c}`;
    
    if (dateMatch) {
      // Increment date for each step from the start cell
      const distance = fillStartCell.r === cellPos.r ? cellPos.c - fillStartCell.c : cellPos.r - fillStartCell.r;
      const newDate = new Date(dateMatch);
      newDate.setDate(newDate.getDate() + distance);
      data[key] = formatDate(newDate);
    } else {
      // Copy the value as-is
      data[key] = sourceValue;
    }
    
    // Copy styles
    if (Object.keys(sourceStyle).length > 0) {
      styles[key] = { ...sourceStyle };
    }
    
    // Update the cell visually
    const input = getCell(cellPos.r, cellPos.c)?.querySelector('input');
    if (input) {
      input.value = data[key];
      applyFormatting(cellPos.r, cellPos.c);
    }
  });
  
  // Save undo state
  undoStack.push({ type: 'fill', beforeState, affectedCells });
  redoStack = []; // Clear redo stack on new action
  
  // Clear the preview highlighting
  fillDragCells.forEach(cellPos => {
    const td = getCell(cellPos.r, cellPos.c);
    if (td) td.classList.remove('fill-preview');
  });
  
  fillHandleActive = false;
  fillStartCell = null;
  fillDragCells = [];
  
  ensureEmptyRows();
  scheduleSave();
  
  document.removeEventListener('mousemove', onFillDrag);
  document.removeEventListener('mouseup', onFillEnd);
}

function tryParseDate(str) {
  if (!str || typeof str !== 'string') return null;
  
  // Try DD-MM-YYYY or DD/MM/YYYY format
  const ddmmyyyy = str.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
  if (ddmmyyyy) {
    const [, day, month, year] = ddmmyyyy;
    const date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
    if (!isNaN(date.getTime())) return date;
  }
  
  return null;
}

function formatDate(date) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}-${month}-${year}`;
}

function performUndo() {
  if (undoStack.length === 0) return;
  
  const action = undoStack.pop();
  
  if (action.type === 'fill') {
    // Save current state for redo
    const afterState = {};
    action.affectedCells.forEach(cellPos => {
      const key = `${cellPos.r},${cellPos.c}`;
      afterState[key] = {
        data: data[key] || '',
        style: styles[key] ? { ...styles[key] } : {}
      };
    });
    redoStack.push({ type: 'fill', beforeState: action.beforeState, afterState, affectedCells: action.affectedCells });
    
    // Restore before state
    action.affectedCells.forEach(cellPos => {
      const key = `${cellPos.r},${cellPos.c}`;
      if (action.beforeState[key]) {
        data[key] = action.beforeState[key].data;
        if (Object.keys(action.beforeState[key].style).length > 0) {
          styles[key] = action.beforeState[key].style;
        } else {
          delete styles[key];
        }
      } else {
        delete data[key];
        delete styles[key];
      }
      
      // Update the cell visually
      const input = getCell(cellPos.r, cellPos.c)?.querySelector('input');
      if (input) {
        input.value = data[key] || '';
        applyFormatting(cellPos.r, cellPos.c);
      }
    });
    
    ensureEmptyRows();
    scheduleSave();
  }
}

function performRedo() {
  if (redoStack.length === 0) return;
  
  const action = redoStack.pop();
  
  if (action.type === 'fill') {
    // Save current state for undo
    const beforeState = {};
    action.affectedCells.forEach(cellPos => {
      const key = `${cellPos.r},${cellPos.c}`;
      beforeState[key] = {
        data: data[key] || '',
        style: styles[key] ? { ...styles[key] } : {}
      };
    });
    undoStack.push({ type: 'fill', beforeState });
    
    // Apply after state
    action.affectedCells.forEach(cellPos => {
      const key = `${cellPos.r},${cellPos.c}`;
      if (action.afterState[key]) {
        data[key] = action.afterState[key].data;
        if (Object.keys(action.afterState[key].style).length > 0) {
          styles[key] = action.afterState[key].style;
        } else {
          delete styles[key];
        }
      } else {
        delete data[key];
        delete styles[key];
      }
      
      // Update the cell visually
      const input = getCell(cellPos.r, cellPos.c)?.querySelector('input');
      if (input) {
        input.value = data[key] || '';
        applyFormatting(cellPos.r, cellPos.c);
      }
    });
    
    ensureEmptyRows();
    scheduleSave();
  }
}

// ─── Scroll Persistence ─────────────────────────────────────────────────────
function scrollKey() { return `scroll_${currentUid}_${currentSheetId}`; }
function saveScroll() {
  if (!currentUid) return; localStorage.setItem(scrollKey(), JSON.stringify({ top: wrapper.scrollTop, left: wrapper.scrollLeft }));
}
function restoreScroll() {
  const raw = localStorage.getItem(scrollKey()); if (!raw) { wrapper.scrollTop = 0; wrapper.scrollLeft = 0; return; }
  const { top, left } = JSON.parse(raw); wrapper.scrollTop = top; wrapper.scrollLeft = left;
}
wrapper.onscroll = () => { if (window.scrollTimer) clearTimeout(window.scrollTimer); window.scrollTimer = setTimeout(saveScroll, 200); };

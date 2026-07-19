/* ============================================================
   Data Filter Add-in — taskpane.js
   GitHub: https://github.com/Luyandantombela/data_filter
   ============================================================ */

'use strict';

// State
let capturedHeaders = [];
let capturedData    = [];    // array of arrays (rows, excluding header)
let capturedAddress = '';
let filterCount     = 0;

/* ── Office Initialise ─────────────────────────────────── */
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    // Ready — nothing extra needed; buttons wire via onclick
  }
});

/* ── Step 1 : Capture the selected range ─────────────────── */
async function captureSelection() {
  clearStatus();
  const btn = document.getElementById('capture-btn');
  btn.disabled = true;
  btn.innerHTML = '<span class="loader"></span> Capturing…';

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(['values', 'address', 'rowCount', 'columnCount']);
      await context.sync();

      if (range.rowCount < 2) {
        showStatus('Please select at least two rows — a header row plus at least one data row.', 'error');
        resetCaptureBtn();
        return;
      }
      if (range.columnCount < 1) {
        showStatus('Please select at least one column.', 'error');
        resetCaptureBtn();
        return;
      }

      const values = range.values;

      // First row = headers
      capturedHeaders = values[0].map(function (h) {
        return h !== null && h !== '' ? String(h) : '(blank)';
      });

      // Remaining rows = data
      capturedData = values.slice(1);
      capturedAddress = range.address;

      const dataRowCount = capturedData.length;

      // Update UI
      document.getElementById('captured-range-label').textContent = capturedAddress;
      document.getElementById('captured-col-count').textContent   = capturedHeaders.length;
      document.getElementById('captured-row-count').textContent   = dataRowCount;

      const infoBox = document.getElementById('captured-info');
      infoBox.classList.add('visible');

      document.getElementById('recapture-btn').style.display = 'inline-flex';
      btn.style.display = 'none';

      // Unlock Step 2 & 3
      unlockFiltersSection();
      unlockApplySection();

      // Clear old filters and add a first one ready to go
      clearAllFilters();
      addFilterRow();
    });
  } catch (err) {
    showStatus('Could not read the selection. ' + (err.message || err), 'error');
    resetCaptureBtn();
  }
}

function resetCaptureBtn() {
  const btn = document.getElementById('capture-btn');
  btn.disabled = false;
  btn.innerHTML = '<svg width="13" height="13" viewBox="0 0 16 16" fill="white"><rect x="1" y="1" width="6" height="6" rx="1" stroke="white" stroke-width="1.5" fill="none"/><rect x="9" y="1" width="6" height="6" rx="1" stroke="white" stroke-width="1.5" fill="none"/><rect x="1" y="9" width="6" height="6" rx="1" stroke="white" stroke-width="1.5" fill="none"/><rect x="9" y="9" width="6" height="6" rx="1" stroke="white" stroke-width="1.5" fill="none"/></svg> Capture Selection';
}

function recapture() {
  // Reset state
  capturedHeaders = [];
  capturedData    = [];
  capturedAddress = '';

  document.getElementById('captured-info').classList.remove('visible');
  document.getElementById('capture-btn').style.display = 'inline-flex';
  document.getElementById('recapture-btn').style.display = 'none';
  resetCaptureBtn();

  lockFiltersSection();
  lockApplySection();
  clearAllFilters();
  clearStatus();
}

/* ── Step 2 : Filter rows ─────────────────────────────── */

function addFilterRow() {
  filterCount++;
  const id = 'filter-' + filterCount;

  // Remove empty-state message if present
  const emptyMsg = document.getElementById('no-filters-msg');
  if (emptyMsg) emptyMsg.remove();

  const row = document.createElement('div');
  row.className = 'filter-row';
  row.id = id;

  // Top controls: column select + value input
  const controls = document.createElement('div');
  controls.className = 'filter-controls';

  const select = document.createElement('select');
  select.title = 'Column';
  const defaultOpt = document.createElement('option');
  defaultOpt.value = '';
  defaultOpt.textContent = '— Column —';
  select.appendChild(defaultOpt);
  select.onchange = function () { refreshApplyBtn(); updateLogicNote(); };

  capturedHeaders.forEach(function (header, idx) {
    const opt = document.createElement('option');
    opt.value = idx;
    opt.textContent = header;
    select.appendChild(opt);
  });

  const input = document.createElement('input');
  input.type        = 'text';
  input.placeholder = 'Value (empty = match blanks)';
  input.title       = 'Value';
  input.oninput = function () { updateLogicNote(); };

  controls.appendChild(select);
  controls.appendChild(input);

  // Footer: ignore toggle + remove button
  const footer = document.createElement('div');
  footer.className = 'filter-footer';

  // Toggle
  const toggleId = 'toggle-' + filterCount;
  const toggleWrap = document.createElement('label');
  toggleWrap.className = 'toggle-wrap';
  toggleWrap.htmlFor   = toggleId;

  const checkbox = document.createElement('input');
  checkbox.type    = 'checkbox';
  checkbox.id      = toggleId;
  checkbox.checked = true; // ON by default

  const track = document.createElement('span');
  track.className = 'toggle-track';

  const toggleLabel = document.createElement('span');
  toggleLabel.className   = 'toggle-label';
  toggleLabel.textContent = 'Ignore casing & spaces';

  toggleWrap.appendChild(checkbox);
  toggleWrap.appendChild(track);
  toggleWrap.appendChild(toggleLabel);

  // Remove button
  const removeBtn = document.createElement('button');
  removeBtn.className   = 'btn btn-danger';
  removeBtn.textContent = '✕';
  removeBtn.title       = 'Remove this filter';
  removeBtn.onclick     = function () { removeFilterRow(id); };

  footer.appendChild(toggleWrap);
  footer.appendChild(removeBtn);

  row.appendChild(controls);
  row.appendChild(footer);

  document.getElementById('filters-list').appendChild(row);

  updateLogicNote();
  refreshApplyBtn();
}

function removeFilterRow(id) {
  const row = document.getElementById(id);
  if (row) row.remove();

  const list = document.getElementById('filters-list');
  if (list.children.length === 0) {
    const emptyMsg = document.createElement('div');
    emptyMsg.className = 'empty-state';
    emptyMsg.id        = 'no-filters-msg';
    emptyMsg.innerHTML = 'No filters added yet. Click <strong>+ Add Filter</strong> below.';
    list.appendChild(emptyMsg);
  }

  updateLogicNote();
  refreshApplyBtn();
}

function clearAllFilters() {
  const list = document.getElementById('filters-list');
  list.innerHTML = '';
  filterCount = 0;

  const emptyMsg = document.createElement('div');
  emptyMsg.className = 'empty-state';
  emptyMsg.id        = 'no-filters-msg';
  emptyMsg.innerHTML = 'No filters added yet. Click <strong>+ Add Filter</strong> below.';
  list.appendChild(emptyMsg);

  updateLogicNote();
  refreshApplyBtn();
}

function getFilters() {
  const rows = document.querySelectorAll('.filter-row');
  const filters = [];
  rows.forEach(function (row) {
    const select   = row.querySelector('select');
    const input    = row.querySelector('input[type="text"]');
    const checkbox = row.querySelector('input[type="checkbox"]');
    const colIdx   = select ? parseInt(select.value, 10) : NaN;
    const value    = input  ? input.value : '';          // keep raw — ignore logic applied at match time
    const ignore   = checkbox ? checkbox.checked : true; // default ON
    if (!isNaN(colIdx) && select.value !== '') {
      filters.push({ colIdx: colIdx, value: value, ignore: ignore });
    }
  });
  return filters;
}

function refreshApplyBtn() {
  const btn = document.getElementById('apply-btn');
  // Enable as soon as at least one filter row has a column selected
  const rows = document.querySelectorAll('.filter-row');
  let hasColumn = false;
  rows.forEach(function (row) {
    const select = row.querySelector('select');
    if (select && select.value !== '') hasColumn = true;
  });
  btn.disabled = !hasColumn;
}

function updateLogicNote() {
  const note = document.getElementById('logic-note');
  if (!note) return;

  const filters = getFilters();
  if (filters.length === 0) {
    note.textContent = '';
    return;
  }

  // Work out same-column groups
  const colCounts = {};
  filters.forEach(function (f) {
    colCounts[f.colIdx] = (colCounts[f.colIdx] || 0) + 1;
  });

  const hasMultiValue = Object.values(colCounts).some(function (c) { return c > 1; });
  const hasMultiCol   = Object.keys(colCounts).length > 1;

  if (hasMultiValue && hasMultiCol) {
    note.innerHTML = 'Rows must match <span>all</span> column conditions. Multiple values for the same column use <span>OR</span> logic.';
  } else if (hasMultiValue) {
    note.innerHTML = 'Multiple values on the same column use <span>OR</span> logic — rows matching any value are included.';
  } else if (hasMultiCol) {
    note.innerHTML = 'All column conditions must be met — filters across different columns use <span>AND</span> logic.';
  } else {
    note.innerHTML = 'Rows matching the filter value will be copied to a new sheet.';
  }
}

/* ── Step 3 : Apply filters & create sheet ────────────── */
async function applyFilters() {
  clearStatus();
  const filters = getFilters();

  if (filters.length === 0) {
    showStatus('Please add at least one filter with a column and value selected.', 'error');
    return;
  }

  const btn = document.getElementById('apply-btn');
  btn.disabled = true;
  btn.innerHTML = '<span class="loader"></span> Filtering…';

  try {
    // Group filters by colIdx.
    // Each group holds an array of { target, ignore } entries.
    // Same column = OR across entries; different columns = AND.
    const grouped = {};
    filters.forEach(function (f) {
      const key = String(f.colIdx);
      if (!grouped[key]) grouped[key] = [];

      const target = f.ignore
        ? f.value.replace(/\s+/g, '').toLowerCase()  // normalise target
        : f.value;                                     // exact target

      grouped[key].push({ target: target, ignore: f.ignore });
    });

    // Filter the data rows
    const matchedRows = capturedData.filter(function (row) {
      // Every column group must match (AND)
      return Object.keys(grouped).every(function (colIdxStr) {
        const colIdx  = parseInt(colIdxStr, 10);
        const entries = grouped[colIdxStr];
        const rawCell = row[colIdx] !== null && row[colIdx] !== undefined
          ? String(row[colIdx])
          : '';

        // Any entry in the group must match (OR within column)
        return entries.some(function (e) {
          const cellVal = e.ignore
            ? rawCell.replace(/\s+/g, '').toLowerCase()
            : rawCell;
          return cellVal === e.target;
        });
      });
    });

    if (matchedRows.length === 0) {
      showStatus('No rows matched your filters. Try different column values and try again.', 'info');
      restoreApplyBtn();
      return;
    }

    // Write to Excel
    await Excel.run(async (context) => {
      const sheets    = context.workbook.worksheets;
      const timestamp = new Date();
      const sheetName = 'Filtered_' +
        padTwo(timestamp.getHours()) + padTwo(timestamp.getMinutes()) + padTwo(timestamp.getSeconds());

      const newSheet = sheets.add(sheetName);
      newSheet.activate();

      // Write header + data
      // safeCell: converts any value Excel might return into something
      // the API can write back (numbers, strings, booleans, empty string).
      // Error-value objects (e.g. { error: '#VALUE!' }) become empty string.
      function safeCell(cell) {
        if (cell === null || cell === undefined) return '';
        if (typeof cell === 'object') return '';   // Excel error objects
        if (typeof cell === 'number' || typeof cell === 'boolean') return cell;
        return String(cell);
      }

      const outputData = [capturedHeaders.map(String)].concat(
        matchedRows.map(function (row) {
          return row.map(safeCell);
        })
      );

      const lastCol   = colLetter(capturedHeaders.length);
      const totalRows = outputData.length;

      const writeRange = newSheet.getRange('A1:' + lastCol + totalRows);
      writeRange.values = outputData;

      // Style the header row
      const headerRange = newSheet.getRange('A1:' + lastCol + '1');
      headerRange.format.fill.color = '#217346';
      headerRange.format.font.color = '#FFFFFF';
      headerRange.format.font.bold  = true;
      headerRange.format.font.size  = 11;

      // Auto-fit columns
      writeRange.format.autofitColumns();

      await context.sync();

      showStatus(
        '✓ Done! Sheet <strong>"' + sheetName + '"</strong> created with <strong>' +
        matchedRows.length + '</strong> matching row' + (matchedRows.length !== 1 ? 's' : '') + '.',
        'success'
      );
    });
  } catch (err) {
    showStatus('Something went wrong: ' + (err.message || err), 'error');
  } finally {
    restoreApplyBtn();
  }
}

function restoreApplyBtn() {
  const btn = document.getElementById('apply-btn');
  btn.innerHTML = '✦ Apply Filters &amp; Create New Sheet';
  btn.disabled  = false;
}

/* ── Helpers ───────────────────────────────────────────── */

function unlockFiltersSection() {
  const sec = document.getElementById('filters-section');
  sec.style.opacity        = '1';
  sec.style.pointerEvents  = 'auto';
}

function lockFiltersSection() {
  const sec = document.getElementById('filters-section');
  sec.style.opacity        = '0.45';
  sec.style.pointerEvents  = 'none';
}

function unlockApplySection() {
  const sec = document.getElementById('apply-section');
  sec.style.opacity        = '1';
  sec.style.pointerEvents  = 'auto';
}

function lockApplySection() {
  const sec = document.getElementById('apply-section');
  sec.style.opacity        = '0.45';
  sec.style.pointerEvents  = 'none';
}

function showStatus(message, type) {
  const box = document.getElementById('status-box');
  box.className  = type;
  box.innerHTML  = message;
}

function clearStatus() {
  const box = document.getElementById('status-box');
  box.className = '';
  box.textContent = '';
  box.style.display = '';
}

/** Convert column index (0-based) to Excel letter(s): 0 → A, 25 → Z, 26 → AA */
function colLetter(count) {
  // count is number of columns; return letter for last column (count-1 index)
  let n = count; // 1-based
  let result = '';
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

function padTwo(n) {
  return n < 10 ? '0' + n : String(n);
}

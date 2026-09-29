(function () {
  /* ── State ──────────────────────────────────────── */
  let headers = [];
  let allRows = [];
  let filtered = [];
  let sortCol = -1, sortDir = 1;
  let fileName = '';

  /* ── Header normalization ───────────────────────── */
  // Fixes common typos that Poshaya / Tally POS exports contain
  const HEADER_FIX = {
    'caterogry':         'Category',
    'sub caterogry':     'Sub Category',
    'cateogry':          'Category',
    'sub cateogry':      'Sub Category',
    'catergory':         'Category',
    'sub catergory':     'Sub Category',
    'catagory':          'Category',
    'sub catagory':      'Sub Category',
    'categroy':          'Category',
    'sub categroy':      'Sub Category',
  };

  function normalizeHeader(raw) {
    if (raw === null || raw === undefined) return '';
    const s = String(raw).trim();
    return HEADER_FIX[s.toLowerCase()] || s;
  }

  /* ── Numeric column detection ───────────────────── */
  function isNumericCol(colIdx) {
    const sample = allRows.slice(0, 20).map(r => r[colIdx])
      .filter(v => v !== null && v !== undefined && v !== '');
    if (sample.length === 0) return false;
    return sample.filter(v =>
      typeof v === 'number' || (!isNaN(Number(v)) && String(v).trim() !== '')
    ).length > sample.length * 0.6;
  }
  isNumericCol.cache = {};

  /* ── GST chip helper ────────────────────────────── */
  function chipify(val, colIdx) {
    const h = (headers[colIdx] || '').toLowerCase();
    if (h === 'voucher type') return `<span class="chip chip-sales">${val}</span>`;
    if (h === 'item tax') {
      const cls = String(val).toUpperCase().includes('18') ? 'chip-gst18' : 'chip-gst5';
      return `<span class="chip ${cls}">${val}</span>`;
    }
    return null;
  }

  function fmtCell(val, colIdx) {
    if (val === null || val === undefined || val === '') return '<span style="color:#D1D5DB">—</span>';
    const chip = chipify(val, colIdx);
    if (chip) return chip;
    if (typeof val === 'number') {
      const h = (headers[colIdx] || '').toLowerCase();
      if (h.includes('%'))  return val.toLocaleString('en-IN', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
      if (h.includes('qty')) return val.toLocaleString('en-IN');
      return val.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return String(val);
  }

  /* ── Column index helpers ───────────────────────── */
  function colIdx(pattern) {
    return headers.findIndex(h => pattern.test(h));
  }

  /* ── Parse sheet ────────────────────────────────── */
  function parseSheet(wb) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    if (!raw || raw.length < 2) throw new Error('Sheet appears empty or has no data rows.');

    // find first non-empty row → header row
    let hdrIdx = 0;
    while (hdrIdx < raw.length && raw[hdrIdx].every(c => c === null)) hdrIdx++;

    // normalize headers (fix typos like "Caterogry")
    headers = raw[hdrIdx].map(normalizeHeader);

    const dataRows = [];
    for (let r = hdrIdx + 1; r < raw.length; r++) {
      const row = raw[r];
      if (row.every(c => c === null || c === '')) continue; // blank row
      dataRows.push(headers.map((_, c) => row[c] !== undefined ? row[c] : null));
    }
    return dataRows;
  }

  /* ── Split data rows from subtotal / grand-total rows ── */
  // Tally/POS exports insert sub-category subtotal rows between sale rows.
  // Example: ["Dress Material :", null, null, ..., 5, 1234.00, ...]
  // A genuine sale row always has a non-empty Voucher No.
  function splitTotals(rows) {
    const vcIdx = colIdx(/^voucher\s*no$/i);
    const dataRows = [];
    let totalRow = null;

    rows.forEach(row => {
      const firstVal = String(row.find(c => c !== null && c !== '') ?? '');

      // Grand total row — capture and skip
      if (/grand\s*total|grant\s*total/i.test(firstVal)) {
        totalRow = row;
        return;
      }

      // Sub-category or category subtotal rows from Tally:
      // they have no Voucher No and their first cell ends with ":"
      // or contains "total" or matches a known grouping pattern
      if (vcIdx >= 0) {
        const vcVal = row[vcIdx];
        if ((vcVal === null || vcVal === '') && /:\s*$|total/i.test(firstVal)) {
          return; // skip subtotal row
        }
      } else {
        // no Voucher No column — fallback: skip if first cell ends with ":"
        if (/:\s*$/.test(firstVal)) return;
      }

      dataRows.push(row);
    });

    return { dataRows, totalRow };
  }

  /* ── Summary cards ──────────────────────────────── */
  function buildSummary(dataRows, totalRow) {
    const grid = document.getElementById('summaryGrid');
    grid.innerHTML = '';

    addCard(grid, 'Items', dataRows.length, '');

    const numericCols = headers.map((_, i) => i).filter(i => isNumericCol.cache[i]);
    const PRIORITY = ['net amount', 'taxable value', 'total discount', 'bill level disc'];

    // prefer named priority columns; fall back to first numeric cols
    const chosen = [];
    PRIORITY.forEach(p => {
      const ci = headers.findIndex(h => h.toLowerCase().includes(p));
      if (ci >= 0 && !chosen.includes(ci)) chosen.push(ci);
    });
    numericCols.forEach(ci => { if (!chosen.includes(ci)) chosen.push(ci); });

    chosen.slice(0, 4).forEach(ci => {
      const h = headers[ci];
      const source = totalRow && (totalRow[ci] !== null && totalRow[ci] !== undefined)
        ? totalRow[ci]
        : dataRows.reduce((acc, r) => acc + (typeof r[ci] === 'number' ? r[ci] : 0), 0);
      if (!source && source !== 0) return;
      const formatted = typeof source === 'number'
        ? '₹ ' + source.toLocaleString('en-IN', { minimumFractionDigits: 2 })
        : source;
      const cls = h.toLowerCase().includes('discount') ? 'red'
                : (h.toLowerCase().includes('amount') || h.toLowerCase().includes('net')) ? 'green'
                : '';
      addCard(grid, h, formatted, cls);
    });
  }

  function addCard(grid, label, value, cls) {
    grid.insertAdjacentHTML('beforeend',
      `<div class="card"><div class="card-label">${label}</div><div class="card-value ${cls}">${value}</div></div>`
    );
  }

  /* ── Table head ─────────────────────────────────── */
  function buildHead() {
    const thead = document.getElementById('tableHead');
    thead.innerHTML = '';
    const tr = document.createElement('tr');
    headers.forEach((h, i) => {
      const th = document.createElement('th');
      th.textContent = h || `Col ${i + 1}`;
      th.dataset.col = i;
      th.addEventListener('click', () => sortBy(i));
      tr.appendChild(th);
    });
    thead.appendChild(tr);
  }

  /* ── Table foot ─────────────────────────────────── */
  let totalRow = null;
  function buildFoot() {
    const tfoot = document.getElementById('tableFoot');
    tfoot.innerHTML = '';
    if (!totalRow) return;
    const tr = document.createElement('tr');
    headers.forEach((h, i) => {
      const td = document.createElement('td');
      if (i === 0) {
        td.textContent = 'Grand Total';
      } else {
        const v = totalRow[i];
        if (typeof v === 'number') {
          td.className = 'num';
          td.textContent = v.toLocaleString('en-IN', { minimumFractionDigits: 2 });
        }
      }
      tr.appendChild(td);
    });
    tfoot.appendChild(tr);
  }

  /* ── Filters ────────────────────────────────────── */
  const FILTER_DEFS = [
    { id: 'partyFilter',       pattern: /^party\s*name$/i,        label: 'All Parties' },
    { id: 'voucherFilter',     pattern: /^voucher\s*no$/i,         label: 'All Vouchers' },
    { id: 'categoryFilter',    pattern: /^category$/i,             label: 'All Categories' },
    { id: 'subCategoryFilter', pattern: /^sub\s*category$/i,       label: 'All Sub-Categories' },
  ];

  function populateFilters(rows) {
    FILTER_DEFS.forEach(({ id, pattern, label }) => {
      const sel = document.getElementById(id);
      sel.innerHTML = `<option value="">${label}</option>`;
      const ci = headers.findIndex(h => pattern.test(h));
      sel.dataset.colIdx = ci;
      if (ci < 0) return;
      [...new Set(rows.map(r => r[ci]).filter(v => v !== null && v !== ''))].sort()
        .forEach(v => sel.insertAdjacentHTML('beforeend', `<option value="${v}">${v}</option>`));
    });
  }

  /* ── Render rows ────────────────────────────────── */
  function render() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';

    if (filtered.length === 0) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = headers.length;
      td.className = 'empty-cell';
      td.textContent = 'No records match your filters.';
      tr.appendChild(td);
      tbody.appendChild(tr);
    } else {
      filtered.forEach(row => {
        const tr = document.createElement('tr');
        row.forEach((cell, i) => {
          const td = document.createElement('td');
          if (typeof cell === 'number' && isNumericCol.cache[i]) td.className = 'num';
          td.innerHTML = fmtCell(cell, i);
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
    }

    document.getElementById('visibleCount').textContent =
      `Showing ${filtered.length} of ${allRows.length} records`;
  }

  /* ── Sort ───────────────────────────────────────── */
  function sortBy(col) {
    const ths = document.querySelectorAll('#tableHead th');
    ths.forEach(th => th.classList.remove('sort-asc', 'sort-desc'));
    sortDir = (sortCol === col) ? sortDir * -1 : 1;
    sortCol = col;
    ths[col].classList.add(sortDir === 1 ? 'sort-asc' : 'sort-desc');
    applySort();
    render();
  }
  function applySort() {
    if (sortCol < 0) return;
    filtered.sort((a, b) => {
      const av = a[sortCol] ?? '', bv = b[sortCol] ?? '';
      if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * sortDir;
      return String(av).localeCompare(String(bv)) * sortDir;
    });
  }

  /* ── Filter + Search ────────────────────────────── */
  function applyFilters() {
    const q = document.getElementById('searchInput').value.toLowerCase();

    filtered = allRows.filter(row => {
      for (const { id } of FILTER_DEFS) {
        const sel = document.getElementById(id);
        const val = sel.value;
        const ci  = parseInt(sel.dataset.colIdx ?? -1);
        if (val && ci >= 0 && row[ci] !== val) return false;
      }
      if (q) return row.some(c => c !== null && String(c).toLowerCase().includes(q));
      return true;
    });

    applySort();
    render();
  }

  /* ── Download Excel ─────────────────────────────── */
  function downloadExcel() {
    const wb = XLSX.utils.book_new();
    const sheetData = [headers, ...filtered];
    if (totalRow) sheetData.push(totalRow);

    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    ws['!cols'] = headers.map((h, i) => {
      const maxLen = Math.max(
        String(h).length,
        ...filtered.slice(0, 50).map(r => String(r[i] ?? '').length)
      );
      return { wch: Math.min(Math.max(maxLen + 2, 10), 40) };
    });

    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c });
      if (ws[addr]) ws[addr].s = { font: { bold: true } };
    }

    XLSX.utils.book_append_sheet(wb, ws, 'Sale Report');
    XLSX.writeFile(wb, (fileName.replace(/\.[^.]+$/, '') || 'sale_report') + '_export.xlsx');
  }

  /* ── Load file ──────────────────────────────────── */
  function loadFile(file) {
    if (!file) return;
    fileName = file.name;
    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array', cellDates: false });
        const rawRows = parseSheet(wb);
        const split  = splitTotals(rawRows);
        allRows  = split.dataRows;
        totalRow = split.totalRow;

        isNumericCol.cache = {};
        headers.forEach((_, i) => { isNumericCol.cache[i] = isNumericCol(i); });

        filtered = [...allRows];

        document.getElementById('reportMeta').textContent =
          `${file.name}  ·  ${allRows.length} rows · ${headers.length} columns`;
        document.getElementById('rowCount').textContent = `${allRows.length} Items`;

        buildHead();
        buildFoot();
        populateFilters(allRows);
        buildSummary(allRows, totalRow);

        // reset toolbar
        document.getElementById('searchInput').value = '';
        FILTER_DEFS.forEach(({ id }) => { document.getElementById(id).value = ''; });

        render();

        document.getElementById('importScreen').classList.add('hidden');
        document.getElementById('reportScreen').classList.remove('hidden');
      } catch (err) {
        showImportError('Could not read file: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function showImportError(msg) {
    const el = document.getElementById('importError');
    el.textContent = msg;
    el.classList.remove('hidden');
  }

  /* ── Wire up import screen ──────────────────────── */
  const fileInput = document.getElementById('fileInput');
  const dropZone  = document.getElementById('dropZone');

  fileInput.addEventListener('change', () => { if (fileInput.files[0]) loadFile(fileInput.files[0]); });
  dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault(); dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) loadFile(e.dataTransfer.files[0]);
  });

  /* ── Wire up report screen ──────────────────────── */
  document.getElementById('btnDownload').addEventListener('click', downloadExcel);
  document.getElementById('btnImportNew').addEventListener('click', () => {
    document.getElementById('reportScreen').classList.add('hidden');
    document.getElementById('importScreen').classList.remove('hidden');
    document.getElementById('importError').classList.add('hidden');
    fileInput.value = '';
  });

  document.getElementById('searchInput').addEventListener('input', applyFilters);
  FILTER_DEFS.forEach(({ id }) => {
    document.getElementById(id).addEventListener('change', applyFilters);
  });
})();

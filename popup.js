(function () {
  /* ── State ──────────────────────────────────────── */
  let headers = [];
  let allRows = [];       // raw data rows (arrays)
  let filtered = [];
  let sortCol = -1, sortDir = 1;
  let fileName = '';

  /* ── Numeric column detection ───────────────────── */
  function isNumericCol(colIdx) {
    const sample = allRows.slice(0, 20).map(r => r[colIdx]).filter(v => v !== null && v !== undefined && v !== '');
    if (sample.length === 0) return false;
    return sample.filter(v => typeof v === 'number' || (!isNaN(Number(v)) && String(v).trim() !== '')).length > sample.length * 0.6;
  }

  /* ── GST chip helper ────────────────────────────── */
  function chipify(val, colIdx, hdr) {
    const h = (hdr || '').toLowerCase();
    if (h === 'voucher type') return `<span class="chip chip-sales">${val}</span>`;
    if (h === 'item tax') {
      const s = String(val).toUpperCase();
      const cls = s.includes('18') ? 'chip-gst18' : 'chip-gst5';
      return `<span class="chip ${cls}">${val}</span>`;
    }
    return null;
  }

  function fmtCell(val, colIdx) {
    if (val === null || val === undefined || val === '') return '<span style="color:#D1D5DB">—</span>';
    const chip = chipify(val, colIdx, headers[colIdx]);
    if (chip) return chip;
    if (typeof val === 'number') {
      const h = (headers[colIdx] || '').toLowerCase();
      if (h.includes('%')) return val.toLocaleString('en-IN', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
      if (h.includes('qty')) return val.toLocaleString('en-IN');
      return val.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return String(val);
  }

  /* ── Parse sheet ────────────────────────────────── */
  function parseSheet(wb) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    if (!raw || raw.length < 2) throw new Error('Sheet appears empty or has no data rows.');

    // find header row (first non-empty row)
    let hdrIdx = 0;
    while (hdrIdx < raw.length && raw[hdrIdx].every(c => c === null)) hdrIdx++;

    headers = raw[hdrIdx].map(h => h === null ? '' : String(h).trim());

    const dataRows = [];
    for (let r = hdrIdx + 1; r < raw.length; r++) {
      const row = raw[r];
      // skip rows that are completely null
      if (row.every(c => c === null || c === '')) continue;
      dataRows.push(headers.map((_, c) => row[c] !== undefined ? row[c] : null));
    }
    return dataRows;
  }

  /* ── Detect totals row ──────────────────────────── */
  function splitTotals(rows) {
    // heuristic: last row whose first non-null cell contains "total" or "grand"
    const last = rows[rows.length - 1];
    if (!last) return { dataRows: rows, totalRow: null };
    const firstVal = last.find(c => c !== null && c !== '');
    if (firstVal && /total|grand/i.test(String(firstVal))) {
      return { dataRows: rows.slice(0, -1), totalRow: last };
    }
    return { dataRows: rows, totalRow: null };
  }

  /* ── Summary cards ──────────────────────────────── */
  function buildSummary(dataRows, totalRow) {
    const grid = document.getElementById('summaryGrid');
    grid.innerHTML = '';

    const numericCols = headers.map((_, i) => i).filter(i => isNumericCol.cache[i]);

    // unique items
    addCard(grid, 'Items', dataRows.length, '');

    // from totals row if available
    if (totalRow) {
      numericCols.slice(0, 4).forEach(ci => {
        const v = totalRow[ci];
        if (v !== null && v !== undefined && v !== '') {
          const h = headers[ci];
          const formatted = typeof v === 'number'
            ? '₹ ' + v.toLocaleString('en-IN', { minimumFractionDigits: 2 })
            : v;
          const cls = h.toLowerCase().includes('discount') ? 'red' :
                      h.toLowerCase().includes('amount') || h.toLowerCase().includes('net') ? 'green' : '';
          addCard(grid, h, formatted, cls);
        }
      });
    } else {
      // auto-sum first few numeric cols
      numericCols.slice(0, 4).forEach(ci => {
        const h = headers[ci];
        const sum = dataRows.reduce((acc, r) => acc + (typeof r[ci] === 'number' ? r[ci] : 0), 0);
        const formatted = '₹ ' + sum.toLocaleString('en-IN', { minimumFractionDigits: 2 });
        const cls = h.toLowerCase().includes('discount') ? 'red' :
                    h.toLowerCase().includes('amount') || h.toLowerCase().includes('net') ? 'green' : '';
        addCard(grid, h, formatted, cls);
      });
    }
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

  /* ── Table foot (totals) ────────────────────────── */
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
  function populateFilters(rows) {
    const pf = document.getElementById('partyFilter');
    const vf = document.getElementById('voucherFilter');
    pf.innerHTML = '<option value="">All Parties</option>';
    vf.innerHTML = '<option value="">All Vouchers</option>';

    // detect col indices by header name
    const partyIdx   = headers.findIndex(h => /party|customer|name/i.test(h));
    const voucherIdx = headers.findIndex(h => /voucher no|bill no|invoice/i.test(h));

    if (partyIdx >= 0) {
      [...new Set(rows.map(r => r[partyIdx]).filter(Boolean))].sort()
        .forEach(p => pf.insertAdjacentHTML('beforeend', `<option value="${p}">${p}</option>`));
    }
    if (voucherIdx >= 0) {
      [...new Set(rows.map(r => r[voucherIdx]).filter(Boolean))].sort()
        .forEach(v => vf.insertAdjacentHTML('beforeend', `<option value="${v}">${v}</option>`));
    }

    pf.dataset.colIdx = partyIdx;
    vf.dataset.colIdx = voucherIdx;
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
    if (sortCol === col) { sortDir *= -1; } else { sortCol = col; sortDir = 1; }
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
    const q    = document.getElementById('searchInput').value.toLowerCase();
    const pf   = document.getElementById('partyFilter');
    const vf   = document.getElementById('voucherFilter');
    const pi   = parseInt(pf.dataset.colIdx ?? -1);
    const vi   = parseInt(vf.dataset.colIdx ?? -1);
    const party   = pf.value;
    const voucher = vf.value;

    filtered = allRows.filter(row => {
      if (party   && pi >= 0 && row[pi]  !== party)   return false;
      if (voucher && vi >= 0 && row[vi]  !== voucher)  return false;
      if (q) return row.some(c => c !== null && String(c).toLowerCase().includes(q));
      return true;
    });
    applySort();
    render();
  }

  /* ── Download Excel ─────────────────────────────── */
  function downloadExcel() {
    const wb = XLSX.utils.book_new();

    // build rows: header + filtered data + optional totals
    const sheetData = [headers, ...filtered];
    if (totalRow) sheetData.push(totalRow);

    const ws = XLSX.utils.aoa_to_sheet(sheetData);

    // column widths
    ws['!cols'] = headers.map((h, i) => {
      const maxLen = Math.max(
        String(h).length,
        ...filtered.slice(0, 50).map(r => String(r[i] ?? '').length)
      );
      return { wch: Math.min(Math.max(maxLen + 2, 10), 40) };
    });

    // style header row bold
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c });
      if (!ws[addr]) continue;
      ws[addr].s = { font: { bold: true }, fill: { fgColor: { rgb: '7C3AED' } } };
    }

    XLSX.utils.book_append_sheet(wb, ws, 'Sale Report');
    const outName = fileName.replace(/\.[^.]+$/, '') + '_export.xlsx';
    XLSX.writeFile(wb, outName || 'sale_report_export.xlsx');
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
        const split   = splitTotals(rawRows);
        allRows  = split.dataRows;
        totalRow = split.totalRow;

        // cache numeric column flags
        isNumericCol.cache = {};
        headers.forEach((_, i) => { isNumericCol.cache[i] = isNumericCol(i); });

        filtered = [...allRows];

        // update UI
        document.getElementById('reportMeta').textContent =
          `${file.name}  ·  ${allRows.length} rows · ${headers.length} columns`;
        document.getElementById('rowCount').textContent = `${allRows.length} Items`;

        buildHead();
        buildFoot();
        populateFilters(allRows);
        buildSummary(allRows, totalRow);

        // reset toolbar
        document.getElementById('searchInput').value = '';
        document.getElementById('partyFilter').value = '';
        document.getElementById('voucherFilter').value = '';

        render();

        // show report screen
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

  dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault(); dropZone.classList.remove('drag-over');
    const f = e.dataTransfer.files[0];
    if (f) loadFile(f);
  });

  /* ── Wire up report screen controls ────────────── */
  document.getElementById('btnDownload').addEventListener('click', downloadExcel);
  document.getElementById('btnImportNew').addEventListener('click', () => {
    document.getElementById('reportScreen').classList.add('hidden');
    document.getElementById('importScreen').classList.remove('hidden');
    document.getElementById('importError').classList.add('hidden');
    fileInput.value = '';
  });

  document.getElementById('searchInput').addEventListener('input', applyFilters);
  document.getElementById('partyFilter').addEventListener('change', applyFilters);
  document.getElementById('voucherFilter').addEventListener('change', applyFilters);
})();

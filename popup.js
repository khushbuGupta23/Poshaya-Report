(function () {
  const { headers, rows, totals, reportDate } = SALE_REPORT;

  // numeric column indices (0-based)
  const NUM_COLS = new Set([12,13,14,15,16,17,18,19,21,23,24,25]);

  // columns that get a colour chip
  function chipify(val, colIdx) {
    if (colIdx === 0) return `<span class="chip chip-sales">${val}</span>`;
    if (colIdx === 20) {
      const cls = String(val).toUpperCase().includes('18') ? 'chip-gst18' : 'chip-gst5';
      return `<span class="chip ${cls}">${val}</span>`;
    }
    return null;
  }

  function fmt(val, colIdx) {
    if (val === null || val === undefined || val === '—') return '<span style="color:#D1D5DB">—</span>';
    const chip = chipify(val, colIdx);
    if (chip) return chip;
    if (NUM_COLS.has(colIdx) && typeof val === 'number') {
      return val === 0 ? '0.00' : val.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return String(val);
  }

  // ── Summary cards ──────────────────────────────────
  document.getElementById('reportDate').textContent = `Date: ${reportDate}`;
  document.getElementById('rowCount').textContent = `${rows.length} Items`;

  const summaryDefs = [
    { label: 'Total Items',    value: rows.length,            cls: '' },
    { label: 'Sales Qty',      value: totals['Sales Qty'],    cls: '' },
    { label: 'Net Amount',     value: '₹ ' + totals['Net Amount'].toLocaleString('en-IN'), cls: 'green' },
    { label: 'Total Discount', value: '₹ ' + totals['Total Discount'].toLocaleString('en-IN'), cls: 'red' },
    { label: 'Taxable Value',  value: '₹ ' + totals['Taxable Value'].toLocaleString('en-IN'), cls: '' },
  ];
  const grid = document.getElementById('summaryGrid');
  summaryDefs.forEach(({ label, value, cls }) => {
    grid.insertAdjacentHTML('beforeend',
      `<div class="card"><div class="card-label">${label}</div><div class="card-value ${cls}">${value}</div></div>`
    );
  });

  // ── Table head ─────────────────────────────────────
  const thead = document.getElementById('tableHead');
  const trHead = document.createElement('tr');
  headers.forEach((h, i) => {
    const th = document.createElement('th');
    th.textContent = h;
    th.dataset.col = i;
    th.addEventListener('click', () => sortBy(i));
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);

  // ── Populate filters ───────────────────────────────
  const partyFilter   = document.getElementById('partyFilter');
  const voucherFilter = document.getElementById('voucherFilter');
  const parties  = [...new Set(rows.map(r => r[3]).filter(Boolean))].sort();
  const vouchers = [...new Set(rows.map(r => r[1]).filter(Boolean))].sort();
  parties.forEach(p  => partyFilter.insertAdjacentHTML('beforeend',  `<option value="${p}">${p}</option>`));
  vouchers.forEach(v => voucherFilter.insertAdjacentHTML('beforeend', `<option value="${v}">${v}</option>`));

  // ── Render ─────────────────────────────────────────
  let sortCol = -1, sortDir = 1;
  let filtered = [...rows];

  function render() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';

    if (filtered.length === 0) {
      tbody.insertAdjacentHTML('beforeend',
        `<tr><td colspan="${headers.length}" class="empty">
          <svg width="40" height="40" viewBox="0 0 24 24" fill="none"><path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2" stroke="#6B7280" stroke-width="1.5"/><rect x="9" y="3" width="6" height="4" rx="1" stroke="#6B7280" stroke-width="1.5"/></svg>
          No records match your filters
        </td></tr>`
      );
    } else {
      filtered.forEach(row => {
        const tr = document.createElement('tr');
        row.forEach((cell, i) => {
          const td = document.createElement('td');
          if (NUM_COLS.has(i)) td.className = 'num';
          td.innerHTML = fmt(cell, i);
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
    }

    document.getElementById('visibleCount').textContent =
      `Showing ${filtered.length} of ${rows.length} records`;
  }

  // ── Totals footer ─────────────────────────────────
  const tfoot = document.getElementById('tableFoot');
  const trFoot = document.createElement('tr');
  headers.forEach((h, i) => {
    const td = document.createElement('td');
    if (i === 0) { td.textContent = 'Grand Total'; }
    else if (totals[h] !== undefined) {
      td.className = 'num';
      td.textContent = totals[h].toLocaleString('en-IN', { minimumFractionDigits: 2 });
    }
    trFoot.appendChild(td);
  });
  tfoot.appendChild(trFoot);

  // ── Sort ───────────────────────────────────────────
  function sortBy(col) {
    const ths = thead.querySelectorAll('th');
    ths.forEach(th => th.classList.remove('sort-asc', 'sort-desc'));
    if (sortCol === col) {
      sortDir *= -1;
    } else {
      sortCol = col; sortDir = 1;
    }
    ths[col].classList.add(sortDir === 1 ? 'sort-asc' : 'sort-desc');
    filtered.sort((a, b) => {
      const av = a[col] ?? '', bv = b[col] ?? '';
      if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * sortDir;
      return String(av).localeCompare(String(bv)) * sortDir;
    });
    render();
  }

  // ── Filter + Search ────────────────────────────────
  function applyFilters() {
    const q      = document.getElementById('searchInput').value.toLowerCase();
    const party  = partyFilter.value;
    const vchr   = voucherFilter.value;
    filtered = rows.filter(row => {
      if (party && row[3] !== party)   return false;
      if (vchr  && row[1] !== vchr)    return false;
      if (q) {
        return row.some(cell => cell !== null && String(cell).toLowerCase().includes(q));
      }
      return true;
    });
    if (sortCol >= 0) {
      filtered.sort((a, b) => {
        const av = a[sortCol] ?? '', bv = b[sortCol] ?? '';
        if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * sortDir;
        return String(av).localeCompare(String(bv)) * sortDir;
      });
    }
    render();
  }

  document.getElementById('searchInput').addEventListener('input', applyFilters);
  partyFilter.addEventListener('change', applyFilters);
  voucherFilter.addEventListener('change', applyFilters);

  render();
})();

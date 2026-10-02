/* =====================================================================
   Zoho POS – Sales & Stock Report (Excel format)
   Sections: 1) field mapping  2) xlsx writer  3) data loading
             4) report builders  5) on-screen grid  6) downloads
   ===================================================================== */

// ---------------------------------------------------------------------
// 1) FIELD MAPPING — change these labels if your Zoho custom fields are
//    named differently.
// ---------------------------------------------------------------------
const FIELDS = {
  // invoice custom fields
  referredBy:   ["Referred by", "Referred By", "Reference By", "Referred"],
  telecaller:   ["Tellecaller Name", "Telecaller Name", "Telecaller", "Tele Caller"],
  userName:     ["User Name", "Cashier", "Billed By"],
  discountName: ["Bill Level Discount Name", "Discount Name", "Scheme", "Coupon"],
  mobile:       ["Mobile No.", "Mobile No", "Mobile", "Phone"],
  // item fields (custom fields on the item)
  category:     ["Category", "Category"],
  subCategory:  ["Sub Category", "Sub Category", "SubCategory", "Sub-Category"],
  class3:       ["class3", "Class 3", "Class3"],
  supplier:     ["Supplier", "Vendor"],
  size:         ["Size"]
};
const DEFAULT_ORG_ID = "60082587105"; // used only when nothing is saved yet

// ---------------------------------------------------------------------
// APPAREL CATEGORY MAP
// If the item's category_name (from Zoho Item report) matches any name
// in this list, the Sales/Stock report Category column shows "Apparel"
// and the original category_name moves to the Sub Category column.
// Add or remove entries to match your Zoho category names exactly.
// ---------------------------------------------------------------------
const APPAREL_CATEGORIES = new Set([
  "Suits", "Suit", "Kurtas", "Kurta", "Sarees", "Saree",
  "Lehengas", "Lehnga", "Lehenga", "Dupatta", "Dupattas",
  "Tops", "Top", "Dress", "Dresses", "Gown", "Gowns",
  "Sharara", "Shararas", "Palazzo", "Palazzos",
  "Skirt", "Skirts", "Blouse", "Blouses",
  "Kurtis", "Kurti", "Anarkali", "Salwar", "Salwar Suit", "Salwar Suits",
  "Co-ord Set", "Co-ord Sets", "Co ord Set", "Coord Set",
  "Jacket", "Jackets", "Shrug", "Shrugs",
  "Nightwear", "Loungewear", "Activewear", "Ethnic Wear",
  "Western Wear", "Fusion Wear", "Indo Western",
  "Boys Wear", "Girls Wear", "Kids Wear", "Kids",
  "Men", "Women", "Apparel", "Clothing", "Garment", "Garments",
  "Fabric", "Fabrics", "Stole", "Stoles", "Scarf", "Scarves"
]);
const APPAREL_LOWER = new Set([...APPAREL_CATEGORIES].map((s) => s.toLowerCase()));

// Heading printed on the Daily Statement. Filled in for Poshaya (as on the Quick Bill slip);
// for any other org the name/address/GSTIN are read from Zoho automatically.
const RECEIPT_HEADER = {
  [DEFAULT_ORG_ID]: {
    name: "VEDSAATVIK ENTERPRISES PVT. LTD. (GN)",
    lines: ["Shop No.124/B/142, Chawla Market Govind Nagar", "Kanpur - 208006", "", "208005"],
    gstin: "09AAKCV5581E1ZX"
  }
};

// ---------------------------------------------------------------------
// 2) XLSX WRITER (no libraries) — reproduces the template styling:
//    Arial 10, blue header (D9E7F9) 30pt high, thin borders, 0.00 numbers,
//    dd-mm-yyyy dates, bold Arial 9 totals.
// ---------------------------------------------------------------------
const XlsxWriter = (() => {
  const enc = new TextEncoder();
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = (b) => { let c = 0xffffffff; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };

  function zip(files) {
    const parts = [], central = []; let offset = 0;
    for (const f of files) {
      const name = enc.encode(f.name), data = enc.encode(f.data), crc = crc32(data), size = data.length;
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(12, 0x21, true);
      h.setUint32(14, crc, true); h.setUint32(18, size, true); h.setUint32(22, size, true); h.setUint16(26, name.length, true);
      parts.push(new Uint8Array(h.buffer), name, data);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(14, 0x21, true);
      c.setUint32(16, crc, true); c.setUint32(20, size, true); c.setUint32(24, size, true); c.setUint16(28, name.length, true); c.setUint32(42, offset, true);
      central.push(new Uint8Array(c.buffer), name);
      offset += 30 + name.length + size;
    }
    const cd = central.reduce((s, a) => s + a.length, 0);
    const e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
    e.setUint32(12, cd, true); e.setUint32(16, offset, true);
    const all = [...parts, ...central, new Uint8Array(e.buffer)];
    const out = new Uint8Array(all.reduce((s, a) => s + a.length, 0)); let p = 0;
    for (const a of all) { out.set(a, p); p += a.length; }
    return out;
  }

  const x = (s) => String(s).replace(/[&<>"]/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[ch])).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");
  const colName = (i) => { let s = ""; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  const serial = (iso) => { const [y, m, d] = iso.split("-").map(Number); return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000; };

  // style ids → see cellXfs below
  const ST = { hdrC: 1, hdrR: 2, txt: 3, num: 4, date: 5, tot: 6, band: 7 };
  const TEXT_STYLES = new Set([ST.hdrC, ST.hdrR, ST.txt, ST.band]);

  function cellXml(ref, v, s) {
    const sAttr = s ? ` s="${s}"` : "";
    if (v === null || v === undefined || v === "") return s ? `<c r="${ref}"${sAttr}/>` : "";
    if (s === ST.date) return /^\d{4}-\d{2}-\d{2}/.test(v) ? `<c r="${ref}"${sAttr}><v>${serial(String(v).slice(0, 10))}</v></c>` : "";
    if (typeof v === "number" && !TEXT_STYLES.has(s)) return isFinite(v) ? `<c r="${ref}"${sAttr}><v>${v}</v></c>` : `<c r="${ref}"${sAttr}/>`;
    return `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${x(v)}</t></is></c>`;
  }

  const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="dd-mm-yyyy"/></numFmts>
<fonts count="3"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
<font><sz val="10"/><color rgb="FF000000"/><name val="Arial"/><family val="2"/></font>
<font><b/><sz val="9"/><color rgb="FF000000"/><name val="Arial"/><family val="2"/></font></fonts>
<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFD9E7F9"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFD4E6FC"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="8">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="49" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="49" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center" wrapText="1"/></xf>
<xf numFmtId="49" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="2" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="164" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="49" fontId="1" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

  // sheet = { name, widths:[..], rows:[{ height, cells:[{v, s}] }], merges:["A2:T2"] }
  function build(sheet) {
    const rowsXml = sheet.rows.map((r, ri) => {
      const rn = ri + 1;
      const ht = r.height ? ` ht="${r.height}" customHeight="1"` : "";
      return `<row r="${rn}"${ht}>${r.cells.map((c, ci) => (c ? cellXml(colName(ci) + rn, c.v, c.s) : "")).join("")}</row>`;
    }).join("");
    const merges = sheet.merges?.length ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>` : "";
    const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<cols>${sheet.widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("")}</cols>
<sheetData>${rowsXml}</sheetData>${merges}
</worksheet>`;
    return zip([
      { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>` },
      { name: "_rels/.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>` },
      { name: "xl/workbook.xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="${x(sheet.name)}" sheetId="1" r:id="rId1"/></sheets>
</workbook>` },
      { name: "xl/_rels/workbook.xml.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>` },
      { name: "xl/styles.xml", data: STYLES },
      { name: "xl/worksheets/sheet1.xml", data: sheetXml }
    ]);
  }
  return { build, colName, ST };
})();
const { colName, ST } = XlsxWriter;

// ---------------------------------------------------------------------
// Column layouts — same order, headings (incl. spelling) and widths as
// your two template files. h = header alignment, t = data style.
// ---------------------------------------------------------------------
const SALES_COLS = [
  { key: "vtype",      label: "Voucher Type",             w: 15,   h: "C", t: "txt" },
  { key: "vno",        label: "Voucher No",               w: 15.6, h: "C", t: "txt" },
  { key: "vdate",      label: "Voucher Date",             w: 14.3, h: "C", t: "date" },
  { key: "party",      label: "Party Name",               w: 28.6, h: "C", t: "txt" },
  { key: "mobile",     label: "Mobile No.",               w: 12.7, h: "C", t: "txt" },
  { key: "referred",   label: "Referred by",              w: 13.1, h: "C", t: "txt" },
  { key: "telecaller", label: "Tellecaller Name",         w: 17.4, h: "C", t: "txt" },
  { key: "sku",        label: "Stock No",                 w: 12.4, h: "C", t: "txt" },
  { key: "desc",       label: "Item Description",         w: 50,   h: "C", t: "txt" },
  { key: "cat",        label: "Category",                w: 12.3, h: "C", t: "txt" },
  { key: "subcat",     label: "Sub Category",            w: 15.6, h: "C", t: "txt" },
  { key: "supplier",   label: "Supplier",                 w: 20.7, h: "C", t: "txt" },
  { key: "qty",        label: "Sales Qty",                w: 14.3, h: "R", t: "num", sum: true },
  { key: "billDisc",   label: "Bill Level Disc. Amt.",    w: 11,   h: "R", t: "num", sum: true },
  { key: "totDisc",    label: "Total Discount",           w: 11,   h: "R", t: "num", sum: true },
  { key: "net",        label: "Net Amount",               w: 11,   h: "R", t: "num", sum: true },
  { key: "retail",     label: "Retail Price",             w: 11,   h: "R", t: "num" },
  { key: "rate",       label: "Item Rate",                w: 11,   h: "R", t: "num" },
  { key: "taxPct",     label: "Tax Percentage",           w: 11,   h: "R", t: "num" },
  { key: "taxAmt",     label: "Tax Amount",               w: 11,   h: "R", t: "num", sum: true },
  { key: "taxable",    label: "Taxable Value",            w: 11,   h: "R", t: "num", sum: true },
  { key: "cgst",       label: "CGST Amount",              w: 11,   h: "R", t: "num", sum: true },
  { key: "sgst",       label: "SGST Amount",              w: 11,   h: "R", t: "num", sum: true },
  { key: "igst",       label: "IGST Amount",              w: 11,   h: "R", t: "num", sum: true },
  { key: "user",       label: "User Name",                w: 17.4, h: "C", t: "txt" },
  { key: "hsn",        label: "HSN Code",                 w: 12.3, h: "C", t: "txt" },
  { key: "discName",   label: "Bill Level Discount Name", w: 28.6, h: "C", t: "txt" }
];

const STOCK_COLS = [
  { key: "_blank",   label: "",                    w: 8.43, h: "",  t: "", screen: false },
  { key: "sku",      label: "Stock No",            w: 12.3, h: "C", t: "txt" },
  { key: "name",     label: "Item Description",    w: 50,   h: "C", t: "txt" },
  { key: "cat",      label: "Category",           w: 12.3, h: "C", t: "txt" },
  { key: "subcat",   label: "Sub Category",       w: 15.6, h: "C", t: "txt" },
  { key: "class3",   label: "class3",              w: 12.3, h: "C", t: "txt" },
  { key: "supplier", label: "Supplier",            w: 43.9, h: "C", t: "txt" },
  { key: "size",     label: "Size",                w: 12.3, h: "C", t: "txt" },
  { key: "taxPct",   label: "Tax Percentage",      w: 14.3, h: "R", t: "num" },
  { key: "retail",   label: "Retail Price",        w: 11,   h: "R", t: "num" },
  { key: "cost",     label: "Cost Price",          w: 11,   h: "R", t: "num" },
  { key: "hsn",      label: "HSN Code",            w: 12.3, h: "C", t: "txt" },
  { key: "costTax",  label: "Cost Price With Tax", w: 14.3, h: "R", t: "num" },
  { key: "openQty",  label: "Opening Qty.",        w: 11,   h: "R", t: "num", sum: true },
  { key: "openVal",  label: "Opening Value",       w: 11,   h: "R", t: "num", sum: true },
  { key: "inQty",    label: "In Qty.",             w: 11,   h: "R", t: "num", sum: true },
  { key: "inVal",    label: "In Value",            w: 11,   h: "R", t: "num", sum: true },
  { key: "outQty",   label: "Out Qty.",            w: 11,   h: "R", t: "num", sum: true },
  { key: "outVal",   label: "Out Value",           w: 11,   h: "R", t: "num", sum: true },
  { key: "closeQty", label: "Closing Qty.",        w: 11,   h: "R", t: "num", sum: true },
  { key: "closeVal", label: "Closing Value",       w: 11,   h: "R", t: "num", sum: true },
  { key: "netAmt",   label: "Net Amount",          w: 13.9, h: "C", t: "txt" }
];

// ---------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const iso = (d) => new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const num = (v) => { const n = parseFloat(String(v ?? "").replace(/[^0-9.\-]/g, "")); return isFinite(n) ? n : 0; };
const norm = (s) => String(s ?? "").toLowerCase().replace(/^cf_/, "").replace(/[^a-z0-9]/g, "");

// read a custom field by any of the given labels (works with label, api_name or cf_ keys)
function cf(obj, labels) {
  if (!obj) return "";
  const wanted = labels.map(norm);
  for (const f of obj.custom_fields || []) {
    if (wanted.includes(norm(f.label)) || wanted.includes(norm(f.api_name)) || wanted.includes(norm(f.placeholder))) {
      const v = f.value_formatted ?? f.value;
      if (v !== undefined && v !== null && v !== "") return String(v);
    }
  }
  for (const [k, v] of Object.entries(obj.custom_field_hash || {})) {
    if (!k.endsWith("_unformatted") && wanted.includes(norm(k)) && v !== "") return String(v);
  }
  for (const [k, v] of Object.entries(obj)) {
    if (k.startsWith("cf_") && !k.endsWith("_unformatted") && wanted.includes(norm(k)) && v !== "" && typeof v !== "object") return String(v);
  }
  return "";
}

function status(msg) { $("statusL").textContent = msg; }
function progress(done, total) { $("bar").style.width = total ? `${Math.round((done / total) * 100)}%` : "0"; }

// ---------------------------------------------------------------------
// 3) DATA LOADING (with local cache of bill/purchase/adjustment details)
// ---------------------------------------------------------------------
const orgId = () => $("orgId").value.trim();

async function api(type, extra) {
  const r = await chrome.runtime.sendMessage({ type, orgId: orgId(), ...extra });
  if (!r.ok) throw new Error(r.error);
  return r.data;
}
const listAll = (path, query = {}, key = path) => api("fetchAll", { path, query, key });

// Get one record's full details; cached until Zoho says it was modified.
async function detail(path, id, stamp, key) {
  const ck = `c:${orgId()}:${path}:${id}`;
  const got = await chrome.storage.local.get(ck);
  if (got[ck] && got[ck].stamp === (stamp || "")) return got[ck].rec;
  const data = await api("fetch", { path: `${path}/${id}`, query: {} });
  const rec = data[key];
  await chrome.storage.local.set({ [ck]: { stamp: stamp || "", rec } });
  return rec;
}

async function detailsFor(label, path, list, idKey, key) {
  // one bulk read of the local cache instead of one storage call per record
  const keys = list.map((r) => `c:${orgId()}:${path}:${r[idKey]}`);
  const cached = keys.length ? await chrome.storage.local.get(keys) : {};
  const out = [];
  let fetched = 0;
  const toWrite = {};
  for (let i = 0; i < list.length; i++) {
    const ck = keys[i], stamp = list[i].last_modified_time || "";
    if (cached[ck] && cached[ck].stamp === stamp) { out.push(cached[ck].rec); continue; }
    fetched++;
    status(`Loading ${label} ${i + 1} of ${list.length} (new: ${fetched})…`);
    progress(i, list.length);
    const data = await api("fetch", { path: `${path}/${list[i][idKey]}`, query: {} });
    const rec = data[key];
    toWrite[ck] = { stamp, rec };
    out.push(rec);
    if (fetched % 20 === 0) { await chrome.storage.local.set(toWrite); for (const k in toWrite) delete toWrite[k]; }
  }
  if (Object.keys(toWrite).length) await chrome.storage.local.set(toWrite);
  progress(0, 0);
  return out;
}

// ---------------------------------------------------------------------
// state
// ---------------------------------------------------------------------
const SHEETS = [
  { name: "Sales Report", cols: SALES_COLS, rows: [], sort: null, filters: {} },
  { name: "Stock Report", cols: STOCK_COLS, rows: [], sort: null, filters: {} }
];
let active = SHEETS[0], view = [], sel = null, selCol = null, loaded = false;
let lastRaw = { invoice: null, item: null };
let stockAll = [];
let summary = null;                  // Daily Statement figures
const SUMMARY_TAB = { name: "All Sales Summary", summary: true };
const TABS = [...SHEETS, SUMMARY_TAB];

// ---------------------------------------------------------------------
// 4) REPORT BUILDERS
// ---------------------------------------------------------------------
function itemInfo(item, lineItem) {
  if (!item) return {};
  const sizeFromAttr = /size/i.test(item.attribute_name1 || "") ? item.attribute_option_name1 : "";

  // Raw category from Item report (category_name is a standard Zoho POS field)
  const rawCat = item.category_name || cf(item, FIELDS.category) ||
                 (lineItem && (lineItem.category_name || lineItem.item_type)) || "";

  // If rawCat is in the Apparel list → Category = "Apparel", Sub Category = rawCat
  // Otherwise → Category = rawCat, Sub Category = item_group_name (size/variant group)
  let cat, subcat;
  const rawTrim = String(rawCat || "").trim();
  if (rawTrim && APPAREL_LOWER.has(rawTrim.toLowerCase())) {
    cat    = "Apparel";
    subcat = rawTrim;
  } else {
    // Not in the Apparel list: keep the Zoho category as Category and repeat it as
    // Sub Category (never the item-group name, which looks like "002130- Kurta Set - XL")
    cat    = rawTrim;
    subcat = cf(item, FIELDS.subCategory) || rawTrim;
  }

  // Supplier → from Manufacturer field (item.manufacturer), not vendor_name which is blank
  const supplier = item.manufacturer || cf(item, FIELDS.supplier) || item.vendor_name || "";

  return {
    sku: item.sku || "",
    name: item.name || "",
    cat,
    subcat,
    class3: cf(item, FIELDS.class3),
    supplier,
    size: sizeFromAttr || cf(item, FIELDS.size),
    taxPct: num(item.tax_percentage),
    retail: num(item.rate),
    cost: num(item.purchase_rate),
    hsn: item.hsn_or_sac || ""
  };
}

// One row per invoice line, with discount and GST split like Quick Bill
function salesRows(invoice, itemsById, mobileOf) {
  const lines = invoice.line_items || [];
  const inclusive = !!invoice.is_inclusive_tax;

  // line-level discount per line
  const lineCalc = lines.map((li) => {
    const qty = num(li.quantity), rate = num(li.rate), gross = qty * rate;
    let lineDisc = li.discount_amount !== undefined && li.discount_amount !== "" ? num(li.discount_amount) : 0;
    if (!lineDisc && li.discount) {
      lineDisc = String(li.discount).includes("%") ? (gross * num(li.discount)) / 100 : num(li.discount);
    }
    return { li, qty, rate, gross, lineDisc, after: gross - lineDisc };
  });

  // bill-level discount, shared across lines in proportion to their value
  const baseSum = lineCalc.reduce((s, c) => s + c.after, 0);
  let billDiscTotal = 0;
  if (invoice.discount && invoice.discount_type !== "item_level") {
    billDiscTotal = String(invoice.discount).includes("%") ? (baseSum * num(invoice.discount)) / 100 : num(invoice.discount);
  }

  return lineCalc.map(({ li, qty, rate, gross, lineDisc, after }) => {
    const info = itemInfo(itemsById[li.item_id], li);
    const billDisc = baseSum ? (billDiscTotal * after) / baseSum : 0;
    const totDisc = lineDisc + billDisc;
    const netBase = gross - totDisc;
    const pct = num(li.tax_percentage);
    const taxable = inclusive ? netBase / (1 + pct / 100) : netBase;
    const tax = inclusive ? netBase - taxable : (taxable * pct) / 100;
    const isIgst = /igst/i.test(`${li.tax_name || ""} ${li.tax_type || ""}`);
    // Use Quick Bill invoice number (cf_autonumber) if available; fall back to Zoho internal number
    const vno = cf(invoice, ["AutoNumber", "Auto Number", "autonumber", "cf_autonumber"]) || invoice.invoice_number;
    return {
      vtype: "Sales",
      vno,
      _invNo: invoice.invoice_number,
      vdate: invoice.date,
      party: invoice.customer_name || "",
      mobile: mobileOf(invoice),
      referred: invoice.salesperson_name || cf(invoice, FIELDS.referredBy) || "",
      telecaller: invoice.cf_telecaller || cf(invoice, FIELDS.telecaller) || (invoice.custom_field_hash || {}).cf_telecaller || "",
      sku: li.sku || info.sku || "",
      desc: li.name || info.name || "",
      cat: info.cat || "",
      subcat: info.subcat || "",
      supplier: info.supplier || "",
      qty: r2(qty),
      billDisc: r2(billDisc),
      totDisc: r2(totDisc),
      net: Math.round(taxable + tax), // rounded to the rupee, as in Quick Bill
      retail: r2(info.retail || rate),
      rate: r2(rate),
      taxPct: r2(pct),
      taxAmt: r2(tax),
      taxable: r2(taxable),
      cgst: isIgst ? 0 : r2(tax / 2),
      sgst: isIgst ? 0 : r2(tax / 2),
      igst: isIgst ? r2(tax) : 0,
      user: invoice.register_name || cf(invoice, FIELDS.userName) || "",
      hsn: li.hsn_or_sac || info.hsn || "",
      discName: cf(invoice, FIELDS.discountName)
    };
  });
}

// Opening / In / Closing per item for [from, to], worked back from current stock
function stockRows(items, invoices, bills, adjustments, from, to) {
  const byName = {};
  items.forEach((it) => (byName[norm(it.name)] = it.item_id));
  const idOf = (li) => li.item_id || byName[norm(li.name)];

  const blank = () => ({ inP: 0, outP: 0, inAfter: 0, outAfter: 0, inValP: 0 });
  const mv = {}; // item_id → quantities in/out during and after the period, and value received
  const add = (id, date, qty, dir, value = 0) => {
    if (!id || !qty || date < from) return;
    const m = (mv[id] ??= blank());
    const k = (dir === "in" ? "in" : "out") + (date <= to ? "P" : "After");
    m[k] += qty;
    if (k === "inP") m.inValP += value;
  };

  invoices.forEach((inv) => (inv.line_items || []).forEach((li) => add(idOf(li), inv.date, num(li.quantity), "out")));
  // In Value = actual purchase value (qty × bill rate), like Quick Bill
  bills.forEach((b) => (b.line_items || []).forEach((li) =>
    add(idOf(li), b.date, num(li.quantity), "in", num(li.quantity) * num(li.rate))));
  adjustments.forEach((a) => {
    if (a.adjustment_type && a.adjustment_type !== "quantity") return;
    (a.line_items || []).forEach((li) => {
      const q = num(li.quantity_adjusted);
      add(idOf(li), a.date, Math.abs(q), q >= 0 ? "in" : "out", q >= 0 ? Math.max(0, num(li.item_total)) : 0);
    });
  });

  return items.map((it) => {
    const info = itemInfo(it);
    const m = mv[it.item_id] || blank();
    const current = num(it.stock_on_hand);
    const closeQty = current - m.inAfter + m.outAfter;
    const openQty = closeQty - m.inP + m.outP;
    return {
      _blank: "",
      ...info,
      taxPct: r2(info.taxPct),
      retail: r2(info.retail),
      cost: r2(info.cost),
      costTax: r2(info.cost * (1 + info.taxPct / 100)),
      openQty: r2(openQty),
      openVal: r2(openQty * info.cost),
      inQty: r2(m.inP),
      inVal: r2(m.inValP),
      outQty: r2(m.outP),
      outVal: r2(m.outP * info.cost),
      closeQty: r2(closeQty),
      closeVal: r2(closeQty * info.cost),
      netAmt: "",
      _moved: m.inP || m.outP
    };
  }).sort((a, b) => String(a.name).localeCompare(String(b.name)));
}

function applyZeroFilter() {
  // zero-stock items are always hidden (no checkbox any more)
  SHEETS[1].rows = stockAll.filter((r) => r.openQty || r.closeQty || r._moved);
}

// ---------------------------------------------------------------------
// RUN
// ---------------------------------------------------------------------
$("run").onclick = async () => {
  if (!orgId()) return status("Enter Organization ID first.");
  const from = $("from").value, to = $("to").value, today = iso(new Date());
  if (!from || !to || from > to) return status("Pick a valid date range.");
  if (to > today) return status("'To' date can't be in the future.");
  await chrome.storage.local.set({ orgId: orgId() });

  $("run").disabled = true;
  const warn = [];
  try {
    // PHASE 1 — everything Sales Report + Summary need. The lists are requested together
    // (the service worker still spaces calls out to respect Zoho's rate limit).
    status("Loading items, invoices and payments…");
    const monthStart = to.slice(0, 8) + "01";
    const [itemsRes, invList, paysRes, mtdRes] = await Promise.all([
      listAll("items"),
      listAll("invoices", { date_start: from, date_end: today }),
      listAll("customerpayments", { date_start: from, date_end: to }),
      monthStart >= from ? Promise.resolve(null) : listAll("invoices", { date_start: monthStart, date_end: to })
    ]);
    const items = itemsRes.rows;
    const itemsById = Object.fromEntries(items.map((i) => [i.item_id, i]));
    lastRaw.item = items[0] || null;
    if (invList.truncated) warn.push("invoices capped at 5000");
    const invs = invList.rows.filter((i) => i.date >= from && !["void", "draft"].includes(i.status));
    const invoices = await detailsFor("bill", "invoices", invs, "invoice_id", "invoice");
    lastRaw.invoice = invoices[0] || null;

    // Mobile numbers: from the bill if present, else the customer record.
    // Many customers → one contacts list call instead of one call per customer.
    const mobiles = {};
    const inRange = invoices.filter((i) => i.date <= to);
    const needContact = [...new Set(inRange
      .filter((i) => !(i.billing_address?.phone || cf(i, FIELDS.mobile)) && i.customer_id)
      .map((i) => i.customer_id))];
    const mobOf = (c) => c?.mobile || c?.phone || c?.contact_persons?.[0]?.mobile || c?.contact_persons?.[0]?.phone || "";
    if (needContact.length > 8) {
      status("Loading customer mobile numbers…");
      try {
        const want = new Set(needContact);
        for (const c of (await listAll("contacts")).rows) if (want.has(c.contact_id)) mobiles[c.contact_id] = mobOf(c);
      } catch { /* fall through to per-customer lookup below */ }
    }
    for (let k = 0; k < needContact.length; k++) {
      if (mobiles[needContact[k]]) continue;
      status(`Loading customer mobile numbers ${k + 1} of ${needContact.length}…`);
      try { mobiles[needContact[k]] = mobOf(await detail("contacts", needContact[k], "v1", "contact")); }
      catch { mobiles[needContact[k]] = ""; }
    }
    const mobileOf = (i) => i.billing_address?.phone || cf(i, FIELDS.mobile) || mobiles[i.customer_id] || "";

    SHEETS[0].rows = inRange
      .sort((a, b) => (a.date + a.invoice_number).localeCompare(b.date + b.invoice_number, "en", { numeric: true }))
      .flatMap((inv) => salesRows(inv, itemsById, mobileOf));

    // Daily Statement: payments by mode, month-to-date sales, shop heading
    const pays = paysRes.rows.filter((p) => p.date >= from && p.date <= to);
    const mtdList = mtdRes ? mtdRes.rows : invList.rows;
    const mtd = mtdList.filter((i) => i.date >= monthStart && i.date <= to && !["void", "draft"].includes(i.status));
    summary = buildSummary(inRange, SHEETS[0].rows, pays, mtd, from, to);
    summary.header = await receiptHeader();

    // Show Sales Report + All Sales Summary now; Stock finishes in the background
    stockAll = []; applyZeroFilter();
    loaded = true;
    renderAll();
    const doneText = () => `${fmtDate(from)} to ${fmtDate(to)} · ${inRange.length} bills · refreshed ${new Date().toLocaleTimeString()}` +
      (warn.length ? ` · ⚠ ${warn.join(", ")}` : "");
    $("statusR").textContent = doneText() + " · Stock loading…";

    // PHASE 2 — purchase bills + adjustments → Stock Report
    const [billList, adjList] = await Promise.all([
      listAll("bills", { date_start: from, date_end: today }),
      listAll("inventoryadjustments", {}, "inventory_adjustments")
    ]);
    const billsBrief = billList.rows.filter((b) => b.date >= from && !["void", "draft"].includes(b.status));
    const bills = await detailsFor("purchase bill", "bills", billsBrief, "bill_id", "bill");
    const adjBrief = adjList.rows.filter((a) => a.date >= from);
    const adjustments = await detailsFor("adjustment", "inventoryadjustments", adjBrief, "inventory_adjustment_id", "inventory_adjustment");

    stockAll = stockRows(items, invoices, bills, adjustments, from, to);
    applyZeroFilter();
    renderAll();
    $("statusR").textContent = doneText();
  } catch (e) {
    status("Error: " + e.message);
  } finally {
    $("run").disabled = false;
    progress(0, 0);
  }
};

// Daily Statement (same lines as the Quick Bill slip)
function buildSummary(invoices, lines, payments, mtdInvoices, from, to) {
  const byBill = {};
  lines.forEach((l) => {
    const b = (byBill[l._invNo] ??= { disc: 0, taxable: 0, tax: 0 });
    b.disc += l.totDisc; b.taxable += l.taxable; b.tax += l.taxAmt;
  });
  let sales = 0, discCnt = 0, disc = 0, taxable = 0, roCnt = 0, ro = 0, credit = 0, creditCnt = 0;
  invoices.forEach((inv) => {
    const b = byBill[inv.invoice_number] || { disc: 0, taxable: 0, tax: 0 };
    const total = num(inv.total);
    sales += total; disc += b.disc; taxable += b.taxable;
    if (b.disc > 0.004) discCnt++;
    // round-off: Zoho's own rounding field if present, else the difference between bill total and item totals
    let r = inv.roundoff_value ?? inv.rounding_off ?? (/round/i.test(inv.adjustment_description || "") ? inv.adjustment : undefined);
    r = r !== undefined && r !== "" ? num(r) : total - (b.taxable + b.tax) - num(inv.shipping_charge);
    r = r2(r);
    if (Math.abs(r) >= 0.01) { roCnt++; ro += r; }
    const bal = num(inv.balance);
    if (bal > 0.004) { credit += bal; creditCnt++; }
  });
  const modes = {};
  payments.forEach((p) => {
    const k = p.payment_mode || "Other";
    (modes[k] ??= { cnt: 0, amt: 0 }); modes[k].cnt++; modes[k].amt += num(p.amount);
  });
  if (credit > 0.004) modes["Credit (Unpaid)"] = { cnt: creditCnt, amt: credit };
  return {
    from, to,
    bill: [
      ["Sales", invoices.length, sales],
      ["Discount", discCnt, disc],
      ["Taxable Amount", "", taxable],
      ["Net Sales", "", sales - disc],
      ["Round Off", roCnt || "", ro],
      ["Net Amount", "", sales - disc + ro, true]
    ],
    collection: Object.entries(modes).sort((a, b) => b[1].amt - a[1].amt).map(([k, v]) => [k, v.cnt, v.amt]),
    mtd: [["Net Sales", mtdInvoices.length, mtdInvoices.reduce((s, i) => s + num(i.total), 0)]]
  };
}

async function receiptHeader() {
  if (RECEIPT_HEADER[orgId()]) return RECEIPT_HEADER[orgId()];
  try {
    const o = (await api("fetch", { path: `organizations/${orgId()}`, query: {} })).organization || {};
    const a = o.address || {};
    return {
      name: o.name || "",
      lines: [[a.street_address1, a.street_address2].filter(Boolean).join(", "), [a.city, a.zip].filter(Boolean).join(" - ")],
      gstin: o.gst_no || o.tax_reg_no || ""
    };
  } catch { return { name: "", lines: [], gstin: "" }; }
}

const inrAmt = (n) => num(n).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const nowStamp = () => { const d = new Date(); return `${fmtDate(iso(d))} ${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`; };

function renderSummary() {
  if (!summary) { $("gridWrap").innerHTML = `<div class="empty">Click ▶ Run to build the All Sales Summary.</div>`; return; }
  const h = summary.header || {};
  const rows = (list) => list.map(([lbl, cnt, amt, strong]) =>
    `<tr class="${strong ? "strong" : ""}"><td class="lbl">${esc(lbl)}</td><td class="cnt">${esc(cnt)}</td><td class="amt">${esc(inrAmt(amt))}</td></tr>`).join("");
  $("gridWrap").innerHTML = `
    <div class="receiptWrap"><div>
      <div class="receipt" id="receipt">
        <div class="c title">All Sales Summary</div>
        <div class="c">From &nbsp;&nbsp; ${fmtDate(summary.from)} &nbsp;To&nbsp; ${fmtDate(summary.to)}</div>
        <div class="c">Print Date &amp; Time: &nbsp;${nowStamp()}</div>
        <div class="c co">${esc(h.name)}</div>
        ${(h.lines || []).map((l) => `<div class="c">${esc(l) || "&nbsp;"}</div>`).join("")}
        ${h.gstin ? `<div class="gst">GSTIN : ${esc(h.gstin)}</div>` : ""}
        <table>
          <tr><th class="n">Name</th><th>Cnt</th><th style="text-align:right">Amount</th></tr>
          <tr class="sec"><td colspan="3">Bill Summary</td></tr>${rows(summary.bill)}
          <tr class="sec"><td colspan="3">Collection Summary</td></tr>${rows(summary.collection.length ? summary.collection : [["No payments", "", 0]])}
          <tr class="sec"><td colspan="3">MTD Sales</td></tr>${rows(summary.mtd)}
        </table>
      </div>
      <div class="note">🖨 Print → choose your receipt printer, or "Save as PDF"</div>
    </div></div>`;
  $("nameBox").textContent = ""; $("fVal").textContent = "";
  status(`All Sales Summary: ${fmtDate(summary.from)} to ${fmtDate(summary.to)}`);
}

// ---------------------------------------------------------------------
// 5) ON-SCREEN GRID
// ---------------------------------------------------------------------
const fmtDate = (v) => (v ? String(v).slice(0, 10).split("-").reverse().join("-") : "");
const show = (v, t) => (t === "num" ? (v === "" || v == null ? "" : num(v).toFixed(2)) : t === "date" ? fmtDate(v) : v ?? "");
const screenCols = () => active.cols.map((c, i) => ({ ...c, i })).filter((c) => c.screen !== false);

function matches(value, f, t) {
  f = f.trim();
  if (!f) return true;
  if (t === "num") {
    const m = f.match(/^(>=|<=|<>|>|<|=)?\s*(-?[\d,]*\.?\d+)$/);
    if (m) {
      const v = num(value), n = Number(m[2].replace(/,/g, ""));
      switch (m[1] || "=") {
        case ">": return v > n;  case "<": return v < n;
        case ">=": return v >= n; case "<=": return v <= n;
        case "<>": return v !== n; default: return v === n;
      }
    }
  }
  const q = f.toLowerCase();
  return String(show(value, t)).toLowerCase().includes(q) || String(value ?? "").toLowerCase().includes(q);
}

function visibleRows(sheet) {
  let rows = sheet.rows.filter((r) => sheet.cols.every((c, i) => matches(r[c.key], sheet.filters[i] || "", c.t)));
  if (sheet.sort) {
    const { i, dir } = sheet.sort, c = sheet.cols[i];
    rows = [...rows].sort((a, b) => {
      const x = a[c.key], y = b[c.key];
      const cmp = c.t === "num" ? num(x) - num(y) : String(x ?? "").localeCompare(String(y ?? ""), "en", { numeric: true });
      return cmp * dir;
    });
  }
  return rows;
}

function renderTabs() {
  $("tabs").innerHTML = TABS.map((s, i) => `<button data-tab="${i}" class="${s === active ? "active" : ""}">${esc(s.name)}</button>`).join("");
}

function renderAll() {
  renderTabs();
  if (active.summary) return renderSummary();
  if (!loaded) return;
  const cols = screenCols();
  const px = (c) => Math.max(60, Math.round(c.w * 7.5));
  const width = 44 + cols.reduce((s, c) => s + px(c), 0);
  $("gridWrap").innerHTML = `
    <table class="grid" style="width:${width}px">
      <colgroup><col style="width:44px">${cols.map((c) => `<col style="width:${px(c)}px">`).join("")}</colgroup>
      <thead>
        <tr class="letters"><th class="rn"></th>${cols.map((c) => `<th data-col="${c.i}">${colName(c.i)}</th>`).join("")}</tr>
        <tr class="hdr"><th class="rn">1</th>${cols.map((c) => {
          const arrow = active.sort?.i === c.i ? `<span class="arrow">${active.sort.dir > 0 ? "▲" : "▼"}</span>` : "";
          return `<th data-sort="${c.i}" class="${c.h === "R" ? "r" : ""}" title="Click to sort">${esc(c.label)}${arrow}</th>`;
        }).join("")}</tr>
        <tr class="flt"><th class="rn">🔍</th>${cols.map((c) =>
          `<th><input data-filter="${c.i}" value="${esc(active.filters[c.i] || "")}" placeholder="${c.t === "num" ? "e.g. >0" : "filter…"}"></th>`).join("")}</tr>
      </thead>
      <tbody id="body"></tbody>
      <tfoot id="foot"></tfoot>
    </table>`;
  renderBody();
}

function renderBody() {
  const cols = screenCols();
  view = visibleRows(active);
  $("body").innerHTML = view.length
    ? view.map((r, ri) => `<tr><th class="rn">${ri + 2}</th>${cols.map((c) =>
        `<td data-r="${ri}" data-c="${c.i}" class="${c.t === "num" ? "r" : ""}">${esc(show(r[c.key], c.t))}</td>`).join("")}</tr>`).join("")
    : `<tr><th class="rn">2</th><td colspan="${cols.length}" style="color:#888">No rows for this period / filter.</td></tr>`;

  $("foot").innerHTML = `<tr><th class="rn">Σ</th>${cols.map((c, k) => {
    if (c.sum) return `<td>${esc(num(view.reduce((s, r) => s + num(r[c.key]), 0)).toFixed(2))}</td>`;
    return `<td>${k === 0 ? "Grand Total : " : ""}</td>`;
  }).join("")}</tr>`;

  sel = null;
  paintSelection();
  status(`${active.name}: showing ${view.length} of ${active.rows.length} rows`);
}

function paintSelection() {
  document.querySelectorAll(".grid .sel, .grid .colsel").forEach((el) => el.classList.remove("sel", "colsel"));
  if (selCol !== null) document.querySelectorAll(`.grid [data-c="${selCol}"], .grid [data-col="${selCol}"]`).forEach((el) => el.classList.add("colsel"));
  if (sel) {
    const td = document.querySelector(`td[data-r="${sel.r}"][data-c="${sel.c}"]`);
    td?.classList.add("sel");
    td?.scrollIntoView({ block: "nearest", inline: "nearest" });
    const c = active.cols[sel.c];
    $("nameBox").textContent = colName(sel.c) + (sel.r + 2);
    $("fVal").textContent = show(view[sel.r]?.[c.key], c.t);
  } else if (selCol !== null) {
    const c = active.cols[selCol];
    $("nameBox").textContent = `${colName(selCol)}:${colName(selCol)}`;
    $("fVal").textContent = c.label;
    const vals = view.map((r) => r[c.key]).filter((v) => v !== "" && v != null);
    if (c.t === "num" && vals.length) {
      const sum = vals.reduce((s, v) => s + num(v), 0);
      $("statusR").textContent = `Average: ${(sum / vals.length).toFixed(2)}    Count: ${vals.length}    Sum: ${sum.toFixed(2)}`;
    } else $("statusR").textContent = `Count: ${vals.length}`;
  } else { $("nameBox").textContent = ""; $("fVal").textContent = ""; }
}

// ---------------------------------------------------------------------
// events
// ---------------------------------------------------------------------
$("tabs").onclick = (e) => {
  const b = e.target.closest("[data-tab]");
  if (!b) return;
  active = TABS[Number(b.dataset.tab)];
  sel = null; selCol = null;
  renderAll();
};

$("gridWrap").addEventListener("click", (e) => {
  const sortTh = e.target.closest("[data-sort]"), letter = e.target.closest("[data-col]"), td = e.target.closest("td[data-r]");
  if (sortTh) {
    const i = Number(sortTh.dataset.sort);
    active.sort = active.sort?.i === i ? (active.sort.dir > 0 ? { i, dir: -1 } : null) : { i, dir: 1 };
    renderAll();
  } else if (letter) { selCol = Number(letter.dataset.col); sel = null; paintSelection(); }
  else if (td) { sel = { r: Number(td.dataset.r), c: Number(td.dataset.c) }; selCol = null; paintSelection(); }
});

$("gridWrap").addEventListener("input", (e) => {
  const f = e.target.closest("[data-filter]");
  if (!f) return;
  active.filters[f.dataset.filter] = f.value;
  renderBody();
});

document.addEventListener("keydown", (e) => {
  if (!sel || e.target.tagName === "INPUT") return;
  const m = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
  if (!m) return;
  e.preventDefault();
  const idx = screenCols().map((c) => c.i);
  const pos = Math.min(Math.max(idx.indexOf(sel.c) + m[1], 0), idx.length - 1);
  sel = { r: Math.min(Math.max(sel.r + m[0], 0), view.length - 1), c: idx[pos] };
  paintSelection();
});

$("clear").onclick = () => { if (active.summary) return; active.filters = {}; active.sort = null; renderAll(); };
$("print").onclick = () => {
  if (active.summary) {
    if (!summary) return status("Run the report first.");
    renderSummary(); // refresh the print time
    document.body.classList.add("print-summary");
    window.print();
    document.body.classList.remove("print-summary");
  } else window.print();
};

$("dlgClose").onclick = () => $("dlg").close();
$("dlgCopy").onclick = () => navigator.clipboard.writeText($("dlgPre").textContent);

// ---------------------------------------------------------------------
// 6) DOWNLOADS — exact template layouts
// ---------------------------------------------------------------------
function download(bytes, filename) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}
const hdrRow = (cols) => ({ height: 30, cells: cols.map((c) => (c.label ? { v: c.label, s: c.h === "R" ? ST.hdrR : ST.hdrC } : null)) });
const dataRow = (cols, r) => ({ height: 18, cells: cols.map((c) => (c.t ? { v: r[c.key], s: ST[c.t] } : null)) });
const totalRow = (cols, rows, label, height, startCol = 0) => ({
  height,
  cells: cols.map((c, i) => {
    if (i < startCol) return null;
    if (c.sum) return { v: r2(rows.reduce((s, r) => s + num(r[c.key]), 0)), s: ST.tot };
    return { v: i === startCol && label ? label : "", s: ST.tot };
  })
});

// Sales: header · one row per item line · "Grand Total : "
function buildSalesXlsx(rows) {
  return XlsxWriter.build({
    name: "Sheet1",
    widths: SALES_COLS.map((c) => c.w),
    rows: [hdrRow(SALES_COLS), ...rows.map((r) => dataRow(SALES_COLS, r)), totalRow(SALES_COLS, rows, "Grand Total : ", 22.5)]
  });
}

// Stock: header · "Net Amount: " band (A:T merged) · items · group total · grand total
function buildStockXlsx(rows) {
  const last = colName(STOCK_COLS.length - 1);
  return XlsxWriter.build({
    name: "Sheet1",
    widths: STOCK_COLS.map((c) => c.w),
    merges: [`A2:${last}2`],
    rows: [
      hdrRow(STOCK_COLS),
      { height: 19.5, cells: STOCK_COLS.map((_, i) => ({ v: i === 0 ? "Net Amount: " : "", s: ST.band })) },
      ...rows.map((r) => dataRow(STOCK_COLS, r)),
      totalRow(STOCK_COLS, rows, "", 18, 1),
      totalRow(STOCK_COLS, rows, "", 22.5, 1)
    ]
  });
}

$("dlSales").onclick = () => {
  if (!loaded) return status("Run the report first.");
  download(buildSalesXlsx(visibleRows(SHEETS[0])), `Sale_Report_${$("from").value}_to_${$("to").value}.xlsx`);
};
$("dlStock").onclick = () => {
  if (!loaded) return status("Run the report first.");
  download(buildStockXlsx(visibleRows(SHEETS[1])), `Stock_Report_${$("from").value}_to_${$("to").value}.xlsx`);
};

// ---------------------------------------------------------------------
// period presets + init
// ---------------------------------------------------------------------
function applyPreset() {
  const p = $("preset").value, t = new Date();
  let from = new Date(t), to = new Date(t);
  if (p === "yesterday") { from.setDate(t.getDate() - 1); to = new Date(from); }
  else if (p === "7") from.setDate(t.getDate() - 6);
  else if (p === "30") from.setDate(t.getDate() - 29);
  else if (p === "month") from = new Date(t.getFullYear(), t.getMonth(), 1);
  else if (p === "lastmonth") { from = new Date(t.getFullYear(), t.getMonth() - 1, 1); to = new Date(t.getFullYear(), t.getMonth(), 0); }
  else if (p === "custom") return;
  $("from").value = iso(from);
  $("to").value = iso(to);
}
$("preset").onchange = applyPreset;
["from", "to"].forEach((id) => ($(id).onchange = () => ($("preset").value = "custom")));
applyPreset();
chrome.storage.local.get("orgId").then(({ orgId: saved }) => { $("orgId").value = saved || DEFAULT_ORG_ID; });
renderTabs();

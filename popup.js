const $ = (id) => document.getElementById(id);

// Which columns to show per module
const COLUMNS = {
  items:    ["name", "sku", "rate", "stock_on_hand"],
  invoices: ["invoice_number", "date", "customer_name", "total", "status"],
  contacts: ["contact_name", "phone", "email"]
};

chrome.storage.local.get("orgId").then(({ orgId }) => { $("orgId").value = orgId || "60082587105"; });

$("login").onclick = async () => {
  const r = await chrome.runtime.sendMessage({ type: "login" });
  $("status").textContent = r.ok ? "Signed in ✓" : "Login failed: " + r.error;
};

$("report").onclick = async () => {
  const orgId = $("orgId").value.trim();
  if (orgId) await chrome.storage.local.set({ orgId });
  chrome.tabs.create({ url: chrome.runtime.getURL("report.html") });
};

$("fetch").onclick = async () => {
  const orgId = $("orgId").value.trim();
  const module = $("module").value;
  if (!orgId) return ($("status").textContent = "Enter Organization ID");
  await chrome.storage.local.set({ orgId });

  $("status").textContent = "Loading…";
  const r = await chrome.runtime.sendMessage({
    type: "fetch", path: module, orgId, query: { page: 1, per_page: 50 }
  });
  if (!r.ok) return ($("status").textContent = "Error: " + r.error);

  const rows = r.data[module] || [];
  const cols = COLUMNS[module];
  $("out").innerHTML =
    "<tr>" + cols.map((c) => `<th>${c}</th>`).join("") + "</tr>" +
    rows.map((row) => "<tr>" + cols.map((c) => `<td>${row[c] ?? ""}</td>`).join("") + "</tr>").join("");
  $("status").textContent = `${rows.length} ${module} loaded` +
    (r.data.page_context?.has_more_page ? " (more pages available)" : "");
};

// ---- CONFIG ----
const CLIENT_ID = "1000.Q025XFVN3YC0SW2NAR4NO8TQSDH7DW";
const ACCOUNTS = "https://accounts.zoho.in";            // .com for US accounts
const API_BASE = "https://api.zakya.in/inventory/v1";   // api.zakya.com for US
const SCOPES = [
  "ZohoPOSAPI.items.READ",
  "ZohoPOSAPI.invoices.READ",
  "ZohoPOSAPI.invoices.CREATE",   // needed for PUT (update) operations
  "ZohoPOSAPI.contacts.READ",
  "ZohoPOSAPI.bills.READ",
  "ZohoPOSAPI.inventoryadjustments.READ",
  "ZohoPOSAPI.customerpayments.READ",
  "ZohoPOSAPI.settings.READ"
].join(",");
const MIN_GAP_MS = 620; // ~95 calls/min, just under Zoho's 100/min limit
// ----------------

// Sign in with Zoho (implicit grant: token comes back in the redirect URL; no client secret needed)
async function login(interactive = true) {
  const redirectUri = chrome.identity.getRedirectURL(); // https://<extension-id>.chromiumapp.org/
  const authUrl =
    `${ACCOUNTS}/oauth/v2/auth?response_type=token` +
    `&client_id=${encodeURIComponent(CLIENT_ID)}` +
    `&scope=${encodeURIComponent(SCOPES)}` +
    `&redirect_uri=${encodeURIComponent(redirectUri)}` +
    (interactive ? "&prompt=consent" : "");

  const resultUrl = await chrome.identity.launchWebAuthFlow({ url: authUrl, interactive });
  const params = new URLSearchParams(new URL(resultUrl).hash.slice(1));
  const token = params.get("access_token");
  if (!token) throw new Error("Zoho did not return a token: " + (params.get("error") || resultUrl));

  const expiresAt = Date.now() + (Number(params.get("expires_in") || 3600) - 60) * 1000;
  await chrome.storage.local.set({ token, expiresAt, tokenScopes: SCOPES });
  return token;
}

async function getToken() {
  const { token, expiresAt, tokenScopes } = await chrome.storage.local.get(["token", "expiresAt", "tokenScopes"]);
  if (token && Date.now() < expiresAt && tokenScopes === SCOPES) return token;
  try { return await login(false); }   // silent re-login if Zoho session is still active
  catch { return await login(true); }
}

// Space calls out so we never exceed the per-minute limit
let nextSlot = 0;
async function throttle() {
  const now = Date.now();
  const wait = Math.max(0, nextSlot - now);
  nextSlot = Math.max(now, nextSlot) + MIN_GAP_MS;
  if (wait) await new Promise((r) => setTimeout(r, wait));
}

// Generic GET to Zoho POS (retries on rate-limit and expired token)
async function posGet(path, orgId, query = {}, attempt = 0) {
  const token = await getToken();
  await throttle();
  const qs = new URLSearchParams({ organization_id: orgId, ...query });
  const res = await fetch(`${API_BASE}/${path}?${qs}`, {
    headers: { Authorization: `Zoho-oauthtoken ${token}` }
  });
  if (res.status === 429 && attempt < 2) {             // too many requests → wait a minute
    await new Promise((r) => setTimeout(r, 61000));
    return posGet(path, orgId, query, attempt + 1);
  }
  if (res.status === 401 && attempt < 1) {             // token expired/revoked → sign in again
    await chrome.storage.local.remove("token");
    return posGet(path, orgId, query, attempt + 1);
  }
  const data = await res.json().catch(() => ({ code: -1, message: `HTTP ${res.status}` }));
  if (data.code !== 0) throw new Error(`${res.status} ${data.message}`);
  return data;
}

// Fetch every page of a list endpoint (per_page max is 200)
async function posGetAll(path, orgId, query = {}, key = path, maxPages = 25) {
  let all = [];
  for (let page = 1; page <= maxPages; page++) {
    const data = await posGet(path, orgId, { ...query, page, per_page: 200 });
    all = all.concat(data[key] || []);
    if (!data.page_context?.has_more_page) return { rows: all, truncated: false };
  }
  return { rows: all, truncated: true };
}

// Pages talk to the service worker via messages
chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  (async () => {
    try {
      if (msg.type === "login") {
        await login(true);
        sendResponse({ ok: true });
      } else if (msg.type === "fetch") {
        sendResponse({ ok: true, data: await posGet(msg.path, msg.orgId, msg.query) });
      } else if (msg.type === "fetchAll") {
        sendResponse({ ok: true, data: await posGetAll(msg.path, msg.orgId, msg.query, msg.key) });
      } else if (msg.type === "put") {
        // PUT to update an existing record (e.g. invoice custom fields)
        const token = await getToken();
        await throttle();
        const qs = new URLSearchParams({ organization_id: msg.orgId });
        const res = await fetch(`${API_BASE}/${msg.path}?${qs}`, {
          method: "PUT",
          headers: { Authorization: `Zoho-oauthtoken ${token}`, "Content-Type": "application/json" },
          body: JSON.stringify(msg.body)
        });
        const data = await res.json().catch(() => ({ code: -1, message: `HTTP ${res.status}` }));
        if (data.code !== 0) throw new Error(`${res.status} ${data.message}`);
        sendResponse({ ok: true, data });
      }
    } catch (e) {
      sendResponse({ ok: false, error: String(e.message || e) });
    }
  })();
  return true; // keep channel open for async response
});

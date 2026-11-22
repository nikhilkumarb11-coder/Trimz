/**
 * Trimz ↔ Google Sheet sync.
 * Same idea as the campus ERP: the screen updates immediately, localStorage
 * keeps a copy, and Apps Script reads/writes the sheet in the background.
 *
 * Paste the web app URL in Settings, or set TRIMZ_API_URL below.
 * The passcode must match SCRIPT_PASSWORD in Code.gs.
 */
const TRIMZ_API_URL = "https://script.google.com/macros/s/AKfycbyQuc-bgQ733X_bBA7OJK5X1Tkh1dlTY4xqFncctnchJXjcbnWW_LNy31vZ8PQ3ioxJ/exec";
const TRIMZ_PASSCODE = "trimz123";
const AUTO_SYNC_MS = 3 * 60 * 1000;

const SHEET_TABS = {
  customers: { sheet: "Customers", cols: ["id", "name", "phone", "gender", "visits", "spent", "last", "points", "notes", "marketingOptIn"] },
  services: { sheet: "Services", cols: ["id", "name", "category", "audience", "price", "priceNote", "duration", "commission", "source"] },
  staff: { sheet: "Staff", cols: ["id", "name", "role", "sales", "commission", "status", "target", "attendance"] },
  inventory: { sheet: "Inventory", cols: ["id", "name", "brand", "qty", "price", "cost", "expiry", "supplier"] },
  appointments: { sheet: "Appointments", cols: ["id", "customer", "service", "staff", "date", "time", "status"] },
  memberships: { sheet: "Memberships", cols: ["id", "name", "customer", "phone", "discount", "valid", "status"] },
  packages: { sheet: "Packages", cols: ["id", "name", "customer", "included", "remaining", "valid"] },
  expenses: { sheet: "Expenses", cols: ["id", "name", "category", "amount", "date", "mode"] },
  bills: { sheet: "Bills", cols: ["id", "no", "customer", "phone", "items", "date", "subtotal", "discount", "tax", "gstIncluded", "gstRate", "total", "mode", "time", "status", "staff", "commission", "lines"] }
};

const SHEET_NUM_COLS = new Set(["visits", "spent", "points", "price", "duration", "commission", "sales", "target", "qty", "cost", "amount", "subtotal", "discount", "tax", "gstRate", "total"]);
const SHEET_BOOL_COLS = new Set(["marketingOptIn", "gstIncluded"]);
const SHEET_JSON_COLS = new Set(["attendance", "lines"]);

let sheetSnapshot = null;
let sheetPushTimer = null;
let sheetAutoTimer = null;
let sheetChain = Promise.resolve();
let isSheetSyncing = false;
let lastServerSyncAt = 0;
let lastSyncFailed = false;
let sheetPullQueued = false;
const pendingSheetDeletes = new Set();

function sheetConfig() {
  return {
    apiUrl: String((typeof localStorage !== "undefined" && localStorage.getItem("trimzApiUrl")) || TRIMZ_API_URL || "").trim(),
    passcode: String((typeof localStorage !== "undefined" && localStorage.getItem("trimzSheetPasscode")) || TRIMZ_PASSCODE || "")
  };
}

function sheetSlug_(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 40) || "x";
}

function suggestSheetId_(type, row) {
  if (type === "bills" && row.no) return String(row.no);
  if (type === "customers") {
    const phone = String(row.phone || "").replace(/\D/g, "");
    if (phone.length >= 10) return "CUS-" + phone.slice(-10);
  }
  if (type === "services") return "SVC-" + sheetSlug_(row.audience) + "-" + sheetSlug_(row.name);
  if (type === "staff") return "STF-" + sheetSlug_(row.name);
  if (type === "memberships" && row.id) return String(row.id);
  return type.slice(0, 3).toUpperCase() + "-" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

function ensureSheetIds() {
  if (typeof data === "undefined" || !data) return;
  Object.keys(SHEET_TABS).forEach(type => {
    const used = new Set();
    (data[type] || []).forEach(row => {
      if (!row || typeof row !== "object") return;
      let id = row.id ? String(row.id) : suggestSheetId_(type, row);
      const base = id;
      let n = 2;
      while (used.has(id)) id = base + "-" + (n++);
      row.id = id;
      used.add(id);
    });
  });
}

function sheetCell_(col, value) {
  if (SHEET_BOOL_COLS.has(col)) return value === true || value === "true" || value === "TRUE" ? "true" : "false";
  if (SHEET_JSON_COLS.has(col)) {
    if (col === "lines") return JSON.stringify(Array.isArray(value) ? value : []);
    if (value && typeof value === "object") return JSON.stringify(value);
    return "{}";
  }
  if (value == null) return "";
  if (typeof value === "object") return JSON.stringify(value);
  return value;
}

function toSheetRow_(type, row) {
  return SHEET_TABS[type].cols.map(col => sheetCell_(col, row[col]));
}

function sheetSignature_(type, row) {
  const text = JSON.stringify(toSheetRow_(type, row));
  let hash = 5381;
  for (let i = 0; i < text.length; i++) hash = ((hash << 5) + hash) ^ text.charCodeAt(i);
  return (hash >>> 0).toString(36);
}

function fromSheetRow_(type, obj) {
  const row = {};
  SHEET_TABS[type].cols.forEach(col => {
    let value = obj[col];
    if (value == null) value = "";
    if (SHEET_JSON_COLS.has(col)) {
      if (typeof value === "string" && /^[\[{]/.test(value.trim())) {
        try { value = JSON.parse(value); } catch (e) { value = col === "lines" ? [] : {}; }
      }
      if (value === "" || value == null) value = col === "lines" ? [] : {};
    } else if (SHEET_BOOL_COLS.has(col)) {
      value = value === true || value === "true" || value === "TRUE";
    } else if (SHEET_NUM_COLS.has(col)) {
      if (value === "" || value == null) value = (col === "duration" || col === "commission") ? null : 0;
      else value = Number(value);
    } else if (col === "id" || col === "no" || col === "phone") {
      value = value == null ? "" : String(value);
    }
    row[col] = value;
  });
  if (!row.id) row.id = suggestSheetId_(type, row);
  return row;
}

function emptySnapshot_() {
  const rows = {};
  Object.keys(SHEET_TABS).forEach(type => { rows[type] = {}; });
  return { gstRate: Number(data.gstRate ?? 5), rows };
}

function snapshotFromData_() {
  const snap = emptySnapshot_();
  Object.keys(SHEET_TABS).forEach(type => {
    (data[type] || []).forEach(row => {
      if (row && row.id) snap.rows[type][String(row.id)] = sheetSignature_(type, row);
    });
  });
  return snap;
}

function persistSheetSnapshot() {
  try { localStorage.setItem("trimzSheetSnapshot", JSON.stringify(sheetSnapshot)); } catch (e) {}
}

function loadSheetSnapshot() {
  try { sheetSnapshot = JSON.parse(localStorage.getItem("trimzSheetSnapshot") || "null"); } catch (e) { sheetSnapshot = null; }
  if (!sheetSnapshot || !sheetSnapshot.rows) sheetSnapshot = null;
}

function persistSalonLocal() {
  try {
    localStorage.setItem("trimzDesk", JSON.stringify(data));
    if (typeof STORE_VERSION !== "undefined") localStorage.setItem("trimzDeskVersion", STORE_VERSION);
  } catch (e) {}
}

function enqueueSheet(work) {
  const run = sheetChain.then(work, work);
  sheetChain = run.then(() => {}, () => {});
  return run;
}

function sheetConfigured() {
  return !!sheetConfig().apiUrl;
}

function updateSyncStatus() {
  const btn = document.getElementById("btnSyncHeader");
  const label = document.getElementById("syncChipLabel");
  if (!btn || !label) return;
  let text = "Not synced yet";
  let tone = "warn";
  const ageMin = lastServerSyncAt ? Math.floor((Date.now() - lastServerSyncAt) / 60000) : -1;
  if (!sheetConfigured()) { text = "Sheet off"; tone = "warn"; }
  else if (isSheetSyncing) { text = "Syncing…"; tone = "busy"; }
  else if (typeof navigator !== "undefined" && navigator.onLine === false) { text = "Offline"; tone = "warn"; }
  else if (lastSyncFailed) { text = "Not synced"; tone = "warn"; }
  else if (ageMin < 0) { text = "Not synced yet"; tone = "warn"; }
  else if (ageMin < 1) { text = "Synced just now"; tone = "ok"; }
  else if (ageMin < 60) { text = "Synced " + ageMin + " min ago"; tone = ageMin < 15 ? "ok" : "warn"; }
  else if (ageMin < 1440) { text = "Synced " + Math.floor(ageMin / 60) + " hr ago"; tone = "warn"; }
  else { text = "Synced " + Math.floor(ageMin / 1440) + " d ago"; tone = "warn"; }
  label.textContent = text;
  btn.dataset.tone = tone;
  btn.title = lastServerSyncAt
    ? "Last synced " + new Date(lastServerSyncAt).toLocaleString("en-IN") + ". Tap to sync now."
    : (sheetConfigured() ? "Tap to sync now" : "Add the Google Sheet web app URL in Settings");
}

function paintAfterSync() {
  const modal = document.getElementById("modal");
  if (modal && modal.classList.contains("open")) return;
  if (typeof current !== "undefined" && current === "billing") return;
  const el = document.activeElement;
  if (el && (el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.tagName === "SELECT")) return;
  if (typeof render === "function") render();
}

async function postSheet(payload) {
  const cfg = sheetConfig();
  if (!cfg.apiUrl) return { success: false, error: "No sheet URL" };
  const res = await fetch(cfg.apiUrl, {
    method: "POST",
    redirect: "follow",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(Object.assign({ password: cfg.passcode }, payload))
  });
  const text = await res.text();
  try { return JSON.parse(text); }
  catch (e) { return { success: false, error: "Apps Script did not return data (HTTP " + res.status + ")." }; }
}

function settingValue_(rows, key) {
  const hit = (rows || []).find(row => String(row.key ?? row.Key ?? row.id ?? "") === key);
  if (!hit) return undefined;
  return hit.value ?? hit.Value;
}

function localRowDirty_(type, row) {
  if (!sheetSnapshot || !sheetSnapshot.rows || !sheetSnapshot.rows[type]) return false;
  const prev = sheetSnapshot.rows[type][String(row.id)];
  if (prev == null) return true;
  return prev !== sheetSignature_(type, row);
}

function mergeFromServer_(serverRows) {
  const fresh = !sheetSnapshot;
  Object.keys(SHEET_TABS).forEach(type => {
    const incoming = serverRows[type] || [];
    const byId = new Map(incoming.map(row => [String(row.id), row]));
    const used = new Set();
    const merged = [];
    (data[type] || []).forEach(row => {
      const id = String(row.id || "");
      const key = type + ":" + id;
      if (pendingSheetDeletes.has(key)) return;
      const serverRow = byId.get(id);
      if (!serverRow) {
        const known = !!(sheetSnapshot && sheetSnapshot.rows[type] && sheetSnapshot.rows[type][id]);
        if (known && !localRowDirty_(type, row)) return;
        if (fresh && incoming.length > 0) return;
        merged.push(row);
        return;
      }
      used.add(id);
      merged.push(!fresh && localRowDirty_(type, row) ? row : serverRow);
    });
    const extras = [];
    incoming.forEach(row => {
      const id = String(row.id);
      if (used.has(id) || pendingSheetDeletes.has(type + ":" + id)) return;
      extras.push(row);
    });
    data[type] = extras.concat(merged);
  });
}

function applyGstFromServer_(settingsRows) {
  const raw = settingValue_(settingsRows, "gstRate");
  if (raw == null || raw === "") return;
  const rate = Number(raw);
  if (!Number.isFinite(rate)) return;
  const dirty = sheetSnapshot && Number(sheetSnapshot.gstRate) !== Number(data.gstRate ?? 5);
  if (!dirty) data.gstRate = rate;
}

async function seedSheet_() {
  if (!sheetSnapshot) sheetSnapshot = emptySnapshot_();
  const types = Object.keys(SHEET_TABS);
  for (let t = 0; t < types.length; t++) {
    const type = types[t];
    const pending = (data[type] || []).filter(row => row && row.id && sheetSnapshot.rows[type][String(row.id)] == null);
    for (let i = 0; i < pending.length; i += 100) {
      const part = pending.slice(i, i + 100);
      const result = await postSheet({
        action: "addRows",
        targetSheet: SHEET_TABS[type].sheet,
        rows: part.map(row => toSheetRow_(type, row))
      });
      if (!result || result.success === false) {
        persistSheetSnapshot();
        throw new Error((result && result.error) || "Could not upload " + type);
      }
      part.forEach(row => { sheetSnapshot.rows[type][String(row.id)] = sheetSignature_(type, row); });
      persistSheetSnapshot();
    }
  }
  const gst = await postSheet({
    action: "upsertRow",
    targetSheet: "Settings",
    rowId: "gstRate",
    rowData: ["gstRate", String(data.gstRate ?? 5)]
  });
  if (!gst || gst.success === false) throw new Error((gst && gst.error) || "Could not save GST rate");
  sheetSnapshot.gstRate = Number(data.gstRate ?? 5);
  persistSheetSnapshot();
}

async function flushSheet() {
  if (!sheetSnapshot || !sheetConfigured() || typeof data === "undefined") return false;
  ensureSheetIds();
  const ops = [];
  Object.keys(SHEET_TABS).forEach(type => {
    const seen = new Set();
    (data[type] || []).forEach(row => {
      if (!row || !row.id) return;
      const id = String(row.id);
      seen.add(id);
      const sig = sheetSignature_(type, row);
      const prev = sheetSnapshot.rows[type] && sheetSnapshot.rows[type][id];
      if (prev === sig) return;
      ops.push({
        type, id, sig, kind: prev == null ? "add" : "update",
        payload: {
          action: prev == null ? "addRow" : "updateRow",
          targetSheet: SHEET_TABS[type].sheet,
          rowId: id,
          rowData: toSheetRow_(type, row)
        }
      });
    });
    Object.keys(sheetSnapshot.rows[type] || {}).forEach(id => {
      if (seen.has(id)) return;
      pendingSheetDeletes.add(type + ":" + id);
      ops.push({
        type, id, kind: "delete",
        payload: { action: "deleteRow", targetSheet: SHEET_TABS[type].sheet, rowId: id }
      });
    });
  });
  if (Number(sheetSnapshot.gstRate) !== Number(data.gstRate ?? 5)) {
    ops.push({
      type: "settings", id: "gstRate", kind: "gst",
      payload: { action: "upsertRow", targetSheet: "Settings", rowId: "gstRate", rowData: ["gstRate", String(data.gstRate ?? 5)] }
    });
  }
  if (!ops.length) return false;

  const outer = isSheetSyncing;
  if (!outer) { isSheetSyncing = true; updateSyncStatus(); }
  let wrote = false;
  function acceptOp_(op) {
    wrote = true;
    lastSyncFailed = false;
    if (op.kind === "delete") {
      delete sheetSnapshot.rows[op.type][op.id];
      pendingSheetDeletes.delete(op.type + ":" + op.id);
    } else if (op.kind === "gst") {
      sheetSnapshot.gstRate = Number(data.gstRate ?? 5);
    } else {
      if (!sheetSnapshot.rows[op.type]) sheetSnapshot.rows[op.type] = {};
      sheetSnapshot.rows[op.type][op.id] = op.sig;
      pendingSheetDeletes.delete(op.type + ":" + op.id);
    }
  }
  try {
    const adds = ops.filter(op => op.kind === "add");
    const singles = ops.filter(op => op.kind !== "add");
    const addsByType = {};
    adds.forEach(op => { (addsByType[op.type] || (addsByType[op.type] = [])).push(op); });
    const addTypes = Object.keys(addsByType);
    for (let t = 0; t < addTypes.length; t++) {
      const group = addsByType[addTypes[t]];
      for (let i = 0; i < group.length; i += 100) {
        const part = group.slice(i, i + 100);
        const result = await postSheet({
          action: "addRows",
          targetSheet: SHEET_TABS[addTypes[t]].sheet,
          rows: part.map(op => op.payload.rowData)
        });
        if (!result || result.success === false) {
          console.error("Sheet write failed:", result && result.error);
          lastSyncFailed = true;
          continue;
        }
        part.forEach(acceptOp_);
      }
    }
    for (let i = 0; i < singles.length; i++) {
      const op = singles[i];
      let result = await postSheet(op.payload);
      if (op.kind === "update" && result && /row not found/i.test(String(result.error || ""))) {
        result = await postSheet(Object.assign({}, op.payload, { action: "addRow" }));
      }
      if (!result || result.success === false) {
        console.error("Sheet write failed:", result && result.error);
        lastSyncFailed = true;
        continue;
      }
      acceptOp_(op);
    }
    persistSheetSnapshot();
    persistSalonLocal();
  } finally {
    if (!outer) { isSheetSyncing = false; updateSyncStatus(); }
  }
  if (wrote) queueSheetPull();
  return wrote;
}

function queueSheetPull() {
  if (sheetPullQueued) return;
  sheetPullQueued = true;
  setTimeout(() => {
    sheetPullQueued = false;
    enqueueSheet(() => pullSheet(false));
  }, 500);
}

function scheduleSheetPush() {
  if (!sheetSnapshot || !sheetConfigured()) return;
  if (sheetPushTimer) clearTimeout(sheetPushTimer);
  sheetPushTimer = setTimeout(() => {
    sheetPushTimer = null;
    enqueueSheet(() => flushSheet());
  }, 800);
}

function sheetErrorText_(status, raw, result) {
  const apiErr = result && (result.error || result.message) ? String(result.error || result.message) : "";
  if (result && /unauthor/i.test(apiErr)) return "Sheet rejected the passcode. Set the same passcode in Settings and in Code.gs, then deploy a new version.";
  if (status === 401 || status === 403 || /need access|sign in|forbidden/i.test(raw || "")) {
    return "Sheet web app is not public. In Apps Script: Deploy → Manage deployments → Edit → Who has access: Anyone → New version → Deploy.";
  }
  if (apiErr) return apiErr;
  return "Could not reach the sheet.";
}

async function pullSheet(showAlert) {
  const cfg = sheetConfig();
  if (!cfg.apiUrl) {
    lastSyncFailed = false;
    updateSyncStatus();
    if (showAlert && typeof toast === "function") toast("Add the Google Sheet web app URL in Settings.");
    return;
  }
  if (typeof data === "undefined") return;
  isSheetSyncing = true;
  updateSyncStatus();
  let syncOk = false;
  try {
    if (sheetPushTimer) {
      clearTimeout(sheetPushTimer);
      sheetPushTimer = null;
      await flushSheet();
    }
    ensureSheetIds();
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 90000);
    const response = await fetch(
      cfg.apiUrl + "?action=getData&password=" + encodeURIComponent(cfg.passcode),
      { signal: controller.signal, cache: "no-store" }
    );
    clearTimeout(timer);
    const raw = await response.text();
    let result;
    try { result = JSON.parse(raw); }
    catch (e) {
      lastSyncFailed = true;
      if (showAlert && typeof toast === "function") toast(sheetErrorText_(response.status, raw, null));
      return;
    }
    if (!result.success) {
      lastSyncFailed = true;
      if (showAlert && typeof toast === "function") toast(sheetErrorText_(response.status, raw, result));
      return;
    }
    const serverRows = {};
    let serverCount = 0;
    Object.keys(SHEET_TABS).forEach(type => {
      serverRows[type] = (result.data[type] || []).map(row => fromSheetRow_(type, row));
      serverCount += serverRows[type].length;
    });
    if (!sheetSnapshot && serverCount === 0) {
      const localCount = Object.keys(SHEET_TABS).reduce((sum, type) => sum + ((data[type] || []).length), 0);
      if (localCount) await seedSheet_();
      sheetSnapshot = snapshotFromData_();
      persistSheetSnapshot();
      persistSalonLocal();
      syncOk = true;
      lastServerSyncAt = Date.now();
      if (showAlert && typeof toast === "function") toast(localCount ? "This device is now the shared salon sheet." : "Sheet connected.");
      return;
    }
    mergeFromServer_(serverRows);
    applyGstFromServer_(result.data.settings);
    ensureSheetIds();
    const previous = sheetSnapshot;
    sheetSnapshot = emptySnapshot_();
    Object.keys(SHEET_TABS).forEach(type => {
      serverRows[type].forEach(row => {
        sheetSnapshot.rows[type][String(row.id)] = sheetSignature_(type, row);
      });
    });
    const rawGst = settingValue_(result.data.settings, "gstRate");
    if (rawGst == null || rawGst === "" || !Number.isFinite(Number(rawGst))) {
      sheetSnapshot.gstRate = previous && previous.gstRate != null ? Number(previous.gstRate) : Number(data.gstRate ?? 5);
    } else {
      sheetSnapshot.gstRate = Number(rawGst);
    }
    persistSheetSnapshot();
    persistSalonLocal();
    syncOk = true;
    lastServerSyncAt = Date.now();
    paintAfterSync();
    scheduleSheetPush();
  } catch (err) {
    console.error("Sheet sync failed:", err);
    lastSyncFailed = true;
    if (showAlert && typeof toast === "function") {
      toast(err && err.name === "AbortError" ? "Sync timed out. Deploy a new Apps Script version and try again." : "Sync failed. Check the connection.");
    }
  } finally {
    isSheetSyncing = false;
    lastSyncFailed = !syncOk;
    updateSyncStatus();
  }
}

function refreshFromSheet(showAlert) {
  return enqueueSheet(() => pullSheet(!!showAlert));
}

function startSheetSync() {
  stopSheetSync();
  loadSheetSnapshot();
  ensureSheetIds();
  persistSalonLocal();
  refreshFromSheet(false);
  sheetAutoTimer = setInterval(() => {
    if (localStorage.getItem("trimzDemoSignedIn") === "yes") refreshFromSheet(false);
  }, AUTO_SYNC_MS);
}

function stopSheetSync() {
  if (sheetAutoTimer) { clearInterval(sheetAutoTimer); sheetAutoTimer = null; }
}

function saveSheetConfig() {
  const urlField = document.getElementById("sheet-api-url");
  const passField = document.getElementById("sheet-passcode");
  const url = urlField ? urlField.value.trim() : "";
  const pass = passField ? passField.value : "";
  if (url && !/^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec\/?$/.test(url)) {
    if (typeof toast === "function") toast("Paste the web app URL that ends with /exec");
    return;
  }
  if (url) localStorage.setItem("trimzApiUrl", url.replace(/\/$/, ""));
  else localStorage.removeItem("trimzApiUrl");
  if (pass) localStorage.setItem("trimzSheetPasscode", pass);
  sheetSnapshot = null;
  try { localStorage.removeItem("trimzSheetSnapshot"); } catch (e) {}
  updateSyncStatus();
  if (typeof toast === "function") toast(url ? "Sheet saved. Syncing this device…" : "Sheet disconnected on this device.");
  if (url && localStorage.getItem("trimzDemoSignedIn") === "yes") startSheetSync();
}

if (typeof window !== "undefined") {
  setInterval(updateSyncStatus, 30000);
  window.addEventListener("online", () => {
    updateSyncStatus();
    if (localStorage.getItem("trimzDemoSignedIn") === "yes") refreshFromSheet(false);
  });
  window.addEventListener("offline", updateSyncStatus);
}

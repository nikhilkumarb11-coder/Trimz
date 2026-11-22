/**
 * TRIMZ Salon Desk — Google Apps Script backend
 *
 * Same pattern as the campus ERP (Code.gs): the website never signs in to
 * Google. This script is the only thing that reads and writes the sheet.
 *
 * SETUP
 * 1. Create a Google Sheet named TRIMZ (any blank sheet is fine).
 * 2. Extensions → Apps Script. Replace ALL code with this file. Save.
 * 3. Deploy → New deployment → Web app
 *      Execute as: Me
 *      Who has access: Anyone
 * 4. Copy the Web App URL (it ends with /exec).
 * 5. In Trimz, open Settings → Google Sheet sync, paste the URL.
 *    The passcode below must match the passcode in Settings.
 * 6. After every edit here: Deploy → Manage deployments → Edit (pencil)
 *    → Version: New version → Deploy. The URL stays the same.
 *
 * "Who has access" must be Anyone. Otherwise the page gets 403 and shows
 * "Not synced".
 *
 * Tabs are created on the first sync if they are missing:
 * Customers, Services, Staff, Inventory, Appointments, Memberships,
 * Packages, Expenses, Bills, Settings.
 * Column A is the record id. Row 1 is the header.
 */

const SCRIPT_PASSWORD = "trimz123";

const HEADERS = {
  Customers: ["id", "name", "phone", "gender", "visits", "spent", "last", "points", "notes", "marketingOptIn"],
  Services: ["id", "name", "category", "audience", "price", "priceNote", "duration", "commission", "source"],
  Staff: ["id", "name", "role", "sales", "commission", "status", "target", "attendance"],
  Inventory: ["id", "name", "brand", "qty", "price", "cost", "expiry", "supplier"],
  Appointments: ["id", "customer", "service", "staff", "date", "time", "status"],
  Memberships: ["id", "name", "customer", "phone", "discount", "valid", "status"],
  Packages: ["id", "name", "customer", "included", "remaining", "valid"],
  Expenses: ["id", "name", "category", "amount", "date", "mode"],
  Bills: ["id", "no", "customer", "phone", "items", "date", "subtotal", "discount", "tax", "gstIncluded", "gstRate", "total", "mode", "time", "status", "staff", "commission", "lines"],
  Settings: ["key", "value"]
};

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function assertPassword_(password) {
  return String(password || "") === SCRIPT_PASSWORD;
}

function isDateValue_(value) {
  return Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime());
}

// Time cells must stay clock times. Turning every Date into yyyy-mm-dd made
// "10:45 pm" come back as 1899-12-30, the blank date Sheets uses for times.
function formatSheetValue_(header, value) {
  if (!isDateValue_(value)) return value;
  const tz = Session.getScriptTimeZone() || "Asia/Kolkata";
  if (header === "time") return Utilities.formatDate(value, tz, "h:mm a");
  return Utilities.formatDate(value, tz, "yyyy-MM-dd");
}

function sheetToObjects_(sheetName) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h || "").trim());
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every(c => c === "" || c === null)) continue;
    const obj = {};
    headers.forEach((h, idx) => {
      if (!h) return;
      obj[h] = formatSheetValue_(h, row[idx]);
    });
    rows.push(obj);
  }
  return rows;
}

function textColumnIndexes_(sh, width) {
  const headers = HEADERS[sh.getName()] || [];
  const cols = [];
  headers.forEach((header, i) => {
    if (i < width && (header === "date" || header === "time" || header === "valid" || header === "expiry")) cols.push(i + 1);
  });
  return cols;
}

function protectTextColumns_(sh, rowData) {
  const headers = HEADERS[sh.getName()] || [];
  return rowData.map((value, i) => {
    const header = headers[i];
    if (header !== "date" && header !== "time" && header !== "valid" && header !== "expiry") return value == null ? "" : value;
    if (value == null || value === "") return "";
    const text = String(value);
    return text.charAt(0) === "'" ? text : "'" + text;
  });
}

function sheetByName_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const headers = HEADERS[name];
  if (headers && sh.getLastRow() === 0) sh.appendRow(headers);
  return sh;
}

function doGet(e) {
  try {
    e = e || { parameter: {} };
    if (!assertPassword_(e.parameter.password)) {
      return jsonOut_({ success: false, error: "Unauthorized access token." });
    }
    const action = e.parameter.action || "getData";
    if (action === "getData") {
      return jsonOut_({
        success: true,
        data: {
          customers: sheetToObjects_("Customers"),
          services: sheetToObjects_("Services"),
          staff: sheetToObjects_("Staff"),
          inventory: sheetToObjects_("Inventory"),
          appointments: sheetToObjects_("Appointments"),
          memberships: sheetToObjects_("Memberships"),
          packages: sheetToObjects_("Packages"),
          expenses: sheetToObjects_("Expenses"),
          bills: sheetToObjects_("Bills"),
          settings: sheetToObjects_("Settings")
        }
      });
    }
    return jsonOut_({ success: false, error: "Unknown action: " + action });
  } catch (err) {
    return jsonOut_({ success: false, error: String(err && err.message ? err.message : err) });
  }
}

function findRow_(sh, rowId) {
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][0] == null ? "" : data[r][0]).trim() === rowId) return r + 1;
  }
  return 0;
}

function writeRow_(sh, rowNumber, rowData) {
  if (!rowData || !rowData.length) return;
  const data = protectTextColumns_(sh, rowData);
  textColumnIndexes_(sh, data.length).forEach(col => sh.getRange(rowNumber, col).setNumberFormat("@"));
  sh.getRange(rowNumber, 1, 1, data.length).setValues([data]);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(20000)) {
      return jsonOut_({ success: false, error: "Sheet is busy. Try again." });
    }
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (!assertPassword_(body.password)) {
      return jsonOut_({ success: false, error: "Unauthorized access token." });
    }
    const sh = sheetByName_(body.targetSheet);
    if (!sh) return jsonOut_({ success: false, error: "Sheet not found: " + body.targetSheet });

    if (body.action === "addRow") {
      writeRow_(sh, Math.max(sh.getLastRow(), 1) + 1, body.rowData || []);
      return jsonOut_({ success: true });
    }

    if (body.action === "addRows") {
      const rows = Array.isArray(body.rows) ? body.rows : [];
      if (!rows.length) return jsonOut_({ success: true, count: 0 });
      const width = rows[0].length;
      const start = Math.max(sh.getLastRow(), 1) + 1;
      const prepared = rows.map(row => protectTextColumns_(sh, row));
      textColumnIndexes_(sh, width).forEach(col => sh.getRange(start, col, rows.length, 1).setNumberFormat("@"));
      sh.getRange(start, 1, rows.length, width).setValues(prepared);
      return jsonOut_({ success: true, count: rows.length });
    }

    if (body.action === "updateRow" || body.action === "upsertRow") {
      const rowId = String(body.rowId || (body.rowData && body.rowData[0]) || "").trim();
      const rowNumber = findRow_(sh, rowId);
      if (!rowNumber) {
        if (body.action === "upsertRow") {
          writeRow_(sh, Math.max(sh.getLastRow(), 1) + 1, body.rowData || []);
          return jsonOut_({ success: true });
        }
        return jsonOut_({ success: false, error: "Row not found: " + rowId });
      }
      writeRow_(sh, rowNumber, body.rowData || []);
      return jsonOut_({ success: true });
    }

    if (body.action === "deleteRow") {
      const rowId = String(body.rowId || "").trim();
      const rowNumber = findRow_(sh, rowId);
      if (!rowNumber) return jsonOut_({ success: true });
      sh.deleteRow(rowNumber);
      return jsonOut_({ success: true });
    }

    return jsonOut_({ success: false, error: "Unknown action: " + body.action });
  } catch (err) {
    return jsonOut_({ success: false, error: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

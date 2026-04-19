/**
 * Spin & Win Google Apps Script Backend
 * Supports:
 * - GET    ?action=list
 * - GET    ?action=settings
 * - GET    ?action=record&recordId=PRZ-...
 * - POST   create
 * - POST   { action: "save_settings", settings: {...} }
 * - PUT    update (or POST + _method: "PUT")
 * - DELETE delete (or POST + _method: "DELETE")
 */

const SHEET_NAME = "SpinWinData";
const SETTINGS_SHEET_NAME = "SpinWinSettings";
const HEADERS = [
  "RecordID",
  "ShopName",
  "CustomerName",
  "CustomerNumber",
  "Amount",
  "Prize",
  "Status",
  "DateTimeISO",
  "DateTimeDisplay",
  "ExpiryISO",
  "ExpiryDisplay",
  "CreatedAtISO",
  "UpdatedAtISO",
  "Source",
  "SavedAtServer"
];
const SETTINGS_HEADERS = ["Key", "ValueJSON", "UpdatedAtISO"];
const SETTINGS_KEY = "app_settings";

function doGet(e) {
  const callback = String((e && e.parameter && e.parameter.callback) || "").trim();
  try {
    const action = String((e && e.parameter && e.parameter.action) || "").toLowerCase();
    const recordIdParam = String(
      (e && e.parameter && (e.parameter.recordId || e.parameter.recordid || e.parameter.id)) || ""
    ).trim();

    if (action === "health") {
      return outputResponse_({
        ok: true,
        message: "Spin & Win backend is running.",
        timestamp: new Date().toISOString()
      }, callback);
    }

    if (action === "" || action === "list") {
      const q = String((e && e.parameter && e.parameter.q) || "").toLowerCase();
      const records = listRecords_().filter((record) => {
        if (!q) return true;
        return String(record.recordId || "").toLowerCase().indexOf(q) !== -1;
      });

      return outputResponse_({
        ok: true,
        message: "Records loaded.",
        count: records.length,
        records: records
      }, callback);
    }

    if (action === "settings") {
      return outputResponse_({
        ok: true,
        message: "Settings loaded.",
        settings: loadSettings_()
      }, callback);
    }

    if (
      action === "record" ||
      action === "get_record" ||
      action === "coupon" ||
      action === "details" ||
      action === "lookup" ||
      (recordIdParam && action === "")
    ) {
      const recordId = recordIdParam;
      if (!recordId) {
        return outputResponse_({
          ok: false,
          message: "recordId is required."
        }, callback);
      }

      const record = getRecordById_(recordId);
      if (!record) {
        return outputResponse_({
          ok: false,
          message: "Record not found.",
          recordId: recordId
        }, callback);
      }

      return outputResponse_({
        ok: true,
        message: "Record loaded.",
        recordId: recordId,
        effectiveStatus: getEffectiveStatus_(record),
        record: record
      }, callback);
    }

    return outputResponse_({
      ok: false,
      message: "Unknown GET action.",
      action: action
    }, callback);
  } catch (err) {
    return outputResponse_({
      ok: false,
      message: err.message || "Unexpected error"
    }, callback);
  }
}

function doPost(e) {
  try {
    const payload = parsePayload_(e);
    const method = String(payload._method || payload.method || "POST").toUpperCase();
    const action = String(payload.action || "").toLowerCase();

    if (action === "save_settings" || action === "settings") {
      return saveSettings_(payload);
    }

    if (method === "PUT" || action === "update") {
      return updateRecord_(payload);
    }
    if (method === "DELETE" || action === "delete") {
      return deleteRecord_(payload);
    }

    return createRecord_(payload);
  } catch (err) {
    return jsonOutput({
      ok: false,
      message: err.message || "Unexpected error"
    });
  }
}

function createRecord_(payload) {
  const record = normalizeRecord_(payload.record || payload);
  const required = ["recordId", "customerName", "prize", "dateTimeIso", "expiryIso"];
  const missing = required.filter((key) => !record[key]);

  if (missing.length) {
    return jsonOutput({
      ok: false,
      message: "Missing required fields",
      missing: missing
    });
  }

  const sheet = getOrCreateSheet_(SHEET_NAME);
  const headerMap = ensureHeaders_(sheet);
  const found = findRowByRecordId_(sheet, headerMap, record.recordId);

  if (found.rowNumber > 0) {
    const merged = mergeRecord_(found.record, record);
    sheet.getRange(found.rowNumber, 1, 1, HEADERS.length).setValues([recordToRow_(merged)]);
    return jsonOutput({
      ok: true,
      message: "Record already exists. Updated existing row.",
      recordId: record.recordId
    });
  }

  sheet.appendRow(recordToRow_(record));
  return jsonOutput({
    ok: true,
    message: "Record created.",
    recordId: record.recordId
  });
}

function updateRecord_(payload) {
  const recordId = String(payload.recordId || (payload.record && payload.record.recordId) || "").trim();
  if (!recordId) {
    return jsonOutput({
      ok: false,
      message: "recordId is required for update."
    });
  }

  const updates = normalizePartialRecord_(payload.updates || payload.record || payload);
  const sheet = getOrCreateSheet_(SHEET_NAME);
  const headerMap = ensureHeaders_(sheet);
  const found = findRowByRecordId_(sheet, headerMap, recordId);

  if (found.rowNumber < 1) {
    return jsonOutput({
      ok: false,
      message: "Record not found.",
      recordId: recordId
    });
  }

  const merged = mergeRecord_(found.record, updates);
  merged.recordId = recordId;
  merged.savedAtServer = new Date().toISOString();

  sheet.getRange(found.rowNumber, 1, 1, HEADERS.length).setValues([recordToRow_(merged)]);
  return jsonOutput({
    ok: true,
    message: "Record updated.",
    recordId: recordId
  });
}

function deleteRecord_(payload) {
  const recordId = String(payload.recordId || "").trim();
  if (!recordId) {
    return jsonOutput({
      ok: false,
      message: "recordId is required for delete."
    });
  }

  const sheet = getOrCreateSheet_(SHEET_NAME);
  const headerMap = ensureHeaders_(sheet);
  const found = findRowByRecordId_(sheet, headerMap, recordId);

  if (found.rowNumber < 1) {
    return jsonOutput({
      ok: false,
      message: "Record not found.",
      recordId: recordId
    });
  }

  sheet.deleteRow(found.rowNumber);
  return jsonOutput({
    ok: true,
    message: "Record deleted.",
    recordId: recordId
  });
}

function listRecords_() {
  const sheet = getOrCreateSheet_(SHEET_NAME);
  const headerMap = ensureHeaders_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const records = rows.map((row) => rowToRecord_(row, headerMap));

  records.sort(function(a, b) {
    return new Date(b.dateTimeIso).getTime() - new Date(a.dateTimeIso).getTime();
  });
  return records;
}

function getRecordById_(recordId) {
  const sheet = getOrCreateSheet_(SHEET_NAME);
  const headerMap = ensureHeaders_(sheet);
  const found = findRowByRecordId_(sheet, headerMap, recordId);
  if (found.rowNumber < 1 || !found.record) return null;
  return found.record;
}

function saveSettings_(payload) {
  const incoming = payload.settings || payload;
  const settings = normalizeSettings_(incoming);
  const sheet = getOrCreateSheet_(SETTINGS_SHEET_NAME);
  const headerMap = ensureSettingsHeaders_(sheet);
  const found = findSettingsRow_(sheet, headerMap, SETTINGS_KEY);
  const updatedAtIso = new Date().toISOString();
  const row = [SETTINGS_KEY, JSON.stringify(settings), updatedAtIso];

  if (found.rowNumber > 0) {
    sheet.getRange(found.rowNumber, 1, 1, SETTINGS_HEADERS.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  return jsonOutput({
    ok: true,
    message: "Settings saved.",
    settings: settings,
    updatedAtIso: updatedAtIso
  });
}

function loadSettings_() {
  const defaults = defaultSettings_();
  const sheet = getOrCreateSheet_(SETTINGS_SHEET_NAME);
  const headerMap = ensureSettingsHeaders_(sheet);
  const found = findSettingsRow_(sheet, headerMap, SETTINGS_KEY);
  if (found.rowNumber < 1 || !found.row) {
    return defaults;
  }

  const raw = String(found.row[headerMap.ValueJSON] || "").trim();
  if (!raw) return defaults;

  try {
    const parsed = JSON.parse(raw);
    return normalizeSettings_(parsed);
  } catch (err) {
    return defaults;
  }
}

function ensureSettingsHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(SETTINGS_HEADERS);
  } else {
    const existing = sheet.getRange(1, 1, 1, SETTINGS_HEADERS.length).getValues()[0];
    const mismatch = SETTINGS_HEADERS.some(function(header, idx) {
      return existing[idx] !== header;
    });
    if (mismatch) {
      sheet.getRange(1, 1, 1, SETTINGS_HEADERS.length).setValues([SETTINGS_HEADERS]);
    }
  }

  const map = {};
  SETTINGS_HEADERS.forEach(function(header, idx) {
    map[header] = idx;
  });
  return map;
}

function findSettingsRow_(sheet, headerMap, key) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rowNumber: -1, row: null };

  const rows = sheet.getRange(2, 1, lastRow - 1, SETTINGS_HEADERS.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rowKey = String(row[headerMap.Key] || "");
    if (rowKey === key) {
      return { rowNumber: i + 2, row: row };
    }
  }
  return { rowNumber: -1, row: null };
}

function normalizeSettings_(settings) {
  settings = settings || {};
  var rawPrizes = Array.isArray(settings.prizes) ? settings.prizes : [];
  var prizes = rawPrizes.map(function(prize, idx) {
    prize = prize || {};
    return {
      id: String(prize.id || ("prize_" + (idx + 1))),
      name: String(prize.name || "Unnamed Prize").trim(),
      probability: Math.max(0, toNumber_(prize.probability, 0)),
      enabled: prize.enabled !== false
    };
  });

  return {
    shopName: String(settings.shopName || "Lucky Shop").trim() || "Lucky Shop",
    shopLogoUrl: String(settings.shopLogoUrl || "").trim(),
    expiryHours: Math.max(0, toNumber_(settings.expiryHours, 24)),
    manualDateEnabled: settings.manualDateEnabled === true || String(settings.manualDateEnabled).toLowerCase() === "true",
    manualDateTime: String(settings.manualDateTime || ""),
    appsScriptUrl: String(settings.appsScriptUrl || "").trim(),
    prizes: prizes.length ? prizes : defaultSettings_().prizes
  };
}

function defaultSettings_() {
  return {
    shopName: "Lucky Shop",
    shopLogoUrl: "",
    expiryHours: 24,
    manualDateEnabled: false,
    manualDateTime: "",
    appsScriptUrl: "",
    prizes: [
      { id: "p1", name: "10% Discount", probability: 40, enabled: true },
      { id: "p2", name: "Free Drink", probability: 25, enabled: true },
      { id: "p3", name: "Buy 1 Get 1", probability: 10, enabled: true },
      { id: "p4", name: "Try Again", probability: 24.9, enabled: true },
      { id: "p5", name: "Grand Prize", probability: 0.1, enabled: true }
    ]
  };
}

function toNumber_(value, fallback) {
  var parsed = Number(value);
  return isNaN(parsed) ? fallback : parsed;
}

function parsePayload_(e) {
  if (!e) return {};

  if (e.postData && e.postData.contents) {
    const raw = e.postData.contents;
    try {
      return JSON.parse(raw);
    } catch (err) {
      return { raw: raw };
    }
  }
  return e.parameter || {};
}

function getOrCreateSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  } else {
    const existing = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
    const mismatch = HEADERS.some(function(header, idx) {
      return existing[idx] !== header;
    });
    if (mismatch) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    }
  }

  const map = {};
  HEADERS.forEach(function(header, idx) {
    map[header] = idx;
  });
  return map;
}

function findRowByRecordId_(sheet, headerMap, recordId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rowNumber: -1, record: null };

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    const row = rows[i];
    const id = String(row[headerMap.RecordID] || "");
    if (id === recordId) {
      return {
        rowNumber: i + 2,
        record: rowToRecord_(row, headerMap)
      };
    }
  }
  return { rowNumber: -1, record: null };
}

function rowToRecord_(row, headerMap) {
  return {
    recordId: String(row[headerMap.RecordID] || ""),
    shopName: String(row[headerMap.ShopName] || ""),
    customerName: String(row[headerMap.CustomerName] || ""),
    customerNumber: String(row[headerMap.CustomerNumber] || ""),
    amount: Number(row[headerMap.Amount] || 0),
    prize: String(row[headerMap.Prize] || ""),
    status: String(row[headerMap.Status] || "Pending"),
    dateTimeIso: String(row[headerMap.DateTimeISO] || ""),
    dateTimeDisplay: String(row[headerMap.DateTimeDisplay] || ""),
    expiryIso: String(row[headerMap.ExpiryISO] || ""),
    expiryDisplay: String(row[headerMap.ExpiryDisplay] || ""),
    createdAtIso: String(row[headerMap.CreatedAtISO] || ""),
    updatedAtIso: String(row[headerMap.UpdatedAtISO] || ""),
    source: String(row[headerMap.Source] || ""),
    savedAtServer: String(row[headerMap.SavedAtServer] || "")
  };
}

function recordToRow_(record) {
  return [
    record.recordId || "",
    record.shopName || "",
    record.customerName || "",
    record.customerNumber || "",
    Number(record.amount || 0),
    record.prize || "",
    normalizeStatus_(record.status),
    record.dateTimeIso || "",
    record.dateTimeDisplay || formatDate_(record.dateTimeIso),
    record.expiryIso || "",
    record.expiryDisplay || formatDate_(record.expiryIso),
    record.createdAtIso || new Date().toISOString(),
    record.updatedAtIso || new Date().toISOString(),
    record.source || "SpinWinFrontend",
    record.savedAtServer || new Date().toISOString()
  ];
}

function normalizeRecord_(record) {
  record = record || {};
  return {
    recordId: String(record.recordId || "").trim(),
    shopName: String(record.shopName || "").trim(),
    customerName: String(record.customerName || "").trim(),
    customerNumber: String(record.customerNumber || "").trim(),
    amount: Number(record.amount || 0),
    prize: String(record.prize || "").trim(),
    status: normalizeStatus_(record.status),
    dateTimeIso: toIso_(record.dateTimeIso || record.dateTime || new Date()),
    dateTimeDisplay: String(record.dateTimeDisplay || ""),
    expiryIso: toIso_(record.expiryIso || record.expiryDate || new Date()),
    expiryDisplay: String(record.expiryDisplay || ""),
    createdAtIso: toIso_(record.createdAtIso || new Date()),
    updatedAtIso: toIso_(record.updatedAtIso || new Date()),
    source: String(record.source || "SpinWinFrontend"),
    savedAtServer: toIso_(new Date())
  };
}

function mergeRecord_(original, updates) {
  original = original || {};
  updates = updates || {};
  return {
    recordId: pickValue_(updates.recordId, original.recordId, ""),
    shopName: pickValue_(updates.shopName, original.shopName, ""),
    customerName: pickValue_(updates.customerName, original.customerName, ""),
    customerNumber: pickValue_(updates.customerNumber, original.customerNumber, ""),
    amount: updates.amount !== undefined ? Number(updates.amount) : Number(original.amount || 0),
    prize: pickValue_(updates.prize, original.prize, ""),
    status: normalizeStatus_(pickValue_(updates.status, original.status, "Pending")),
    dateTimeIso: toIso_(pickValue_(updates.dateTimeIso, original.dateTimeIso, new Date())),
    dateTimeDisplay: pickValue_(updates.dateTimeDisplay, original.dateTimeDisplay, formatDate_(pickValue_(updates.dateTimeIso, original.dateTimeIso, new Date()))),
    expiryIso: toIso_(pickValue_(updates.expiryIso, original.expiryIso, new Date())),
    expiryDisplay: pickValue_(updates.expiryDisplay, original.expiryDisplay, formatDate_(pickValue_(updates.expiryIso, original.expiryIso, new Date()))),
    createdAtIso: toIso_(pickValue_(original.createdAtIso, updates.createdAtIso, new Date())),
    updatedAtIso: toIso_(pickValue_(updates.updatedAtIso, new Date(), new Date())),
    source: pickValue_(updates.source, original.source, "SpinWinFrontend"),
    savedAtServer: toIso_(new Date())
  };
}

function normalizePartialRecord_(record) {
  record = record || {};
  const partial = {};

  if (record.recordId !== undefined) partial.recordId = String(record.recordId).trim();
  if (record.shopName !== undefined) partial.shopName = String(record.shopName).trim();
  if (record.customerName !== undefined) partial.customerName = String(record.customerName).trim();
  if (record.customerNumber !== undefined) partial.customerNumber = String(record.customerNumber).trim();
  if (record.amount !== undefined) partial.amount = Number(record.amount || 0);
  if (record.prize !== undefined) partial.prize = String(record.prize).trim();
  if (record.status !== undefined) partial.status = normalizeStatus_(record.status);
  if (record.dateTimeIso !== undefined || record.dateTime !== undefined) {
    partial.dateTimeIso = toIso_(record.dateTimeIso || record.dateTime);
  }
  if (record.dateTimeDisplay !== undefined) partial.dateTimeDisplay = String(record.dateTimeDisplay);
  if (record.expiryIso !== undefined || record.expiryDate !== undefined) {
    partial.expiryIso = toIso_(record.expiryIso || record.expiryDate);
  }
  if (record.expiryDisplay !== undefined) partial.expiryDisplay = String(record.expiryDisplay);
  if (record.createdAtIso !== undefined) partial.createdAtIso = toIso_(record.createdAtIso);
  if (record.updatedAtIso !== undefined) partial.updatedAtIso = toIso_(record.updatedAtIso);
  if (record.source !== undefined) partial.source = String(record.source);

  return partial;
}

function normalizeStatus_(status) {
  const value = String(status || "Pending").toLowerCase();
  if (value === "approved") return "Completed";
  if (value === "completed") return "Completed";
  if (value === "rejected") return "Rejected";
  if (value === "expired") return "Expired";
  return "Pending";
}

function getEffectiveStatus_(record) {
  const manual = normalizeStatus_(record && record.status);
  if (manual === "Completed" || manual === "Rejected" || manual === "Expired") {
    return manual;
  }

  const expiry = new Date(record && record.expiryIso).getTime();
  if (!isNaN(expiry) && expiry < Date.now()) {
    return "Expired";
  }
  return "Pending";
}

function pickValue_(...values) {
  for (var i = 0; i < values.length; i++) {
    if (values[i] !== undefined && values[i] !== null && values[i] !== "") {
      return values[i];
    }
  }
  return "";
}

function toIso_(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return new Date().toISOString();
  return date.toISOString();
}

function formatDate_(value) {
  const date = new Date(value);
  if (isNaN(date.getTime())) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

function jsonOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function outputResponse_(data, callback) {
  if (!callback) return jsonOutput(data);

  const safeCallback = callback.replace(/[^0-9A-Za-z_.$]/g, "");
  const payload = safeCallback + "(" + JSON.stringify(data) + ");";
  return ContentService
    .createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

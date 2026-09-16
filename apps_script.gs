/**
 * Spin & Win Google Apps Script Backend
 * 
 * Supports:
 * - GET    ?action=list                  (List prize voucher records)
 * - GET    ?action=settings              (Load wheel settings & multi-conditions)
 * - GET    ?action=record&recordId=...   (Lookup single voucher record)
 * - GET    ?action=members               (List all Members Club customers & balances)
 * - GET    ?action=member_history&id=... (List ledger transaction history)
 * - POST   { action: "create", record }  (Create/update prize voucher record)
 * - POST   { action: "save_settings" }   (Save wheel settings & multi-conditions)
 * - POST   { action: "save_member" }     (Create or update a loyalty member)
 * - POST   { action: "delete_member" }   (Delete a loyalty member)
 * - POST   { action: "add_credits" }     (Record purchase & credit loyalty points)
 * - POST   { action: "deduct_credits" }  (Redeem bill discount or cash payout)
 * - POST   { action: "sync_members" }    (Batch sync all members & ledger)
 * - PUT    update (or POST + _method: "PUT")
 * - DELETE delete (or POST + _method: "DELETE")
 */

const SHEET_NAME = "SpinWinData";
const SETTINGS_SHEET_NAME = "SpinWinSettings";
const MEMBERS_SHEET_NAME = "SpinWinMembers";
const LEDGER_SHEET_NAME = "SpinWinLedger";

const HEADERS = [
  "RecordID",
  "ShopName",
  "CustomerName",
  "CustomerNumber",
  "Amount",
  "Prize",
  "PurchasedItem",
  "ConditionID",
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

const MEMBERS_HEADERS = [
  "MemberID",
  "CustomerName",
  "Phone",
  "MemberType",
  "Credits",
  "CashWorth",
  "TotalEarned",
  "TotalRedeemed",
  "TotalCashPaid",
  "Notes",
  "CreatedAtISO",
  "UpdatedAtISO",
  "Source",
  "SavedAtServer"
];

const LEDGER_HEADERS = [
  "TxID",
  "MemberID",
  "CustomerName",
  "Phone",
  "Type",
  "Credits",
  "CashValue",
  "PurchaseAmount",
  "BalanceAfter",
  "Note",
  "DateISO",
  "SavedAtServer"
];

// 100 Credits = ₹1.00
const CREDIT_RATE = 0.01;

/**
 * Runs automatically when the Google Sheet is opened.
 * Adds a custom menu to initialize or format all sheets.
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("🎰 Spin & Win")
      .addItem("🚀 Initialize / Verify All Sheets", "initAllSheets")
      .addItem("👥 Setup Members & Ledger Sheets", "initMembersAndLedgerSheets")
      .addItem("🧹 Find Duplicate Tabs", "findDuplicateSheets")
      .addItem("🔍 Check for Duplicate Rows", "findDuplicateRows")
      .addToUi();
  } catch (err) {
    // Silent catch if running headless or via Web App
  }
}

/**
 * Initializes all required sheets with proper headers, styling, and frozen rows.
 */
function initAllSheets() {
  initMembersAndLedgerSheets();
  const vSheet = getOrCreateSheet_(SHEET_NAME);
  ensureHeaders_(vSheet, HEADERS);
  const sSheet = getOrCreateSheet_(SETTINGS_SHEET_NAME);
  ensureHeaders_(sSheet, SETTINGS_HEADERS);
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast("All Spin & Win sheets initialized successfully!", "🎰 Setup Complete", 5);
  } catch (e) {}
}

/**
 * Specifically ensures SpinWinMembers and SpinWinLedger sheets are created with styled headers.
 */
function initMembersAndLedgerSheets() {
  const mSheet = getOrCreateSheet_(MEMBERS_SHEET_NAME);
  ensureHeaders_(mSheet, MEMBERS_HEADERS);
  const lSheet = getOrCreateSheet_(LEDGER_SHEET_NAME);
  ensureHeaders_(lSheet, LEDGER_HEADERS);
}

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

    // 1. Members Club: List all members
    if (action === "members" || action === "list_members") {
      const q = String((e && e.parameter && e.parameter.q) || "").toLowerCase();
      const members = listMembers_().filter(function(m) {
        if (!q) return true;
        const nameMatch = String(m.name || "").toLowerCase().indexOf(q) !== -1;
        const phoneMatch = String(m.phone || "").indexOf(q) !== -1;
        const idMatch = String(m.id || "").toLowerCase().indexOf(q) !== -1;
        return nameMatch || phoneMatch || idMatch;
      });

      return outputResponse_({
        ok: true,
        message: "Members loaded.",
        count: members.length,
        members: members
      }, callback);
    }

    // 2. Members Club: List ledger history
    if (action === "member_history" || action === "ledger") {
      const memberId = String((e && e.parameter && (e.parameter.memberId || e.parameter.id)) || "").trim();
      const history = listLedger_(memberId);
      return outputResponse_({
        ok: true,
        message: "Ledger history loaded.",
        count: history.length,
        history: history
      }, callback);
    }

    // 3. Wheel Settings
    if (action === "settings") {
      return outputResponse_({
        ok: true,
        message: "Settings loaded.",
        settings: loadSettings_()
      }, callback);
    }

    // 3.5 Public live coupon status page (what a scanned QR opens).
    // Returns a human-friendly HTML page instead of raw JSON.
    if (
      action === "coupon_page" || action === "couponpage" ||
      action === "coupon" || action === "voucher" ||
      action === "scan" || action === "qr" || action === "status"
    ) {
      const recordId = recordIdParam;
      if (!recordId) {
        return couponOutput_(couponPageHtml_(null, "", "recordId parameter is required."));
      }
      const record = getRecordById_(recordId);
      return couponOutput_(couponPageHtml_(record, recordId, ""));
    }

    // Also allow the JSON record lookup to render as the HTML page when the
    // caller explicitly asks for a page view (e.g. manually built link).
    if (
      action === "get_record" && String((e && e.parameter && (e.parameter.view || e.parameter.format)) || "").toLowerCase() === "html"
    ) {
      const recordId = recordIdParam;
      const record = recordId ? getRecordById_(recordId) : null;
      return couponOutput_(couponPageHtml_(record, recordId || "", ""));
    }

    // 4. Single Voucher Record Lookup
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

    // 5. Default: List Prize Voucher Records
    if (action === "" || action === "list") {
      const q = String((e && e.parameter && e.parameter.q) || "").toLowerCase();
      const records = listRecords_().filter(function(record) {
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

    // 1. Settings save
    if (action === "save_settings" || action === "settings") {
      return saveSettings_(payload);
    }

    // 2. Members Club actions
    if (action === "save_member" || action === "create_member" || action === "update_member") {
      return saveMember_(payload);
    }
    if (action === "delete_member") {
      return deleteMember_(payload);
    }
    if (action === "add_credits") {
      return addMemberCredits_(payload);
    }
    if (action === "deduct_credits") {
      return deductMemberCredits_(payload);
    }
    if (action === "sync_members") {
      return syncMembersBatch_(payload);
    }

    // 3. Prize Voucher Record updates/deletions
    if (method === "PUT" || action === "update") {
      return updateRecord_(payload);
    }
    if (method === "DELETE" || action === "delete") {
      return deleteRecord_(payload);
    }

    // 4. Default: Create Prize Voucher Record
    return createRecord_(payload);
  } catch (err) {
    return jsonOutput({
      ok: false,
      message: err.message || "Unexpected error"
    });
  }
}

// ==========================================
// Prize Records Operations (SpinWinData)
// ==========================================

function createRecord_(payload) {
  const record = normalizeRecord_(payload.record || payload);
  const required = ["recordId", "customerName", "prize", "dateTimeIso", "expiryIso"];
  const missing = required.filter(function(key) { return !record[key]; });

  if (missing.length) {
    return jsonOutput({
      ok: false,
      message: "Missing required fields",
      missing: missing
    });
  }

  const sheet = getOrCreateSheet_(SHEET_NAME);
  const headerMap = ensureHeaders_(sheet, HEADERS);
  const found = findRowByRecordId_(sheet, headerMap, record.recordId);

  if (found.rowNumber > 0) {
    const merged = mergeRecord_(found.record, record);
    sheet.getRange(found.rowNumber, 1, 1, HEADERS.length).setValues([recordToRow_(merged, headerMap)]);
    return jsonOutput({
      ok: true,
      message: "Record already exists. Updated existing row.",
      recordId: record.recordId
    });
  }

  appendRowFast_(sheet, recordToRow_(record, headerMap));
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
  const headerMap = ensureHeaders_(sheet, HEADERS);
  const found = findRowByRecordId_(sheet, headerMap, recordId);

  if (found.rowNumber < 1) {
    return jsonOutput({
      ok: false,
      message: "Record not found.",
      recordId: recordId
    });
  }

  const merged = mergeRecord_(found.record, updates);
  sheet.getRange(found.rowNumber, 1, 1, HEADERS.length).setValues([recordToRow_(merged, headerMap)]);

  return jsonOutput({
    ok: true,
    message: "Record updated.",
    recordId: recordId,
    record: merged
  });
}

function deleteRecord_(payload) {
  const recordId = String(payload.recordId || (payload.record && payload.record.recordId) || "").trim();
  if (!recordId) {
    return jsonOutput({
      ok: false,
      message: "recordId is required for deletion."
    });
  }

  const sheet = getOrCreateSheet_(SHEET_NAME);
  const headerMap = ensureHeaders_(sheet, HEADERS);
  const found = findRowByRecordId_(sheet, headerMap, recordId);

  if (found.rowNumber < 1) {
    return jsonOutput({
      ok: false,
      message: "Record not found for delete.",
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
  const headerMap = ensureHeaders_(sheet, HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const lastCol = sheet.getLastColumn();
  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const records = rows.map(function(row) {
    return rowToRecord_(row, headerMap);
  });

  records.sort(function(a, b) {
    return new Date(b.dateTimeIso).getTime() - new Date(a.dateTimeIso).getTime();
  });
  return records;
}

function getRecordById_(recordId) {
  const sheet = getOrCreateSheet_(SHEET_NAME);
  const headerMap = ensureHeaders_(sheet, HEADERS);
  const found = findRowByRecordId_(sheet, headerMap, recordId);
  if (found.rowNumber < 1 || !found.record) return null;
  return found.record;
}

// ==========================================
// Members Club Operations (SpinWinMembers & SpinWinLedger)
// ==========================================

function listMembers_() {
  const sheet = getOrCreateSheet_(MEMBERS_SHEET_NAME);
  const headerMap = ensureHeaders_(sheet, MEMBERS_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const lastCol = sheet.getLastColumn();
  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const members = rows.map(function(row) {
    return rowToMember_(row, headerMap);
  }).filter(function(m) { return !!m.id; });

  // Attach ledger transactions to each member
  const allLedger = listLedger_();
  const ledgerByMember = {};
  allLedger.forEach(function(tx) {
    if (!ledgerByMember[tx.memberId]) ledgerByMember[tx.memberId] = [];
    ledgerByMember[tx.memberId].push(tx);
  });

  members.forEach(function(m) {
    m.transactions = ledgerByMember[m.id] || [];
  });

  return members;
}

function saveMember_(payload) {
  const m = payload.member || payload;
  const member = normalizeMember_(m);
  if (!member.name) {
    return jsonOutput({ ok: false, message: "Member name is required." });
  }

  const sheet = getOrCreateSheet_(MEMBERS_SHEET_NAME);
  const headerMap = ensureHeaders_(sheet, MEMBERS_HEADERS);
  const found = findRowByMemberId_(sheet, headerMap, member.id);

  if (found.rowNumber > 0) {
    sheet.getRange(found.rowNumber, 1, 1, MEMBERS_HEADERS.length).setValues([memberToRow_(member, headerMap)]);
    return jsonOutput({
      ok: true,
      message: "Member updated successfully.",
      member: member
    });
  }

  appendRowFast_(sheet, memberToRow_(member, headerMap));

  // If initial credits > 0, log an initial transaction in the ledger.
  // Reuse the exact transaction id attached by the frontend (or a stable
  // id derived from the member) so a later batch re-sync never writes a
  // second, duplicate "Initial" row.
  if (member.credits > 0) {
    var initialTx = null;
    if (Array.isArray(m.transactions)) {
      for (var t = 0; t < m.transactions.length; t++) {
        var cand = m.transactions[t];
        if (!cand) continue;
        if (
          String(cand.type || "").toUpperCase() === "CREDIT" &&
          Math.round(Number(cand.credits) || 0) === member.credits &&
          Math.round(Number(cand.balanceAfter) || 0) === member.credits
        ) {
          initialTx = cand;
          break;
        }
      }
    }

    const tx = {
      id: (initialTx && String(initialTx.id || "").trim())
        ? String(initialTx.id).trim()
        : ("tx_init_" + member.id),
      memberId: member.id,
      customerName: member.name,
      phone: member.phone,
      type: "CREDIT",
      credits: member.credits,
      cashValue: round2_(member.credits * CREDIT_RATE),
      purchaseAmount: 0,
      balanceAfter: member.credits,
      note: (initialTx && String(initialTx.note || "").trim()) || "Initial joining bonus credits",
      dateIso: toIso_((initialTx && initialTx.dateIso) || member.createdAtIso),
      savedAtServer: toIso_(new Date())
    };
    appendLedgerRow_(tx);
  }

  return jsonOutput({
    ok: true,
    message: "Member created successfully.",
    member: member
  });
}

function deleteMember_(payload) {
  const memberId = String(payload.memberId || payload.id || "").trim();
  if (!memberId) {
    return jsonOutput({ ok: false, message: "memberId is required." });
  }

  const sheet = getOrCreateSheet_(MEMBERS_SHEET_NAME);
  const headerMap = ensureHeaders_(sheet, MEMBERS_HEADERS);
  const found = findRowByMemberId_(sheet, headerMap, memberId);

  if (found.rowNumber < 1) {
    return jsonOutput({ ok: false, message: "Member not found." });
  }

  sheet.deleteRow(found.rowNumber);
  return jsonOutput({
    ok: true,
    message: "Member deleted successfully.",
    memberId: memberId
  });
}

function addMemberCredits_(payload) {
  const memberId = String(payload.memberId || payload.id || "").trim();
  const creditsToAdd = Math.round(Number(payload.credits) || 0);
  const purchaseAmount = round2_(Number(payload.purchaseAmount) || 0);
  const note = String(payload.note || "").trim();
  const txId = String(payload.txId || ("tx_" + Date.now())).trim();

  if (!memberId || creditsToAdd <= 0) {
    return jsonOutput({ ok: false, message: "Valid memberId and credits (> 0) required." });
  }

  const sheet = getOrCreateSheet_(MEMBERS_SHEET_NAME);
  const headerMap = ensureHeaders_(sheet, MEMBERS_HEADERS);
  const found = findRowByMemberId_(sheet, headerMap, memberId);

  var member = found.member;
  if (found.rowNumber < 1 || !member) {
    // Auto-create member in Google Sheets if not already present
    member = {
      id: memberId,
      name: String(payload.name || payload.customerName || "Member " + memberId).trim(),
      phone: String(payload.phone || payload.customerPhone || "").trim(),
      memberType: String(payload.memberType || "Spin Winner").trim(),
      credits: creditsToAdd,
      cashWorth: round2_(creditsToAdd * CREDIT_RATE),
      totalEarned: creditsToAdd,
      totalRedeemed: 0,
      totalCashPaid: 0,
      notes: String(payload.notes || "Auto-registered via Add Credits").trim(),
      createdAtIso: toIso_(new Date()),
      updatedAtIso: toIso_(new Date()),
      source: "Auto / Credits Action",
      savedAtServer: toIso_(new Date())
    };
    appendRowFast_(sheet, memberToRow_(member, headerMap));
  } else {
    member.credits = Math.round((member.credits || 0) + creditsToAdd);
    member.cashWorth = round2_(member.credits * CREDIT_RATE);
    member.totalEarned = Math.round((member.totalEarned || 0) + creditsToAdd);
    member.updatedAtIso = toIso_(new Date());
    sheet.getRange(found.rowNumber, 1, 1, MEMBERS_HEADERS.length).setValues([memberToRow_(member, headerMap)]);
  }

  // Log in SpinWinLedger
  const tx = {
    id: txId,
    memberId: member.id,
    customerName: member.name,
    phone: member.phone,
    type: "CREDIT",
    credits: creditsToAdd,
    cashValue: round2_(creditsToAdd * CREDIT_RATE),
    purchaseAmount: purchaseAmount,
    balanceAfter: member.credits,
    note: note || (purchaseAmount > 0 ? "Purchase ₹" + purchaseAmount : "Points added"),
    dateIso: toIso_(new Date()),
    savedAtServer: toIso_(new Date())
  };
  appendLedgerRow_(tx);

  return jsonOutput({
    ok: true,
    message: "Credits added successfully.",
    member: member,
    newBalance: member.credits
  });
}

function deductMemberCredits_(payload) {
  const memberId = String(payload.memberId || payload.id || "").trim();
  const creditsToDeduct = Math.round(Number(payload.credits) || 0);
  const redeemMode = String(payload.redeemMode || "purchase_discount").trim();
  const billAmount = round2_(Number(payload.billAmount) || 0);
  const note = String(payload.note || "").trim();
  const txId = String(payload.txId || ("tx_" + Date.now())).trim();

  if (!memberId || creditsToDeduct <= 0) {
    return jsonOutput({ ok: false, message: "Valid memberId and credits (> 0) required." });
  }

  const sheet = getOrCreateSheet_(MEMBERS_SHEET_NAME);
  const headerMap = ensureHeaders_(sheet, MEMBERS_HEADERS);
  const found = findRowByMemberId_(sheet, headerMap, memberId);

  if (found.rowNumber < 1 || !found.member) {
    return jsonOutput({ ok: false, message: "Member not found in Google Sheets." });
  }

  const member = found.member;
  if ((member.credits || 0) < creditsToDeduct) {
    return jsonOutput({
      ok: false,
      message: "Insufficient credit balance. Member only has " + member.credits + " credits."
    });
  }

  const cashValue = round2_(creditsToDeduct * CREDIT_RATE);
  member.credits = Math.max(0, Math.round((member.credits || 0) - creditsToDeduct));
  member.cashWorth = round2_(member.credits * CREDIT_RATE);
  member.totalRedeemed = Math.round((member.totalRedeemed || 0) + creditsToDeduct);
  member.totalCashPaid = round2_((member.totalCashPaid || 0) + cashValue);
  member.updatedAtIso = toIso_(new Date());

  sheet.getRange(found.rowNumber, 1, 1, MEMBERS_HEADERS.length).setValues([memberToRow_(member, headerMap)]);

  const isCash = redeemMode === "cash_payout";
  const typeStr = isCash ? "DEBIT - CASH" : "DEBIT - DISCOUNT";
  const defaultNote = isCash
    ? "Cash Payout ₹" + cashValue.toFixed(2)
    : "Bill Discount ₹" + cashValue.toFixed(2) + (billAmount > 0 ? " on ₹" + billAmount : "");

  // Log in SpinWinLedger
  const tx = {
    id: txId,
    memberId: member.id,
    customerName: member.name,
    phone: member.phone,
    type: typeStr,
    credits: -creditsToDeduct,
    cashValue: cashValue,
    purchaseAmount: billAmount,
    balanceAfter: member.credits,
    note: note || defaultNote,
    dateIso: toIso_(new Date()),
    savedAtServer: toIso_(new Date())
  };
  appendLedgerRow_(tx);

  return jsonOutput({
    ok: true,
    message: isCash ? "Cash payout recorded." : "Purchase bill discount applied.",
    member: member,
    newBalance: member.credits
  });
}

/**
 * Batched member + ledger sync.
 * Reads the members and ledger sheets exactly ONCE, diffs everything in
 * memory, then writes members and ledger in a handful of range operations.
 * (Previous version re-scanned the ID columns once per member/txAID - O(N^2).)
 */
function syncMembersBatch_(payload) {
  const members = Array.isArray(payload.members) ? payload.members : [];
  if (!members.length) {
    return jsonOutput({ ok: true, message: "Batch members synced.", count: 0 });
  }

  const mSheet = getOrCreateSheet_(MEMBERS_SHEET_NAME);
  const mHeader = ensureHeaders_(mSheet, MEMBERS_HEADERS);
  const mCols = MEMBERS_HEADERS.length;
  const mLastRow = mSheet.getLastRow();

  // Single pass over the full members range
  const existingIdToRow = {};
  const mIdCol = (mHeader.MemberID !== undefined ? mHeader.MemberID : 0) + 1;
  if (mLastRow >= 2) {
    const mRows = mSheet.getRange(2, 1, mLastRow - 1, mCols).getValues();
    for (var mi = 0; mi < mRows.length; mi++) {
      const id = String(mRows[mi][mIdCol - 1] || "");
      if (id) existingIdToRow[id] = mi + 2; // 1-based sheet row
    }
  }

  // Single pass over the ledger TxID column for dedup
  const lSheet = getOrCreateSheet_(LEDGER_SHEET_NAME);
  const lHeader = ensureHeaders_(lSheet, LEDGER_HEADERS);
  const lCols = LEDGER_HEADERS.length;
  const lLastRow = lSheet.getLastRow();
  const existingTxIds = {};
  const lTxIdx = (lHeader.TxID !== undefined ? lHeader.TxID : 0) + 1;
  if (lLastRow >= 2) {
    const tIds = lSheet.getRange(2, lTxIdx, lLastRow - 1, 1).getValues();
    for (var li = 0; li < tIds.length; li++) {
      const t = String(tIds[li][0] || "");
      if (t) existingTxIds[t] = true;
    }
  }

  const updates = [];
  const newMemberRows = [];
  const newLedgerRows = [];

  members.forEach(function(mRaw) {
    const member = normalizeMember_(mRaw);
    const rowNum = existingIdToRow[member.id];
    if (rowNum) {
      updates.push([rowNum, memberToRow_(member, mHeader)]);
    } else {
      newMemberRows.push(memberToRow_(member, mHeader));
    }

    if (Array.isArray(mRaw.transactions)) {
      mRaw.transactions.forEach(function(tx) {
        const txId = String((tx && tx.id) || "").trim();
        if (!txId || existingTxIds[txId]) return;
        existingTxIds[txId] = true;
        newLedgerRows.push(ledgerToRow_({
          id: txId,
          memberId: member.id,
          customerName: member.name,
          phone: member.phone,
          type: tx.type || "CREDIT",
          credits: Number(tx.credits || 0),
          cashValue: round2_(Number(tx.cashValue || 0)),
          purchaseAmount: round2_(Number(tx.purchaseAmount || 0)),
          balanceAfter: Math.round(Number(tx.balanceAfter || 0)),
          note: tx.note || "",
          dateIso: toIso_(tx.dateIso || new Date()),
          savedAtServer: toIso_(new Date())
        }, lHeader));
      });
    }
  });

  // Member updates (scattered rows = per-row setValues)
  updates.forEach(function(p) {
    mSheet.getRange(p[0], 1, 1, mCols).setValues([p[1]]);
  });

  // All brand-new members + new ledger rows append in one batch each
  if (newMemberRows.length) {
    mSheet.getRange(mLastRow + 1, 1, newMemberRows.length, mCols).setValues(newMemberRows);
  }
  if (newLedgerRows.length) {
    lSheet.getRange(lLastRow + 1, 1, newLedgerRows.length, lCols).setValues(newLedgerRows);
  }

  return jsonOutput({
    ok: true,
    message: "Batch members synced.",
    count: members.length
  });
}

function listLedger_(filterMemberId) {
  const sheet = getOrCreateSheet_(LEDGER_SHEET_NAME);
  const headerMap = ensureHeaders_(sheet, LEDGER_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const lastCol = sheet.getLastColumn();
  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var txs = rows.map(function(row) {
    return rowToLedger_(row, headerMap);
  }).filter(function(tx) {
    return !!tx.id;
  });

  if (filterMemberId) {
    txs = txs.filter(function(tx) { return tx.memberId === filterMemberId; });
  }

  txs.sort(function(a, b) {
    return new Date(b.dateIso).getTime() - new Date(a.dateIso).getTime();
  });
  return txs;
}

function appendLedgerRow_(tx) {
  const sheet = getOrCreateSheet_(LEDGER_SHEET_NAME);
  const headerMap = ensureHeaders_(sheet, LEDGER_HEADERS);

  // Never write two rows with the same transaction id (idempotent retries).
  const txId = String((tx && tx.id) || "").trim();
  if (txId) {
    const lastRow = sheet.getLastRow();
    const idCol = (headerMap.TxID !== undefined ? headerMap.TxID : 0) + 1;
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === txId) return; // already logged
      }
    }
  }

  appendRowFast_(sheet, ledgerToRow_(tx, headerMap));
}

function appendLedgerRowIfNotExists_(tx, member) {
  const sheet = getOrCreateSheet_(LEDGER_SHEET_NAME);
  const headerMap = ensureHeaders_(sheet, LEDGER_HEADERS);
  const txId = String(tx.id || "").trim();
  if (!txId) return;

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, (headerMap.TxID || 0) + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === txId) return; // Already logged
    }
  }

  const rowData = {
    id: txId,
    memberId: member ? member.id : tx.memberId,
    customerName: member ? member.name : tx.customerName,
    phone: member ? member.phone : tx.phone,
    type: tx.type || "CREDIT",
    credits: Number(tx.credits || 0),
    cashValue: round2_(Number(tx.cashValue || 0)),
    purchaseAmount: round2_(Number(tx.purchaseAmount || 0)),
    balanceAfter: Math.round(Number(tx.balanceAfter || 0)),
    note: tx.note || "",
    dateIso: toIso_(tx.dateIso || new Date()),
    savedAtServer: toIso_(new Date())
  };
  appendRowFast_(sheet, ledgerToRow_(rowData, headerMap));
}

// ==========================================
// Settings Operations (SpinWinSettings)
// ==========================================

function saveSettings_(payload) {
  const incoming = payload.settings || payload;
  const settings = normalizeSettings_(incoming);
  const sheet = getOrCreateSheet_(SETTINGS_SHEET_NAME);
  const headerMap = ensureHeaders_(sheet, SETTINGS_HEADERS);
  const found = findSettingsRow_(sheet, headerMap, SETTINGS_KEY);
  const updatedAtIso = new Date().toISOString();
  const row = [SETTINGS_KEY, JSON.stringify(settings), updatedAtIso];

  if (found.rowNumber > 0) {
    sheet.getRange(found.rowNumber, 1, 1, SETTINGS_HEADERS.length).setValues([row]);
  } else {
    appendRowFast_(sheet, row);
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
  const headerMap = ensureHeaders_(sheet, SETTINGS_HEADERS);
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

  // Preserve Multi-Condition Wheel Setups
  var rawConditions = Array.isArray(settings.conditions) ? settings.conditions : [];
  var conditions = rawConditions.map(function(cond, cIdx) {
    cond = cond || {};
    var cPrizes = Array.isArray(cond.prizes) ? cond.prizes : [];
    return {
      id: String(cond.id || ("cond_" + (cIdx + 1))),
      itemName: String(cond.itemName || "Unnamed Item").trim(),
      label: String(cond.label || cond.itemName || ("Wheel " + (cIdx + 1))).trim(),
      icon: String(cond.icon || "🎁").trim(),
      minAmount: Math.max(0, toNumber_(cond.minAmount, 0)),
      prizes: cPrizes.map(function(p, pIdx) {
        p = p || {};
        return {
          id: String(p.id || ("prize_" + (cIdx + 1) + "_" + (pIdx + 1))),
          name: String(p.name || "Unnamed Prize").trim(),
          probability: Math.max(0, toNumber_(p.probability, 0)),
          enabled: p.enabled !== false
        };
      })
    };
  });

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
    spinDuration: Math.max(2, Math.min(15, toNumber_(settings.spinDuration, 5))),
    manualDateEnabled: settings.manualDateEnabled === true || String(settings.manualDateEnabled).toLowerCase() === "true",
    manualDateTime: String(settings.manualDateTime || ""),
    appsScriptUrl: String(settings.appsScriptUrl || "").trim(),
    conditions: conditions.length ? conditions : defaultSettings_().conditions,
    prizes: prizes.length ? prizes : defaultSettings_().prizes
  };
}

function defaultSettings_() {
  return {
    shopName: "Lucky Shop",
    shopLogoUrl: "",
    expiryHours: 24,
    spinDuration: 5,
    manualDateEnabled: false,
    manualDateTime: "",
    appsScriptUrl: "",
    conditions: [
      {
        id: "cond_jeans",
        itemName: "Jeans Pant",
        label: "Jeans Pant (Wheel 1)",
        icon: "👖",
        minAmount: 0,
        prizes: [
          { id: "p1", name: "Leather Belt Free", probability: 30, enabled: true },
          { id: "p2", name: "20% Off Denim", probability: 25, enabled: true },
          { id: "p3", name: "Cap Free", probability: 15, enabled: true },
          { id: "p4", name: "Try Again", probability: 29.9, enabled: true },
          { id: "p5", name: "Grand Prize", probability: 0.1, enabled: true }
        ]
      },
      {
        id: "cond_shirt",
        itemName: "Shirt",
        label: "Shirt (Wheel 2)",
        icon: "👔",
        minAmount: 0,
        prizes: [
          { id: "p6", name: "Free Tie", probability: 35, enabled: true },
          { id: "p7", name: "15% Off Shirt", probability: 30, enabled: true },
          { id: "p8", name: "Socks Pair Free", probability: 20, enabled: true },
          { id: "p9", name: "Try Again", probability: 14.9, enabled: true },
          { id: "p10", name: "Grand Prize", probability: 0.1, enabled: true }
        ]
      }
    ],
    prizes: [
      { id: "p1", name: "10% Discount", probability: 40, enabled: true },
      { id: "p2", name: "Free Drink", probability: 25, enabled: true },
      { id: "p3", name: "Buy 1 Get 1", probability: 10, enabled: true },
      { id: "p4", name: "Try Again", probability: 24.9, enabled: true },
      { id: "p5", name: "Grand Prize", probability: 0.1, enabled: true }
    ]
  };
}

// ==========================================
// Sheet & Row Helpers
// ==========================================

// Per-execution cache.
// Every web-app request runs in a fresh execution, so these module-level
// caches only live for the lifetime of a single request — but they collapse
// many repeated getActiveSpreadsheet()/getSheetByName()/header reads inside
// that request into a handful of Sheets API calls (the main latency source).
var _ssCache_ = {};
var _headerCache_ = {};

function getSpreadsheet_() {
  if (!_ssCache_.ss) _ssCache_.ss = SpreadsheetApp.getActiveSpreadsheet();
  return _ssCache_.ss;
}

function canonicalSheetName_(name) {
  return String(name || "").trim().toLowerCase();
}

function getOrCreateSheet_(sheetName) {
  if (_ssCache_[sheetName]) return _ssCache_[sheetName];

  const ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    // Reuse any existing tab whose name only differs by case or whitespace,
    // so a stray "SpinWinData " / "spinwinmembers" tab never causes us to
    // create a second, duplicate data tab.
    const wanted = canonicalSheetName_(sheetName);
    const all = ss.getSheets();
    for (var i = 0; i < all.length; i++) {
      if (canonicalSheetName_(all[i].getName()) === wanted) {
        sheet = all[i];
        break;
      }
    }
  }

  if (!sheet) sheet = ss.insertSheet(sheetName);
  _ssCache_[sheetName] = sheet;
  return sheet;
}

/**
 * Menu helper: scans the workbook for any duplicate Spin & Win tabs
 * (same logical sheet name appearing more than once, ignoring case/space)
 * and reports them so they can be deleted.
 */
function findDuplicateSheets() {
  const ss = getSpreadsheet_();
  const expected = [SHEET_NAME, SETTINGS_SHEET_NAME, MEMBERS_SHEET_NAME, LEDGER_SHEET_NAME];
  const wanted = {};
  expected.forEach(function(n) { wanted[canonicalSheetName_(n)] = true; });

  const seen = {};
  const result = [];
  ss.getSheets().forEach(function(s) {
    const key = canonicalSheetName_(s.getName());
    if (!wanted[key]) return;
    if (seen[key]) {
      result.push(s.getName());
    } else {
      seen[key] = true;
    }
  });

  if (result.length) {
    ss.toast(
      "Duplicate tabs found: " + result.join(", ") + ". Delete the extra copies to avoid confusion.",
      "🧹 Duplicate Tabs",
      10
    );
    Logger.log("Duplicate Spin & Win tabs found: " + result.join(", "));
  } else {
    ss.toast("No duplicate Spin & Win tabs found. All good!", "🧹 Duplicate Tabs", 4);
  }
}

/**
 * Menu helper: scans the Members (MemberID) and Ledger (TxID) tabs for any
 * repeat ids — those would be genuine duplicate data rows. Reports them so
 * the extra copies can be deleted safely (balances are kept in Members).
 */
function findDuplicateRows() {
  const ss = getSpreadsheet_();

  var mSheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  if (mSheet) {
    const mHeader = ensureHeaders_(mSheet, MEMBERS_HEADERS);
    const lastRow = mSheet.getLastRow();
    if (lastRow >= 2) {
      const idCol = (mHeader.MemberID !== undefined ? mHeader.MemberID : 0) + 1;
      const ids = mSheet.getRange(2, idCol, lastRow - 1, 1).getValues();
      const seen = {}; const dups = [];
      for (var i = 0; i < ids.length; i++) {
        const v = String(ids[i][0] || "").trim();
        if (!v) continue;
        if (seen[v]) dups.push(v); else seen[v] = true;
      }
      if (dups.length) {
        ss.toast("Found " + dups.length + " duplicate member(s) in Members tab: " + dups.slice(0, 5).join(", ") + (dups.length > 5 ? " …" : "") + ".", "🔍 Duplicate Check", 12);
        Logger.log("Duplicate MemberIDs: " + dups.join(", "));
      }
    }
  }

  var lSheet = ss.getSheetByName(LEDGER_SHEET_NAME);
  if (lSheet) {
    const lHeader = ensureHeaders_(lSheet, LEDGER_HEADERS);
    const lastRow = lSheet.getLastRow();
    if (lastRow >= 2) {
      const idCol = (lHeader.TxID !== undefined ? lHeader.TxID : 0) + 1;
      const ids = lSheet.getRange(2, idCol, lastRow - 1, 1).getValues();
      const seen = {}; const dups = [];
      for (var j = 0; j < ids.length; j++) {
        const v = String(ids[j][0] || "").trim();
        if (!v) continue;
        if (seen[v]) dups.push(v); else seen[v] = true;
      }
      if (dups.length) {
        ss.toast("Found " + dups.length + " duplicate transaction(s) in Ledger tab: " + dups.slice(0, 5).join(", ") + (dups.length > 5 ? " …" : "") + ". Delete the extra copies.", "🔍 Duplicate Check", 12);
        Logger.log("Duplicate TxIDs: " + dups.join(", "));
      } else if (!mSheet) {
        ss.toast("No duplicate ledger rows. Each TxID is unique.", "🔍 Duplicate Check", 5);
      }
    }
  }

  if (mSheet && lSheet) {
    ss.toast("Duplicate row check complete. Details in View → Logs.", "🔍 Duplicate Check", 5);
  }
}

// appendRow() is slow; a single setValues() into the first free row is faster.
function appendRowFast_(sheet, row) {
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
}

function ensureHeaders_(sheet, expectedHeaders) {
  const sig = sheet.getName() + "|" + expectedHeaders.length;
  if (_headerCache_[sig]) return _headerCache_[sig];

  let map;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(expectedHeaders);
    map = {};
    expectedHeaders.forEach(function(h, idx) { map[h] = idx; });
  } else {
    const lastCol = sheet.getLastColumn();
    const existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    map = {};
    existing.forEach(function(h, idx) { map[String(h).trim()] = idx; });

    // Append any missing columns safely
    expectedHeaders.forEach(function(h) {
      if (map[h] === undefined) {
        const newCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCol).setValue(h);
        map[h] = newCol - 1;
      }
    });
  }

  _headerCache_[sig] = map;
  return map;
}

function findRowByRecordId_(sheet, headerMap, recordId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rowNumber: -1, record: null };

  const idCol = (headerMap.RecordID !== undefined ? headerMap.RecordID : 0) + 1;
  const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === recordId) {
      const row = sheet.getRange(i + 2, 1, 1, sheet.getLastColumn()).getValues()[0];
      return {
        rowNumber: i + 2,
        record: rowToRecord_(row, headerMap)
      };
    }
  }
  return { rowNumber: -1, record: null };
}

function findRowByMemberId_(sheet, headerMap, memberId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rowNumber: -1, member: null };

  const idCol = (headerMap.MemberID !== undefined ? headerMap.MemberID : 0) + 1;
  const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === memberId) {
      const row = sheet.getRange(i + 2, 1, 1, sheet.getLastColumn()).getValues()[0];
      return {
        rowNumber: i + 2,
        member: rowToMember_(row, headerMap)
      };
    }
  }
  return { rowNumber: -1, member: null };
}

function rowToRecord_(row, headerMap) {
  return {
    recordId: String(row[headerMap.RecordID] || ""),
    shopName: String(row[headerMap.ShopName] || ""),
    customerName: String(row[headerMap.CustomerName] || ""),
    customerNumber: String(row[headerMap.CustomerNumber] || ""),
    amount: Number(row[headerMap.Amount] || 0),
    prize: String(row[headerMap.Prize] || ""),
    purchasedItem: String(headerMap.PurchasedItem !== undefined ? row[headerMap.PurchasedItem] || "" : ""),
    conditionId: String(headerMap.ConditionID !== undefined ? row[headerMap.ConditionID] || "" : ""),
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

function recordToRow_(record, headerMap) {
  const row = new Array(Object.keys(headerMap).length).fill("");
  row[headerMap.RecordID] = record.recordId || "";
  row[headerMap.ShopName] = record.shopName || "";
  row[headerMap.CustomerName] = record.customerName || "";
  row[headerMap.CustomerNumber] = record.customerNumber || "";
  row[headerMap.Amount] = Number(record.amount || 0);
  row[headerMap.Prize] = record.prize || "";
  if (headerMap.PurchasedItem !== undefined) row[headerMap.PurchasedItem] = record.purchasedItem || "";
  if (headerMap.ConditionID !== undefined) row[headerMap.ConditionID] = record.conditionId || "";
  row[headerMap.Status] = record.status || "Pending";
  row[headerMap.DateTimeISO] = record.dateTimeIso || "";
  row[headerMap.DateTimeDisplay] = record.dateTimeDisplay || formatDate_(record.dateTimeIso);
  row[headerMap.ExpiryISO] = record.expiryIso || "";
  row[headerMap.ExpiryDisplay] = record.expiryDisplay || formatDate_(record.expiryIso);
  row[headerMap.CreatedAtISO] = record.createdAtIso || toIso_(new Date());
  row[headerMap.UpdatedAtISO] = record.updatedAtIso || toIso_(new Date());
  row[headerMap.Source] = record.source || "SpinWinFrontend";
  row[headerMap.SavedAtServer] = record.savedAtServer || toIso_(new Date());
  return row;
}

function rowToMember_(row, headerMap) {
  const credits = Math.round(Number(row[headerMap.Credits] || 0));
  return {
    id: String(row[headerMap.MemberID] || ""),
    name: String(row[headerMap.CustomerName] || ""),
    phone: String(row[headerMap.Phone] || ""),
    memberType: String(row[headerMap.MemberType] || "Direct Member"),
    credits: credits,
    cashWorth: round2_(Number(row[headerMap.CashWorth] !== undefined ? row[headerMap.CashWorth] : credits * CREDIT_RATE)),
    totalEarned: Math.round(Number(row[headerMap.TotalEarned] || credits)),
    totalRedeemed: Math.round(Number(row[headerMap.TotalRedeemed] || 0)),
    totalCashPaid: round2_(Number(row[headerMap.TotalCashPaid] || 0)),
    notes: String(row[headerMap.Notes] || ""),
    createdAtIso: String(row[headerMap.CreatedAtISO] || ""),
    updatedAtIso: String(row[headerMap.UpdatedAtISO] || ""),
    source: String(row[headerMap.Source] || "SpinWinFrontend")
  };
}

function memberToRow_(member, headerMap) {
  const row = new Array(Object.keys(headerMap).length).fill("");
  row[headerMap.MemberID] = member.id || "";
  row[headerMap.CustomerName] = member.name || "";
  row[headerMap.Phone] = member.phone || "";
  row[headerMap.MemberType] = member.memberType || "Direct Member";
  row[headerMap.Credits] = Math.round(Number(member.credits || 0));
  row[headerMap.CashWorth] = round2_(Number(member.cashWorth !== undefined ? member.cashWorth : member.credits * CREDIT_RATE));
  row[headerMap.TotalEarned] = Math.round(Number(member.totalEarned || 0));
  row[headerMap.TotalRedeemed] = Math.round(Number(member.totalRedeemed || 0));
  row[headerMap.TotalCashPaid] = round2_(Number(member.totalCashPaid || 0));
  row[headerMap.Notes] = member.notes || "";
  row[headerMap.CreatedAtISO] = member.createdAtIso || toIso_(new Date());
  row[headerMap.UpdatedAtISO] = member.updatedAtIso || toIso_(new Date());
  row[headerMap.Source] = member.source || "SpinWinFrontend";
  row[headerMap.SavedAtServer] = toIso_(new Date());
  return row;
}

function rowToLedger_(row, headerMap) {
  return {
    id: String(row[headerMap.TxID] || ""),
    memberId: String(row[headerMap.MemberID] || ""),
    customerName: String(row[headerMap.CustomerName] || ""),
    phone: String(row[headerMap.Phone] || ""),
    type: String(row[headerMap.Type] || "CREDIT"),
    credits: Number(row[headerMap.Credits] || 0),
    cashValue: round2_(Number(row[headerMap.CashValue] || 0)),
    purchaseAmount: round2_(Number(row[headerMap.PurchaseAmount] || 0)),
    balanceAfter: Math.round(Number(row[headerMap.BalanceAfter] || 0)),
    note: String(row[headerMap.Note] || ""),
    dateIso: String(row[headerMap.DateISO] || "")
  };
}

function ledgerToRow_(tx, headerMap) {
  const row = new Array(Object.keys(headerMap).length).fill("");
  row[headerMap.TxID] = tx.id || ("tx_" + Date.now());
  row[headerMap.MemberID] = tx.memberId || "";
  row[headerMap.CustomerName] = tx.customerName || "";
  row[headerMap.Phone] = tx.phone || "";
  row[headerMap.Type] = tx.type || "CREDIT";
  row[headerMap.Credits] = Number(tx.credits || 0);
  row[headerMap.CashValue] = round2_(Number(tx.cashValue || 0));
  row[headerMap.PurchaseAmount] = round2_(Number(tx.purchaseAmount || 0));
  row[headerMap.BalanceAfter] = Math.round(Number(tx.balanceAfter || 0));
  row[headerMap.Note] = tx.note || "";
  row[headerMap.DateISO] = tx.dateIso || toIso_(new Date());
  row[headerMap.SavedAtServer] = toIso_(new Date());
  return row;
}

function normalizeRecord_(record) {
  record = record || {};
  return {
    recordId: String(record.recordId || record.id || ("DPM" + Date.now())).trim(),
    shopName: String(record.shopName || "Lucky Shop").trim(),
    customerName: String(record.customerName || "Unknown").trim(),
    customerNumber: String(record.customerNumber || "").trim(),
    amount: Number(record.amount || 0),
    prize: String(record.prize || "").trim(),
    purchasedItem: String(record.purchasedItem || "").trim(),
    conditionId: String(record.conditionId || "").trim(),
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
    purchasedItem: pickValue_(updates.purchasedItem, original.purchasedItem, ""),
    conditionId: pickValue_(updates.conditionId, original.conditionId, ""),
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
  if (record.purchasedItem !== undefined) partial.purchasedItem = String(record.purchasedItem).trim();
  if (record.conditionId !== undefined) partial.conditionId = String(record.conditionId).trim();
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

function normalizeMember_(m) {
  m = m || {};
  const now = new Date();
  const credits = Math.max(0, Math.round(Number(m.credits) || 0));
  return {
    id: String(m.id || ("mem_" + now.getTime())).trim(),
    name: String(m.name || "Member").trim(),
    phone: String(m.phone || "").trim(),
    memberType: String(m.memberType || "Direct Member").trim(),
    credits: credits,
    cashWorth: round2_(Number(m.cashWorth !== undefined ? m.cashWorth : credits * CREDIT_RATE)),
    totalEarned: Math.max(0, Math.round(Number(m.totalEarned) || credits)),
    totalRedeemed: Math.max(0, Math.round(Number(m.totalRedeemed) || 0)),
    totalCashPaid: round2_(Number(m.totalCashPaid) || 0),
    notes: String(m.notes || "").trim(),
    createdAtIso: toIso_(m.createdAtIso || now),
    updatedAtIso: toIso_(m.updatedAtIso || now),
    source: String(m.source || "SpinWinFrontend")
  };
}

function normalizeStatus_(status) {
  const value = String(status || "Pending").toLowerCase();
  if (value === "approved" || value === "completed") return "Completed";
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

function couponOutput_(html) {
  return ContentService.createTextOutput(html).setMimeType(ContentService.MimeType.HTML);
}

function escHtml_(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

/**
 * Renders the mobile-friendly "Scan for Live Status" page that a scanned QR
 * opens. Shows the voucher's current status, prize, expiry, and auto-refreshes.
 */
function couponPageHtml_(record, recordId, errorMessage) {
  const tz = Session.getScriptTimeZone();
  const statusMeta = {
    "Pending":   { label: "Valid - Awaiting Redemption", color: "#0f766e", bg: "#d1fae5", icon: "✅" },
    "Completed": { label: "Completed",                   color: "#1d4ed8", bg: "#dbeafe", icon: "🎉" },
    "Rejected":  { label: "Rejected",                    color: "#b91c1c", bg: "#fee2e2", icon: "⛔" },
    "Expired":   { label: "Expired",                     color: "#4b5563", bg: "#f3f4f6", icon: "⏰" }
  };

  var statusLabel = "Not Found";
  var statusColor = "#6b7280";
  var statusBg = "#f3f4f6";
  var statusIcon = "❓";
  var shopName = "", prize = "", customer = "", customerNumber = "—", amount = "0.00";
  var dateText = "—", expiryText = "—", expiryEpoch = 0;
  var escRecordId = escHtml_(recordId);

  if (!record) {
    statusLabel = errorMessage ? "Error" : "Voucher Not Found";
    statusIcon = "⚠️";
  } else {
    const meta = statusMeta[getEffectiveStatus_(record)] || statusMeta["Pending"];
    statusLabel = meta.label;
    statusColor = meta.color;
    statusBg = meta.bg;
    statusIcon = meta.icon;
    shopName = escHtml_(record.shopName || "");
    prize = escHtml_(record.prize || "");
    customer = escHtml_(record.customerName || "");
    customerNumber = escHtml_((record.customerNumber || "") === "" ? "—" : record.customerNumber);
    amount = escHtml_(Number(record.amount || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
    escRecordId = escHtml_(record.recordId || "");
    if (record.dateTimeIso) dateText = escHtml_(formatDate_(record.dateTimeIso));
    if (record.expiryIso) {
      const d = new Date(record.expiryIso);
      if (!isNaN(d.getTime())) {
        expiryEpoch = d.getTime();
        expiryText = escHtml_(Utilities.formatDate(d, tz, "dd MMM yyyy, hh:mm a"));
      }
    }
  }

  const title = record ? ("Voucher Status - " + statusLabel) : "Voucher Status";
  const errorHtml = errorMessage ? '<div class="err">' + escHtml_(errorMessage) + '</div>' : "";
  const checkedAt = escHtml_(Utilities.formatDate(new Date(), tz, "dd MMM yyyy, hh:mm:ss a"));

  return '' +
'<!DOCTYPE html>\n' +
'<html lang="en"><head><meta charset="UTF-8">\n' +
'<meta name="viewport" content="width=device-width, initial-scale=1.0">\n' +
'<title>' + title + '</title>\n' +
'<style>\n' +
'  *{margin:0;padding:0;box-sizing:border-box;}\n' +
'  body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#0f172a;color:#0f172a;min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:24px 16px;}\n' +
'  .card{background:#ffffff;border-radius:20px;max-width:420px;width:100%;padding:28px 24px;box-shadow:0 20px 60px rgba(0,0,0,.35);}\n' +
'  .brand{text-align:center;font-size:13px;letter-spacing:2px;text-transform:uppercase;color:#94a3b8;margin-bottom:6px;}\n' +
'  h1{font-size:22px;line-height:1.3;text-align:center;margin-bottom:18px;color:#1e293b;}\n' +
'  .badge{display:flex;align-items:center;justify-content:center;gap:10px;border-radius:999px;padding:12px 18px;font-size:17px;font-weight:700;margin-bottom:20px;}\n' +
'  .row{display:flex;justify-content:space-between;gap:12px;padding:11px 0;border-bottom:1px solid #eef2f7;font-size:15px;}\n' +
'  .row:last-child{border-bottom:none;}\n' +
'  .key{color:#64748b;}\n' +
'  .val{font-weight:600;color:#1e293b;text-align:right;}\n' +
'  .prize{font-size:16px;font-weight:700;color:#b45309;}\n' +
'  .err{background:#fee2e2;color:#b91c1c;border-radius:12px;padding:12px 14px;font-size:14px;margin-bottom:16px;text-align:center;}\n' +
'  .foot{text-align:center;font-size:11px;color:#94a3b8;margin-top:18px;line-height:1.6;}\n' +
'  .note{text-align:center;font-size:13px;color:#475569;margin-top:6px;}\n' +
'  @keyframes pulse{0%,100%{opacity:1}50%{opacity:.45}}\n' +
'  .live{display:inline-block;width:9px;height:9px;border-radius:50%;background:#10b981;margin-right:6px;animation:pulse 1.6s infinite;vertical-align:middle;}\n' +
'</style></head><body>\n' +
'<div class="card">\n' +
'  <div class="brand">' + (shopName || "Spin &amp; Win") + '</div>\n' +
'  <h1>' + statusIcon + ' Voucher Status</h1>\n' +
'  <div class="badge" style="color:' + statusColor + ';background:' + statusBg + ';">' + statusIcon + ' ' + statusLabel + '</div>\n' +
  errorHtml +
'  <div class="row"><span class="key">Voucher ID</span><span class="val">' + escRecordId + '</span></div>\n' +
'  <div class="row"><span class="key">Prize / Reward</span><span class="val prize">' + (prize || "—") + '</span></div>\n' +
'  <div class="row"><span class="key">Customer</span><span class="val">' + (customer || "—") + '</span></div>\n' +
'  <div class="row"><span class="key">Phone</span><span class="val">' + customerNumber + '</span></div>\n' +
'  <div class="row"><span class="key">Purchase Amount</span><span class="val">₹' + amount + '</span></div>\n' +
'  <div class="row"><span class="key">Date / Time</span><span class="val">' + dateText + '</span></div>\n' +
'  <div class="row"><span class="key">Valid Until</span><span class="val" id="expiry">' + expiryText + '</span></div>\n' +
'  <div class="foot">' + (expiryEpoch > 0
    ? '<span id="countdown"></span>\n'
    : '') + '<span class="live"></span>Live status &middot; last updated ' + checkedAt + '<br>This page refreshes automatically every 20 seconds.</div>\n' +
'</div>\n' +
'<script>\n' +
'(function(){\n' +
'  var exp = ' + expiryEpoch + ';\n' +
'  var cd = document.getElementById("countdown");\n' +
'  if (exp && cd) {\n' +
'    cd.textContent = "Expires in: " + fmt(exp - Date.now());\n' +
'    setInterval(function(){ var d = exp - Date.now(); cd.textContent = d > 0 ? "Expires in: " + fmt(d) : "Expired now. Reloading..."; }, 1000);\n' +
'  }\n' +
'  setInterval(function(){ if (!document.hidden) location.reload(); }, 20000);\n' +
'  function fmt(ms){\n' +
'    if (ms <= 0) return "Just expired";\n' +
'    var s = Math.floor(ms/1000), h = Math.floor(s/3600), m = Math.floor((s%3600)/60), sec = s%60;\n' +
'    if (h>0) return h+"h "+m+"m";\n' +
'    if (m>0) return m+"m "+sec+"s";\n' +
'    return sec+"s";\n' +
'  }\n' +
'})();\n' +
'</script>\n' +
'</body></html>';
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

function toNumber_(value, fallback) {
  var parsed = Number(value);
  return isNaN(parsed) ? fallback : parsed;
}

function round2_(value) {
  return Math.round(toNumber_(value, 0) * 100) / 100;
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

const SETTINGS_KEY = "spinwin_settings_v1";
const RECORDS_KEY = "spinwin_history_v1";
const ACTIVE_TAB_KEY = "spinwin_active_tab_v1";
const APPS_SCRIPT_URL_KEY = "spinwin_apps_script_url_v1";

const STATUS_PENDING = "Pending";
const STATUS_COMPLETED = "Completed";
const STATUS_REJECTED = "Rejected";
const STATUS_EXPIRED = "Expired";

const DEFAULT_SETTINGS = {
  shopName: "Lucky Shop",
  shopLogoUrl: "",
  expiryHours: 24,
  manualDateEnabled: false,
  manualDateTime: "",
  appsScriptUrl: "",
  prizes: [
    { id: uid(), name: "10% Discount", probability: 40, enabled: true },
    { id: uid(), name: "Free Drink", probability: 25, enabled: true },
    { id: uid(), name: "Buy 1 Get 1", probability: 10, enabled: true },
    { id: uid(), name: "Try Again", probability: 24.9, enabled: true },
    { id: uid(), name: "Grand Prize", probability: 0.1, enabled: true }
  ]
};

const state = {
  settings: loadSettings(),
  records: loadRecords(),
  wheelAngle: 0,
  spinning: false,
  spinLockedForEntry: false,
  lastEntrySignature: "",
  currentResultId: null,
  selectedRecordIds: new Set(),
  logoImage: null
};

const el = {
  shopNameDisplay: byId("shopNameDisplay"),
  customerName: byId("customerName"),
  customerNumber: byId("customerNumber"),
  purchaseAmount: byId("purchaseAmount"),
  spinBtn: byId("spinBtn"),
  saveBtn: byId("saveBtn"),
  downloadBtn: byId("downloadBtn"),
  shareBtn: byId("shareBtn"),
  statusText: byId("statusText"),
  wheelCanvas: byId("wheelCanvas"),
  couponCanvas: byId("couponCanvas"),
  resultEmpty: byId("resultEmpty"),
  resultDetails: byId("resultDetails"),
  resultRecordId: byId("resultRecordId"),
  resultShop: byId("resultShop"),
  resultCustomer: byId("resultCustomer"),
  resultCustomerNumber: byId("resultCustomerNumber"),
  resultAmount: byId("resultAmount"),
  resultPrize: byId("resultPrize"),
  resultStatus: byId("resultStatus"),
  resultDate: byId("resultDate"),
  resultExpiry: byId("resultExpiry"),
  shopNameInput: byId("shopNameInput"),
  shopLogoUrlInput: byId("shopLogoUrlInput"),
  expiryHoursInput: byId("expiryHoursInput"),
  appsScriptUrlInput: byId("appsScriptUrlInput"),
  manualDateToggle: byId("manualDateToggle"),
  manualDateInput: byId("manualDateInput"),
  prizeTableBody: byId("prizeTableBody"),
  addPrizeBtn: byId("addPrizeBtn"),
  saveSettingsBtn: byId("saveSettingsBtn"),
  resetDataBtn: byId("resetDataBtn"),
  prizeRowTemplate: byId("prizeRowTemplate"),
  recordSearchInput: byId("recordSearchInput"),
  recordFilterSortSelect: byId("recordFilterSortSelect"),
  refreshRecordsBtn: byId("refreshRecordsBtn"),
  recordsSummary: byId("recordsSummary"),
  recordsTableBody: byId("recordsTableBody"),
  selectAllRecords: byId("selectAllRecords"),
  bulkActionSelect: byId("bulkActionSelect"),
  applyBulkActionBtn: byId("applyBulkActionBtn"),
  clearSelectionBtn: byId("clearSelectionBtn"),
  editModal: byId("editModal"),
  editRecordForm: byId("editRecordForm"),
  editRecordId: byId("editRecordId"),
  editCustomerName: byId("editCustomerName"),
  editCustomerNumber: byId("editCustomerNumber"),
  editPurchaseAmount: byId("editPurchaseAmount"),
  editPrizeWon: byId("editPrizeWon"),
  editExpiryDate: byId("editExpiryDate"),
  cancelEditBtn: byId("cancelEditBtn"),
  toastContainer: byId("toastContainer"),
  bulkActionsBar: byId("bulkActionsBar"),
  tabMenuToggle: byId("tabMenuToggle"),
  tabNav: byId("tabNav"),
  tabButtons: Array.from(document.querySelectorAll("[data-tab-target]")),
  tabPanels: Array.from(document.querySelectorAll("[data-tab-panel]"))
};

const wheelCtx = el.wheelCanvas.getContext("2d");
const couponCtx = el.couponCanvas.getContext("2d");

init();

function init() {
  bindEvents();
  initTabs();
  persistAppsScriptUrl(state.settings.appsScriptUrl);
  syncSettingsToInputs();
  if (state.settings.shopLogoUrl) {
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => {
      state.logoImage = img;
    };
    img.src = state.settings.shopLogoUrl;
  }
  renderPrizeTable();
  drawWheel();
  drawCouponPlaceholder();
  renderRecordsTable();
  renderResult();
  refreshUI();

  if (getAppsScriptUrl()) {
    void hydrateFromSheets();
  }
}

async function hydrateFromSheets() {
  await loadSettingsFromSheets(false);
  await loadRecordsFromSheets(false);
}

function bindEvents() {
  el.tabButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      setActiveTab(btn.dataset.tabTarget);
    });
  });

  if (el.tabMenuToggle && el.tabNav) {
    el.tabMenuToggle.addEventListener("click", () => {
      const shouldOpen = !el.tabNav.classList.contains("open");
      toggleTabMenu(shouldOpen);
    });
  }

  document.addEventListener("click", (event) => {
    if (!el.tabNav || !el.tabNav.classList.contains("open")) return;
    const target = event.target;
    if (!(target instanceof Element)) return;
    if (target.closest(".tab-nav-wrap") || target.closest("#tabMenuToggle")) return;
    toggleTabMenu(false);
  });

  window.addEventListener("resize", () => {
    if (window.innerWidth > 760) {
      toggleTabMenu(false);
    }
  });

  el.spinBtn.addEventListener("click", onSpinClick);
  el.saveBtn.addEventListener("click", onSaveClick);
  el.downloadBtn.addEventListener("click", onDownloadCoupon);
  el.shareBtn.addEventListener("click", onShareWhatsApp);

  el.addPrizeBtn.addEventListener("click", () => {
    syncPrizeDraftFromTable();
    state.settings.prizes.push({
      id: uid(),
      name: "",
      probability: 0,
      enabled: true
    });
    renderPrizeTable();
    persistSettings();
    refreshUI();
  });

  el.saveSettingsBtn.addEventListener("click", () => void onSaveSettings());
  el.resetDataBtn.addEventListener("click", onResetData);
  el.customerName.addEventListener("input", onEntryChange);
  el.customerNumber.addEventListener("input", onEntryChange);
  el.purchaseAmount.addEventListener("input", onEntryChange);
  el.shopNameInput.addEventListener("input", () => {
    el.shopNameDisplay.textContent = safeText(el.shopNameInput.value, DEFAULT_SETTINGS.shopName);
  });

  el.manualDateToggle.addEventListener("change", () => {
    el.manualDateInput.disabled = !el.manualDateToggle.checked;
  });

  el.recordSearchInput.addEventListener("input", renderRecordsTable);
  el.recordFilterSortSelect.addEventListener("change", renderRecordsTable);
  el.refreshRecordsBtn.addEventListener("click", async () => {
    await loadSettingsFromSheets(false);
    await loadRecordsFromSheets(true);
  });
  el.recordsTableBody.addEventListener("click", onRecordsActionClick);
  el.recordsTableBody.addEventListener("change", onRecordSelectionChange);
  if (el.selectAllRecords) {
    el.selectAllRecords.addEventListener("change", onSelectAllVisibleChange);
  }
  if (el.applyBulkActionBtn) {
    el.applyBulkActionBtn.addEventListener("click", () => void onApplyBulkAction());
  }
  if (el.clearSelectionBtn) {
    el.clearSelectionBtn.addEventListener("click", onClearSelection);
  }

  el.editRecordForm.addEventListener("submit", onEditRecordSubmit);
  el.cancelEditBtn.addEventListener("click", closeEditModal);
  el.editModal.addEventListener("click", (event) => {
    const target = event.target;
    if (target && target.getAttribute("data-close-modal") === "true") {
      closeEditModal();
    }
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeEditModal();
  });
}

function initTabs() {
  if (!el.tabButtons.length || !el.tabPanels.length) return;
  const preferred = safeText(localStorage.getItem(ACTIVE_TAB_KEY), "customer-spin");
  setActiveTab(preferred, false);
}

function setActiveTab(tabId, persist = true) {
  if (!el.tabButtons.length || !el.tabPanels.length) return;
  const available = el.tabPanels.map((panel) => panel.dataset.tabPanel);
  const nextTab = available.includes(tabId) ? tabId : available[0];
  if (!nextTab) return;

  if (nextTab === "settings" && persist) {
    const password = prompt("Enter password to access settings:");
    if (password !== "111111") {
      alert("Incorrect password.");
      return;
    }
  }

  el.tabButtons.forEach((btn) => {
    const isActive = btn.dataset.tabTarget === nextTab;
    btn.classList.toggle("is-active", isActive);
    btn.setAttribute("aria-selected", String(isActive));
  });

  el.tabPanels.forEach((panel) => {
    const isActive = panel.dataset.tabPanel === nextTab;
    panel.classList.toggle("hidden", !isActive);
    panel.classList.toggle("is-active", isActive);
  });

  if (persist) {
    localStorage.setItem(ACTIVE_TAB_KEY, nextTab);
  }
  toggleTabMenu(false);

  if (nextTab === "customer-spin") {
    drawWheel();
  }
}

function toggleTabMenu(forceOpen) {
  if (!el.tabNav || !el.tabMenuToggle) return;
  const open = typeof forceOpen === "boolean" ? forceOpen : !el.tabNav.classList.contains("open");
  el.tabNav.classList.toggle("open", open);
  el.tabMenuToggle.setAttribute("aria-expanded", String(open));
}

function onEntryChange() {
  const signature = getEntrySignature();
  if (signature !== state.lastEntrySignature) {
    state.spinLockedForEntry = false;
    state.currentResultId = null;
    state.lastEntrySignature = signature;
    renderResult();
    drawCouponPlaceholder();
    updateStatus("Entry changed. Ready for a new spin.");
  }
  refreshUI();
}

async function onSaveSettings() {
  const prizes = collectPrizesFromTable();
  if (!prizes.length) {
    alert("Please add at least one prize.");
    return;
  }

  state.settings.shopName = safeText(el.shopNameInput.value, "Lucky Shop");
  state.settings.shopLogoUrl = (el.shopLogoUrlInput.value || "").trim();
  if (state.settings.shopLogoUrl) {
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => {
      state.logoImage = img;
      renderResult();
    };
    img.src = state.settings.shopLogoUrl;
  } else {
    state.logoImage = null;
  }
  state.settings.expiryHours = Math.max(0, parseInt(el.expiryHoursInput.value || "0", 10));
  state.settings.appsScriptUrl = (el.appsScriptUrlInput.value || "").trim();
  state.settings.manualDateEnabled = !!el.manualDateToggle.checked;
  state.settings.manualDateTime = el.manualDateInput.value || "";
  state.settings.prizes = prizes;

  persistSettings();
  renderPrizeTable();
  drawWheel();
  renderRecordsTable();
  renderResult();
  refreshUI();

  if (!getAppsScriptUrl()) {
    updateStatus("Settings saved locally.");
    showToast("Settings saved locally. Add Apps Script URL to sync with Google Sheets.", "info");
    return;
  }

  try {
    updateStatus("Saving settings to Google Sheets...");
    const res = await syncSaveSettings(state.settings);
    if (!res.ok) {
      throw new Error(res.message || "Settings sync failed.");
    }

    if (res.settings) {
      applySettingsFromServer(res.settings, true);
    }

    updateStatus("Settings saved to Google Sheets.");
    showToast("Settings saved to Google Sheets.", "success");
  } catch (err) {
    console.error(err);
    updateStatus("Settings saved locally. Google Sheets sync failed.");
    showToast(`Settings saved locally, but sync failed: ${err.message}`, "error");
  }
}

function onResetData() {
  const ok = confirm("Reset all local settings and prize records?");
  if (!ok) return;

  localStorage.removeItem(SETTINGS_KEY);
  localStorage.removeItem(RECORDS_KEY);

  state.settings = loadSettings();
  state.records = [];
  state.currentResultId = null;
  state.spinning = false;
  state.spinLockedForEntry = false;
  state.lastEntrySignature = "";
  state.wheelAngle = 0;
  state.selectedRecordIds.clear();

  syncSettingsToInputs();
  renderPrizeTable();
  drawWheel();
  drawCouponPlaceholder();
  renderRecordsTable();
  renderResult();
  refreshUI();
  updateStatus("Local data reset complete.");
  showToast("Local settings and records were reset.", "info");
}

async function onSpinClick() {
  if (state.spinning) return;

  const customerName = safeText(el.customerName.value);
  const customerNumber = safeText(el.customerNumber.value);
  const amount = parseFloat(el.purchaseAmount.value || "0");

  if (!customerName) {
    alert("Please enter customer name.");
    return;
  }
  if (!Number.isFinite(amount) || amount < 0) {
    alert("Please enter a valid purchase amount (>= 0).");
    return;
  }
  if (state.spinLockedForEntry) {
    alert("This customer entry already spun once. Change name or amount for a new spin.");
    return;
  }

  const enabledPrizes = getEnabledPrizes(state.settings.prizes);
  const weightedPrizes = enabledPrizes.filter((p) => num(p.probability) > 0);

  if (!enabledPrizes.length) {
    alert("No enabled prizes. Please enable at least one prize in settings.");
    return;
  }
  if (!weightedPrizes.length) {
    alert("All enabled prizes have 0% probability. Please set at least one prize above 0.");
    return;
  }

  const winner = chooseWeightedPrize(weightedPrizes);
  const wheelPrizes = enabledPrizes;
  const winnerIndex = wheelPrizes.findIndex((p) => p.id === winner.id);
  if (winnerIndex < 0) {
    alert("Prize selection issue. Please save settings and try again.");
    return;
  }

  state.spinning = true;
  refreshUI();
  updateStatus("Spinning...");

  await animateSpin(winnerIndex, wheelPrizes.length);

  const spunDate = getSpinDate();
  const expiryDate = addHours(spunDate, state.settings.expiryHours);
  const nowIso = new Date().toISOString();
  const record = {
    recordId: createRecordId(),
    shopName: state.settings.shopName,
    customerName,
    customerNumber,
    amount: round2(amount),
    prize: winner.name,
    status: STATUS_PENDING,
    dateTimeIso: spunDate.toISOString(),
    expiryIso: expiryDate.toISOString(),
    createdAtIso: nowIso,
    updatedAtIso: nowIso,
    synced: false
  };

  state.records.unshift(record);
  state.currentResultId = record.recordId;
  state.spinLockedForEntry = true;
  state.lastEntrySignature = getEntrySignature();
  state.spinning = false;

  persistRecords();
  renderRecordsTable();
  renderResult();
  drawCouponFromRecord(record);
  refreshUI();

  updateStatus(`Result: ${record.prize}`);
  showToast(`Spin complete. Record ID: ${record.recordId}`, "success");
}

async function onSaveClick() {
  const record = getCurrentRecord();
  if (!record) {
    alert("Spin first before saving.");
    return;
  }

  if (!getAppsScriptUrl()) {
    alert("Please add Apps Script Web App URL in settings.");
    return;
  }

  try {
    updateStatus("Saving to Google Sheets...");
    const res = await syncCreateRecord(record);
    if (!res.ok) {
      throw new Error(res.message || "Save failed");
    }

    markRecordSynced(record.recordId, true);
    updateStatus("Saved to Google Sheets successfully.");
    refreshUI();

    if (res.mode === "no-cors") {
      showToast("Request sent (no-cors). Check your Google Sheet to confirm.", "info");
    } else {
      showToast("Record saved to Google Sheets.", "success");
    }
  } catch (err) {
    console.error(err);
    updateStatus("Save failed. Check Apps Script deployment and URL.");
    showToast(`Save failed: ${err.message}`, "error");
  }
}

function onDownloadCoupon() {
  const record = getCurrentRecord();
  if (!record) {
    alert("Spin first to generate coupon.");
    return;
  }

  const fileName = `spin-win-${sanitizeFileName(record.customerName)}-${Date.now()}.png`;
  const dataUrl = el.couponCanvas.toDataURL("image/png");
  const a = document.createElement("a");
  a.href = dataUrl;
  a.download = fileName;
  a.click();
}

async function onShareWhatsApp() {
  const record = getCurrentRecord();
  if (!record) {
    alert("Spin first before sharing.");
    return;
  }

  const summary = buildShareText(record);
  const blob = await canvasToBlob(el.couponCanvas);
  const fileName = `spin-win-${sanitizeFileName(record.customerName)}.png`;
  const file = blob ? new File([blob], fileName, { type: "image/png" }) : null;

  if (navigator.share && file && navigator.canShare && navigator.canShare({ files: [file] })) {
    try {
      await navigator.share({
        title: `${record.shopName} - Spin & Win`,
        text: summary,
        files: [file]
      });
      return;
    } catch (err) {
      console.warn("Web share canceled or failed", err);
    }
  }

  const waUrl = `https://wa.me/?text=${encodeURIComponent(summary)}`;
  window.open(waUrl, "_blank", "noopener");
}

function onRecordsActionClick(event) {
  const btn = event.target.closest("button[data-action]");
  if (!btn) return;

  const action = btn.dataset.action;
  const recordId = btn.dataset.id;
  if (!recordId) return;

  if (action === "edit") {
    const record = getRecordById(recordId);
    if (!record) return;
    openEditModal(record);
    return;
  }

  if (action === "complete") {
    void updateRecordStatus(recordId, STATUS_COMPLETED);
    return;
  }

  if (action === "reject") {
    void updateRecordStatus(recordId, STATUS_REJECTED);
    return;
  }

  if (action === "delete") {
    void deleteRecord(recordId);
  }
}

function onRecordSelectionChange(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;
  if (!target.classList.contains("record-select-checkbox")) return;

  const recordId = safeText(target.dataset.id);
  if (!recordId) return;

  if (target.checked) {
    state.selectedRecordIds.add(recordId);
  } else {
    state.selectedRecordIds.delete(recordId);
  }
  updateBulkSelectionUI();
}

function onSelectAllVisibleChange() {
  if (!el.selectAllRecords) return;
  const visibleIds = getVisibleRecords().map((record) => record.recordId);
  if (el.selectAllRecords.checked) {
    visibleIds.forEach((id) => state.selectedRecordIds.add(id));
  } else {
    visibleIds.forEach((id) => state.selectedRecordIds.delete(id));
  }
  renderRecordsTable();
}

function onClearSelection() {
  state.selectedRecordIds.clear();
  renderRecordsTable();
}

async function onApplyBulkAction() {
  const action = safeText(el.bulkActionSelect && el.bulkActionSelect.value);
  const selectedIds = getSelectedExistingRecordIds();
  if (!action) {
    showToast("Choose a bulk action first.", "info");
    return;
  }
  if (!selectedIds.length) {
    showToast("Select at least one record.", "info");
    return;
  }

  if (action === "delete") {
    const confirmed = confirm(`Delete ${selectedIds.length} selected record(s)?`);
    if (!confirmed) return;
    await runBulkDelete(selectedIds);
    return;
  }

  const nextStatus = normalizeManualStatus(action);
  await runBulkStatusUpdate(selectedIds, nextStatus);
}

function getSelectedExistingRecordIds() {
  const existing = new Set(state.records.map((record) => record.recordId));
  return Array.from(state.selectedRecordIds).filter((id) => existing.has(id));
}

async function runBulkStatusUpdate(recordIds, nextStatus) {
  const nowIso = new Date().toISOString();
  let changed = 0;

  state.records = state.records.map((record) => {
    if (!recordIds.includes(record.recordId)) return record;
    changed += 1;
    return normalizeRecord({
      ...record,
      status: nextStatus,
      updatedAtIso: nowIso,
      synced: false
    });
  });

  if (!changed) return;

  persistRecords();
  renderResult();
  if (state.currentResultId) {
    const current = getCurrentRecord();
    if (current) drawCouponFromRecord(current);
  }
  renderRecordsTable();
  refreshUI();

  if (!getAppsScriptUrl()) {
    showToast(`Updated ${changed} record(s) locally. Add Apps Script URL to sync.`, "info");
    state.selectedRecordIds.clear();
    renderRecordsTable();
    return;
  }

  const syncPayload = { status: nextStatus, updatedAtIso: nowIso };
  const results = await Promise.allSettled(
    recordIds.map((recordId) => syncUpdateRecord(recordId, syncPayload))
  );

  let synced = 0;
  const syncedIds = [];
  results.forEach((result, idx) => {
    if (result.status === "fulfilled" && result.value && result.value.ok) {
      syncedIds.push(recordIds[idx]);
      synced += 1;
    }
  });

  if (syncedIds.length) {
    const syncedSet = new Set(syncedIds);
    state.records = state.records.map((record) => (
      syncedSet.has(record.recordId)
        ? { ...record, synced: true, updatedAtIso: new Date().toISOString() }
        : record
    ));
    persistRecords();
    renderRecordsTable();
    refreshUI();
  }

  const failed = recordIds.length - synced;
  if (failed > 0) {
    showToast(`Bulk status done. Synced ${synced}, failed ${failed}.`, "error");
  } else {
    showToast(`Bulk status updated for ${synced} record(s).`, "success");
  }

  state.selectedRecordIds.clear();
  renderRecordsTable();
}

async function runBulkDelete(recordIds) {
  const removeSet = new Set(recordIds);
  state.records = state.records.filter((record) => !removeSet.has(record.recordId));

  if (state.currentResultId && removeSet.has(state.currentResultId)) {
    state.currentResultId = null;
    renderResult();
    drawCouponPlaceholder();
  }

  persistRecords();
  renderRecordsTable();
  refreshUI();

  if (!getAppsScriptUrl()) {
    showToast(`Deleted ${recordIds.length} record(s) locally. Add Apps Script URL to sync.`, "info");
    state.selectedRecordIds.clear();
    renderRecordsTable();
    return;
  }

  const results = await Promise.allSettled(
    recordIds.map((recordId) => syncDeleteRecord(recordId))
  );

  let synced = 0;
  results.forEach((result) => {
    if (result.status === "fulfilled" && result.value && result.value.ok) {
      synced += 1;
    }
  });

  const failed = recordIds.length - synced;
  if (failed > 0) {
    showToast(`Bulk delete done. Synced ${synced}, failed ${failed}.`, "error");
  } else {
    showToast(`Deleted ${synced} record(s).`, "success");
  }

  state.selectedRecordIds.clear();
  renderRecordsTable();
}

async function onEditRecordSubmit(event) {
  event.preventDefault();
  const recordId = safeText(el.editRecordId.value);
  if (!recordId) return;

  const customerName = safeText(el.editCustomerName.value);
  const customerNumber = safeText(el.editCustomerNumber.value);
  const amount = num(el.editPurchaseAmount.value);
  const prize = safeText(el.editPrizeWon.value);
  const expiryIso = localInputToIso(el.editExpiryDate.value);

  if (!customerName) {
    alert("Customer name is required.");
    return;
  }
  if (!Number.isFinite(amount) || amount <= 0) {
    alert("Purchase amount must be greater than 0.");
    return;
  }
  if (!prize) {
    alert("Prize is required.");
    return;
  }
  if (!expiryIso) {
    alert("Please choose a valid expiry date.");
    return;
  }

  const updates = {
    customerName,
    customerNumber,
    amount: round2(amount),
    prize,
    expiryIso,
    updatedAtIso: new Date().toISOString()
  };

  applyLocalRecordUpdate(recordId, updates);
  closeEditModal();
  showToast(`Record ${recordId} updated locally.`, "success");

  if (!getAppsScriptUrl()) {
    showToast("Add Apps Script URL to sync edit changes.", "info");
    return;
  }

  try {
    const res = await syncUpdateRecord(recordId, updates);
    if (!res.ok) {
      throw new Error(res.message || "Update sync failed.");
    }
    markRecordSynced(recordId, true);
    showToast(`Record ${recordId} synced to Google Sheets.`, "success");
  } catch (err) {
    console.error(err);
    showToast(`Update sync failed: ${err.message}`, "error");
  }
}

async function updateRecordStatus(recordId, nextStatus) {
  const updates = {
    status: normalizeManualStatus(nextStatus),
    updatedAtIso: new Date().toISOString()
  };

  applyLocalRecordUpdate(recordId, updates);
  showToast(`Status set to ${updates.status}.`, "success");

  if (!getAppsScriptUrl()) {
    showToast("Add Apps Script URL to sync status updates.", "info");
    return;
  }

  try {
    const res = await syncUpdateRecord(recordId, updates);
    if (!res.ok) {
      throw new Error(res.message || "Status sync failed.");
    }
    markRecordSynced(recordId, true);
    showToast("Status synced to Google Sheets.", "success");
  } catch (err) {
    console.error(err);
    showToast(`Status sync failed: ${err.message}`, "error");
  }
}

async function deleteRecord(recordId) {
  const record = getRecordById(recordId);
  if (!record) return;

  const ok = confirm(`Delete record ${recordId} permanently?`);
  if (!ok) return;

  removeLocalRecord(recordId);
  showToast(`Record ${recordId} deleted locally.`, "info");

  if (!getAppsScriptUrl()) {
    showToast("Add Apps Script URL to sync deletions.", "info");
    return;
  }

  try {
    const res = await syncDeleteRecord(recordId);
    if (!res.ok) {
      throw new Error(res.message || "Delete sync failed.");
    }
    showToast(`Record ${recordId} deleted from Google Sheets.`, "success");
  } catch (err) {
    console.error(err);
    showToast(`Delete sync failed: ${err.message}`, "error");
  }
}

async function loadRecordsFromSheets(showToastMessage) {
  if (!getAppsScriptUrl()) {
    if (showToastMessage) showToast("Add Apps Script URL first.", "info");
    return;
  }

  try {
    const res = await apiListRecords();
    if (!res.ok) {
      throw new Error(res.message || "Failed to load records");
    }

    const serverRecords = normalizeRecords(Array.isArray(res.records) ? res.records : [])
      .map((record) => ({ ...record, synced: true }));

    const localUnsynced = state.records.filter((record) => !record.synced);
    const map = new Map();
    serverRecords.forEach((record) => map.set(record.recordId, record));
    localUnsynced.forEach((record) => {
      if (!map.has(record.recordId)) {
        map.set(record.recordId, record);
      }
    });

    state.records = Array.from(map.values()).sort((a, b) => getRecordTime(b) - getRecordTime(a));
    persistRecords();
    renderRecordsTable();
    refreshUI();

    if (showToastMessage) {
      showToast(`Loaded ${serverRecords.length} records from Google Sheets.`, "success");
    }
  } catch (err) {
    console.error(err);
    showToast(`Sync failed: ${err.message}`, "error");
  }
}

async function loadSettingsFromSheets(showToastMessage) {
  if (!getAppsScriptUrl()) {
    if (showToastMessage) showToast("Add Apps Script URL first.", "info");
    return;
  }

  try {
    const res = await apiGetSettings();
    if (!res.ok) {
      throw new Error(res.message || "Failed to load settings.");
    }

    if (!res.settings || typeof res.settings !== "object") {
      if (showToastMessage) {
        showToast("No saved settings found in Google Sheets yet.", "info");
      }
      return;
    }

    applySettingsFromServer(res.settings, true);
    if (showToastMessage) {
      showToast("Settings loaded from Google Sheets.", "success");
    }
  } catch (err) {
    console.error(err);
    if (showToastMessage) {
      showToast(`Settings sync failed: ${err.message}`, "error");
    }
  }
}

function applySettingsFromServer(serverSettings, keepCurrentUrl = true) {
  const currentUrl = safeText(state.settings.appsScriptUrl);
  const next = { ...(serverSettings || {}) };
  if (keepCurrentUrl && !safeText(next.appsScriptUrl)) {
    next.appsScriptUrl = currentUrl;
  }

  state.settings = mergeSettings(next);
  persistSettings();
  syncSettingsToInputs();
  renderPrizeTable();
  drawWheel();
  renderResult();
  refreshUI();
}

function syncSettingsToInputs() {
  el.shopNameInput.value = state.settings.shopName || "";
  el.shopLogoUrlInput.value = state.settings.shopLogoUrl || "";
  el.expiryHoursInput.value = String(num(state.settings.expiryHours));
  el.appsScriptUrlInput.value = state.settings.appsScriptUrl || "";
  el.manualDateToggle.checked = !!state.settings.manualDateEnabled;
  el.manualDateInput.value = state.settings.manualDateTime || "";
  el.manualDateInput.disabled = !state.settings.manualDateEnabled;
}

function renderPrizeTable() {
  el.prizeTableBody.innerHTML = "";
  state.settings.prizes.forEach((prize) => {
    const row = el.prizeRowTemplate.content.firstElementChild.cloneNode(true);
    row.dataset.id = prize.id;
    row.querySelector(".prize-name-input").value = prize.name;
    row.querySelector(".prize-prob-input").value = String(num(prize.probability));
    row.querySelector(".prize-enabled-input").checked = !!prize.enabled;
    row.querySelector(".remove-prize-btn").addEventListener("click", () => {
      syncPrizeDraftFromTable();
      state.settings.prizes = state.settings.prizes.filter((p) => p.id !== prize.id);
      renderPrizeTable();
      persistSettings();
      drawWheel();
      refreshUI();
    });
    el.prizeTableBody.appendChild(row);
  });
}

function collectPrizesFromTable() {
  const rows = Array.from(el.prizeTableBody.querySelectorAll("tr"));
  return rows.map((row) => {
    const id = row.dataset.id || uid();
    const name = safeText(row.querySelector(".prize-name-input").value, "Unnamed Prize");
    const probabilityRaw = row.querySelector(".prize-prob-input").value;
    const probability = Math.max(0, num(probabilityRaw));
    const enabled = !!row.querySelector(".prize-enabled-input").checked;
    return { id, name, probability, enabled };
  });
}

function syncPrizeDraftFromTable() {
  const hasRows = el.prizeTableBody && el.prizeTableBody.querySelector("tr");
  if (!hasRows) return;
  state.settings.prizes = collectPrizesFromTable();
}

function refreshUI() {
  const hasCustomer = !!safeText(el.customerName.value);
  const amount = parseFloat(el.purchaseAmount.value || "0");
  const hasAmount = Number.isFinite(amount) && amount > 0;
  const canSpin = hasCustomer && hasAmount && !state.spinning && !state.spinLockedForEntry;
  const currentRecord = getCurrentRecord();

  el.spinBtn.disabled = !canSpin;
  el.saveBtn.disabled = !currentRecord;
  el.downloadBtn.disabled = !currentRecord;
  el.shareBtn.disabled = !currentRecord;
  el.shopNameDisplay.textContent = state.settings.shopName || "Lucky Shop";
  el.saveBtn.textContent = currentRecord && currentRecord.synced
    ? "Re-Sync to Google Sheets"
    : "Save";
}

function renderResult() {
  const record = getCurrentRecord();
  if (!record) {
    el.resultEmpty.classList.remove("hidden");
    el.resultDetails.classList.add("hidden");
    return;
  }

  el.resultEmpty.classList.add("hidden");
  el.resultDetails.classList.remove("hidden");

  el.resultRecordId.textContent = record.recordId;
  el.resultShop.textContent = record.shopName || state.settings.shopName;
  el.resultCustomer.textContent = record.customerName;
  el.resultCustomerNumber.textContent = record.customerNumber || "-";
  el.resultAmount.textContent = formatAmount(record.amount);
  el.resultPrize.textContent = record.prize;
  el.resultStatus.textContent = getEffectiveStatus(record);
  el.resultDate.textContent = formatIsoDateTime(record.dateTimeIso);
  el.resultExpiry.textContent = formatIsoDateTime(record.expiryIso);
}

function renderRecordsTable() {
  pruneSelectionForMissingRecords();
  const visible = getVisibleRecords();
  const counts = getStatusCounts(state.records);
  const selectedCount = getSelectedExistingRecordIds().length;
  el.recordsSummary.textContent =
    `Total: ${state.records.length} | Selected: ${selectedCount} | Pending: ${counts.pending} | Completed: ${counts.completed} | Rejected: ${counts.rejected} | Expired: ${counts.expired}`;

  el.recordsTableBody.innerHTML = "";
  if (!visible.length) {
    const row = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 10;
    td.textContent = "No matching prize records.";
    td.className = "centered";
    row.appendChild(td);
    el.recordsTableBody.appendChild(row);
    updateBulkSelectionUI();
    return;
  }

  visible.forEach((record) => {
    const tr = document.createElement("tr");
    const effectiveStatus = getEffectiveStatus(record);
    const isSelected = state.selectedRecordIds.has(record.recordId);
    if (isSelected) {
      tr.classList.add("record-row-selected");
    }

    const selectTd = document.createElement("td");
    selectTd.className = "centered";
    const check = document.createElement("input");
    check.type = "checkbox";
    check.className = "record-select-checkbox";
    check.dataset.id = record.recordId;
    check.checked = isSelected;
    check.setAttribute("aria-label", `Select record ${record.recordId}`);
    selectTd.appendChild(check);
    tr.appendChild(selectTd);

    appendCell(tr, record.recordId);
    appendCell(tr, record.customerName);
    appendCell(tr, record.customerNumber || "-");
    appendCell(tr, formatAmount(record.amount));
    appendCell(tr, record.prize);

    const statusTd = document.createElement("td");
    const statusPill = document.createElement("span");
    statusPill.className = `status-pill status-${effectiveStatus.toLowerCase()}`;
    statusPill.textContent = effectiveStatus;
    statusTd.appendChild(statusPill);
    tr.appendChild(statusTd);

    appendCell(tr, formatIsoDateTime(record.dateTimeIso));
    appendCell(tr, formatIsoDateTime(record.expiryIso));

    const actionTd = document.createElement("td");
    const wrap = document.createElement("div");
    wrap.className = "record-actions";

    wrap.appendChild(createActionButton("Edit", "btn-ghost", "edit", record.recordId));
    wrap.appendChild(createActionButton("Completed", "btn-secondary", "complete", record.recordId, effectiveStatus === STATUS_COMPLETED));
    wrap.appendChild(createActionButton("Rejected", "btn-danger", "reject", record.recordId, effectiveStatus === STATUS_REJECTED));
    wrap.appendChild(createActionButton("Delete", "btn-danger", "delete", record.recordId));

    actionTd.appendChild(wrap);
    tr.appendChild(actionTd);
    el.recordsTableBody.appendChild(tr);
  });
  updateBulkSelectionUI();
}

function pruneSelectionForMissingRecords() {
  const existing = new Set(state.records.map((record) => record.recordId));
  Array.from(state.selectedRecordIds).forEach((id) => {
    if (!existing.has(id)) {
      state.selectedRecordIds.delete(id);
    }
  });
}

function updateBulkSelectionUI() {
  const selectedCount = getSelectedExistingRecordIds().length;
  if (el.applyBulkActionBtn) {
    el.applyBulkActionBtn.disabled = selectedCount < 1;
  }

  if (el.bulkActionsBar) {
    const showBar = selectedCount > 0;
    el.bulkActionsBar.classList.toggle("hidden", !showBar);
    el.bulkActionsBar.classList.toggle("visible", showBar);
  }

  if (!el.selectAllRecords) return;
  const visible = getVisibleRecords();
  const visibleIds = visible.map((record) => record.recordId);
  const selectedVisibleCount = visibleIds.filter((id) => state.selectedRecordIds.has(id)).length;

  const hasVisible = visibleIds.length > 0;
  const allVisibleSelected = hasVisible && selectedVisibleCount === visibleIds.length;
  const partiallySelected = selectedVisibleCount > 0 && !allVisibleSelected;

  el.selectAllRecords.checked = allVisibleSelected;
  el.selectAllRecords.indeterminate = partiallySelected;
}

function appendCell(tr, text) {
  const td = document.createElement("td");
  td.textContent = text;
  tr.appendChild(td);
}

function createActionButton(label, toneClass, action, recordId, disabled = false) {
  const btn = document.createElement("button");
  btn.className = `btn ${toneClass} btn-small`;
  btn.textContent = label;
  btn.dataset.action = action;
  btn.dataset.id = recordId;
  btn.disabled = !!disabled;
  return btn;
}

function getVisibleRecords() {
  const search = safeText(el.recordSearchInput.value).toLowerCase();
  const mode = el.recordFilterSortSelect.value || "recent_old";

  let rows = state.records.filter((record) => {
    if (!search) return true;
    return record.recordId.toLowerCase().includes(search);
  });

  if (mode === "completed") {
    rows = rows.filter((record) => getEffectiveStatus(record) === STATUS_COMPLETED);
  } else if (mode === "pending") {
    rows = rows.filter((record) => getEffectiveStatus(record) === STATUS_PENDING);
  } else if (mode === "rejected") {
    rows = rows.filter((record) => getEffectiveStatus(record) === STATUS_REJECTED);
  } else if (mode === "expired") {
    rows = rows.filter((record) => getEffectiveStatus(record) === STATUS_EXPIRED);
  }

  if (mode === "old_recent") {
    rows.sort((a, b) => getRecordTime(a) - getRecordTime(b));
  } else if (mode === "big_small") {
    rows.sort((a, b) => num(b.amount) - num(a.amount));
  } else if (mode === "small_big") {
    rows.sort((a, b) => num(a.amount) - num(b.amount));
  } else {
    rows.sort((a, b) => getRecordTime(b) - getRecordTime(a));
  }

  return rows;
}

function getStatusCounts(records) {
  const counts = { pending: 0, completed: 0, rejected: 0, expired: 0 };
  records.forEach((record) => {
    const status = getEffectiveStatus(record);
    if (status === STATUS_COMPLETED) counts.completed += 1;
    else if (status === STATUS_REJECTED) counts.rejected += 1;
    else if (status === STATUS_EXPIRED) counts.expired += 1;
    else counts.pending += 1;
  });
  return counts;
}

function getEffectiveStatus(record) {
  const manual = normalizeManualStatus(record.status);
  if (manual === STATUS_COMPLETED || manual === STATUS_REJECTED) return manual;
  if (manual === STATUS_EXPIRED) return STATUS_EXPIRED;

  const expiry = new Date(record.expiryIso).getTime();
  if (Number.isFinite(expiry) && expiry < Date.now()) return STATUS_EXPIRED;
  return STATUS_PENDING;
}

function openEditModal(record) {
  el.editRecordId.value = record.recordId;
  el.editCustomerName.value = record.customerName;
  el.editCustomerNumber.value = record.customerNumber || "";
  el.editPurchaseAmount.value = String(num(record.amount));
  el.editPrizeWon.value = record.prize;
  el.editExpiryDate.value = isoToLocalInput(record.expiryIso);
  el.editModal.classList.remove("hidden");
  el.editModal.setAttribute("aria-hidden", "false");
}

function closeEditModal() {
  if (el.editModal.classList.contains("hidden")) return;
  el.editModal.classList.add("hidden");
  el.editModal.setAttribute("aria-hidden", "true");
}

function applyLocalRecordUpdate(recordId, updates) {
  const idx = state.records.findIndex((record) => record.recordId === recordId);
  if (idx < 0) return;

  const merged = {
    ...state.records[idx],
    ...updates,
    status: updates.status ? normalizeManualStatus(updates.status) : normalizeManualStatus(state.records[idx].status),
    updatedAtIso: updates.updatedAtIso || new Date().toISOString(),
    synced: false
  };
  state.records[idx] = normalizeRecord(merged);
  persistRecords();

  if (state.currentResultId === recordId) {
    renderResult();
    drawCouponFromRecord(state.records[idx]);
  }
  renderRecordsTable();
  refreshUI();
}

function removeLocalRecord(recordId) {
  state.records = state.records.filter((record) => record.recordId !== recordId);
  persistRecords();

  if (state.currentResultId === recordId) {
    state.currentResultId = null;
    renderResult();
    drawCouponPlaceholder();
  }
  renderRecordsTable();
  refreshUI();
}

function getCurrentRecord() {
  if (!state.currentResultId) return null;
  return getRecordById(state.currentResultId);
}

function getRecordById(recordId) {
  return state.records.find((record) => record.recordId === recordId) || null;
}

function markRecordSynced(recordId, synced) {
  const record = getRecordById(recordId);
  if (!record) return;
  record.synced = !!synced;
  record.updatedAtIso = new Date().toISOString();
  persistRecords();
  renderRecordsTable();
  refreshUI();
}

function getSpinDate() {
  if (!state.settings.manualDateEnabled) return new Date();
  const manual = state.settings.manualDateTime;
  if (!manual) return new Date();
  const parsed = new Date(manual);
  if (Number.isNaN(parsed.getTime())) return new Date();
  return parsed;
}

function drawWheel() {
  const prizes = getEnabledPrizes(state.settings.prizes);
  const ctx = wheelCtx;
  const w = el.wheelCanvas.width;
  const h = el.wheelCanvas.height;
  const cx = w / 2;
  const cy = h / 2;
  const radius = Math.min(cx, cy) - 8;

  ctx.clearRect(0, 0, w, h);

  if (!prizes.length) {
    ctx.fillStyle = "#fff5ea";
    ctx.beginPath();
    ctx.arc(cx, cy, radius, 0, Math.PI * 2);
    ctx.fill();
    ctx.fillStyle = "#9aa3af";
    ctx.font = "700 24px Sora, sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText("No Enabled Prizes", cx, cy);
    return;
  }

  const arc = (Math.PI * 2) / prizes.length;
  for (let i = 0; i < prizes.length; i++) {
    const start = state.wheelAngle + i * arc;
    const end = start + arc;
    const color = colorForSlice(i);

    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, radius, start, end);
    ctx.closePath();
    ctx.fillStyle = color;
    ctx.fill();

    ctx.save();
    ctx.translate(cx, cy);
    ctx.rotate(start + arc / 2);
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillStyle = "#1f2937";
    const fontSize = prizes.length <= 6 ? 22 : prizes.length <= 8 ? 19 : 16;
    ctx.font = `700 ${fontSize}px DM Sans, sans-serif`;

    const maxTextWidth = radius * 0.55;
    const lines = wrapPrizeTextLines(ctx, prizes[i].name, maxTextWidth, 3);
    const lineHeight = Math.round(fontSize * 1.05);
    const startY = -((lines.length - 1) * lineHeight) / 2;
    const textX = radius * 0.62;
    lines.forEach((line, lineIdx) => {
      ctx.fillText(line, textX, startY + lineIdx * lineHeight);
    });
    ctx.restore();
  }

  ctx.beginPath();
  ctx.arc(cx, cy, 42, 0, Math.PI * 2);
  ctx.fillStyle = "#fff";
  ctx.fill();
  ctx.lineWidth = 7;
  ctx.strokeStyle = "#c5d7f2";
  ctx.stroke();

  ctx.fillStyle = "#173b7a";
  ctx.font = "800 16px Sora, sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText("SPIN", cx, cy);
}

function animateSpin(winnerIndex, totalSegments) {
  return new Promise((resolve) => {
    const pointerAngle = -Math.PI / 2;
    const arc = (Math.PI * 2) / totalSegments;
    const current = normalizeAngle(state.wheelAngle);
    const winnerCenter = pointerAngle - (winnerIndex + 0.5) * arc;
    const extraSpins = randInt(6, 9);

    let target = winnerCenter + extraSpins * Math.PI * 2;
    while (target <= current) {
      target += Math.PI * 2;
    }

    const duration = 4300;
    const start = performance.now();

    const step = (now) => {
      const t = Math.min(1, (now - start) / duration);
      const eased = easeOutCubic(t);
      state.wheelAngle = current + (target - current) * eased;
      drawWheel();
      if (t < 1) {
        requestAnimationFrame(step);
      } else {
        state.wheelAngle = normalizeAngle(target);
        drawWheel();
        resolve();
      }
    };

    requestAnimationFrame(step);
  });
}

function drawCouponPlaceholder() {
  const ctx = couponCtx;
  const w = el.couponCanvas.width;
  const h = el.couponCanvas.height;
  ctx.clearRect(0, 0, w, h);

  const grad = ctx.createLinearGradient(0, 0, w, h);
  grad.addColorStop(0, "#edf4ff");
  grad.addColorStop(1, "#eef6ff");
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, w, h);

  ctx.fillStyle = "#ffffff";
  ctx.strokeStyle = "#c9d8f3";
  ctx.lineWidth = 2;
  ctx.fillRect(32, 32, w - 64, h - 64);
  ctx.strokeRect(32, 32, w - 64, h - 64);

  ctx.fillStyle = "#173b7a";
  ctx.font = "700 50px Sora, sans-serif";
  ctx.fillText("Spin & Win Coupon", 64, 116);

  ctx.fillStyle = "#334155";
  ctx.font = "700 28px DM Sans, sans-serif";
  ctx.fillText("Details and QR will appear after spin.", 64, 174);

  ctx.fillStyle = "#e2e8f0";
  ctx.fillRect(w - 284, 188, 220, 220);
  ctx.fillStyle = "#64748b";
  ctx.font = "700 22px DM Sans, sans-serif";
  ctx.fillText("QR Preview", w - 244, 304);
}

function drawCouponFromRecord(record) {
  if (!record) {
    drawCouponPlaceholder();
    return;
  }

  const ctx = couponCtx;
  const w = el.couponCanvas.width;
  const h = el.couponCanvas.height;
  const status = getEffectiveStatus(record);

  ctx.clearRect(0, 0, w, h);

  const grad = ctx.createLinearGradient(0, 0, w, h);
  grad.addColorStop(0, "#edf4ff");
  grad.addColorStop(0.5, "#f8fbff");
  grad.addColorStop(1, "#eef6ff");
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, w, h);

  const cardX = 28;
  const cardY = 24;
  const cardW = w - 56;
  const cardH = h - 48;
  ctx.fillStyle = "#ffffff";
  ctx.strokeStyle = "#c9d8f3";
  ctx.lineWidth = 2;
  ctx.fillRect(cardX, cardY, cardW, cardH);
  ctx.strokeRect(cardX, cardY, cardW, cardH);

  if (state.logoImage) {
    ctx.globalAlpha = 0.1;
    const logoW = 200;
    const logoH = 200;
    const logoX = (w - logoW) / 2;
    const logoY = (h - logoH) / 2;
    ctx.drawImage(state.logoImage, logoX, logoY, logoW, logoH);
    ctx.globalAlpha = 1;
  }

  ctx.fillStyle = "#173b7a";
  ctx.fillRect(cardX, cardY, cardW, 88);

  ctx.fillStyle = "#ffffff";
  ctx.font = "700 40px Sora, sans-serif";
  const shopTitle = fitTextWithEllipsis(ctx, record.shopName || state.settings.shopName, cardW - 260);
  ctx.fillText(shopTitle, cardX + 26, cardY + 56);

  ctx.fillStyle = "#1f2937";
  ctx.font = "700 30px Sora, sans-serif";
  ctx.fillText("Spin & Win Prize Coupon", cardX + 28, cardY + 138);

  const statusColor = status === STATUS_COMPLETED
    ? "#0b6b3f"
    : status === STATUS_REJECTED
      ? "#8c1f1f"
      : status === STATUS_EXPIRED
        ? "#4b5563"
        : "#8a6700";

  const statusBg = status === STATUS_COMPLETED
    ? "#dff8ea"
    : status === STATUS_REJECTED
      ? "#ffe4e4"
      : status === STATUS_EXPIRED
        ? "#e8ebf0"
        : "#fff5cd";

  const statusX = cardX + cardW - 208;
  const statusY = cardY + 104;
  ctx.fillStyle = statusBg;
  ctx.fillRect(statusX, statusY, 176, 44);
  ctx.fillStyle = statusColor;
  ctx.font = "700 22px DM Sans, sans-serif";
  ctx.fillText(`Status: ${status}`, statusX + 14, statusY + 29);

  const qrSize = 212;
  const qrX = cardX + cardW - qrSize - 34;
  const qrY = cardY + 174;

  const leftX = cardX + 32;
  const leftW = qrX - leftX - 24;
  const details = [
    { label: "Unique ID", value: record.recordId },
    { label: "Customer Name", value: record.customerName },
    { label: "Customer Number", value: record.customerNumber || "-" },
    { label: "Purchase Amount", value: formatAmount(record.amount) },
    { label: "Prize Won", value: record.prize },
    { label: "Date & Time", value: formatIsoDateTime(record.dateTimeIso) },
    { label: "Expiry Date", value: formatIsoDateTime(record.expiryIso) }
  ];

  let rowY = cardY + 198;
  const rowGap = 44;
  details.forEach((item) => {
    ctx.font = "700 15px DM Sans, sans-serif";
    const labelText = `${item.label}: `;
    const labelWidth = ctx.measureText(labelText).width;
    ctx.fillStyle = "#64748b";
    ctx.fillText(labelText, leftX, rowY);

    ctx.font = "700 17px DM Sans, sans-serif";
    const valueWidth = Math.max(90, leftW - labelWidth);
    const valueText = fitTextWithEllipsis(ctx, String(item.value || "-"), valueWidth);
    ctx.fillStyle = "#0f172a";
    ctx.fillText(valueText, leftX + labelWidth, rowY);
    rowY += rowGap;
  });

  const qrData = buildQrPayload(record, status);
  const qrOk = drawQrOnCanvas(ctx, qrData.payload, qrX, qrY, qrSize);

  ctx.fillStyle = "#334155";
  ctx.font = "700 18px DM Sans, sans-serif";
  ctx.fillText(
    qrData.isLive ? "" : "Scan for full coupon details",
    qrX,
    qrY + qrSize + 28
  );

  if (!qrOk) {
    ctx.fillStyle = "#b91c1c";
    ctx.font = "700 14px DM Sans, sans-serif";
    ctx.fillText("QR library not loaded", qrX, qrY + qrSize + 50);
  }

  ctx.fillStyle = "#0f9d8a";
  ctx.font = "700 20px DM Sans, sans-serif";
  ctx.fillText("Show this coupon at cashier to claim your reward.", cardX + 32, cardY + cardH - 28);
}

function buildShareText(record) {
  return [
    `${record.shopName || state.settings.shopName} - Spin & Win`,
    `ID: ${record.recordId}`,
    `Customer: ${record.customerName}`,
    `Number: ${record.customerNumber || "-"}`,
    `Amount: ${formatAmount(record.amount)}`,
    `Prize: ${record.prize}`,
    `Status: ${getEffectiveStatus(record)}`,
    `Date: ${formatIsoDateTime(record.dateTimeIso)}`,
    `Expiry: ${formatIsoDateTime(record.expiryIso)}`
  ].join("\n");
}

function buildQrPayload(record, status) {
  const liveUrl = buildLiveCouponUrl(record.recordId);
  if (liveUrl) {
    return { payload: liveUrl, isLive: true };
  }

  return {
    payload: [
      "Spin & Win Coupon",
      `Shop: ${record.shopName || state.settings.shopName}`,
      `ID: ${record.recordId}`,
      `Customer: ${record.customerName}`,
      `Number: ${record.customerNumber || "-"}`,
      `Amount: ${formatAmount(record.amount)}`,
      `Prize: ${record.prize}`,
      `Status: ${status}`,
      `Date: ${formatIsoDateTime(record.dateTimeIso)}`,
      `Expiry: ${formatIsoDateTime(record.expiryIso)}`
    ].join("\n"),
    isLive: false
  };
}

function buildLiveCouponUrl(recordId) {
  const baseUrl = getAppsScriptUrl();
  const id = safeText(recordId);
  if (!baseUrl || !id) return "";
  return appendQueryParams(baseUrl, { action: "get_record", recordId: id });
}

function drawQrOnCanvas(ctx, payload, x, y, size) {
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(x, y, size, size);
  ctx.strokeStyle = "#e2e8f0";
  ctx.lineWidth = 2;
  ctx.strokeRect(x, y, size, size);

  if (typeof window.qrcode !== "function") return false;

  try {
    const qr = window.qrcode(0, "M");
    qr.addData(payload);
    qr.make();

    const modules = qr.getModuleCount();
    const cell = Math.floor(size / modules);
    const drawSize = cell * modules;
    const offsetX = x + Math.floor((size - drawSize) / 2);
    const offsetY = y + Math.floor((size - drawSize) / 2);

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(offsetX, offsetY, drawSize, drawSize);
    ctx.fillStyle = "#0f172a";

    for (let row = 0; row < modules; row++) {
      for (let col = 0; col < modules; col++) {
        if (qr.isDark(row, col)) {
          ctx.fillRect(offsetX + col * cell, offsetY + row * cell, cell, cell);
        }
      }
    }
    return true;
  } catch (err) {
    console.warn("QR render failed", err);
    return false;
  }
}

async function apiListRecords() {
  const url = getAppsScriptUrl();
  if (!url) throw new Error("Apps Script URL not set.");

  const fetchUrl = appendQueryParams(url, { action: "list" });
  try {
    const response = await fetch(fetchUrl, { method: "GET" });
    return parseApiResponse(response);
  } catch (err) {
    return jsonpRequest(url, { action: "list" });
  }
}

async function apiGetSettings() {
  const url = getAppsScriptUrl();
  if (!url) throw new Error("Apps Script URL not set.");

  const fetchUrl = appendQueryParams(url, { action: "settings" });
  try {
    const response = await fetch(fetchUrl, { method: "GET" });
    return parseApiResponse(response);
  } catch (err) {
    return jsonpRequest(url, { action: "settings" });
  }
}

async function syncCreateRecord(record) {
  return apiMutate("POST", { action: "create", record });
}

async function syncSaveSettings(settings) {
  return apiMutate("POST", { action: "save_settings", settings });
}

async function syncUpdateRecord(recordId, updates) {
  return apiMutate("PUT", { action: "update", recordId, updates });
}

async function syncDeleteRecord(recordId) {
  return apiMutate("DELETE", { action: "delete", recordId });
}

async function apiMutate(method, payload) {
  const url = getAppsScriptUrl();
  if (!url) throw new Error("Apps Script URL not set.");

  const headers = { "Content-Type": "text/plain;charset=utf-8" };
  const body = JSON.stringify(payload);

  if (method === "POST") {
    try {
      const response = await fetch(url, { method: "POST", headers, body });
      return parseApiResponse(response);
    } catch (err) {
      await fetch(url, { method: "POST", mode: "no-cors", headers, body });
      return { ok: true, mode: "no-cors", message: "Request sent in no-cors mode." };
    }
  }

  try {
    const directResponse = await fetch(url, { method, headers, body });
    const direct = await parseApiResponse(directResponse);
    if (direct.ok) return direct;

    if (direct.status !== 405 && direct.status !== 404 && direct.status !== 501) {
      return direct;
    }
  } catch (err) {
    console.warn(`${method} direct call failed, trying POST override`, err);
  }

  const overridePayload = JSON.stringify({ ...payload, _method: method });
  try {
    const response = await fetch(url, { method: "POST", headers, body: overridePayload });
    return parseApiResponse(response);
  } catch (err) {
    await fetch(url, { method: "POST", mode: "no-cors", headers, body: overridePayload });
    return { ok: true, mode: "no-cors", message: "Request sent in no-cors mode." };
  }
}

async function parseApiResponse(response) {
  const text = await response.text();
  let data = {};
  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    data = { raw: text };
  }

  const ok = response.ok && data.ok !== false;
  return {
    ok,
    status: response.status,
    mode: "cors",
    ...data
  };
}

function jsonpRequest(url, params) {
  return new Promise((resolve, reject) => {
    const callbackName = `spinwin_jsonp_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
    const timeoutMs = 10000;
    const script = document.createElement("script");

    const cleanup = () => {
      delete window[callbackName];
      script.remove();
    };

    const timeout = window.setTimeout(() => {
      cleanup();
      reject(new Error("JSONP request timeout."));
    }, timeoutMs);

    window[callbackName] = (data) => {
      window.clearTimeout(timeout);
      cleanup();
      resolve({
        ok: data && data.ok !== false,
        status: 200,
        mode: "jsonp",
        ...(data || {})
      });
    };

    const query = { ...params, callback: callbackName };
    script.src = appendQueryParams(url, query);
    script.onerror = () => {
      window.clearTimeout(timeout);
      cleanup();
      reject(new Error("JSONP request failed."));
    };

    document.body.appendChild(script);
  });
}

function appendQueryParams(url, params) {
  const hasQuery = url.includes("?");
  const search = new URLSearchParams(params);
  return `${url}${hasQuery ? "&" : "?"}${search.toString()}`;
}

function getEnabledPrizes(prizes) {
  return (prizes || []).filter((p) => p.enabled);
}

function chooseWeightedPrize(prizes) {
  const total = prizes.reduce((sum, p) => sum + num(p.probability), 0);
  let r = Math.random() * total;
  for (let i = 0; i < prizes.length; i++) {
    r -= num(prizes[i].probability);
    if (r <= 0) return prizes[i];
  }
  return prizes[prizes.length - 1];
}

function getEntrySignature() {
  return `${safeText(el.customerName.value)}|${safeText(el.customerNumber.value)}|${num(el.purchaseAmount.value).toFixed(2)}`;
}

function updateStatus(text) {
  el.statusText.textContent = text;
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    if (!raw) return mergeSettings({});
    const parsed = JSON.parse(raw);
    return mergeSettings(parsed);
  } catch {
    return mergeSettings({});
  }
}

function mergeSettings(settings) {
  settings = settings || {};
  const merged = cloneDefaultSettings();
  const fallbackAppsScriptUrl = getBootstrapAppsScriptUrl();
  merged.shopName = safeText(settings.shopName, merged.shopName);
  merged.shopLogoUrl = safeText(settings.shopLogoUrl, merged.shopLogoUrl);
  merged.expiryHours = Math.max(0, num(settings.expiryHours));
  merged.manualDateEnabled = !!settings.manualDateEnabled;
  merged.manualDateTime = typeof settings.manualDateTime === "string" ? settings.manualDateTime : "";
  merged.appsScriptUrl = safeText(
    typeof settings.appsScriptUrl === "string" ? settings.appsScriptUrl : "",
    fallbackAppsScriptUrl
  );
  merged.prizes = Array.isArray(settings.prizes) && settings.prizes.length
    ? settings.prizes.map((prize) => ({
      id: prize.id || uid(),
      name: safeText(prize.name, "Unnamed Prize"),
      probability: Math.max(0, num(prize.probability)),
      enabled: prize.enabled !== false
    }))
    : merged.prizes;
  return merged;
}

function persistSettings() {
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(state.settings));
  persistAppsScriptUrl(state.settings.appsScriptUrl);
}

function loadRecords() {
  try {
    const raw = localStorage.getItem(RECORDS_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return normalizeRecords(parsed);
  } catch {
    return [];
  }
}

function normalizeRecords(records) {
  return records
    .map(normalizeRecord)
    .filter((record) => !!record.recordId)
    .sort((a, b) => getRecordTime(b) - getRecordTime(a));
}

function normalizeRecord(record) {
  const now = new Date();
  const created = toIso(record.createdAtIso || record.createdAt || now);
  const dateTimeIso = toIso(record.dateTimeIso || record.dateTime || created);
  const expiryBase = toIso(record.expiryIso || record.expiryDateIso || record.expiryDate || addHours(new Date(dateTimeIso), 24));

  return {
    recordId: safeText(record.recordId || record.id, createRecordId()),
    shopName: safeText(record.shopName, DEFAULT_SETTINGS.shopName),
    customerName: safeText(record.customerName, "Unknown"),
    customerNumber: safeText(record.customerNumber, ""),
    amount: round2(num(record.amount)),
    prize: safeText(record.prize, "Unknown Prize"),
    status: normalizeManualStatus(record.status),
    dateTimeIso,
    expiryIso: expiryBase,
    createdAtIso: created,
    updatedAtIso: toIso(record.updatedAtIso || record.updatedAt || now),
    synced: !!record.synced
  };
}

function persistRecords() {
  localStorage.setItem(RECORDS_KEY, JSON.stringify(state.records));
}

function formatDateTime(date) {
  return new Intl.DateTimeFormat(undefined, {
    year: "numeric",
    month: "short",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit"
  }).format(date);
}

function formatIsoDateTime(iso) {
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return "-";
  return formatDateTime(date);
}

function formatAmount(amount) {
  return num(amount).toLocaleString(undefined, {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function addHours(date, hours) {
  return new Date(date.getTime() + num(hours) * 60 * 60 * 1000);
}

function showToast(message, type = "info") {
  const toast = document.createElement("div");
  toast.className = `toast toast-${type}`;
  toast.textContent = message;
  el.toastContainer.appendChild(toast);
  window.setTimeout(() => toast.remove(), 3200);
}

function getAppsScriptUrl() {
  return safeText(state.settings.appsScriptUrl, getBootstrapAppsScriptUrl());
}

function getBootstrapAppsScriptUrl() {
  const fromQuery = getAppsScriptUrlFromQuery();
  if (fromQuery) return fromQuery;
  try {
    return safeText(localStorage.getItem(APPS_SCRIPT_URL_KEY));
  } catch {
    return "";
  }
}

function getAppsScriptUrlFromQuery() {
  try {
    const query = new URLSearchParams(window.location.search);
    return safeText(query.get("appsScriptUrl"));
  } catch {
    return "";
  }
}

function persistAppsScriptUrl(url) {
  const clean = safeText(url);
  try {
    if (clean) {
      localStorage.setItem(APPS_SCRIPT_URL_KEY, clean);
    } else {
      localStorage.removeItem(APPS_SCRIPT_URL_KEY);
    }
  } catch {
    // Ignore storage errors in restricted environments
  }
}

function getRecordTime(record) {
  return new Date(record.dateTimeIso || record.createdAtIso).getTime() || 0;
}

function normalizeManualStatus(status) {
  const value = safeText(status, STATUS_PENDING).toLowerCase();
  if (value === "approved") return STATUS_COMPLETED;
  if (value === STATUS_COMPLETED.toLowerCase()) return STATUS_COMPLETED;
  if (value === STATUS_REJECTED.toLowerCase()) return STATUS_REJECTED;
  if (value === STATUS_EXPIRED.toLowerCase()) return STATUS_EXPIRED;
  return STATUS_PENDING;
}

function createRecordId() {
  const now = new Date();
  const stamp = now.toISOString().replace(/[-:TZ.]/g, "").slice(0, 8);
  const suffix = Math.random().toString(36).slice(2, 6).toUpperCase();
  return `DPM${stamp}${suffix}`;
}

function toIso(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return new Date().toISOString();
  return date.toISOString();
}

function isoToLocalInput(iso) {
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return "";
  const tzOffset = date.getTimezoneOffset() * 60000;
  return new Date(date.getTime() - tzOffset).toISOString().slice(0, 16);
}

function localInputToIso(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return date.toISOString();
}

function byId(id) {
  return document.getElementById(id);
}

function num(value) {
  const number = parseFloat(value);
  return Number.isFinite(number) ? number : 0;
}

function round2(value) {
  return Math.round(num(value) * 100) / 100;
}

function safeText(value, fallback = "") {
  const text = String(value || "").trim();
  return text || fallback;
}

function uid() {
  return Math.random().toString(36).slice(2, 10);
}

function normalizeAngle(angle) {
  const full = Math.PI * 2;
  return ((angle % full) + full) % full;
}

function randInt(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

function easeOutCubic(t) {
  return 1 - Math.pow(1 - t, 3);
}

function colorForSlice(index) {
  const colors = ["#dbe7ff", "#c5daf8", "#b6d0f3", "#a8c5ee", "#9bbbe8", "#89abdd"];
  return colors[index % colors.length];
}

function wrapPrizeTextLines(ctx, text, maxWidth, maxLines = 3) {
  const words = String(text || "").trim().split(/\s+/).filter(Boolean);
  if (!words.length) return ["-"];

  const lines = [];
  let current = "";
  let idx = 0;

  while (idx < words.length) {
    const candidate = current ? `${current} ${words[idx]}` : words[idx];
    if (ctx.measureText(candidate).width <= maxWidth) {
      current = candidate;
      idx += 1;
      continue;
    }

    if (!current) {
      lines.push(fitTextWithEllipsis(ctx, words[idx], maxWidth));
      idx += 1;
    } else {
      lines.push(current);
      current = "";
    }

    if (lines.length >= maxLines - 1) break;
  }

  const hasRemaining = idx < words.length;
  const tail = hasRemaining
    ? [current, ...words.slice(idx)].filter(Boolean).join(" ")
    : current;

  if (tail) {
    lines.push(hasRemaining ? fitTextWithEllipsis(ctx, tail, maxWidth) : tail);
  }

  return lines.slice(0, maxLines);
}

function fitTextWithEllipsis(ctx, text, maxWidth) {
  const value = String(text || "").trim();
  if (!value) return "";
  if (ctx.measureText(value).width <= maxWidth) return value;

  let cut = value;
  while (cut.length > 1 && ctx.measureText(`${cut}...`).width > maxWidth) {
    cut = cut.slice(0, -1);
  }
  return `${cut}...`;
}

function trimText(text, maxLen) {
  const value = String(text || "");
  return value.length <= maxLen ? value : `${value.slice(0, maxLen - 1)}...`;
}

function sanitizeFileName(name) {
  return String(name || "customer").replace(/[^a-z0-9_-]+/gi, "-").replace(/-+/g, "-");
}

function canvasToBlob(canvas) {
  return new Promise((resolve) => {
    canvas.toBlob((blob) => resolve(blob), "image/png");
  });
}

function cloneDefaultSettings() {
  return JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
}

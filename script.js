const SETTINGS_KEY = "spinwin_settings_v2";
const RECORDS_KEY = "spinwin_history_v1";
const ACTIVE_TAB_KEY = "spinwin_active_tab_v1";
const APPS_SCRIPT_URL_KEY = "spinwin_apps_script_url_v1";
const MEMBERS_KEY = "spinwin_members_v1";
const CREDIT_RATE_PER_100 = 1; // 100 Credits = ₹1.00

const STATUS_PENDING = "Pending";
const STATUS_COMPLETED = "Completed";
const STATUS_REJECTED = "Rejected";
const STATUS_EXPIRED = "Expired";

const DEFAULT_SETTINGS = {
  shopName: "Lucky Shop",
  shopLogoUrl: "",
  expiryHours: 24,
  spinDuration: 5,
  manualDateEnabled: false,
  manualDateTime: "",
  appsScriptUrl: "https://script.google.com/macros/s/AKfycbxSIzCpUu7VCgAOViM7XpsYKQWFTY-Ug5EoiucocVIfw9YniK8JuJcj9FLDwBqRRlTb_w/exec",
  conditions: [
    {
      id: uid(),
      itemName: "Jeans Pant",
      label: "Jeans Pant (Wheel 1)",
      icon: "👖",
      minAmount: 0,
      prizes: [
        { id: uid(), name: "Leather Belt Free", probability: 30, enabled: true },
        { id: uid(), name: "20% Off Denim", probability: 25, enabled: true },
        { id: uid(), name: "Cap Free", probability: 15, enabled: true },
        { id: uid(), name: "Try Again", probability: 29.9, enabled: true },
        { id: uid(), name: "Grand Prize", probability: 0.1, enabled: true }
      ]
    },
    {
      id: uid(),
      itemName: "Shirt",
      label: "Shirt (Wheel 2)",
      icon: "👔",
      minAmount: 0,
      prizes: [
        { id: uid(), name: "Silk Tie Free", probability: 25, enabled: true },
        { id: uid(), name: "15% Off Next Shirt", probability: 30, enabled: true },
        { id: uid(), name: "Cufflinks Free", probability: 15, enabled: true },
        { id: uid(), name: "Try Again", probability: 29.9, enabled: true },
        { id: uid(), name: "Grand Prize", probability: 0.1, enabled: true }
      ]
    }
  ],
  // Keep flat prizes for backward compatibility / migration
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
  wheelHovered: false,
  spinLockedForEntry: false,
  lastEntrySignature: "",
  currentResultId: null,
  selectedRecordIds: new Set(),
  selectedMemberIds: new Set(),
  logoImage: null,
  recordsPage: 1,
  recordsPageSize: 10,
  soundMuted: localStorage.getItem("spinwin_sound_muted") === "true",
  selectedConditionId: null,
  members: loadMembers(),
  memberSearchQuery: "",
  memberFilter: "all",
  membersPage: 1,
  membersPageSize: 10,
  activeHistoryMemberId: null,
  activeRecordForCredit: null
};

const el = {
  shopNameDisplay: byId("shopNameDisplay"),
  customerName: byId("customerName"),
  customerNumber: byId("customerNumber"),
  purchaseAmount: byId("purchaseAmount"),
  saveBtn: byId("saveBtn"),
  downloadBtn: byId("downloadBtn"),
  shareBtn: byId("shareBtn"),
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
  testAppsScriptBtn: byId("testAppsScriptBtn"),
  appsScriptStatusMsg: byId("appsScriptStatusMsg"),
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
  exportRecordsBtn: byId("exportRecordsBtn"),
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
  tabPanels: Array.from(document.querySelectorAll("[data-tab-panel]")),
  soundToggleBtn: byId("soundToggleBtn"),
  soundIcon: byId("soundIcon"),
  wheelPointer: byId("wheelPointer"),
  confettiCanvas: byId("confettiCanvas"),
  wheelWrap: document.querySelector(".wheel-wrap"),
  spinDurationInput: byId("spinDurationInput"),
  spinTimePills: byId("spinTimePills"),
  winnerModal: byId("winnerModal"),
  winnerPrizeName: byId("winnerPrizeName"),
  winnerCustomerMsg: byId("winnerCustomerMsg"),
  winnerRecordId: byId("winnerRecordId"),
  winnerViewCouponBtn: byId("winnerViewCouponBtn"),
  winnerSpinAgainBtn: byId("winnerSpinAgainBtn"),
  recordsPaginationBar: byId("recordsPaginationBar"),
  paginationRangeText: byId("paginationRangeText"),
  pageSizeSelect: byId("pageSizeSelect"),
  firstPageBtn: byId("firstPageBtn"),
  prevPageBtn: byId("prevPageBtn"),
  nextPageBtn: byId("nextPageBtn"),
  lastPageBtn: byId("lastPageBtn"),
  paginationPages: byId("paginationPages"),
  spinConditionSelect: byId("spinConditionSelect"),
  conditionTabsBar: byId("conditionTabsBar"),
  conditionEditorCard: byId("conditionEditorCard"),
  conditionBadgeActive: byId("conditionBadgeActive"),
  conditionEditorTitle: byId("conditionEditorTitle"),
  deleteConditionBtn: byId("deleteConditionBtn"),
  conditionItemNameInput: byId("conditionItemNameInput"),
  conditionLabelInput: byId("conditionLabelInput"),
  conditionIconInput: byId("conditionIconInput"),
  conditionMinAmountInput: byId("conditionMinAmountInput"),
  addConditionBtn: byId("addConditionBtn"),
  conditionProbSummary: byId("conditionProbSummary"),
  winnerItemPill: byId("winnerItemPill"),
  winnerItemText: byId("winnerItemText"),

  // Members Section Elements
  statTotalMembers: byId("statTotalMembers"),
  statMembersSubtitle: byId("statMembersSubtitle"),
  statTotalCredits: byId("statTotalCredits"),
  statCreditsWorth: byId("statCreditsWorth"),
  statTotalEarned: byId("statTotalEarned"),
  statEarnedWorth: byId("statEarnedWorth"),
  statTotalCashPaid: byId("statTotalCashPaid"),
  statRedeemedCredits: byId("statRedeemedCredits"),
  memberSearchInput: byId("memberSearchInput"),
  clearMemberSearchBtn: byId("clearMemberSearchBtn"),
  memberFilterSelect: byId("memberFilterSelect"),
  membersTableBody: byId("membersTableBody"),
  selectAllMembers: byId("selectAllMembers"),
  memberBulkActionsBar: byId("memberBulkActionsBar"),
  memberBulkActionSelect: byId("memberBulkActionSelect"),
  applyMemberBulkActionBtn: byId("applyMemberBulkActionBtn"),
  clearMemberSelectionBtn: byId("clearMemberSelectionBtn"),
  membersEmptyState: byId("membersEmptyState"),
  membersEmptyMsg: byId("membersEmptyMsg"),
  emptyAddMemberBtn: byId("emptyAddMemberBtn"),
  openNewMemberModalBtn: byId("openNewMemberModalBtn"),
  openAddCreditModalBtn: byId("openAddCreditModalBtn"),
  openDeductModalBtn: byId("openDeductModalBtn"),
  pushMembersBtn: byId("pushMembersBtn"),
  refreshMembersBtn: byId("refreshMembersBtn"),
  exportMembersBtn: byId("exportMembersBtn"),
  membersPaginationBar: byId("membersPaginationBar"),
  membersPaginationRangeText: byId("membersPaginationRangeText"),
  membersPageSizeSelect: byId("membersPageSizeSelect"),
  membersFirstPageBtn: byId("membersFirstPageBtn"),
  membersPrevPageBtn: byId("membersPrevPageBtn"),
  membersNextPageBtn: byId("membersNextPageBtn"),
  membersLastPageBtn: byId("membersLastPageBtn"),
  membersPaginationPages: byId("membersPaginationPages"),

  // Member Modal
  memberModal: byId("memberModal"),
  memberModalTitle: byId("memberModalTitle"),
  memberForm: byId("memberForm"),
  memberFormId: byId("memberFormId"),
  memberFormName: byId("memberFormName"),
  memberFormPhone: byId("memberFormPhone"),
  memberFormType: byId("memberFormType"),
  memberFormInitialCredits: byId("memberFormInitialCredits"),
  memberInitialCreditsGroup: byId("memberInitialCreditsGroup"),
  initialCreditsHint: byId("initialCreditsHint"),
  memberFormNotes: byId("memberFormNotes"),
  saveMemberBtn: byId("saveMemberBtn"),

  // Credit Modal
  creditModal: byId("creditModal"),
  creditForm: byId("creditForm"),
  creditMemberSelect: byId("creditMemberSelect"),
  creditQuickAddMemberBtn: byId("creditQuickAddMemberBtn"),
  creditCurrentBalanceRow: byId("creditCurrentBalanceRow"),
  creditCurrentBalanceText: byId("creditCurrentBalanceText"),
  creditPurchaseAmount: byId("creditPurchaseAmount"),
  creditAmountInput: byId("creditAmountInput"),
  creditCashWorthPreview: byId("creditCashWorthPreview"),
  creditNoteInput: byId("creditNoteInput"),

  // Deduct Modal
  deductModal: byId("deductModal"),
  deductForm: byId("deductForm"),
  deductMemberSelect: byId("deductMemberSelect"),
  deductAvailableText: byId("deductAvailableText"),
  deductRedeemMode: byId("deductRedeemMode"),
  modePurchaseDiscountBtn: byId("modePurchaseDiscountBtn"),
  modeCashPayoutBtn: byId("modeCashPayoutBtn"),
  purchaseDiscountFields: byId("purchaseDiscountFields"),
  deductBillAmountInput: byId("deductBillAmountInput"),
  deductCreditsInput: byId("deductCreditsInput"),
  deductCashInput: byId("deductCashInput"),
  deductValueLabel: byId("deductValueLabel"),
  discountNetBanner: byId("discountNetBanner"),
  deductNetPayableDisplay: byId("deductNetPayableDisplay"),
  deductBreakdownDisplay: byId("deductBreakdownDisplay"),
  cashPayoutBanner: byId("cashPayoutBanner"),
  deductPayoutDisplay: byId("deductPayoutDisplay"),
  deductNoteInput: byId("deductNoteInput"),
  deductSendWhatsAppCheck: byId("deductSendWhatsAppCheck"),
  confirmDeductBtn: byId("confirmDeductBtn"),

  // Member History Modal
  memberHistoryModal: byId("memberHistoryModal"),
  historyMemberName: byId("historyMemberName"),
  historyMemberSub: byId("historyMemberSub"),
  historyCurrentCredits: byId("historyCurrentCredits"),
  historyCurrentCash: byId("historyCurrentCash"),
  historyLifetimeEarned: byId("historyLifetimeEarned"),
  historyLifetimeCashed: byId("historyLifetimeCashed"),
  historyTableBody: byId("historyTableBody"),

  // Cross-Tab Integration Elements
  customerMemberBadge: byId("customerMemberBadge"),
  customerMemberBadgeText: byId("customerMemberBadgeText"),
  customerMemberQuickCreditBtn: byId("customerMemberQuickCreditBtn"),
  winnerAddMemberBtn: byId("winnerAddMemberBtn"),

  // Prize History Direct Credit Modal Elements
  recordCreditModal: byId("recordCreditModal"),
  recordCreditCustName: byId("recordCreditCustName"),
  recordCreditCustPhone: byId("recordCreditCustPhone"),
  recordCreditMemberStatusPill: byId("recordCreditMemberStatusPill"),
  recordCreditBalanceRow: byId("recordCreditBalanceRow"),
  recordCreditCurrentBal: byId("recordCreditCurrentBal"),
  recordCreditCashWorth: byId("recordCreditCashWorth"),
  recordCreditNotMemberNotice: byId("recordCreditNotMemberNotice"),
  recordCreditRecId: byId("recordCreditRecId"),
  recordCreditRecItem: byId("recordCreditRecItem"),
  recordCreditRecAmount: byId("recordCreditRecAmount"),
  recordCreditRecPrize: byId("recordCreditRecPrize"),
  recordCreditAddBtn: byId("recordCreditAddBtn"),
  recordCreditDeductBtn: byId("recordCreditDeductBtn"),
  recordCreditViewMemberBtn: byId("recordCreditViewMemberBtn"),
  recordCreditAddTitle: byId("recordCreditAddTitle"),
  recordCreditAddSub: byId("recordCreditAddSub"),
  recordCreditDeductTitle: byId("recordCreditDeductTitle"),
  recordCreditDeductSub: byId("recordCreditDeductSub")
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

  // Initialize conditions
  initConditions();

  // Initialize members club
  renderMembersSection();

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
  await Promise.allSettled([
    loadSettingsFromSheets(false),
    loadRecordsFromSheets(false),
    loadMembersFromSheets(false)
  ]);
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

  if (el.wheelCanvas) {
    el.wheelCanvas.addEventListener("click", onWheelClick);
    el.wheelCanvas.addEventListener("mouseenter", () => {
      if (!state.spinning) {
        state.wheelHovered = true;
        drawWheel();
      }
    });
    el.wheelCanvas.addEventListener("mouseleave", () => {
      state.wheelHovered = false;
      drawWheel();
    });
  }

  if (el.wheelWrap) {
    el.wheelWrap.addEventListener("click", (e) => {
      if (e.target !== el.wheelCanvas) {
        onWheelClick(e);
      }
    });
  }

  if (el.spinTimePills) {
    el.spinTimePills.addEventListener("click", (e) => {
      const btn = e.target.closest(".time-pill-btn");
      if (!btn) return;
      const dur = parseInt(btn.dataset.duration, 10);
      if (dur > 0) {
        setSpinDuration(dur);
      }
    });
  }

  if (el.spinDurationInput) {
    el.spinDurationInput.addEventListener("input", () => {
      const val = Math.max(2, Math.min(15, parseInt(el.spinDurationInput.value || "5", 10)));
      setSpinDuration(val, false);
    });
  }

  if (el.winnerViewCouponBtn) {
    el.winnerViewCouponBtn.addEventListener("click", () => {
      closeWinnerModal();
      scrollToCouponPreview();
    });
  }

  if (el.winnerSpinAgainBtn) {
    el.winnerSpinAgainBtn.addEventListener("click", () => {
      closeWinnerModal();
      resetFormForNewSpin();
    });
  }

  if (el.winnerModal) {
    el.winnerModal.addEventListener("click", (event) => {
      const target = event.target;
      if (target && target.getAttribute("data-close-winner-modal") === "true") {
        closeWinnerModal();
      }
    });
  }

  if (el.winnerAddMemberBtn) {
    el.winnerAddMemberBtn.addEventListener("click", onAddWinnerToMembers);
  }

  if (el.customerMemberQuickCreditBtn) {
    el.customerMemberQuickCreditBtn.addEventListener("click", () => {
      const memId = el.customerMemberQuickCreditBtn.dataset.memberId;
      if (memId) {
        setActiveTab("members");
        openCreditModal(memId, el.purchaseAmount.value);
      }
    });
  }

  bindMembersEvents();

  el.saveBtn.addEventListener("click", onSaveClick);
  el.downloadBtn.addEventListener("click", onDownloadCoupon);
  el.shareBtn.addEventListener("click", onShareWhatsApp);

  el.addPrizeBtn.addEventListener("click", () => {
    syncConditionDraftFromEditor();
    const cond = getActiveCondition();
    if (cond) {
      cond.prizes.push({
        id: uid(),
        name: "",
        probability: 0,
        enabled: true
      });
    }
    renderPrizeTable();
    persistSettings();
    refreshUI();
  });

  // Condition management events
  if (el.addConditionBtn) {
    el.addConditionBtn.addEventListener("click", onAddCondition);
  }
  if (el.deleteConditionBtn) {
    el.deleteConditionBtn.addEventListener("click", onDeleteCondition);
  }
  if (el.conditionTabsBar) {
    el.conditionTabsBar.addEventListener("click", (e) => {
      const chip = e.target.closest(".condition-tab-chip");
      if (!chip) return;
      const condId = chip.dataset.conditionId;
      if (condId) {
        switchSettingsCondition(condId);
      }
    });
  }
  if (el.conditionItemNameInput) {
    el.conditionItemNameInput.addEventListener("input", syncConditionFieldsToActive);
  }
  if (el.conditionLabelInput) {
    el.conditionLabelInput.addEventListener("input", syncConditionFieldsToActive);
  }
  if (el.conditionIconInput) {
    el.conditionIconInput.addEventListener("input", syncConditionFieldsToActive);
  }
  if (el.conditionMinAmountInput) {
    el.conditionMinAmountInput.addEventListener("input", syncConditionFieldsToActive);
  }

  // Customer Spin tab: condition selector
  if (el.spinConditionSelect) {
    el.spinConditionSelect.addEventListener("change", () => {
      const condId = el.spinConditionSelect.value;
      if (condId) {
        state.selectedConditionId = condId;
        drawWheel();
        showToast(`Wheel switched to ${getConditionLabel(condId)}`, "info");
      }
    });
  }

  el.saveSettingsBtn.addEventListener("click", () => void onSaveSettings());
  if (el.testAppsScriptBtn) {
    el.testAppsScriptBtn.addEventListener("click", () => void testAppsScriptConnection());
  }
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

  el.recordSearchInput.addEventListener("input", () => {
    state.recordsPage = 1;
    renderRecordsTable();
  });
  el.recordFilterSortSelect.addEventListener("change", () => {
    state.recordsPage = 1;
    renderRecordsTable();
  });

  if (el.pageSizeSelect) {
    el.pageSizeSelect.addEventListener("change", () => {
      const size = parseInt(el.pageSizeSelect.value, 10);
      state.recordsPageSize = size > 0 ? size : 10;
      state.recordsPage = 1;
      renderRecordsTable();
    });
  }

  if (el.firstPageBtn) {
    el.firstPageBtn.addEventListener("click", () => goToRecordsPage(1));
  }
  if (el.prevPageBtn) {
    el.prevPageBtn.addEventListener("click", () => goToRecordsPage(state.recordsPage - 1));
  }
  if (el.nextPageBtn) {
    el.nextPageBtn.addEventListener("click", () => goToRecordsPage(state.recordsPage + 1));
  }
  if (el.lastPageBtn) {
    el.lastPageBtn.addEventListener("click", () => goToRecordsPage(getLastRecordsPage()));
  }
  if (el.exportRecordsBtn) {
    el.exportRecordsBtn.addEventListener("click", exportRecordsToCSV);
  }
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
    if (event.key === "Escape") {
      closeEditModal();
      closeWinnerModal();
      closeAllMemberModals();
    }
  });

  if (el.soundToggleBtn) {
    el.soundToggleBtn.addEventListener("click", () => {
      state.soundMuted = !state.soundMuted;
      localStorage.setItem("spinwin_sound_muted", String(state.soundMuted));
      updateSoundButtonUI();
    });
    updateSoundButtonUI();
  }
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
  checkCustomerMemberMatch();
  refreshUI();
}

async function onSaveSettings() {
  // Sync current condition editor fields + prize table into active condition
  syncConditionDraftFromEditor();

  // Validate that all conditions have at least one prize
  const conditions = state.settings.conditions || [];
  if (!conditions.length) {
    showToast("Please add at least one spin condition.", "error");
    return;
  }
  for (const cond of conditions) {
    if (!cond.prizes || !cond.prizes.length) {
      showToast(`Condition "${cond.itemName || "Unnamed"}" has no prizes. Add at least one.`, "error");
      return;
    }
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
  state.settings.spinDuration = Math.max(2, Math.min(15, parseInt(el.spinDurationInput ? el.spinDurationInput.value : "5", 10) || 5));
  state.settings.appsScriptUrl = (el.appsScriptUrlInput.value || "").trim();
  state.settings.manualDateEnabled = !!el.manualDateToggle.checked;
  state.settings.manualDateTime = el.manualDateInput.value || "";

  // Build flat prizes from all conditions for backward compatibility
  state.settings.prizes = getAllPrizesFlat();

  persistSettings();
  renderConditionTabs();
  renderConditionEditor();
  renderPrizeTable();
  populateSpinConditionSelect();
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
  state.selectedConditionId = null;

  syncSettingsToInputs();
  initConditions();
  renderPrizeTable();
  drawWheel();
  drawCouponPlaceholder();
  renderRecordsTable();
  renderResult();
  refreshUI();
  updateStatus("Local data reset complete.");
  showToast("Local settings and records were reset.", "info");
}

function onWheelClick(event) {
  if (state.spinning) return;
  void onSpinClick();
}

async function onSpinClick() {
  if (state.spinning) return;

  const customerName = safeText(el.customerName.value);
  const customerNumber = safeText(el.customerNumber.value);
  const amount = parseFloat(el.purchaseAmount.value || "0");

  if (!customerName) {
    highlightInputError(el.customerName);
    showToast("Please enter customer name to spin!", "error");
    updateStatus("Enter customer name above to spin!");
    return;
  }
  if (!Number.isFinite(amount) || amount < 0) {
    highlightInputError(el.purchaseAmount);
    showToast("Please enter a valid purchase amount (>= 0).", "error");
    updateStatus("Enter purchase amount above to spin!");
    return;
  }
  if (state.spinLockedForEntry) {
    showToast("This customer entry already spun once. Change name or amount for a new spin.", "info");
    updateStatus("This customer entry already spun once. Enter new details for a new spin.");
    return;
  }

  // Get selected condition
  const activeCond = getSpinCondition();
  if (!activeCond) {
    showToast("Please select a purchased item / spin condition.", "error");
    return;
  }

  // Check min purchase amount for condition
  if (activeCond.minAmount && amount < activeCond.minAmount) {
    showToast(`Minimum purchase amount for "${activeCond.itemName}" is ${formatAmount(activeCond.minAmount)}.`, "error");
    return;
  }

  const enabledPrizes = getEnabledPrizes(activeCond.prizes);
  const weightedPrizes = enabledPrizes.filter((p) => num(p.probability) > 0);

  if (!enabledPrizes.length) {
    showToast("No enabled prizes for this condition. Enable at least one prize in settings.", "error");
    return;
  }
  if (!weightedPrizes.length) {
    showToast("All enabled prizes have 0% probability. Please set prize probability.", "error");
    return;
  }

  const winner = chooseWeightedPrize(weightedPrizes);
  const wheelPrizes = enabledPrizes;
  const winnerIndex = wheelPrizes.findIndex((p) => p.id === winner.id);
  if (winnerIndex < 0) {
    showToast("Prize selection issue. Please save settings and try again.", "error");
    return;
  }

  state.spinning = true;
  if (el.wheelWrap) el.wheelWrap.classList.add("is-spinning");
  refreshUI();
  updateStatus("Spinning lucky wheel...");

  await animateSpin(winnerIndex, wheelPrizes.length);
  stopSpinMusic();
  playWinFanfare();
  triggerConfetti();

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
    purchasedItem: activeCond.itemName || "—",
    conditionId: activeCond.id,
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
  if (el.wheelWrap) el.wheelWrap.classList.remove("is-spinning");

  persistRecords();
  renderRecordsTable();
  renderResult();
  drawCouponFromRecord(record);
  refreshUI();

  updateStatus(`Prize Won: ${record.prize}! Record ID: ${record.recordId}`);
  showToast(`Spin complete! Won: ${record.prize}`, "success");
  openWinnerModal(record);
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
  // Clean the customer phone number for WhatsApp (digits only, no +, spaces, dashes)
  const rawNumber = safeText(record.customerNumber);
  const cleanNumber = rawNumber.replace(/[^0-9]/g, "");
  const waUrl = cleanNumber
    ? `https://wa.me/${cleanNumber}?text=${encodeURIComponent(summary)}`
    : `https://wa.me/?text=${encodeURIComponent(summary)}`;
  window.open(waUrl, "_blank", "noopener");
}

function onRecordsActionClick(event) {
  const btn = event.target.closest("button[data-action]");
  if (!btn) return;

  const action = btn.dataset.action;
  const recordId = btn.dataset.id;
  if (!recordId) return;

  if (action === "manage-credits") {
    openRecordCreditModal(recordId);
    return;
  }

  if (action === "add-credit") {
    state.activeRecordForCredit = getRecordById(recordId);
    if (state.activeRecordForCredit) onRecordCreditAddClick();
    return;
  }

  if (action === "deduct-credit") {
    state.activeRecordForCredit = getRecordById(recordId);
    if (state.activeRecordForCredit) onRecordCreditDeductClick();
    return;
  }

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

async function testAppsScriptConnection() {
  const url = (el.appsScriptUrlInput ? el.appsScriptUrlInput.value : "").trim() || getAppsScriptUrl();
  if (!url) {
    showToast("Please enter a Google Apps Script URL first.", "error");
    return;
  }

  showToast("Testing connection to Google Apps Script...", "info");
  if (el.appsScriptStatusMsg) {
    el.appsScriptStatusMsg.className = "connection-status-msg";
    el.appsScriptStatusMsg.style.background = "rgba(99, 102, 241, 0.15)";
    el.appsScriptStatusMsg.style.border = "1px solid rgba(99, 102, 241, 0.35)";
    el.appsScriptStatusMsg.style.color = "#a5b4fc";
    el.appsScriptStatusMsg.textContent = "⏳ Testing connection to Google Apps Script...";
    el.appsScriptStatusMsg.classList.remove("hidden");
  }

  try {
    const testUrl = appendQueryParams(url, { action: "members" });
    let res = null;
    try {
      const fetchRes = await fetch(testUrl, { method: "GET" });
      res = await parseApiResponse(fetchRes);
    } catch (e) {
      res = await jsonpRequest(url, { action: "members" });
    }

    if (res && res.ok) {
      if (el.appsScriptStatusMsg) {
        el.appsScriptStatusMsg.style.background = "rgba(16, 185, 129, 0.15)";
        el.appsScriptStatusMsg.style.border = "1px solid rgba(16, 185, 129, 0.35)";
        el.appsScriptStatusMsg.style.color = "#34d399";
        el.appsScriptStatusMsg.innerHTML = "✅ <strong>Connection Successful!</strong> Google Apps Script backend is active with full Members Club & Ledger support.";
      }
      showToast("Connection successful! Full backend support active.", "success");
      return;
    }

    if (res && res.message && (res.message.includes("Unknown") || res.message.includes("action"))) {
      if (el.appsScriptStatusMsg) {
        el.appsScriptStatusMsg.style.background = "rgba(239, 68, 68, 0.18)";
        el.appsScriptStatusMsg.style.border = "1px solid rgba(239, 68, 68, 0.35)";
        el.appsScriptStatusMsg.style.color = "#fca5a5";
        el.appsScriptStatusMsg.innerHTML = "❌ <strong>Outdated Script Deployed:</strong> This URL is still running the old script version without Members support.<br><br>👉 <strong>Fix:</strong> In Google Apps Script, click <strong>Deploy &gt; Manage deployments &gt; Edit (pencil) &gt; Version: New version &gt; Deploy</strong>.<br>Or if you created a <em>New deployment</em>, copy that new URL and paste it above.";
      }
      showToast("Outdated Apps Script version detected. Please redeploy.", "error", 6000);
      return;
    }

    if (el.appsScriptStatusMsg) {
      el.appsScriptStatusMsg.style.background = "rgba(245, 158, 11, 0.15)";
      el.appsScriptStatusMsg.style.border = "1px solid rgba(245, 158, 11, 0.35)";
      el.appsScriptStatusMsg.style.color = "#fbbf24";
      el.appsScriptStatusMsg.textContent = `⚠️ Response received: ${res && res.message ? res.message : "Check script deployment permissions (Who has access: Anyone)."}`;
    }
  } catch (err) {
    if (el.appsScriptStatusMsg) {
      el.appsScriptStatusMsg.style.background = "rgba(239, 68, 68, 0.18)";
      el.appsScriptStatusMsg.style.border = "1px solid rgba(239, 68, 68, 0.35)";
      el.appsScriptStatusMsg.style.color = "#fca5a5";
      el.appsScriptStatusMsg.textContent = `❌ Connection error: ${err.message}`;
    }
    showToast(`Connection error: ${err.message}`, "error");
  }
}

async function pushAllMembersToSheets(showToastMessage = true) {
  if (!getAppsScriptUrl()) {
    if (showToastMessage) showToast("Add Apps Script URL in Settings first.", "error");
    return;
  }

  if (!state.members.length) {
    if (showToastMessage) showToast("No members found to save to Google Sheets.", "info");
    return;
  }

  if (showToastMessage) {
    showToast(`Saving ${state.members.length} member(s) to Google Sheets...`, "info");
  }

  try {
    const res = await syncAllMembersToSheets();
    if (!res || !res.ok) {
      if (res && res.message && (res.message.includes("Missing required fields") || res.message.includes("Unknown"))) {
        showToast("⚠️ Outdated Web App: Deploy a New Version in Google Sheets.", "error", 8000);
        throw new Error("Apps Script Web App is outdated. Please deploy the updated apps_script.gs in Google Sheets.");
      }
      throw new Error((res && res.message) || "Failed to save members to Google Sheets.");
    }

    if (showToastMessage) {
      showToast(`Saved ${state.members.length} member(s) to Google Sheets successfully!`, "success");
    }
    markMembersSyncedAll();
  } catch (err) {
    console.error("pushAllMembersToSheets error:", err);
    if (showToastMessage) {
      showToast(`Save to Sheet failed: ${err.message}`, "error");
    }
  }
}

async function loadMembersFromSheets(showToastMessage) {
  if (!getAppsScriptUrl()) {
    if (showToastMessage) showToast("Add Apps Script URL in Settings to sync with Google Sheets.", "info");
    return;
  }

  if (showToastMessage) {
    showToast("Syncing members with Google Sheets...", "info");
  }

  try {
    // Proactively sync only locally-modified members (not yet confirmed in Sheets)
    const unsyncedLocalMembers = state.members.filter((m) => !m.synced);
    if (unsyncedLocalMembers.length > 0) {
      try {
        const pushRes = await syncAllMembersToSheets(unsyncedLocalMembers);
        if (!pushRes || !pushRes.ok) {
          if (pushRes && pushRes.message && (pushRes.message.includes("Missing required fields") || pushRes.message.includes("Unknown"))) {
            showToast("⚠️ Outdated Web App: Deploy a New Version in Google Sheets.", "error", 8000);
            throw new Error("Apps Script Web App is outdated. Deploy the updated apps_script.gs in Google Sheets.");
          }
        }
      } catch (pushErr) {
        console.warn("Pre-sync local members:", pushErr);
        if (pushErr.message && pushErr.message.includes("outdated")) throw pushErr;
      }
    }

    const res = await apiListMembers();
    if (!res.ok) {
      if (res.message && (res.message.includes("Unknown") || res.message.includes("action"))) {
        showToast("⚠️ Outdated Web App: Deploy a New Version in Google Sheets.", "error", 8000);
        throw new Error("Apps Script Web App is outdated. Deploy the updated apps_script.gs in Google Sheets.");
      }
      throw new Error(res.message || "Failed to load members from Google Sheets.");
    }

    const serverMembers = Array.isArray(res.members) ? res.members : [];

    // Merge server members with any local unsynced members
    const map = new Map();
    serverMembers.forEach((m) => {
      map.set(m.id, normalizeMember(m));
    });

    for (const localM of state.members) {
      if (!map.has(localM.id)) {
        map.set(localM.id, localM);
        try {
          await syncSaveMember(localM);
        } catch (e) { }
      }
    }

    state.members = Array.from(map.values()).sort((a, b) => new Date(b.updatedAtIso || b.createdAtIso).getTime() - new Date(a.updatedAtIso || a.createdAtIso).getTime());
    markMembersSyncedAll();
    persistMembers();
    renderMembersSection();
    refreshUI();

    if (showToastMessage) {
      showToast(`Synced ${state.members.length} member(s) with Google Sheets.`, "success");
    }
  } catch (err) {
    console.warn("Members cloud sync warning:", err);
    if (showToastMessage) {
      showToast(`Members sync: ${err.message}`, "error");
    }
  }
}

function applySettingsFromServer(serverSettings, keepCurrentUrl = true) {
  const currentUrl = safeText(state.settings.appsScriptUrl);
  const currentConditions = state.settings.conditions; // Preserve local conditions
  const next = { ...(serverSettings || {}) };
  if (keepCurrentUrl && !safeText(next.appsScriptUrl)) {
    next.appsScriptUrl = currentUrl;
  }
  // Preserve local conditions if server doesn't have them
  // (Google Sheets may not store conditions, only flat prizes)
  if (!Array.isArray(next.conditions) || !next.conditions.length) {
    next.conditions = currentConditions;
  }

  state.settings = mergeSettings(next);
  persistSettings();
  syncSettingsToInputs();
  initConditions();
  renderPrizeTable();
  populateSpinConditionSelect();
  drawWheel();
  renderResult();
  refreshUI();
}

function syncSettingsToInputs() {
  el.shopNameInput.value = state.settings.shopName || "";
  el.shopLogoUrlInput.value = state.settings.shopLogoUrl || "";
  el.expiryHoursInput.value = String(num(state.settings.expiryHours));
  if (el.spinDurationInput) {
    el.spinDurationInput.value = String(state.settings.spinDuration || 5);
  }
  updateSpinPillsUI();
  el.appsScriptUrlInput.value = state.settings.appsScriptUrl || "";
  el.manualDateToggle.checked = !!state.settings.manualDateEnabled;
  el.manualDateInput.value = state.settings.manualDateTime || "";
  el.manualDateInput.disabled = !state.settings.manualDateEnabled;
}

function renderPrizeTable() {
  el.prizeTableBody.innerHTML = "";
  const cond = getActiveCondition();
  const prizes = cond ? cond.prizes : [];
  prizes.forEach((prize) => {
    const row = el.prizeRowTemplate.content.firstElementChild.cloneNode(true);
    row.dataset.id = prize.id;
    row.querySelector(".prize-name-input").value = prize.name;
    row.querySelector(".prize-prob-input").value = String(num(prize.probability));
    row.querySelector(".prize-enabled-input").checked = !!prize.enabled;
    row.querySelector(".remove-prize-btn").addEventListener("click", () => {
      syncConditionDraftFromEditor();
      const activeCond = getActiveCondition();
      if (activeCond) {
        activeCond.prizes = activeCond.prizes.filter((p) => p.id !== prize.id);
      }
      renderPrizeTable();
      persistSettings();
      drawWheel();
      refreshUI();
    });
    el.prizeTableBody.appendChild(row);
  });
  updateConditionProbSummary();
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
  const cond = getActiveCondition();
  if (cond) {
    cond.prizes = collectPrizesFromTable();
  }
}

function refreshUI() {
  const currentRecord = getCurrentRecord();

  el.saveBtn.disabled = !currentRecord;
  el.downloadBtn.disabled = !currentRecord;
  el.shareBtn.disabled = !currentRecord;
  el.shopNameDisplay.textContent = state.settings.shopName || "Lucky Shop";
  el.saveBtn.innerHTML = currentRecord && currentRecord.synced
    ? `<span class="btn-icon">🔄</span> Re-Sync to Sheets`
    : `<span class="btn-icon">💾</span> Save to Sheets`;
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
  const allFiltered = getAllFilteredRecords();
  const totalCount = allFiltered.length;
  const pageSize = state.recordsPageSize || 10;
  const totalPages = Math.max(1, Math.ceil(totalCount / pageSize));

  if (state.recordsPage > totalPages) state.recordsPage = totalPages;
  if (state.recordsPage < 1) state.recordsPage = 1;

  const startIndex = (state.recordsPage - 1) * pageSize;
  const endIndex = Math.min(startIndex + pageSize, totalCount);
  const visible = allFiltered.slice(startIndex, endIndex);

  const counts = getStatusCounts(state.records);
  const selectedCount = getSelectedExistingRecordIds().length;
  el.recordsSummary.textContent =
    `Total: ${state.records.length} | Filtered: ${totalCount} | Selected: ${selectedCount} | Pending: ${counts.pending} | Completed: ${counts.completed} | Rejected: ${counts.rejected} | Expired: ${counts.expired}`;

  el.recordsTableBody.innerHTML = "";
  if (!visible.length) {
    const row = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 11;
    td.textContent = totalCount === 0 ? "No matching prize records found." : "No records on this page.";
    td.className = "centered";
    row.appendChild(td);
    el.recordsTableBody.appendChild(row);
    updateBulkSelectionUI();
    renderPaginationControls(totalCount, 0, 0, totalPages);
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

    // ID
    const tdId = document.createElement("td");
    tdId.innerHTML = `<span class="text-truncate mono" style="max-width:90px;" title="${record.recordId}">${record.recordId}</span>`;
    tr.appendChild(tdId);

    // Customer cell with interactive member loyalty badge
    const member = (record.customerNumber ? findMemberByPhone(record.customerNumber) : null) || findMemberByName(record.customerName);
    const custTd = document.createElement("td");
    const custWrap = document.createElement("div");
    custWrap.className = "customer-cell-wrap";

    const custNameSpan = document.createElement("span");
    custNameSpan.className = "customer-name-text text-truncate";
    custNameSpan.style.maxWidth = "110px";
    custNameSpan.textContent = record.customerName;
    custNameSpan.title = record.customerName;
    custWrap.appendChild(custNameSpan);

    const badgeBtn = document.createElement("button");
    badgeBtn.type = "button";
    badgeBtn.dataset.action = "manage-credits";
    badgeBtn.dataset.id = record.recordId;
    if (member) {
      badgeBtn.className = "member-quick-badge is-member";
      badgeBtn.title = `Active Member: ${member.credits} Credits (≈ ₹${creditsToCash(member.credits).toFixed(2)}). Click to manage loyalty credits.`;
      badgeBtn.innerHTML = `👑 ${member.credits} pts`;
    } else {
      badgeBtn.className = "member-quick-badge not-member";
      badgeBtn.title = "Not in Members Club yet. Click to auto-register & manage credits.";
      badgeBtn.innerHTML = `+ Member`;
    }
    custWrap.appendChild(badgeBtn);
    custTd.appendChild(custWrap);
    tr.appendChild(custTd);

    // Phone with Direct Call 📞 and WhatsApp 💬
    const rawPhone = (record.customerNumber || "").trim();
    const tdPhone = document.createElement("td");
    if (rawPhone && rawPhone !== "-") {
      const clean = cleanDigits(rawPhone);
      const telNum = clean ? (clean.length === 10 ? `+91${clean}` : `+${clean}`) : rawPhone;
      const waNum = clean ? (clean.length === 10 ? `91${clean}` : clean) : rawPhone;

      const phoneWrap = document.createElement("div");
      phoneWrap.className = "phone-cell-wrap";

      const callA = document.createElement("a");
      callA.className = "phone-icon-btn btn-call";
      callA.href = `tel:${telNum}`;
      callA.title = `Call ${record.customerName} directly (${rawPhone})`;
      callA.innerHTML = `📞`;

      const waA = document.createElement("a");
      waA.className = "phone-icon-btn btn-wa";
      waA.href = `https://wa.me/${waNum}`;
      waA.target = "_blank";
      waA.title = `Chat with ${record.customerName} on WhatsApp (${rawPhone})`;
      waA.innerHTML = `💬`;

      const numSpan = document.createElement("span");
      numSpan.className = "phone-number-text text-truncate";
      numSpan.style.maxWidth = "85px";
      numSpan.textContent = rawPhone;

      phoneWrap.appendChild(callA);
      phoneWrap.appendChild(waA);
      phoneWrap.appendChild(numSpan);
      tdPhone.appendChild(phoneWrap);
    } else {
      tdPhone.innerHTML = `<span class="muted">—</span>`;
    }
    tr.appendChild(tdPhone);

    // Purchased Item
    const tdItem = document.createElement("td");
    const itemText = record.purchasedItem || "—";
    tdItem.innerHTML = `<span class="text-truncate" style="max-width:70px;" title="${itemText}">${itemText}</span>`;
    tr.appendChild(tdItem);

    // Amount
    const tdAmount = document.createElement("td");
    tdAmount.innerHTML = `<span style="font-weight:700; white-space:nowrap; font-size:0.8rem;">₹${formatAmount(record.amount)}</span>`;
    tr.appendChild(tdAmount);

    // Prize
    const tdPrize = document.createElement("td");
    tdPrize.innerHTML = `<span class="text-truncate" style="max-width:80px;" title="${record.prize}">${record.prize}</span>`;
    tr.appendChild(tdPrize);

    // Status
    const statusTd = document.createElement("td");
    const statusPill = document.createElement("span");
    statusPill.className = `status-pill status-${effectiveStatus.toLowerCase()}`;
    const statusShort = effectiveStatus === STATUS_COMPLETED ? "✓ Done" : effectiveStatus === STATUS_REJECTED ? "✕ Rej" : effectiveStatus === STATUS_EXPIRED ? "⌛ Exp" : "⏳ Pend";
    statusPill.textContent = statusShort;
    statusPill.title = `Status: ${effectiveStatus}`;
    statusTd.appendChild(statusPill);
    tr.appendChild(statusTd);

    // Date & Expiry (Short Date Format)
    const dt = new Date(record.dateTimeIso);
    const expDt = new Date(record.expiryIso);
    const shortDtStr = Number.isNaN(dt.getTime()) ? "-" : dt.toLocaleDateString(undefined, { day: "2-digit", month: "short" });
    const shortExpStr = Number.isNaN(expDt.getTime()) ? "-" : expDt.toLocaleDateString(undefined, { day: "2-digit", month: "short" });

    const tdDate = document.createElement("td");
    tdDate.innerHTML = `<span class="muted" style="white-space:nowrap; font-size:0.76rem;" title="Date: ${formatIsoDateTime(record.dateTimeIso)}">${shortDtStr}</span>`;
    tr.appendChild(tdDate);

    const tdExpiry = document.createElement("td");
    tdExpiry.innerHTML = `<span class="muted" style="white-space:nowrap; font-size:0.76rem;" title="Expiry: ${formatIsoDateTime(record.expiryIso)}">${shortExpStr}</span>`;
    tr.appendChild(tdExpiry);

    // Compact Actions (Icon Buttons)
    const actionTd = document.createElement("td");
    const wrap = document.createElement("div");
    wrap.className = "record-actions";

    const creditBtnTone = member ? "btn-credit-member" : "btn-credit-new";
    const creditTitle = member ? `Active Member (${member.credits} pts). Click to manage loyalty credits` : "Click to register member & manage credits";
    const creditBtn = createActionButton("⭐", creditBtnTone, "manage-credits", record.recordId);
    creditBtn.title = creditTitle;
    wrap.appendChild(creditBtn);

    const editBtn = createActionButton("✏️", "btn-ghost", "edit", record.recordId);
    editBtn.title = "Edit Prize Record";
    wrap.appendChild(editBtn);

    const completeBtn = createActionButton("✓", "btn-secondary", "complete", record.recordId, effectiveStatus === STATUS_COMPLETED);
    completeBtn.title = effectiveStatus === STATUS_COMPLETED ? "Already Completed" : "Mark as Completed";
    wrap.appendChild(completeBtn);

    const rejectBtn = createActionButton("✕", "btn-danger", "reject", record.recordId, effectiveStatus === STATUS_REJECTED);
    rejectBtn.title = effectiveStatus === STATUS_REJECTED ? "Already Rejected" : "Mark as Rejected";
    wrap.appendChild(rejectBtn);

    const deleteBtn = createActionButton("🗑️", "btn-danger", "delete", record.recordId);
    deleteBtn.title = "Delete Record";
    wrap.appendChild(deleteBtn);

    actionTd.appendChild(wrap);
    tr.appendChild(actionTd);
    el.recordsTableBody.appendChild(tr);
  });

  updateBulkSelectionUI();
  renderPaginationControls(totalCount, startIndex + 1, endIndex, totalPages);
}

function renderPaginationControls(totalCount, start, end, totalPages) {
  if (!el.recordsPaginationBar) return;

  if (el.paginationRangeText) {
    if (totalCount === 0) {
      el.paginationRangeText.innerHTML = "Showing <strong>0</strong> of <strong>0</strong> records";
    } else {
      el.paginationRangeText.innerHTML = `Showing <strong>${start}–${end}</strong> of <strong>${totalCount}</strong> records (Page ${state.recordsPage}/${totalPages})`;
    }
  }

  if (el.pageSizeSelect) {
    el.pageSizeSelect.value = String(state.recordsPageSize || 10);
  }

  const isFirst = state.recordsPage <= 1;
  const isLast = state.recordsPage >= totalPages;

  if (el.firstPageBtn) el.firstPageBtn.disabled = isFirst;
  if (el.prevPageBtn) el.prevPageBtn.disabled = isFirst;
  if (el.nextPageBtn) el.nextPageBtn.disabled = isLast;
  if (el.lastPageBtn) el.lastPageBtn.disabled = isLast;

  if (!el.paginationPages) return;
  el.paginationPages.innerHTML = "";

  const pages = generatePaginationPageNumbers(state.recordsPage, totalPages);
  pages.forEach((item) => {
    if (item === "...") {
      const span = document.createElement("span");
      span.className = "pagination-ellipsis";
      span.textContent = "…";
      el.paginationPages.appendChild(span);
    } else {
      const pageNum = item;
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = `page-btn ${pageNum === state.recordsPage ? "is-active" : ""}`;
      btn.textContent = String(pageNum);
      btn.setAttribute("aria-label", `Page ${pageNum}`);
      if (pageNum === state.recordsPage) {
        btn.setAttribute("aria-current", "page");
      }
      btn.addEventListener("click", () => goToRecordsPage(pageNum));
      el.paginationPages.appendChild(btn);
    }
  });
}

function generatePaginationPageNumbers(current, total) {
  if (total <= 7) {
    return Array.from({ length: total }, (_, i) => i + 1);
  }

  const pages = [];
  pages.push(1);

  if (current > 3) {
    pages.push("...");
  }

  const start = Math.max(2, current - 1);
  const end = Math.min(total - 1, current + 1);

  for (let i = start; i <= end; i++) {
    pages.push(i);
  }

  if (current < total - 2) {
    pages.push("...");
  }

  pages.push(total);
  return pages;
}

function getLastRecordsPage() {
  const all = getAllFilteredRecords();
  const pageSize = state.recordsPageSize || 10;
  return Math.max(1, Math.ceil(all.length / pageSize));
}

function goToRecordsPage(page) {
  const totalPages = getLastRecordsPage();
  const target = Math.max(1, Math.min(totalPages, page));
  if (target === state.recordsPage) return;
  state.recordsPage = target;
  renderRecordsTable();
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

function getAllFilteredRecords() {
  const search = safeText(el.recordSearchInput.value).toLowerCase();
  const mode = el.recordFilterSortSelect.value || "recent_old";

  let rows = state.records.filter((record) => {
    if (!search) return true;
    return (
      record.recordId.toLowerCase().includes(search) ||
      record.customerName.toLowerCase().includes(search) ||
      (record.customerNumber && record.customerNumber.toLowerCase().includes(search)) ||
      record.prize.toLowerCase().includes(search)
    );
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

function getVisibleRecords() {
  const all = getAllFilteredRecords();
  const pageSize = state.recordsPageSize || 10;
  const totalPages = Math.max(1, Math.ceil(all.length / pageSize));
  if (state.recordsPage > totalPages) state.recordsPage = totalPages;
  if (state.recordsPage < 1) state.recordsPage = 1;

  const start = (state.recordsPage - 1) * pageSize;
  return all.slice(start, start + pageSize);
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
  const prizes = getEnabledPrizes(getActiveSpinPrizes());
  const ctx = wheelCtx;

  // Crisp HiDPI rendering while keeping one fixed logical drawing space
  const LOGICAL = 480;
  const dpr = Math.min(window.devicePixelRatio || 1, 2.5);
  const backing = Math.round(LOGICAL * dpr);
  if (el.wheelCanvas.width !== backing || el.wheelCanvas.height !== backing) {
    el.wheelCanvas.width = backing;
    el.wheelCanvas.height = backing;
  }

  const cx = LOGICAL / 2;
  const cy = LOGICAL / 2;
  const outerBezelRadius = Math.min(cx, cy) - 6;
  const rimWidth = 22;
  const sliceRadius = outerBezelRadius - rimWidth;
  const isSpinning = state.spinning;
  const isHovered = state.wheelHovered;

  ctx.clearRect(0, 0, el.wheelCanvas.width, el.wheelCanvas.height);
  ctx.save();
  ctx.scale(backing / LOGICAL, backing / LOGICAL);

  if (!prizes.length) {
    ctx.fillStyle = "#0f172a";
    ctx.beginPath();
    ctx.arc(cx, cy, outerBezelRadius, 0, Math.PI * 2);
    ctx.fill();
    ctx.fillStyle = "#94a3b8";
    ctx.font = "700 22px Sora, sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText("No Enabled Prizes", cx, cy);
    ctx.restore();
    return;
  }

  // 0. Cast shadow that lifts the wheel off the surface
  drawWheelShadow(ctx, cx, cy, outerBezelRadius);

  // 1. Metallic 3D rim (torus illusion)
  drawMetallicRim(ctx, cx, cy, outerBezelRadius, rimWidth, sliceRadius);

  // 2. Wheel slices with depth relief
  const arc = (Math.PI * 2) / prizes.length;
  const sliceStart = state.wheelAngle;
  for (let i = 0; i < prizes.length; i++) {
    const start = sliceStart + i * arc;
    const end = start + arc;
    const sliceColor = colorForSlice(i);

    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, sliceRadius, start, end);
    ctx.closePath();
    ctx.fillStyle = sliceColor;
    ctx.fill();

    // Relief: bright near hub, shaded toward the rim for a 3D pop
    const relief = ctx.createRadialGradient(cx, cy, sliceRadius * 0.08, cx, cy, sliceRadius);
    relief.addColorStop(0, "rgba(255, 255, 255, 0.24)");
    relief.addColorStop(0.5, "rgba(255, 255, 255, 0.06)");
    relief.addColorStop(0.82, "rgba(0, 0, 0, 0.05)");
    relief.addColorStop(1, "rgba(0, 0, 0, 0.34)");
    ctx.fillStyle = relief;
    ctx.fill();

    // Divider: recessed dark groove + bright golden pin
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.lineTo(cx + Math.cos(start) * sliceRadius, cy + Math.sin(start) * sliceRadius);
    ctx.strokeStyle = "rgba(0, 0, 0, 0.4)";
    ctx.lineWidth = 3;
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(cx + Math.cos(start) * sliceRadius * 0.4, cy + Math.sin(start) * sliceRadius * 0.4);
    ctx.lineTo(cx + Math.cos(start) * sliceRadius, cy + Math.sin(start) * sliceRadius);
    ctx.strokeStyle = "rgba(251, 191, 36, 0.7)";
    ctx.lineWidth = 1.5;
    ctx.stroke();

    // Slice text (radial-aligned, high contrast)
    const sliceMidAngle = start + arc / 2;
    ctx.save();
    ctx.translate(cx, cy);
    ctx.rotate(sliceMidAngle);
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";

    const fontSize = prizes.length <= 6 ? 20 : prizes.length <= 8 ? 17 : 14;
    ctx.font = `700 ${fontSize}px "Plus Jakarta Sans", sans-serif`;
    ctx.fillStyle = "#ffffff";
    ctx.shadowColor = "rgba(0, 0, 0, 0.9)";
    ctx.shadowBlur = 6;
    ctx.shadowOffsetX = 1;
    ctx.shadowOffsetY = 2;

    const maxTextWidth = sliceRadius * 0.52;
    const lines = wrapPrizeTextLines(ctx, prizes[i].name, maxTextWidth, 3);
    const lineHeight = Math.round(fontSize * 1.15);
    const startY = -((lines.length - 1) * lineHeight) / 2;
    const textX = sliceRadius * 0.64;
    lines.forEach((line, lineIdx) => {
      ctx.fillText(line, textX, startY + lineIdx * lineHeight);
    });
    ctx.restore();
  }

  // 3. Global lighting: soft top gloss, shaded bottom
  ctx.save();
  ctx.beginPath();
  ctx.arc(cx, cy, sliceRadius, 0, Math.PI * 2);
  ctx.clip();
  const faceLight = ctx.createLinearGradient(0, cy - sliceRadius, 0, cy + sliceRadius);
  faceLight.addColorStop(0, "rgba(255, 255, 255, 0.18)");
  faceLight.addColorStop(0.35, "rgba(255, 255, 255, 0.04)");
  faceLight.addColorStop(0.5, "rgba(0, 0, 0, 0)");
  faceLight.addColorStop(1, "rgba(0, 0, 0, 0.28)");
  ctx.fillStyle = faceLight;
  ctx.fillRect(cx - sliceRadius, cy - sliceRadius, sliceRadius * 2, sliceRadius * 2);
  ctx.restore();

  // 4. Glossy glass sheen
  drawGlossySheen(ctx, cx, cy, sliceRadius);

  // 5. Chrome/gold studs on the rim
  drawRimStuds(ctx, cx, cy, outerBezelRadius, rimWidth, isSpinning);

  // 6. Realistic 3D center hub
  drawCenterHub(ctx, cx, cy, isHovered, isSpinning);

  ctx.restore();
}

function drawWheelShadow(ctx, cx, cy, radius) {
  ctx.save();
  ctx.translate(cx + 8, cy + 24);
  ctx.scale(1, 0.55);
  const shadow = ctx.createRadialGradient(0, 0, radius * 0.35, 0, 0, radius * 1.02);
  shadow.addColorStop(0, "rgba(0, 0, 0, 0.42)");
  shadow.addColorStop(0.6, "rgba(0, 0, 0, 0.28)");
  shadow.addColorStop(1, "rgba(0, 0, 0, 0)");
  ctx.fillStyle = shadow;
  ctx.beginPath();
  ctx.arc(0, 0, radius, 0, Math.PI * 2);
  ctx.fill();
  ctx.restore();
}

function drawMetallicRim(ctx, cx, cy, outer, rimWidth, sliceRadius) {
  // Donut band with diagonal metallic gold gradient
  const band = ctx.createLinearGradient(cx - outer, cy - outer, cx + outer, cy + outer);
  band.addColorStop(0, "#fef3c7");
  band.addColorStop(0.22, "#fbbf24");
  band.addColorStop(0.45, "#d97706");
  band.addColorStop(0.58, "#fef08a");
  band.addColorStop(0.78, "#b45309");
  band.addColorStop(1, "#f59e0b");
  ctx.beginPath();
  ctx.arc(cx, cy, outer, 0, Math.PI * 2);
  ctx.arc(cx, cy, sliceRadius, 0, Math.PI * 2, true);
  ctx.fillStyle = band;
  ctx.fill("evenodd");

  ctx.save();
  ctx.lineCap = "round";

  // Torus highlight (light catching the top-left curve)
  ctx.beginPath();
  ctx.arc(cx, cy, outer - rimWidth * 0.5, 0, Math.PI * 2);
  ctx.lineWidth = rimWidth * 0.85;
  const highlight = ctx.createRadialGradient(cx - outer * 0.45, cy - outer * 0.5, rimWidth * 0.5, cx, cy, outer);
  highlight.addColorStop(0, "rgba(255, 255, 255, 0.6)");
  highlight.addColorStop(0.42, "rgba(255, 255, 255, 0.12)");
  highlight.addColorStop(1, "rgba(255, 255, 255, 0)");
  ctx.strokeStyle = highlight;
  ctx.stroke();

  // Bottom shade for volume
  ctx.beginPath();
  ctx.arc(cx, cy, outer - rimWidth * 0.5, 0, Math.PI * 2);
  const shading = ctx.createRadialGradient(cx + outer * 0.4, cy + outer * 0.45, rimWidth * 0.5, cx, cy, outer);
  shading.addColorStop(0, "rgba(0, 0, 0, 0)");
  shading.addColorStop(0.6, "rgba(0, 0, 0, 0.25)");
  shading.addColorStop(1, "rgba(0, 0, 0, 0.45)");
  ctx.strokeStyle = shading;
  ctx.stroke();

  // Inner recessed groove where slices meet the rim
  ctx.beginPath();
  ctx.arc(cx, cy, sliceRadius + 1, 0, Math.PI * 2);
  ctx.strokeStyle = "rgba(0, 0, 0, 0.5)";
  ctx.lineWidth = 3;
  ctx.stroke();
  ctx.beginPath();
  ctx.arc(cx, cy, sliceRadius - 1, 0, Math.PI * 2);
  ctx.strokeStyle = "rgba(255, 255, 255, 0.12)";
  ctx.lineWidth = 1.5;
  ctx.stroke();

  ctx.restore();
}

function drawRimStuds(ctx, cx, cy, outer, rimWidth, spinning) {
  const bandRadius = outer - rimWidth / 2;
  const count = Math.max(28, Math.round((Math.PI * 2 * bandRadius) / 26));
  const phase = spinning ? Math.floor(performance.now() / 50) : 0;

  for (let p = 0; p < count; p++) {
    const angle = (Math.PI * 2 * p) / count;
    const px = cx + Math.cos(angle) * bandRadius;
    const py = cy + Math.sin(angle) * bandRadius;
    const isLit = spinning && (p + phase) % 4 === 0;

    const stud = ctx.createRadialGradient(px - 1.6, py - 1.6, 0.6, px, py, 4.6);
    if (isLit) {
      stud.addColorStop(0, "#ffffff");
      stud.addColorStop(0.5, "#fff7d6");
      stud.addColorStop(1, "#f59e0b");
    } else {
      stud.addColorStop(0, "#fffbe6");
      stud.addColorStop(0.55, "#fbbf24");
      stud.addColorStop(1, "#92400e");
    }

    ctx.beginPath();
    ctx.arc(px, py, 4.2, 0, Math.PI * 2);
    ctx.fillStyle = stud;
    ctx.fill();

    if (isLit) {
      ctx.shadowColor = "rgba(255, 255, 255, 0.95)";
      ctx.shadowBlur = 12;
      ctx.beginPath();
      ctx.arc(px, py, 3.4, 0, Math.PI * 2);
      ctx.fill();
      ctx.shadowBlur = 0;
    }
  }
}

function drawGlossySheen(ctx, cx, cy, radius) {
  ctx.save();
  ctx.beginPath();
  ctx.arc(cx, cy, radius, 0, Math.PI * 2);
  ctx.clip();

  // Large soft light source from the top-left
  const sheen = ctx.createRadialGradient(cx - radius * 0.42, cy - radius * 0.52, radius * 0.05, cx, cy, radius * 1.15);
  sheen.addColorStop(0, "rgba(255, 255, 255, 0.28)");
  sheen.addColorStop(0.35, "rgba(255, 255, 255, 0.09)");
  sheen.addColorStop(0.7, "rgba(255, 255, 255, 0)");
  sheen.addColorStop(1, "rgba(255, 255, 255, 0)");
  ctx.fillStyle = sheen;
  ctx.fillRect(cx - radius, cy - radius, radius * 2, radius * 2);

  // Crescent top reflection
  ctx.beginPath();
  ctx.arc(cx, cy, radius * 0.93, Math.PI * 1.05, Math.PI * 1.6);
  ctx.lineWidth = radius * 0.09;
  ctx.lineCap = "round";
  const crescent = ctx.createLinearGradient(cx - radius, cy - radius, cx + radius, cy + radius);
  crescent.addColorStop(0, "rgba(255, 255, 255, 0.42)");
  crescent.addColorStop(1, "rgba(255, 255, 255, 0)");
  ctx.strokeStyle = crescent;
  ctx.stroke();

  ctx.restore();
}

function drawCenterHub(ctx, cx, cy, hovered, spinning) {
  // Recessed socket shadow
  ctx.save();
  ctx.lineCap = "round";

  ctx.beginPath();
  ctx.arc(cx, cy, 54, 0, Math.PI * 2);
  const socket = ctx.createRadialGradient(cx, cy, 8, cx, cy, 54);
  socket.addColorStop(0, "rgba(0, 0, 0, 0.55)");
  socket.addColorStop(1, "rgba(0, 0, 0, 0)");
  ctx.fillStyle = socket;
  ctx.fill();

  // Metallic outer ring
  ctx.beginPath();
  ctx.arc(cx, cy, 48, 0, Math.PI * 2);
  const ring = ctx.createLinearGradient(cx - 48, cy - 48, cx + 48, cy + 48);
  ring.addColorStop(0, "#fef08a");
  ring.addColorStop(0.35, "#f59e0b");
  ring.addColorStop(0.65, "#b45309");
  ring.addColorStop(1, "#fde68a");
  ctx.fillStyle = ring;
  ctx.shadowColor = "rgba(0, 0, 0, 0.65)";
  ctx.shadowBlur = 14;
  ctx.fill();
  ctx.shadowBlur = 0;

  // Ring top highlight
  ctx.beginPath();
  ctx.arc(cx, cy, 41, Math.PI * 1.08, Math.PI * 1.58);
  ctx.strokeStyle = "rgba(255, 255, 255, 0.5)";
  ctx.lineWidth = 5;
  ctx.stroke();

  // Inner dark bezel (recessed)
  ctx.beginPath();
  ctx.arc(cx, cy, 37, 0, Math.PI * 2);
  const bezel = ctx.createRadialGradient(cx - 6, cy - 8, 4, cx, cy, 37);
  bezel.addColorStop(0, "#1e293b");
  bezel.addColorStop(1, "#0b1120");
  ctx.fillStyle = bezel;
  ctx.fill();

  // Glass button
  ctx.beginPath();
  ctx.arc(cx, cy, 31, 0, Math.PI * 2);
  const btn = ctx.createRadialGradient(cx - 10, cy - 12, 2, cx, cy, 31);
  btn.addColorStop(0, hovered ? "#ffffff" : "#fde68a");
  btn.addColorStop(0.55, hovered ? "#fbbf24" : "#f59e0b");
  btn.addColorStop(1, "#b45309");
  ctx.fillStyle = btn;
  ctx.fill();

  // Button specular dot
  ctx.beginPath();
  ctx.ellipse(cx - 11, cy - 13, 8, 4.5, -0.6, 0, Math.PI * 2);
  ctx.fillStyle = "rgba(255, 255, 255, 0.55)";
  ctx.fill();

  // Button rim light
  ctx.beginPath();
  ctx.arc(cx, cy, 31, 0, Math.PI * 2);
  ctx.strokeStyle = hovered ? "rgba(255, 255, 255, 0.7)" : "rgba(255, 255, 255, 0.22)";
  ctx.lineWidth = 1.5;
  ctx.stroke();

  // Label
  ctx.fillStyle = "#1e1100";
  ctx.font = "800 15px Sora, sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(spinning ? "···" : "SPIN", cx, cy);

  ctx.restore();
}

function animateSpin(winnerIndex, totalSegments) {
  return new Promise((resolve) => {
    const pointerAngle = -Math.PI / 2;
    const arc = (Math.PI * 2) / totalSegments;
    const current = normalizeAngle(state.wheelAngle);
    const winnerCenter = pointerAngle - (winnerIndex + 0.5) * arc;

    const durationSeconds = state.settings.spinDuration || 5;
    const duration = durationSeconds * 1000;
    const extraSpins = Math.round(durationSeconds * 1.5) + randInt(1, 3);

    let target = winnerCenter + extraSpins * Math.PI * 2;
    while (target <= current) {
      target += Math.PI * 2;
    }

    const start = performance.now();
    let lastSegmentPassed = -1;

    // Launch spin sound & upbeat soundtrack
    playSpinStartSound();
    startSpinMusic(duration);

    // Physical "kick" as the wheel is launched
    triggerSpinKick();

    const step = (now) => {
      const t = Math.min(1, (now - start) / duration);
      const eased = easeOutQuart(t);
      state.wheelAngle = current + (target - current) * eased;

      // Slice tick calculation
      const relativeAngle = normalizeAngle(pointerAngle - state.wheelAngle);
      const currentSegment = Math.floor(relativeAngle / arc);
      if (currentSegment !== lastSegmentPassed) {
        lastSegmentPassed = currentSegment;
        playTickSound();
        triggerPointerWobble();
      }

      drawWheel();
      if (t < 1) {
        requestAnimationFrame(step);
      } else {
        // Damped-spring "settle" wobble so the wheel feels physical
        const settleStart = performance.now();
        const settleDuration = 380;
        const settleAmp = 0.04;
        const settleDecay = 6.5;
        const settleFreq = 15;

        const settle = (stamp) => {
          const s = Math.min(1, (stamp - settleStart) / settleDuration);
          const decay = settleAmp * Math.exp(-settleDecay * s);
          const wave = decay * Math.sin(settleFreq * s);
          state.wheelAngle = target + wave;
          drawWheel();
          if (s < 1) {
            requestAnimationFrame(settle);
          } else {
            state.wheelAngle = normalizeAngle(target);
            drawWheel();
            settleSpinBounce();
            resolve();
          }
        };
        requestAnimationFrame(settle);
      }
    };

    requestAnimationFrame(step);
  });
}

function easeOutQuart(t) {
  return 1 - Math.pow(1 - t, 4);
}

function triggerSpinKick() {
  const wrap = document.querySelector(".wheel-wrap");
  if (!wrap) return;
  wrap.classList.remove("spin-kick", "is-spinning");
  void wrap.offsetWidth;
  wrap.classList.add("spin-kick", "is-spinning");
  clearTimeout(triggerSpinKick._t);
  triggerSpinKick._t = setTimeout(() => wrap.classList.remove("spin-kick"), 650);
}

function settleSpinBounce() {
  const wrap = document.querySelector(".wheel-wrap");
  if (!wrap) return;
  wrap.classList.remove("spin-bounce", "is-spinning");
  void wrap.offsetWidth;
  wrap.classList.add("spin-bounce");
  setTimeout(() => wrap.classList.remove("spin-bounce"), 600);
}

function drawCouponPlaceholder() {
  const ctx = couponCtx;
  const w = el.couponCanvas.width;
  const h = el.couponCanvas.height;
  ctx.clearRect(0, 0, w, h);

  // Dark Luxury Voucher Background
  const grad = ctx.createLinearGradient(0, 0, w, h);
  grad.addColorStop(0, "#080d1a");
  grad.addColorStop(0.5, "#0f172a");
  grad.addColorStop(1, "#060911");
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, w, h);

  // Border frame
  ctx.strokeStyle = "rgba(245, 158, 11, 0.35)";
  ctx.lineWidth = 3;
  ctx.strokeRect(28, 28, w - 56, h - 56);

  // Inner subtle border
  ctx.strokeStyle = "rgba(255, 255, 255, 0.08)";
  ctx.lineWidth = 1;
  ctx.strokeRect(36, 36, w - 72, h - 72);

  // Header
  ctx.fillStyle = "#fbbf24";
  ctx.font = "800 44px Sora, sans-serif";
  ctx.fillText("VIP REWARDS VOUCHER", 64, 115);

  ctx.fillStyle = "#94a3b8";
  ctx.font = "600 24px 'Plus Jakarta Sans', sans-serif";
  ctx.fillText("Spin the lucky wheel to unlock your exclusive prize coupon.", 64, 170);

  // QR preview placeholder
  ctx.fillStyle = "rgba(255, 255, 255, 0.05)";
  ctx.fillRect(w - 280, 185, 216, 216);
  ctx.strokeStyle = "rgba(255, 255, 255, 0.15)";
  ctx.strokeRect(w - 280, 185, 216, 216);

  ctx.fillStyle = "#64748b";
  ctx.font = "700 20px 'Plus Jakarta Sans', sans-serif";
  ctx.textAlign = "center";
  ctx.fillText("QR Code", w - 172, 285);
  ctx.fillText("Preview", w - 172, 315);
  ctx.textAlign = "left";
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

  // Background
  const grad = ctx.createLinearGradient(0, 0, w, h);
  grad.addColorStop(0, "#080d1a");
  grad.addColorStop(0.5, "#0f172a");
  grad.addColorStop(1, "#060911");
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, w, h);

  const cardX = 24;
  const cardY = 24;
  const cardW = w - 48;
  const cardH = h - 48;

  // Golden luxury double border
  ctx.strokeStyle = "#fbbf24";
  ctx.lineWidth = 3;
  ctx.strokeRect(cardX, cardY, cardW, cardH);

  ctx.strokeStyle = "rgba(251, 191, 36, 0.25)";
  ctx.lineWidth = 1;
  ctx.strokeRect(cardX + 8, cardY + 8, cardW - 16, cardH - 16);

  // Optional Logo Watermark
  if (state.logoImage) {
    ctx.save();
    ctx.globalAlpha = 0.12;
    const logoSize = 220;
    ctx.drawImage(state.logoImage, (w - logoSize) / 2, (h - logoSize) / 2, logoSize, logoSize);
    ctx.restore();
  }

  // Header Banner with Gold Gradient
  const headerH = 92;
  const headerGrad = ctx.createLinearGradient(cardX, cardY, cardX + cardW, cardY);
  headerGrad.addColorStop(0, "#92400e");
  headerGrad.addColorStop(0.5, "#d97706");
  headerGrad.addColorStop(1, "#b45309");
  ctx.fillStyle = headerGrad;
  ctx.fillRect(cardX, cardY, cardW, headerH);

  ctx.fillStyle = "#ffffff";
  ctx.font = "800 36px Sora, sans-serif";
  const shopTitle = fitTextWithEllipsis(ctx, record.shopName || state.settings.shopName, cardW - 280);
  ctx.fillText(shopTitle, cardX + 28, cardY + 58);

  // Title below banner
  ctx.fillStyle = "#f8fafc";
  ctx.font = "800 28px Sora, sans-serif";
  ctx.fillText("OFFICIAL REWARD VOUCHER", cardX + 28, cardY + 144);

  // Status Badge
  const statusColor = status === STATUS_COMPLETED
    ? "#10b981"
    : status === STATUS_REJECTED
      ? "#ef4444"
      : status === STATUS_EXPIRED
        ? "#94a3b8"
        : "#f59e0b";

  const statusBg = status === STATUS_COMPLETED
    ? "rgba(16, 185, 129, 0.2)"
    : status === STATUS_REJECTED
      ? "rgba(239, 68, 68, 0.2)"
      : status === STATUS_EXPIRED
        ? "rgba(148, 163, 184, 0.2)"
        : "rgba(245, 158, 11, 0.2)";

  const statusW = 180;
  const statusH = 40;
  const statusX = cardX + cardW - statusW - 28;
  const statusY = cardY + 116;

  ctx.fillStyle = statusBg;
  ctx.fillRect(statusX, statusY, statusW, statusH);
  ctx.strokeStyle = statusColor;
  ctx.lineWidth = 1.5;
  ctx.strokeRect(statusX, statusY, statusW, statusH);

  ctx.fillStyle = statusColor;
  ctx.font = "700 18px 'Plus Jakarta Sans', sans-serif";
  ctx.textAlign = "center";
  ctx.fillText(`STATUS: ${status.toUpperCase()}`, statusX + statusW / 2, statusY + 26);
  ctx.textAlign = "left";

  // QR Code Box (white frame)
  const qrSize = 210;
  const qrX = cardX + cardW - qrSize - 32;
  const qrY = cardY + 180;

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(qrX - 10, qrY - 10, qrSize + 20, qrSize + 20);
  ctx.strokeStyle = "rgba(251, 191, 36, 0.5)";
  ctx.lineWidth = 2;
  ctx.strokeRect(qrX - 10, qrY - 10, qrSize + 20, qrSize + 20);

  // Render QR
  const qrData = buildQrPayload(record, status);
  const qrOk = drawQrOnCanvas(ctx, qrData.payload, qrX, qrY, qrSize);

  ctx.fillStyle = "#94a3b8";
  ctx.font = "600 16px 'Plus Jakarta Sans', sans-serif";
  ctx.textAlign = "center";
  ctx.fillText(qrData.isLive ? "Scan for Live Status" : "Scan to Verify Voucher", qrX + qrSize / 2, qrY + qrSize + 32);
  ctx.textAlign = "left";

  if (!qrOk) {
    ctx.fillStyle = "#ef4444";
    ctx.font = "700 14px 'Plus Jakarta Sans', sans-serif";
    ctx.fillText("QR code generation unavailable", qrX, qrY + qrSize + 54);
  }

  // Prize Highlight Banner on Left
  const leftX = cardX + 32;
  const prizeBoxY = cardY + 172;
  const prizeBoxW = qrX - leftX - 40;

  ctx.fillStyle = "rgba(245, 158, 11, 0.12)";
  ctx.fillRect(leftX, prizeBoxY, prizeBoxW, 68);
  ctx.strokeStyle = "rgba(245, 158, 11, 0.4)";
  ctx.lineWidth = 1;
  ctx.strokeRect(leftX, prizeBoxY, prizeBoxW, 68);

  ctx.fillStyle = "#fbbf24";
  ctx.font = "700 15px 'Plus Jakarta Sans', sans-serif";
  ctx.fillText("PRIZE WON:", leftX + 16, prizeBoxY + 26);

  ctx.fillStyle = "#ffffff";
  ctx.font = "800 24px Sora, sans-serif";
  const prizeText = fitTextWithEllipsis(ctx, record.prize, prizeBoxW - 32);
  ctx.fillText(prizeText, leftX + 16, prizeBoxY + 54);

  // Metadata items
  const details = [
    { label: "Unique ID", value: record.recordId },
    { label: "Customer Name", value: record.customerName },
    { label: "Customer Number", value: record.customerNumber || "-" },
    { label: "Purchased Item", value: record.purchasedItem || "-" },
    { label: "Purchase Amount", value: formatAmount(record.amount) },
    { label: "Issued Date", value: formatIsoDateTime(record.dateTimeIso) },
    { label: "Expiry Date", value: formatIsoDateTime(record.expiryIso) }
  ];

  let rowY = prizeBoxY + 104;
  const rowGap = 40;
  details.forEach((item) => {
    ctx.font = "600 16px 'Plus Jakarta Sans', sans-serif";
    const labelText = `${item.label}: `;
    const labelWidth = ctx.measureText(labelText).width;
    ctx.fillStyle = "#94a3b8";
    ctx.fillText(labelText, leftX, rowY);

    ctx.font = "700 17px 'Plus Jakarta Sans', sans-serif";
    const valueWidth = Math.max(80, prizeBoxW - labelWidth);
    const valueText = fitTextWithEllipsis(ctx, String(item.value || "-"), valueWidth);
    ctx.fillStyle = item.label === "Unique ID" ? "#fef08a" : "#f8fafc";
    ctx.fillText(valueText, leftX + labelWidth, rowY);
    rowY += rowGap;
  });

  // Bottom Footer
  ctx.fillStyle = "rgba(245, 158, 11, 0.85)";
  ctx.font = "700 18px 'Plus Jakarta Sans', sans-serif";
  ctx.textAlign = "center";
  ctx.fillText("★ Official Store Reward Voucher • Present at Cashier Counter to Redeem ★", cardX + cardW / 2, cardY + cardH - 24);
  ctx.textAlign = "left";
}

function buildShareText(record) {
  const lines = [
    `${record.shopName || state.settings.shopName} - Spin & Win`,
    `ID: ${record.recordId}`,
    `Customer: ${record.customerName}`,
    `Number: ${record.customerNumber || "-"}`
  ];
  if (record.purchasedItem) {
    lines.push(`Purchased: ${record.purchasedItem}`);
  }
  lines.push(
    `Amount: ${formatAmount(record.amount)}`,
    `Prize: ${record.prize}`,
    `Status: ${getEffectiveStatus(record)}`,
    `Date: ${formatIsoDateTime(record.dateTimeIso)}`,
    `Expiry: ${formatIsoDateTime(record.expiryIso)}`
  );
  return lines.join("\n");
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
  return appendQueryParams(baseUrl, { action: "coupon_page", recordId: id });
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

async function apiListMembers() {
  const url = getAppsScriptUrl();
  if (!url) throw new Error("Apps Script URL not set.");

  const fetchUrl = appendQueryParams(url, { action: "members" });
  try {
    const response = await fetch(fetchUrl, { method: "GET" });
    return parseApiResponse(response);
  } catch (err) {
    return jsonpRequest(url, { action: "members" });
  }
}

async function syncSaveMember(member) {
  return apiMutate("POST", { action: "save_member", member });
}

async function syncDeleteMember(memberId) {
  return apiMutate("POST", { action: "delete_member", memberId });
}

async function syncAddMemberCredits(memberId, credits, purchaseAmount, note, txId, name, phone, memberType) {
  return apiMutate("POST", { action: "add_credits", memberId, credits, purchaseAmount, note, txId, name, phone, memberType });
}

async function syncDeductMemberCredits(memberId, credits, note, redeemMode, billAmount, txId) {
  return apiMutate("POST", { action: "deduct_credits", memberId, credits, note, redeemMode, billAmount, txId });
}

async function syncAllMembersToSheets(list) {
  return apiMutate("POST", { action: "sync_members", members: list != null ? list : state.members });
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
  if (!el.statusText) return;
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
  merged.spinDuration = Math.max(2, Math.min(15, num(settings.spinDuration) || 5));
  merged.manualDateEnabled = !!settings.manualDateEnabled;
  merged.manualDateTime = typeof settings.manualDateTime === "string" ? settings.manualDateTime : "";
  merged.appsScriptUrl = safeText(
    typeof settings.appsScriptUrl === "string" ? settings.appsScriptUrl : "",
    merged.appsScriptUrl
  );

  // Conditions migration and merge
  if (Array.isArray(settings.conditions) && settings.conditions.length) {
    merged.conditions = settings.conditions.map((cond) => ({
      id: cond.id || uid(),
      itemName: safeText(cond.itemName, "Unnamed Item"),
      label: safeText(cond.label, cond.itemName || "Unnamed"),
      icon: safeText(cond.icon, "🎁"),
      minAmount: Math.max(0, num(cond.minAmount)),
      prizes: Array.isArray(cond.prizes) ? cond.prizes.map((p) => ({
        id: p.id || uid(),
        name: safeText(p.name, "Unnamed Prize"),
        probability: Math.max(0, num(p.probability)),
        enabled: p.enabled !== false
      })) : []
    }));
  } else if (Array.isArray(settings.prizes) && settings.prizes.length) {
    // Migration: old flat prizes -> single "Default Wheel" condition
    merged.conditions = [{
      id: uid(),
      itemName: "All Items",
      label: "Default Wheel",
      icon: "🎁",
      minAmount: 0,
      prizes: settings.prizes.map((p) => ({
        id: p.id || uid(),
        name: safeText(p.name, "Unnamed Prize"),
        probability: Math.max(0, num(p.probability)),
        enabled: p.enabled !== false
      }))
    }];
  }
  // else: use default conditions from cloneDefaultSettings()

  // Build flat prizes from all conditions for backward compatibility
  merged.prizes = getAllPrizesFromConditions(merged.conditions);

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
    purchasedItem: safeText(record.purchasedItem, ""),
    conditionId: safeText(record.conditionId, ""),
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
    day: "2-digit",
    month: "short",
    year: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
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
  if (value == null) return fallback;
  const text = String(value).trim();
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
  const colors = [
    "#4338ca", // Royal Indigo
    "#0284c7", // Electric Sapphire
    "#059669", // Emerald Green
    "#d97706", // Amber Gold
    "#db2777", // Ruby Pink
    "#7c3aed", // Vivid Violet
    "#ea580c", // Coral Sunset
    "#0d9488"  // Deep Teal
  ];
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

function downloadCSV(filename, headers, rows) {
  const escapeCell = (cell) => {
    if (cell == null) return '""';
    const str = String(cell).replace(/"/g, '""');
    return `"${str}"`;
  };

  const csvContent = [
    headers.map(escapeCell).join(","),
    ...rows.map((row) => row.map(escapeCell).join(","))
  ].join("\r\n");

  const blob = new Blob(["\uFEFF" + csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.setAttribute("href", url);
  link.setAttribute("download", filename);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function exportRecordsToCSV() {
  const records = getAllFilteredRecords();
  if (!records.length) {
    showToast("No prize history records found to export.", "info");
    return;
  }

  const headers = [
    "Record ID",
    "Date & Time",
    "Customer Name",
    "Phone Number",
    "Purchase Amount (₹)",
    "Purchased Item",
    "Prize Won",
    "Expiry Date",
    "Status"
  ];

  const rows = records.map((r) => [
    r.recordId || "",
    formatIsoDateTime(r.dateTimeIso),
    r.customerName || "",
    formatDisplayPhone(r.customerNumber),
    formatAmount(r.amount),
    r.purchasedItem || "Wheel",
    r.prize || "",
    formatIsoDateTime(r.expiryIso),
    getEffectiveStatus(r)
  ]);

  const dateStr = new Date().toISOString().slice(0, 10);
  downloadCSV(`spinwin_prize_history_${dateStr}.csv`, headers, rows);
  showToast(`Exported ${records.length} prize record(s) to CSV!`, "success");
}

function exportMembersToCSV() {
  const members = getFilteredMembers();
  if (!members.length) {
    showToast("No members found to export.", "info");
    return;
  }

  const headers = [
    "Member ID",
    "Member Name",
    "Phone Number",
    "Member Type",
    "Active Credits",
    "Cash Worth (₹)",
    "Lifetime Earned (pts)",
    "Total Redeemed (pts)",
    "Total Cash Paid (₹)",
    "Notes",
    "Joined Date"
  ];

  const rows = members.map((m) => [
    m.id || "",
    m.name || "",
    formatDisplayPhone(m.phone),
    m.memberType || "Direct Member",
    m.credits || 0,
    creditsToCash(m.credits).toFixed(2),
    m.totalEarned || 0,
    m.totalRedeemed || 0,
    (m.totalCashPaid || 0).toFixed(2),
    m.notes || "",
    formatIsoDateTime(m.createdAtIso)
  ]);

  const dateStr = new Date().toISOString().slice(0, 10);
  downloadCSV(`spinwin_members_club_${dateStr}.csv`, headers, rows);
  showToast(`Exported ${members.length} member(s) to CSV!`, "success");
}

// ==========================================
// Audio Synthesizer Engine (Web Audio API)
// ==========================================
let audioCtx = null;

function getAudioContext() {
  if (!audioCtx) {
    const AudioContextClass = window.AudioContext || window.webkitAudioContext;
    if (AudioContextClass) {
      audioCtx = new AudioContextClass();
    }
  }
  if (audioCtx && audioCtx.state === "suspended") {
    void audioCtx.resume();
  }
  return audioCtx;
}

let activeSpinMusicStopFn = null;

function playSpinStartSound() {
  if (state.soundMuted) return;
  try {
    const ctx = getAudioContext();
    if (!ctx) return;
    const now = ctx.currentTime;

    // Aerodynamic whoosh launch sound
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    const filter = ctx.createBiquadFilter();

    osc.type = "sawtooth";
    osc.frequency.setValueAtTime(130, now);
    osc.frequency.exponentialRampToValueAtTime(700, now + 0.35);

    filter.type = "lowpass";
    filter.frequency.setValueAtTime(260, now);
    filter.frequency.exponentialRampToValueAtTime(2800, now + 0.35);

    gain.gain.setValueAtTime(0.001, now);
    gain.gain.linearRampToValueAtTime(0.18, now + 0.07);
    gain.gain.exponentialRampToValueAtTime(0.001, now + 0.45);

    osc.connect(filter);
    filter.connect(gain);
    gain.connect(ctx.destination);

    osc.start(now);
    osc.stop(now + 0.5);
  } catch (e) {
    // Ignore audio restrictions
  }
}

function startSpinMusic(durationMs) {
  stopSpinMusic();
  if (state.soundMuted) return;

  try {
    const ctx = getAudioContext();
    if (!ctx) return;

    let isPlaying = true;
    const activeNodes = [];
    const startTime = ctx.currentTime;
    const totalDuration = durationMs / 1000;

    // Upbeat energetic game-show arpeggio scale notes in Hz
    const notes = [
      261.63, 329.63, 392.00, 440.00,
      523.25, 587.33, 659.25, 783.99,
      659.25, 523.25, 440.00, 392.00
    ];
    const bassNotes = [130.81, 164.81, 196.00, 146.83]; // C3, E3, G3, D3

    let noteIdx = 0;
    let nextNoteTime = startTime + 0.04;

    function scheduleNotes() {
      if (!isPlaying) return;
      const currentCtxTime = ctx.currentTime;
      const elapsed = currentCtxTime - startTime;
      const progress = Math.min(1, elapsed / totalDuration);

      if (progress >= 0.98) {
        return;
      }

      // Tempo sync: fast upbeat (105ms) in early spin, slowing down towards end (up to 320ms) for dramatic suspense
      const noteInterval = progress < 0.65
        ? 0.105
        : 0.105 + Math.pow((progress - 0.65) / 0.35, 2) * 0.22;

      while (nextNoteTime < currentCtxTime + 0.32 && isPlaying) {
        if (nextNoteTime > startTime + totalDuration - 0.08) break;

        const noteProgress = (nextNoteTime - startTime) / totalDuration;
        const noteFreq = notes[noteIdx % notes.length];

        // Synthesize lead pulse
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        const filter = ctx.createBiquadFilter();

        osc.type = noteIdx % 2 === 0 ? "sawtooth" : "square";
        osc.frequency.setValueAtTime(noteFreq, nextNoteTime);

        filter.type = "lowpass";
        filter.frequency.setValueAtTime(Math.max(400, 1800 - noteProgress * 800), nextNoteTime);
        filter.Q.setValueAtTime(3, nextNoteTime);

        const volume = 0.075 * (1 - noteProgress * 0.3);
        gain.gain.setValueAtTime(0.001, nextNoteTime);
        gain.gain.linearRampToValueAtTime(volume, nextNoteTime + 0.015);
        gain.gain.exponentialRampToValueAtTime(0.001, nextNoteTime + noteInterval * 0.85);

        osc.connect(filter);
        filter.connect(gain);
        gain.connect(ctx.destination);

        osc.start(nextNoteTime);
        osc.stop(nextNoteTime + noteInterval);
        activeNodes.push(osc);

        // Every 4th note: punchy synth bass note
        if (noteIdx % 4 === 0) {
          const bassOsc = ctx.createOscillator();
          const bassGain = ctx.createGain();
          const bassFreq = bassNotes[(noteIdx / 4) % bassNotes.length];

          bassOsc.type = "triangle";
          bassOsc.frequency.setValueAtTime(bassFreq, nextNoteTime);

          bassGain.gain.setValueAtTime(0.12, nextNoteTime);
          bassGain.gain.exponentialRampToValueAtTime(0.001, nextNoteTime + noteInterval * 1.8);

          bassOsc.connect(bassGain);
          bassGain.connect(ctx.destination);

          bassOsc.start(nextNoteTime);
          bassOsc.stop(nextNoteTime + noteInterval * 2);
          activeNodes.push(bassOsc);
        }

        noteIdx++;
        nextNoteTime += noteInterval;
      }

      if (isPlaying && elapsed < totalDuration) {
        setTimeout(scheduleNotes, 75);
      }
    }

    scheduleNotes();

    activeSpinMusicStopFn = () => {
      isPlaying = false;
      activeNodes.forEach((node) => {
        try {
          node.stop();
          node.disconnect();
        } catch (_) { }
      });
      activeNodes.length = 0;
    };
  } catch (err) {
    console.warn("Spin audio error:", err);
  }
}

function stopSpinMusic() {
  if (typeof activeSpinMusicStopFn === "function") {
    activeSpinMusicStopFn();
    activeSpinMusicStopFn = null;
  }
}

function playTickSound() {
  if (state.soundMuted) return;
  try {
    const ctx = getAudioContext();
    if (!ctx) return;
    const now = ctx.currentTime;

    // Crisp mechanical clicker sound
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();

    const detune = (Math.random() - 0.5) * 60;
    osc.type = "triangle";
    osc.frequency.setValueAtTime(720 + detune, now);
    osc.frequency.exponentialRampToValueAtTime(160, now + 0.028);

    gain.gain.setValueAtTime(0.15, now);
    gain.gain.exponentialRampToValueAtTime(0.001, now + 0.028);

    osc.connect(gain);
    gain.connect(ctx.destination);

    osc.start(now);
    osc.stop(now + 0.032);
  } catch (e) {
    // Ignore audio restrictions
  }
}

function playWinFanfare() {
  if (state.soundMuted) return;
  try {
    const ctx = getAudioContext();
    if (!ctx) return;
    const now = ctx.currentTime;

    // Victory fanfare chords (Grand game show victory style)
    const chords = [
      { time: 0.0, freqs: [261.63, 329.63, 392.00], dur: 0.22, vol: 0.18 },
      { time: 0.22, freqs: [349.23, 440.00, 523.25], dur: 0.22, vol: 0.20 },
      { time: 0.44, freqs: [392.00, 493.88, 587.33], dur: 0.28, vol: 0.22 },
      { time: 0.72, freqs: [523.25, 659.25, 783.99, 1046.50], dur: 1.2, vol: 0.28 }
    ];

    chords.forEach(({ time, freqs, dur, vol }) => {
      const chordStart = now + time;
      freqs.forEach((freq) => {
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();

        osc.type = "triangle";
        osc.frequency.setValueAtTime(freq, chordStart);

        gain.gain.setValueAtTime(0.001, chordStart);
        gain.gain.linearRampToValueAtTime(vol / freqs.length, chordStart + 0.03);
        gain.gain.exponentialRampToValueAtTime(0.0001, chordStart + dur);

        osc.connect(gain);
        gain.connect(ctx.destination);

        osc.start(chordStart);
        osc.stop(chordStart + dur);
      });
    });

    // Sparkling victory chimes on finale chord
    const chimes = [1046.50, 1318.51, 1567.98, 2093.00, 2637.02];
    chimes.forEach((freq, idx) => {
      const chimeStart = now + 0.85 + idx * 0.08;
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();

      osc.type = "sine";
      osc.frequency.setValueAtTime(freq, chimeStart);

      gain.gain.setValueAtTime(0.001, chimeStart);
      gain.gain.linearRampToValueAtTime(0.12, chimeStart + 0.02);
      gain.gain.exponentialRampToValueAtTime(0.0001, chimeStart + 0.6);

      osc.connect(gain);
      gain.connect(ctx.destination);

      osc.start(chimeStart);
      osc.stop(chimeStart + 0.65);
    });
  } catch (e) {
    // Ignore
  }
}

function triggerPointerWobble() {
  if (!el.wheelPointer) return;
  el.wheelPointer.classList.remove("tick-wobble");
  void el.wheelPointer.offsetWidth; // Force reflow
  el.wheelPointer.classList.add("tick-wobble");
}

function updateSoundButtonUI() {
  if (!el.soundIcon) return;
  el.soundIcon.textContent = state.soundMuted ? "🔇" : "🔊";
  if (el.soundToggleBtn) {
    el.soundToggleBtn.setAttribute("title", state.soundMuted ? "Unmute Sound" : "Mute Sound");
  }
}

// ==========================================
// Confetti Celebration Particle Engine
// ==========================================
function triggerConfetti() {
  const canvas = el.confettiCanvas;
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  if (!ctx) return;

  canvas.width = window.innerWidth;
  canvas.height = window.innerHeight;

  const count = 90;
  const particles = [];
  const colors = ["#fbbf24", "#f59e0b", "#6366f1", "#10b981", "#ec4899", "#38bdf8", "#ffffff"];

  for (let i = 0; i < count; i++) {
    particles.push({
      x: window.innerWidth / 2 + (Math.random() - 0.5) * 200,
      y: window.innerHeight * 0.38,
      vx: (Math.random() - 0.5) * 16,
      vy: Math.random() * -14 - 6,
      size: Math.random() * 8 + 6,
      color: colors[Math.floor(Math.random() * colors.length)],
      rotation: Math.random() * 360,
      rotSpeed: (Math.random() - 0.5) * 12,
      gravity: 0.38,
      opacity: 1,
      shape: Math.random() > 0.4 ? "rect" : "circle"
    });
  }

  const startTime = performance.now();
  const duration = 2800;

  function renderFrame(now) {
    const elapsed = now - startTime;
    const progress = Math.min(1, elapsed / duration);
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    particles.forEach((p) => {
      p.x += p.vx;
      p.y += p.vy;
      p.vy += p.gravity;
      p.rotation += p.rotSpeed;
      p.opacity = Math.max(0, 1 - progress);

      ctx.save();
      ctx.globalAlpha = p.opacity;
      ctx.translate(p.x, p.y);
      ctx.rotate((p.rotation * Math.PI) / 180);
      ctx.fillStyle = p.color;

      if (p.shape === "rect") {
        ctx.fillRect(-p.size / 2, -p.size / 2, p.size, p.size * 0.6);
      } else {
        ctx.beginPath();
        ctx.arc(0, 0, p.size / 2, 0, Math.PI * 2);
        ctx.fill();
      }
      ctx.restore();
    });

    if (progress < 1) {
      requestAnimationFrame(renderFrame);
    } else {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
    }
  }

  requestAnimationFrame(renderFrame);
}

// ==========================================
// Winner Celebration Modal & Spin Time Helpers
// ==========================================
function openWinnerModal(record) {
  if (!el.winnerModal) return;
  if (el.winnerPrizeName) el.winnerPrizeName.textContent = record.prize;
  if (el.winnerCustomerMsg) el.winnerCustomerMsg.textContent = `Awarded to ${record.customerName}`;
  if (el.winnerRecordId) el.winnerRecordId.textContent = record.recordId;
  if (el.winnerItemText && record.purchasedItem) {
    el.winnerItemText.textContent = `Purchased: ${record.purchasedItem}`;
  }
  if (el.winnerItemPill) {
    el.winnerItemPill.classList.toggle("hidden", !record.purchasedItem);
  }
  el.winnerModal.classList.remove("hidden");
  el.winnerModal.setAttribute("aria-hidden", "false");
}

function closeWinnerModal() {
  if (!el.winnerModal || el.winnerModal.classList.contains("hidden")) return;
  el.winnerModal.classList.add("hidden");
  el.winnerModal.setAttribute("aria-hidden", "true");
}

function scrollToCouponPreview() {
  const target = el.resultDetails || el.couponCanvas;
  if (target) {
    target.scrollIntoView({ behavior: "smooth", block: "center" });
  }
}

function resetFormForNewSpin() {
  el.customerName.value = "";
  el.customerNumber.value = "";
  el.purchaseAmount.value = "";
  state.spinLockedForEntry = false;
  state.currentResultId = null;
  state.lastEntrySignature = "";
  if (el.customerMemberBadge) el.customerMemberBadge.classList.add("hidden");
  renderResult();
  drawCouponPlaceholder();
  refreshUI();
  updateStatus("Ready for next lucky customer.");
  el.customerName.focus();
}

function setSpinDuration(seconds, syncInput = true) {
  const duration = Math.max(2, Math.min(15, parseInt(seconds, 10) || 5));
  state.settings.spinDuration = duration;
  persistSettings();
  if (syncInput && el.spinDurationInput) {
    el.spinDurationInput.value = String(duration);
  }
  updateSpinPillsUI();
  showToast(`Spin time: ${duration}s`, "info");
}

function updateSpinPillsUI() {
  if (!el.spinTimePills) return;
  const currentDur = state.settings.spinDuration || 5;
  const pills = el.spinTimePills.querySelectorAll(".time-pill-btn");
  pills.forEach((btn) => {
    const dur = parseInt(btn.dataset.duration, 10);
    btn.classList.toggle("is-active", dur === currentDur);
  });
}

function highlightInputError(inputEl) {
  if (!inputEl) return;
  inputEl.classList.remove("shake-input");
  void inputEl.offsetWidth; // Force reflow
  inputEl.classList.add("shake-input");
  inputEl.focus();
  setTimeout(() => {
    inputEl.classList.remove("shake-input");
  }, 600);
}

// ==========================================
// Multi-Condition Spin Management System
// ==========================================

function initConditions() {
  const conditions = state.settings.conditions || [];
  if (conditions.length > 0 && !state.selectedConditionId) {
    state.selectedConditionId = conditions[0].id;
  }
  renderConditionTabs();
  renderConditionEditor();
  populateSpinConditionSelect();
}

function getActiveCondition() {
  const conditions = state.settings.conditions || [];
  if (!state.selectedConditionId && conditions.length) {
    state.selectedConditionId = conditions[0].id;
  }
  return conditions.find((c) => c.id === state.selectedConditionId) || conditions[0] || null;
}

function getSpinCondition() {
  const condId = el.spinConditionSelect ? el.spinConditionSelect.value : state.selectedConditionId;
  const conditions = state.settings.conditions || [];
  return conditions.find((c) => c.id === condId) || conditions[0] || null;
}

function getActiveSpinPrizes() {
  const cond = getSpinCondition();
  return cond ? cond.prizes : [];
}

function getConditionLabel(condId) {
  const conditions = state.settings.conditions || [];
  const cond = conditions.find((c) => c.id === condId);
  return cond ? `${cond.icon || "🎁"} ${cond.label || cond.itemName}` : "Unknown";
}

function getAllPrizesFlat() {
  return getAllPrizesFromConditions(state.settings.conditions);
}

function getAllPrizesFromConditions(conditions) {
  const all = [];
  (conditions || []).forEach((cond) => {
    (cond.prizes || []).forEach((p) => all.push(p));
  });
  return all;
}

function renderConditionTabs() {
  if (!el.conditionTabsBar) return;
  el.conditionTabsBar.innerHTML = "";
  const conditions = state.settings.conditions || [];
  const activeId = state.selectedConditionId || (conditions[0] && conditions[0].id);

  conditions.forEach((cond, idx) => {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = `condition-tab-chip${cond.id === activeId ? " is-active" : ""}`;
    chip.dataset.conditionId = cond.id;
    chip.setAttribute("role", "tab");
    chip.setAttribute("aria-selected", String(cond.id === activeId));
    chip.innerHTML = `<span class="chip-icon">${cond.icon || "🎁"}</span> <span class="chip-label">${safeText(cond.label || cond.itemName, `Wheel ${idx + 1}`)}</span>`;
    el.conditionTabsBar.appendChild(chip);
  });
}

function renderConditionEditor() {
  const cond = getActiveCondition();
  if (!cond) {
    if (el.conditionEditorCard) el.conditionEditorCard.classList.add("hidden");
    return;
  }
  if (el.conditionEditorCard) el.conditionEditorCard.classList.remove("hidden");
  if (el.conditionEditorTitle) {
    el.conditionEditorTitle.textContent = `${cond.icon || "🎁"} ${cond.label || cond.itemName || "Condition"} Settings`;
  }
  if (el.conditionItemNameInput) el.conditionItemNameInput.value = cond.itemName || "";
  if (el.conditionLabelInput) el.conditionLabelInput.value = cond.label || "";
  if (el.conditionIconInput) el.conditionIconInput.value = cond.icon || "";
  if (el.conditionMinAmountInput) el.conditionMinAmountInput.value = cond.minAmount ? String(cond.minAmount) : "";

  // Show/hide delete button (at least 1 condition must remain)
  if (el.deleteConditionBtn) {
    el.deleteConditionBtn.classList.toggle("hidden", (state.settings.conditions || []).length <= 1);
  }
}

function switchSettingsCondition(condId) {
  // Save current condition editor state first
  syncConditionDraftFromEditor();
  state.selectedConditionId = condId;
  renderConditionTabs();
  renderConditionEditor();
  renderPrizeTable();
}

function syncConditionFieldsToActive() {
  const cond = getActiveCondition();
  if (!cond) return;
  if (el.conditionItemNameInput) cond.itemName = (el.conditionItemNameInput.value || "").trim();
  if (el.conditionLabelInput) cond.label = (el.conditionLabelInput.value || "").trim();
  if (el.conditionIconInput) cond.icon = (el.conditionIconInput.value || "").trim();
  if (el.conditionMinAmountInput) cond.minAmount = Math.max(0, num(el.conditionMinAmountInput.value));
  // Update tab label live
  renderConditionTabs();
}

function syncConditionDraftFromEditor() {
  syncConditionFieldsToActive();
  syncPrizeDraftFromTable();
}

function onAddCondition() {
  syncConditionDraftFromEditor();
  const conditions = state.settings.conditions || [];
  const newId = uid();
  const newIndex = conditions.length + 1;
  conditions.push({
    id: newId,
    itemName: "",
    label: `New Wheel ${newIndex}`,
    icon: "🎁",
    minAmount: 0,
    prizes: [
      { id: uid(), name: "Try Again", probability: 50, enabled: true },
      { id: uid(), name: "Special Prize", probability: 50, enabled: true }
    ]
  });
  state.settings.conditions = conditions;
  state.selectedConditionId = newId;
  persistSettings();
  renderConditionTabs();
  renderConditionEditor();
  renderPrizeTable();
  populateSpinConditionSelect();
  showToast("New condition added. Configure the item name, prizes, and probabilities.", "success");
}

function onDeleteCondition() {
  const conditions = state.settings.conditions || [];
  if (conditions.length <= 1) {
    showToast("Cannot delete the last condition. At least one must remain.", "error");
    return;
  }
  const cond = getActiveCondition();
  if (!cond) return;
  const ok = confirm(`Delete condition "${cond.icon} ${cond.itemName || cond.label}"? This will remove all its prizes.`);
  if (!ok) return;

  state.settings.conditions = conditions.filter((c) => c.id !== cond.id);
  state.selectedConditionId = state.settings.conditions[0] ? state.settings.conditions[0].id : null;
  persistSettings();
  renderConditionTabs();
  renderConditionEditor();
  renderPrizeTable();
  populateSpinConditionSelect();
  drawWheel();
  showToast(`Condition "${cond.itemName || cond.label}" deleted.`, "info");
}

function populateSpinConditionSelect() {
  if (!el.spinConditionSelect) return;
  const conditions = state.settings.conditions || [];
  el.spinConditionSelect.innerHTML = "";
  conditions.forEach((cond) => {
    const opt = document.createElement("option");
    opt.value = cond.id;
    opt.textContent = `${cond.icon || "🎁"} ${cond.label || cond.itemName || "Wheel"}`;
    el.spinConditionSelect.appendChild(opt);
  });
  // Restore selected condition or default to first
  if (state.selectedConditionId && conditions.some((c) => c.id === state.selectedConditionId)) {
    el.spinConditionSelect.value = state.selectedConditionId;
  } else if (conditions.length) {
    state.selectedConditionId = conditions[0].id;
    el.spinConditionSelect.value = conditions[0].id;
  }
}

function updateConditionProbSummary() {
  if (!el.conditionProbSummary) return;
  const cond = getActiveCondition();
  if (!cond || !cond.prizes) {
    el.conditionProbSummary.textContent = "Total Probability: 0%";
    return;
  }
  const total = cond.prizes.reduce((sum, p) => sum + num(p.probability), 0);
  const display = Math.round(total * 100) / 100;
  const isValid = Math.abs(total - 100) < 0.1;
  el.conditionProbSummary.textContent = `Total Probability: ${display}%`;
  el.conditionProbSummary.style.color = isValid ? "" : "#ef4444";
}

/* ==========================================================================
   Members Club & Loyalty Credit Engine
   Rule: 100 Credits = ₹1.00 (1 Credit = ₹0.01)
   ========================================================================== */

function creditsToCash(credits) {
  const numCredits = Number(credits) || 0;
  return Math.round((numCredits * (CREDIT_RATE_PER_100 / 100)) * 100) / 100;
}

function cashToCredits(cash) {
  const numCash = Number(cash) || 0;
  return Math.round((numCash / (CREDIT_RATE_PER_100 / 100)));
}

function loadMembers() {
  try {
    const raw = localStorage.getItem(MEMBERS_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.map(normalizeMember).filter((m) => !!m.id);
  } catch {
    return [];
  }
}

function persistMembers() {
  localStorage.setItem(MEMBERS_KEY, JSON.stringify(state.members));
}

function normalizeMember(m) {
  const now = new Date().toISOString();
  return {
    id: safeText(m.id, uid()),
    name: safeText(m.name, "Member"),
    phone: safeText(m.phone, ""),
    memberType: safeText(m.memberType, "Direct Member"),
    credits: Math.max(0, Math.round(Number(m.credits) || 0)),
    totalEarned: Math.max(0, Math.round(Number(m.totalEarned) || 0)),
    totalRedeemed: Math.max(0, Math.round(Number(m.totalRedeemed) || 0)),
    totalCashPaid: round2(Number(m.totalCashPaid) || 0),
    notes: safeText(m.notes, ""),
    transactions: Array.isArray(m.transactions) ? m.transactions.map(normalizeMemberTransaction) : [],
    createdAtIso: safeText(m.createdAtIso, now),
    updatedAtIso: safeText(m.updatedAtIso, now),
    synced: !!m.synced
  };
}

function markMembersSyncedAll() {
  let changed = false;
  state.members.forEach((m) => {
    if (!m.synced) {
      m.synced = true;
      changed = true;
    }
  });
  if (changed) persistMembers();
}

function normalizeMemberTransaction(tx) {
  return {
    id: safeText(tx.id, uid()),
    type: tx.type === "DEBIT" ? "DEBIT" : "CREDIT",
    credits: Math.round(Number(tx.credits) || 0),
    cashValue: round2(Number(tx.cashValue) || 0),
    purchaseAmount: round2(Number(tx.purchaseAmount) || 0),
    note: safeText(tx.note, ""),
    dateIso: safeText(tx.dateIso, new Date().toISOString()),
    balanceAfter: Math.round(Number(tx.balanceAfter) || 0)
  };
}

function formatDisplayPhone(str) {
  const text = safeText(str).trim();
  if (!text || text.toUpperCase().includes("ERROR") || text === "-" || text === "null" || text === "undefined" || text === "0") {
    return "";
  }
  return text;
}

function cleanDigits(str) {
  return String(str || "").replace(/\D+/g, "");
}

function findMemberByPhone(phone) {
  const clean = cleanDigits(phone);
  if (!clean || clean.length < 5) return null;
  return state.members.find((m) => {
    const mClean = cleanDigits(m.phone);
    return mClean && (mClean === clean || mClean.endsWith(clean) || clean.endsWith(mClean));
  });
}

function findMemberByName(name) {
  const clean = safeText(name).trim().toLowerCase();
  if (!clean || clean.length < 2) return null;
  return state.members.find((m) => m.name.trim().toLowerCase() === clean);
}

function addOrUpdateMember({ id, name, phone, memberType, initialCredits, notes }) {
  const cleanPhone = cleanDigits(phone);
  const now = new Date().toISOString();

  if (id) {
    const member = state.members.find((m) => m.id === id);
    if (!member) return null;
    member.name = safeText(name, member.name);
    member.phone = safeText(phone, member.phone);
    member.memberType = safeText(memberType, member.memberType);
    member.notes = safeText(notes, member.notes);
    member.updatedAtIso = now;
    persistMembers();
    renderMembersSection();
    showToast(`Member "${member.name}" updated successfully.`, "success");
    if (getAppsScriptUrl()) {
      syncSaveMember(member).then(() => {
        member.synced = true;
        persistMembers();
      }).catch((err) => console.warn("Cloud member update warning:", err));
    }
    return member;
  }

  if (cleanPhone) {
    const existing = findMemberByPhone(cleanPhone);
    if (existing) {
      showToast(`A member with phone ${phone} already exists (${existing.name}).`, "error");
      return null;
    }
  }

  const credits = Math.max(0, Math.round(Number(initialCredits) || 0));
  const newMember = {
    id: uid(),
    name: safeText(name, "New Member"),
    phone: safeText(phone, ""),
    memberType: safeText(memberType, "Direct Member"),
    credits,
    totalEarned: credits,
    totalRedeemed: 0,
    totalCashPaid: 0,
    notes: safeText(notes, ""),
    transactions: [],
    createdAtIso: now,
    updatedAtIso: now
  };

  if (credits > 0) {
    newMember.transactions.push({
      id: uid(),
      type: "CREDIT",
      credits,
      cashValue: creditsToCash(credits),
      purchaseAmount: 0,
      note: "Initial Welcome Bonus",
      dateIso: now,
      balanceAfter: credits
    });
  }

  state.members.unshift(newMember);
  persistMembers();
  renderMembersSection();
  renderRecordsTable();
  showToast(`Member "${newMember.name}" registered successfully!`, "success");
  if (getAppsScriptUrl()) {
    syncSaveMember(newMember).then(() => {
      newMember.synced = true;
      persistMembers();
    }).catch((err) => console.warn("Cloud member create warning:", err));
  }
  return newMember;
}

function addMemberCredits(memberId, creditsToAdd, purchaseAmount, note) {
  const member = state.members.find((m) => m.id === memberId);
  if (!member) return false;

  const credits = Math.max(1, Math.round(Number(creditsToAdd) || 0));
  const cashVal = creditsToCash(credits);
  const now = new Date().toISOString();

  member.credits += credits;
  member.totalEarned += credits;
  member.updatedAtIso = now;

  const tx = {
    id: uid(),
    type: "CREDIT",
    credits,
    cashValue: cashVal,
    purchaseAmount: round2(Number(purchaseAmount) || 0),
    note: safeText(note, purchaseAmount ? `Purchase Bill ₹${purchaseAmount}` : "Loyalty Credit"),
    dateIso: now,
    balanceAfter: member.credits
  };
  member.transactions.unshift(tx);

  persistMembers();
  renderMembersSection();
  checkCustomerMemberMatch();
  renderRecordsTable();
  showToast(`+${credits} Credits (₹${cashVal.toFixed(2)}) awarded to ${member.name}!`, "success");

  if (getAppsScriptUrl()) {
    syncAddMemberCredits(memberId, credits, purchaseAmount, note, tx.id, member.name, member.phone, member.memberType).then(() => {
      member.synced = true;
      persistMembers();
    }).catch((err) => console.warn("Cloud add credits warning:", err));
  }
  return true;
}

function deductMemberCredits(memberId, creditsToDeduct, note, sendReceipt = false, redeemMode = "purchase_discount", billAmount = 0) {
  const member = state.members.find((m) => m.id === memberId);
  if (!member) return { success: false, error: "Member not found" };

  const credits = Math.max(1, Math.round(Number(creditsToDeduct) || 0));
  if (member.credits < credits) {
    return {
      success: false,
      error: `Insufficient balance! ${member.name} has only ${member.credits} credits (Max Value: ₹${creditsToCash(member.credits).toFixed(2)}).`
    };
  }

  const cashVal = creditsToCash(credits);
  const now = new Date().toISOString();
  const isDiscount = redeemMode === "purchase_discount";

  member.credits -= credits;
  member.totalRedeemed += credits;
  if (!isDiscount) {
    member.totalCashPaid = round2(member.totalCashPaid + cashVal);
  }
  member.updatedAtIso = now;

  let defaultNote = "";
  if (isDiscount) {
    defaultNote = billAmount > 0
      ? `Purchase Bill Discount: -₹${cashVal.toFixed(2)} off ₹${num(billAmount).toFixed(2)}`
      : `Purchase Item Discount (-₹${cashVal.toFixed(2)})`;
  } else {
    defaultNote = "Cash Reward Redemption across counter";
  }

  const tx = {
    id: uid(),
    type: "DEBIT",
    redeemType: isDiscount ? "DISCOUNT" : "CASH",
    credits,
    cashValue: cashVal,
    purchaseAmount: isDiscount ? round2(num(billAmount)) : 0,
    note: safeText(note, defaultNote),
    dateIso: now,
    balanceAfter: member.credits
  };
  member.transactions.unshift(tx);

  persistMembers();
  renderMembersSection();
  checkCustomerMemberMatch();
  renderRecordsTable();

  if (sendReceipt && member.phone) {
    sendRedemptionWhatsApp(member, cashVal, credits, member.credits, isDiscount, billAmount);
  }

  if (isDiscount) {
    const netPayable = Math.max(0, num(billAmount) - cashVal);
    showToast(`Deducted ${credits} Credits. Applied ₹${cashVal.toFixed(2)} discount! Customer pays ₹${netPayable.toFixed(2)}.`, "success");
  } else {
    showToast(`Deducted ${credits} Credits. Pay ₹${cashVal.toFixed(2)} cash to ${member.name}.`, "success");
  }

  if (getAppsScriptUrl()) {
    syncDeductMemberCredits(memberId, credits, note, redeemMode, billAmount, tx.id).then(() => {
      member.synced = true;
      persistMembers();
    }).catch((err) => console.warn("Cloud deduct credits warning:", err));
  }

  return { success: true, cashValue: cashVal, newBalance: member.credits, member };
}

function deleteMember(memberId) {
  const member = state.members.find((m) => m.id === memberId);
  if (!member) return;
  const ok = confirm(`Delete member "${member.name}"? This will permanently remove their record and ${member.credits} credits.`);
  if (!ok) return;

  state.members = state.members.filter((m) => m.id !== memberId);
  persistMembers();
  renderMembersSection();
  checkCustomerMemberMatch();
  renderRecordsTable();
  showToast(`Member "${member.name}" deleted.`, "info");

  if (getAppsScriptUrl()) {
    syncDeleteMember(memberId).catch((err) => console.warn("Cloud delete member warning:", err));
  }
}

function sendRedemptionWhatsApp(member, discountOrCash, creditsDeducted, newBalance, isDiscount, billAmount) {
  const cleanPhone = cleanDigits(member.phone);
  if (!cleanPhone) return;

  const phoneParam = cleanPhone.length === 10 ? `91${cleanPhone}` : cleanPhone;
  const shopName = state.settings.shopName || "Our Shop";
  let message = "";

  if (isDiscount) {
    const netBill = Math.max(0, num(billAmount) - discountOrCash);
    message = `*${shopName} - Purchase Reward Redemption*\n\n` +
      `Hello *${member.name}*,\n` +
      `You redeemed *${creditsDeducted} Credits* for a *₹${discountOrCash.toFixed(2)} Discount* on your purchase.\n` +
      (billAmount > 0 ? `🧾 Original Bill: ₹${num(billAmount).toFixed(2)}\n💰 Discount Applied: -₹${discountOrCash.toFixed(2)}\n🏷️ Final Amount Paid: *₹${netBill.toFixed(2)}*\n\n` : `\n`) +
      `💳 Remaining Credit Balance: *${newBalance} Credits* (₹${creditsToCash(newBalance).toFixed(2)})\n` +
      `Thank you for shopping with us! 🛍️`;
  } else {
    message = `*${shopName} - Cash Reward Redemption Receipt*\n\n` +
      `Hello *${member.name}*,\n` +
      `You have redeemed *${creditsDeducted} Credits* for *₹${discountOrCash.toFixed(2)} Cash* across counter.\n\n` +
      `💳 Remaining Credit Balance: *${newBalance} Credits* (₹${creditsToCash(newBalance).toFixed(2)})\n` +
      `Thank you for being our valuable customer! 🛍️`;
  }

  const url = `https://wa.me/${encodeURIComponent(phoneParam)}?text=${encodeURIComponent(message)}`;
  window.open(url, "_blank");
}

/* ==========================================================================
   Members UI Rendering
   ========================================================================== */

function renderMembersSection() {
  renderMembersStats();
  renderMembersTable();
  populateMemberSelectDropdowns();
}

function renderMembersStats() {
  const members = state.members || [];
  const totalCount = members.length;
  const winnersCount = members.filter((m) => m.memberType === "Spin Winner").length;
  const directCount = totalCount - winnersCount;

  let totalCredits = 0;
  let totalEarned = 0;
  let totalRedeemed = 0;
  let totalCashPaid = 0;

  members.forEach((m) => {
    totalCredits += m.credits;
    totalEarned += m.totalEarned;
    totalRedeemed += m.totalRedeemed;
    totalCashPaid += m.totalCashPaid;
  });

  if (el.statTotalMembers) el.statTotalMembers.textContent = String(totalCount);
  if (el.statMembersSubtitle) el.statMembersSubtitle.textContent = `${winnersCount} Winners · ${directCount} Direct`;
  if (el.statTotalCredits) el.statTotalCredits.textContent = totalCredits.toLocaleString();
  if (el.statCreditsWorth) el.statCreditsWorth.textContent = `Worth ≈ ₹${creditsToCash(totalCredits).toFixed(2)} cash`;
  if (el.statTotalEarned) el.statTotalEarned.textContent = totalEarned.toLocaleString();
  if (el.statEarnedWorth) el.statEarnedWorth.textContent = `Issued value ≈ ₹${creditsToCash(totalEarned).toFixed(2)}`;
  if (el.statTotalCashPaid) el.statTotalCashPaid.textContent = `₹${totalCashPaid.toFixed(2)}`;
  if (el.statRedeemedCredits) el.statRedeemedCredits.textContent = `${totalRedeemed.toLocaleString()} credits redeemed`;
}

function getFilteredMembers() {
  const members = state.members || [];
  const query = (state.memberSearchQuery || "").trim().toLowerCase();
  const filter = state.memberFilter || "all";

  return members.filter((m) => {
    // Filter type
    if (filter === "winner" && m.memberType !== "Spin Winner") return false;
    if (filter === "direct" && m.memberType === "Spin Winner") return false;
    if (filter === "has_balance" && m.credits <= 0) return false;

    // Search query
    if (query) {
      const matchName = m.name.toLowerCase().includes(query);
      const matchPhone = m.phone.toLowerCase().includes(query);
      const matchNotes = (m.notes || "").toLowerCase().includes(query);
      if (!matchName && !matchPhone && !matchNotes) return false;
    }

    return true;
  });
}

function getVisibleMembers() {
  const all = getFilteredMembers();
  const pageSize = state.membersPageSize || 10;
  const totalPages = Math.max(1, Math.ceil(all.length / pageSize));
  if (state.membersPage > totalPages) state.membersPage = totalPages;
  if (state.membersPage < 1) state.membersPage = 1;

  const start = (state.membersPage - 1) * pageSize;
  return all.slice(start, start + pageSize);
}

// =====================================
// Members Bulk Selection (like Prize History)
// =====================================

function onMemberSelectionChange(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;
  if (!target.classList.contains("member-select-checkbox")) return;

  const memberId = safeText(target.dataset.id);
  if (!memberId) return;

  if (target.checked) {
    state.selectedMemberIds.add(memberId);
  } else {
    state.selectedMemberIds.delete(memberId);
  }
  updateMemberBulkSelectionUI();
}

function onSelectAllMembersChange() {
  if (!el.selectAllMembers) return;
  const visibleIds = getVisibleMembers().map((m) => m.id);
  if (el.selectAllMembers.checked) {
    visibleIds.forEach((id) => state.selectedMemberIds.add(id));
  } else {
    visibleIds.forEach((id) => state.selectedMemberIds.delete(id));
  }
  renderMembersTable();
}

function onClearMemberSelection() {
  state.selectedMemberIds.clear();
  renderMembersTable();
}

async function onApplyMemberBulkAction() {
  const action = safeText(el.memberBulkActionSelect && el.memberBulkActionSelect.value);
  const selectedIds = getSelectedExistingMemberIds();
  if (!action) {
    showToast("Choose a bulk action first.", "info");
    return;
  }
  if (!selectedIds.length) {
    showToast("Select at least one member.", "info");
    return;
  }

  if (action === "delete") {
    const names = selectedIds.map((id) => {
      const found = state.members.find((m) => m.id === id);
      return found ? found.name : id;
    });
    const confirmed = confirm(
      `Delete ${selectedIds.length} selected member(s)?\n${names.slice(0, 5).join(", ")}${names.length > 5 ? " …" : ""}\n\nThis will permanently remove their records and credits.`
    );
    if (!confirmed) return;
    await runBulkDeleteMembers(selectedIds);
    return;
  }

  if (action === "sync") {
    await runBulkSyncMembers(selectedIds);
  }
}

function getSelectedExistingMemberIds() {
  const existing = new Set(state.members.map((m) => m.id));
  return Array.from(state.selectedMemberIds).filter((id) => existing.has(id));
}

function pruneMemberSelectionForMissingMembers() {
  const existing = new Set(state.members.map((m) => m.id));
  Array.from(state.selectedMemberIds).forEach((id) => {
    if (!existing.has(id)) {
      state.selectedMemberIds.delete(id);
    }
  });
}

function updateMemberBulkSelectionUI() {
  const selectedCount = getSelectedExistingMemberIds().length;
  if (el.applyMemberBulkActionBtn) {
    el.applyMemberBulkActionBtn.disabled = selectedCount < 1;
  }

  if (el.memberBulkActionsBar) {
    const showBar = selectedCount > 0;
    el.memberBulkActionsBar.classList.toggle("hidden", !showBar);
    el.memberBulkActionsBar.classList.toggle("visible", showBar);
  }

  if (!el.selectAllMembers) return;
  const visible = getVisibleMembers();
  const visibleIds = visible.map((m) => m.id);
  const selectedVisibleCount = visibleIds.filter((id) => state.selectedMemberIds.has(id)).length;

  const hasVisible = visibleIds.length > 0;
  const allVisibleSelected = hasVisible && selectedVisibleCount === visibleIds.length;
  const partiallySelected = selectedVisibleCount > 0 && !allVisibleSelected;

  el.selectAllMembers.checked = allVisibleSelected;
  el.selectAllMembers.indeterminate = partiallySelected;
}

async function runBulkDeleteMembers(memberIds) {
  const removeSet = new Set(memberIds);
  state.members = state.members.filter((m) => !removeSet.has(m.id));
  state.selectedMemberIds.clear();
  persistMembers();
  renderMembersSection();
  checkCustomerMemberMatch();
  renderRecordsTable();

  if (!getAppsScriptUrl()) {
    showToast(`Deleted ${memberIds.length} member(s) locally. Add Apps Script URL to sync.`, "info");
    return;
  }

  const results = await Promise.allSettled(
    memberIds.map((memberId) => syncDeleteMember(memberId))
  );

  let synced = 0;
  results.forEach((result) => {
    if (result.status === "fulfilled" && result.value && result.value.ok) {
      synced += 1;
    }
  });

  const failed = memberIds.length - synced;
  if (failed > 0) {
    showToast(`Bulk delete done. Synced ${synced}, failed ${failed} to Sheets.`, "error");
  } else {
    showToast(`Deleted ${synced} member(s).`, "success");
  }
}

async function runBulkSyncMembers(memberIds) {
  const selectedMembers = memberIds
    .map((id) => state.members.find((m) => m.id === id))
    .filter((m) => !!m);

  if (!selectedMembers.length) {
    showToast("No members to sync.", "info");
    return;
  }

  try {
    const res = await syncAllMembersToSheets(selectedMembers);
    if (!res || !res.ok) {
      throw new Error((res && res.message) || "Failed to sync members to Google Sheets.");
    }
    selectedMembers.forEach((m) => {
      m.synced = true;
    });
    persistMembers();
    state.selectedMemberIds.clear();
    renderMembersTable();
    showToast(`Synced ${selectedMembers.length} member(s) to Google Sheets.`, "success");
  } catch (err) {
    showToast(`Sync failed: ${err.message}`, "error");
  }
}

function getLastMembersPage() {
  const all = getFilteredMembers();
  const pageSize = state.membersPageSize || 10;
  return Math.max(1, Math.ceil(all.length / pageSize));
}

function goToMembersPage(page) {
  const totalPages = getLastMembersPage();
  const target = Math.max(1, Math.min(totalPages, page));
  if (target === state.membersPage) return;
  state.membersPage = target;
  renderMembersTable();
}

function renderMembersPaginationControls(totalCount, start, end, totalPages) {
  if (!el.membersPaginationBar) return;

  if (el.membersPaginationRangeText) {
    if (totalCount === 0) {
      el.membersPaginationRangeText.innerHTML = "Showing <strong>0</strong> of <strong>0</strong> members";
    } else {
      el.membersPaginationRangeText.innerHTML = `Showing <strong>${start}–${end}</strong> of <strong>${totalCount}</strong> members (Page ${state.membersPage}/${totalPages})`;
    }
  }

  if (el.membersPageSizeSelect) {
    el.membersPageSizeSelect.value = String(state.membersPageSize || 10);
  }

  const isFirst = state.membersPage <= 1;
  const isLast = state.membersPage >= totalPages;

  if (el.membersFirstPageBtn) el.membersFirstPageBtn.disabled = isFirst;
  if (el.membersPrevPageBtn) el.membersPrevPageBtn.disabled = isFirst;
  if (el.membersNextPageBtn) el.membersNextPageBtn.disabled = isLast;
  if (el.membersLastPageBtn) el.membersLastPageBtn.disabled = isLast;

  if (!el.membersPaginationPages) return;
  el.membersPaginationPages.innerHTML = "";

  const pages = generatePaginationPageNumbers(state.membersPage, totalPages);
  pages.forEach((item) => {
    if (item === "...") {
      const span = document.createElement("span");
      span.className = "pagination-ellipsis";
      span.textContent = "…";
      el.membersPaginationPages.appendChild(span);
    } else {
      const pageNum = item;
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = `page-btn ${pageNum === state.membersPage ? "is-active" : ""}`;
      btn.textContent = String(pageNum);
      btn.setAttribute("aria-label", `Page ${pageNum}`);
      if (pageNum === state.membersPage) {
        btn.setAttribute("aria-current", "page");
      }
      btn.addEventListener("click", () => goToMembersPage(pageNum));
      el.membersPaginationPages.appendChild(btn);
    }
  });
}

function renderMembersTable() {
  if (!el.membersTableBody) return;
  pruneMemberSelectionForMissingMembers();
  const filtered = getFilteredMembers();
  const totalCount = filtered.length;
  const pageSize = state.membersPageSize || 10;
  const totalPages = Math.max(1, Math.ceil(totalCount / pageSize));

  if (state.membersPage > totalPages) state.membersPage = totalPages;
  if (state.membersPage < 1) state.membersPage = 1;

  const startIndex = (state.membersPage - 1) * pageSize;
  const visible = filtered.slice(startIndex, startIndex + pageSize);
  const endIndex = Math.min(startIndex + visible.length, totalCount);

  el.membersTableBody.innerHTML = "";

  if (!totalCount) {
    if (el.membersEmptyState) {
      el.membersEmptyState.classList.remove("hidden");
      if (el.membersEmptyMsg) {
        const query = (state.memberSearchQuery || "").trim();
        const filter = state.memberFilter || "all";
        el.membersEmptyMsg.textContent = query || filter !== "all"
          ? "No members match your search or filter criteria."
          : "Add customers who win prizes or purchase membership to start tracking credits!";
      }
    }
    renderMembersPaginationControls(0, 0, 0, 1);
    updateMemberBulkSelectionUI();
    return;
  }

  if (el.membersEmptyState) el.membersEmptyState.classList.add("hidden");

  visible.forEach((m) => {
    const tr = document.createElement("tr");
    const isSelected = state.selectedMemberIds.has(m.id);
    if (isSelected) {
      tr.classList.add("record-row-selected");
    }

    // Selection checkbox
    const selectTd = document.createElement("td");
    selectTd.className = "centered";
    const check = document.createElement("input");
    check.type = "checkbox";
    check.className = "member-select-checkbox";
    check.dataset.id = m.id;
    check.checked = isSelected;
    check.setAttribute("aria-label", `Select member ${m.name || m.id}`);
    selectTd.appendChild(check);
    tr.appendChild(selectTd);

    // Member Name & Avatar
    const tdName = document.createElement("td");
    const nameCell = document.createElement("div");
    nameCell.className = "member-name-cell";

    const initial = (m.name || "M").trim().charAt(0).toUpperCase();
    const avatar = document.createElement("div");
    avatar.className = "member-avatar";
    avatar.textContent = initial;

    const nameTextWrap = document.createElement("div");
    nameTextWrap.style.minWidth = "0";

    const nameText = document.createElement("div");
    nameText.className = "member-name-text text-truncate";
    nameText.textContent = m.name;
    nameText.title = m.name;

    nameTextWrap.appendChild(nameText);
    if (m.notes) {
      const notesSub = document.createElement("div");
      notesSub.className = "muted text-truncate";
      notesSub.style.fontSize = "0.72rem";
      notesSub.style.maxWidth = "120px";
      notesSub.textContent = m.notes;
      notesSub.title = m.notes;
      nameTextWrap.appendChild(notesSub);
    }

    nameCell.appendChild(avatar);
    nameCell.appendChild(nameTextWrap);
    tdName.appendChild(nameCell);
    tr.appendChild(tdName);

    // Phone with Direct Call 📞 and WhatsApp 💬
    const tdPhone = document.createElement("td");
    const rawPhone = (m.phone || "").trim();
    if (rawPhone && rawPhone !== "-") {
      const clean = cleanDigits(rawPhone);
      const telNum = clean ? (clean.length === 10 ? `+91${clean}` : `+${clean}`) : rawPhone;
      const waNum = clean ? (clean.length === 10 ? `91${clean}` : clean) : rawPhone;

      const phoneWrap = document.createElement("div");
      phoneWrap.className = "phone-cell-wrap";

      const callA = document.createElement("a");
      callA.className = "phone-icon-btn btn-call";
      callA.href = `tel:${telNum}`;
      callA.title = `Call ${m.name} directly (${rawPhone})`;
      callA.innerHTML = `📞`;

      const waA = document.createElement("a");
      waA.className = "phone-icon-btn btn-wa";
      waA.href = `https://wa.me/${waNum}`;
      waA.target = "_blank";
      waA.title = `Chat with ${m.name} on WhatsApp (${rawPhone})`;
      waA.innerHTML = `💬`;

      const numSpan = document.createElement("span");
      numSpan.className = "phone-number-text";
      numSpan.textContent = rawPhone;

      phoneWrap.appendChild(callA);
      phoneWrap.appendChild(waA);
      phoneWrap.appendChild(numSpan);
      tdPhone.appendChild(phoneWrap);
    } else {
      tdPhone.innerHTML = `<span class="muted">-</span>`;
    }
    tr.appendChild(tdPhone);

    // Member Type
    const tdType = document.createElement("td");
    const typeBadge = document.createElement("span");
    const typeLower = (m.memberType || "").toLowerCase();
    if (typeLower.includes("winner")) {
      typeBadge.className = "badge-type badge-type-winner";
      typeBadge.title = m.memberType;
      typeBadge.innerHTML = `<span>👑</span> <span>Winner</span>`;
    } else if (typeLower.includes("vip")) {
      typeBadge.className = "badge-type badge-type-vip";
      typeBadge.title = m.memberType;
      typeBadge.innerHTML = `<span>⭐</span> <span>VIP</span>`;
    } else {
      typeBadge.className = "badge-type badge-type-direct";
      typeBadge.title = m.memberType;
      typeBadge.innerHTML = `<span>💎</span> <span>Direct</span>`;
    }
    tdType.appendChild(typeBadge);
    tr.appendChild(tdType);

    // Credit Balance
    const tdBalance = document.createElement("td");
    tdBalance.className = "centered";
    const creditPill = document.createElement("span");
    creditPill.className = "credit-balance-pill";
    creditPill.textContent = `⭐ ${m.credits.toLocaleString()}`;
    tdBalance.appendChild(creditPill);
    tr.appendChild(tdBalance);

    // Cash Worth
    const tdCash = document.createElement("td");
    tdCash.className = "centered";
    const cashPill = document.createElement("span");
    cashPill.className = "cash-worth-pill";
    cashPill.textContent = `₹${creditsToCash(m.credits).toFixed(2)}`;
    tdCash.appendChild(cashPill);
    tr.appendChild(tdCash);

    // Lifetime Earned
    const tdLifetime = document.createElement("td");
    tdLifetime.className = "centered";
    tdLifetime.innerHTML = `<span style="font-size:0.8rem; color:var(--text-secondary); white-space:nowrap;">+${m.totalEarned.toLocaleString()} pts</span>`;
    tr.appendChild(tdLifetime);

    // Joined Date
    const tdDate = document.createElement("td");
    const joinedDate = new Date(m.createdAtIso);
    const dateStr = joinedDate.toLocaleDateString(undefined, { day: "2-digit", month: "short", year: "2-digit" });
    tdDate.innerHTML = `<span class="muted" style="font-size:0.78rem; white-space:nowrap;" title="${joinedDate.toLocaleString()}">${dateStr}</span>`;
    tr.appendChild(tdDate);

    // Actions
    const tdActions = document.createElement("td");
    tdActions.className = "centered";

    const actWrap = document.createElement("div");
    actWrap.className = "member-actions-group";

    const addCreditBtn = document.createElement("button");
    addCreditBtn.type = "button";
    addCreditBtn.className = "action-btn-icon btn-add-pts";
    addCreditBtn.innerHTML = `+`;
    addCreditBtn.title = `Record Purchase & Add Credits for ${m.name}`;
    addCreditBtn.addEventListener("click", () => openCreditModal(m.id));

    const cashOutBtn = document.createElement("button");
    cashOutBtn.type = "button";
    cashOutBtn.className = "action-btn-icon btn-cash-out";
    cashOutBtn.innerHTML = `-`;
    cashOutBtn.title = "Deduct Credits (Purchase Discount or Cash Payout)";
    cashOutBtn.addEventListener("click", () => openDeductModal(m.id));

    const histBtn = document.createElement("button");
    histBtn.type = "button";
    histBtn.className = "action-btn-sm btn-view-hist";
    histBtn.innerHTML = `<span>📜</span>`;
    histBtn.title = "View Transaction History Ledger";
    histBtn.addEventListener("click", () => openMemberHistoryModal(m.id));

    const editBtn = document.createElement("button");
    editBtn.type = "button";
    editBtn.className = "btn-icon-del";
    editBtn.innerHTML = `✏️`;
    editBtn.title = "Edit Member Details";
    editBtn.addEventListener("click", () => openEditMemberModal(m.id));

    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "btn-icon-del";
    delBtn.innerHTML = `🗑️`;
    delBtn.title = "Delete Member";
    delBtn.addEventListener("click", () => deleteMember(m.id));

    actWrap.appendChild(addCreditBtn);
    actWrap.appendChild(cashOutBtn);
    actWrap.appendChild(histBtn);
    actWrap.appendChild(editBtn);
    actWrap.appendChild(delBtn);

    tdActions.appendChild(actWrap);
    tr.appendChild(tdActions);

    el.membersTableBody.appendChild(tr);
  });

  updateMemberBulkSelectionUI();
  renderMembersPaginationControls(totalCount, totalCount > 0 ? startIndex + 1 : 0, endIndex, totalPages);
}

function populateMemberSelectDropdowns() {
  const members = state.members || [];

  [el.creditMemberSelect, el.deductMemberSelect].forEach((select) => {
    if (!select) return;
    const currentVal = select.value;
    select.innerHTML = "";

    if (!members.length) {
      const opt = document.createElement("option");
      opt.value = "";
      opt.textContent = "-- No registered members --";
      select.appendChild(opt);
      return;
    }

    const defaultOpt = document.createElement("option");
    defaultOpt.value = "";
    defaultOpt.textContent = "-- Select Member --";
    select.appendChild(defaultOpt);

    members.forEach((m) => {
      const opt = document.createElement("option");
      opt.value = m.id;
      const cash = creditsToCash(m.credits);
      opt.textContent = `${m.name} (${m.phone || "No phone"}) · ${m.credits} pts (₹${cash.toFixed(2)})`;
      select.appendChild(opt);
    });

    if (currentVal && members.some((m) => m.id === currentVal)) {
      select.value = currentVal;
    }
  });
}

function checkCustomerMemberMatch() {
  if (!el.customerMemberBadge) return;
  const phone = safeText(el.customerNumber ? el.customerNumber.value : "");
  const name = safeText(el.customerName ? el.customerName.value : "");

  let match = null;
  if (phone) {
    match = findMemberByPhone(phone);
  }
  if (!match && name && name.length >= 3) {
    match = findMemberByName(name);
  }

  if (match) {
    const cashVal = creditsToCash(match.credits);
    if (el.customerMemberBadgeText) {
      el.customerMemberBadgeText.textContent = `👑 Member: ${match.name} | ${match.credits} Credits (≈ ₹${cashVal.toFixed(2)})`;
    }
    el.customerMemberBadge.classList.remove("hidden");
    if (el.customerMemberQuickCreditBtn) {
      el.customerMemberQuickCreditBtn.dataset.memberId = match.id;
    }
  } else {
    el.customerMemberBadge.classList.add("hidden");
  }
}

function onAddWinnerToMembers() {
  const record = getCurrentRecord();
  if (!record) return;

  const match = (record.customerNumber ? findMemberByPhone(record.customerNumber) : null) || findMemberByName(record.customerName);
  closeWinnerModal();
  setActiveTab("members");

  if (match) {
    openCreditModal(match.id, record.amount, `Won ${record.prize} (${record.purchasedItem || "Wheel"})`);
    showToast(`Existing member: ${match.name}. Record purchase credits!`, "info");
  } else {
    openAddMemberModal({
      name: record.customerName,
      phone: record.customerNumber,
      memberType: "Spin Winner",
      notes: `Won: ${record.prize} on ${record.purchasedItem || "Wheel"}`
    });
  }
}

/* ==========================================================================
   Members Modal Controls
   ========================================================================== */

function openAddMemberModal(prefill = {}) {
  if (!el.memberModal) return;
  if (el.memberModalTitle) el.memberModalTitle.textContent = "Add New Member";
  if (el.memberFormId) el.memberFormId.value = "";
  if (el.memberFormName) el.memberFormName.value = prefill.name || "";
  if (el.memberFormPhone) el.memberFormPhone.value = prefill.phone || "";
  if (el.memberFormType) el.memberFormType.value = prefill.memberType || "Direct Member";
  if (el.memberFormInitialCredits) el.memberFormInitialCredits.value = prefill.initialCredits || "";
  if (el.memberInitialCreditsGroup) el.memberInitialCreditsGroup.classList.remove("hidden");
  if (el.memberFormNotes) el.memberFormNotes.value = prefill.notes || "";

  updateInitialCreditsHint();
  el.memberModal.classList.remove("hidden");
  el.memberModal.setAttribute("aria-hidden", "false");
  if (el.memberFormName) el.memberFormName.focus();
}

function openEditMemberModal(id) {
  const member = state.members.find((m) => m.id === id);
  if (!member || !el.memberModal) return;

  if (el.memberModalTitle) el.memberModalTitle.textContent = `Edit Member: ${member.name}`;
  if (el.memberFormId) el.memberFormId.value = member.id;
  if (el.memberFormName) el.memberFormName.value = member.name;
  if (el.memberFormPhone) el.memberFormPhone.value = member.phone;
  if (el.memberFormType) el.memberFormType.value = member.memberType;
  if (el.memberInitialCreditsGroup) el.memberInitialCreditsGroup.classList.add("hidden");
  if (el.memberFormNotes) el.memberFormNotes.value = member.notes || "";

  el.memberModal.classList.remove("hidden");
  el.memberModal.setAttribute("aria-hidden", "false");
  if (el.memberFormName) el.memberFormName.focus();
}

function closeMemberModal() {
  if (!el.memberModal) return;
  el.memberModal.classList.add("hidden");
  el.memberModal.setAttribute("aria-hidden", "true");
}

function updateInitialCreditsHint() {
  if (!el.initialCreditsHint || !el.memberFormInitialCredits) return;
  const credits = Math.max(0, parseInt(el.memberFormInitialCredits.value, 10) || 0);
  const cash = creditsToCash(credits);
  el.initialCreditsHint.textContent = `${credits} Credits = ₹${cash.toFixed(2)} cash value`;
}

function openCreditModal(preselectedMemberId = null, purchaseAmount = null, note = "") {
  if (!el.creditModal) return;
  populateMemberSelectDropdowns();

  if (el.creditMemberSelect) {
    if (preselectedMemberId && state.members.some((m) => m.id === preselectedMemberId)) {
      el.creditMemberSelect.value = preselectedMemberId;
    } else if (state.members.length) {
      el.creditMemberSelect.value = state.members[0].id;
    }
  }

  if (el.creditPurchaseAmount) {
    el.creditPurchaseAmount.value = purchaseAmount ? String(num(purchaseAmount)) : "";
  }

  if (el.creditAmountInput) {
    // Default 100 credits
    el.creditAmountInput.value = "100";
  }

  if (el.creditNoteInput) {
    el.creditNoteInput.value = note || "";
  }

  updateCreditModalPreview();
  el.creditModal.classList.remove("hidden");
  el.creditModal.setAttribute("aria-hidden", "false");
  if (el.creditAmountInput) el.creditAmountInput.focus();
}

function closeCreditModal() {
  if (!el.creditModal) return;
  el.creditModal.classList.add("hidden");
  el.creditModal.setAttribute("aria-hidden", "true");
}

function updateCreditModalPreview() {
  if (!el.creditMemberSelect) return;
  const memberId = el.creditMemberSelect.value;
  const member = state.members.find((m) => m.id === memberId);

  if (member) {
    if (el.creditCurrentBalanceText) {
      el.creditCurrentBalanceText.textContent = `${member.credits} Credits (₹${creditsToCash(member.credits).toFixed(2)})`;
    }
  }

  const credits = Math.max(0, parseInt(el.creditAmountInput ? el.creditAmountInput.value : "0", 10) || 0);
  const cash = creditsToCash(credits);
  if (el.creditCashWorthPreview) {
    el.creditCashWorthPreview.textContent = `₹${cash.toFixed(2)}`;
  }
}

function openDeductModal(preselectedMemberId = null, purchaseBillAmount = null, note = "") {
  if (!el.deductModal) return;
  populateMemberSelectDropdowns();

  if (el.deductMemberSelect) {
    if (preselectedMemberId && state.members.some((m) => m.id === preselectedMemberId)) {
      el.deductMemberSelect.value = preselectedMemberId;
    } else if (state.members.length) {
      el.deductMemberSelect.value = state.members[0].id;
    }
  }

  if (el.deductBillAmountInput) {
    el.deductBillAmountInput.value = purchaseBillAmount ? String(num(purchaseBillAmount)) : "";
  }
  if (el.deductCreditsInput) el.deductCreditsInput.value = "100";
  if (el.deductCashInput) el.deductCashInput.value = (creditsToCash(100)).toFixed(2);
  if (el.deductNoteInput) el.deductNoteInput.value = note || "";

  setRedeemMode(purchaseBillAmount ? "purchase_discount" : "purchase_discount");

  el.deductModal.classList.remove("hidden");
  el.deductModal.setAttribute("aria-hidden", "false");
  if (el.deductCreditsInput) el.deductCreditsInput.focus();
}

function setRedeemMode(mode) {
  if (el.deductRedeemMode) el.deductRedeemMode.value = mode;
  const isDiscount = mode === "purchase_discount";

  if (el.modePurchaseDiscountBtn) el.modePurchaseDiscountBtn.classList.toggle("is-active", isDiscount);
  if (el.modeCashPayoutBtn) el.modeCashPayoutBtn.classList.toggle("is-active", !isDiscount);
  if (el.purchaseDiscountFields) el.purchaseDiscountFields.classList.toggle("hidden", !isDiscount);
  if (el.discountNetBanner) el.discountNetBanner.classList.toggle("hidden", !isDiscount);
  if (el.cashPayoutBanner) el.cashPayoutBanner.classList.toggle("hidden", isDiscount);

  if (el.deductValueLabel) {
    el.deductValueLabel.textContent = isDiscount ? "Discount Amount (₹)" : "Cash Payout (₹)";
  }
  if (el.confirmDeductBtn) {
    el.confirmDeductBtn.textContent = isDiscount ? "Confirm Purchase Discount" : "Confirm Cash Payout";
  }
  updateDeductModalPreview("credits");
}

function closeDeductModal() {
  if (!el.deductModal) return;
  el.deductModal.classList.add("hidden");
  el.deductModal.setAttribute("aria-hidden", "true");
}

function updateDeductModalPreview(source = "credits") {
  if (!el.deductMemberSelect) return;
  const memberId = el.deductMemberSelect.value;
  const member = state.members.find((m) => m.id === memberId);

  if (member && el.deductAvailableText) {
    const maxValue = creditsToCash(member.credits);
    el.deductAvailableText.textContent = `${member.credits} Credits (Max Value: ₹${maxValue.toFixed(2)})`;
  }

  const mode = el.deductRedeemMode ? el.deductRedeemMode.value : "purchase_discount";
  const isDiscount = mode === "purchase_discount";

  let credits = 0;
  let cash = 0;

  if (source === "credits") {
    credits = Math.max(0, parseInt(el.deductCreditsInput ? el.deductCreditsInput.value : "0", 10) || 0);
    cash = creditsToCash(credits);
    if (el.deductCashInput) el.deductCashInput.value = cash.toFixed(2);
  } else if (source === "cash") {
    cash = Math.max(0, parseFloat(el.deductCashInput ? el.deductCashInput.value : "0") || 0);
    credits = cashToCredits(cash);
    if (el.deductCreditsInput) el.deductCreditsInput.value = String(credits);
  }

  const billAmount = Math.max(0, parseFloat(el.deductBillAmountInput ? el.deductBillAmountInput.value : "0") || 0);
  const netPayable = Math.max(0, billAmount - cash);

  if (el.deductNetPayableDisplay) {
    el.deductNetPayableDisplay.textContent = `₹${netPayable.toFixed(2)}`;
  }
  if (el.deductBreakdownDisplay) {
    el.deductBreakdownDisplay.textContent = `Bill: ₹${billAmount.toFixed(2)} - Discount (${credits} pts): -₹${cash.toFixed(2)}`;
  }
  if (el.deductPayoutDisplay) {
    el.deductPayoutDisplay.textContent = `₹${cash.toFixed(2)}`;
  }
}

function openMemberHistoryModal(memberId) {
  const member = state.members.find((m) => m.id === memberId);
  if (!member || !el.memberHistoryModal) return;

  state.activeHistoryMemberId = memberId;
  if (el.historyMemberName) el.historyMemberName.textContent = `${member.name}'s Ledger`;
  if (el.historyMemberSub) {
    el.historyMemberSub.textContent = `Phone: ${member.phone || "None"} · Type: ${member.memberType} · Joined: ${new Date(member.createdAtIso).toLocaleDateString()}`;
  }

  if (el.historyCurrentCredits) el.historyCurrentCredits.textContent = `${member.credits} pts`;
  if (el.historyCurrentCash) el.historyCurrentCash.textContent = `₹${creditsToCash(member.credits).toFixed(2)}`;
  if (el.historyLifetimeEarned) el.historyLifetimeEarned.textContent = `+${member.totalEarned} pts`;
  if (el.historyLifetimeCashed) el.historyLifetimeCashed.textContent = `-₹${member.totalCashPaid.toFixed(2)}`;

  if (el.historyTableBody) {
    el.historyTableBody.innerHTML = "";
    const txs = member.transactions || [];

    if (!txs.length) {
      el.historyTableBody.innerHTML = `<tr><td colspan="6" class="centered muted" style="padding:24px;">No transactions recorded yet.</td></tr>`;
    } else {
      txs.forEach((tx) => {
        const tr = document.createElement("tr");
        const dt = new Date(tx.dateIso);
        const dateStr = `${dt.toLocaleDateString()} ${dt.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}`;

        const isCredit = tx.type === "CREDIT";
        const typeBadge = isCredit
          ? `<span class="badge-tx-credit">+ CREDIT</span>`
          : `<span class="badge-tx-debit">- DEBIT</span>`;

        const creditStr = isCredit ? `+${tx.credits}` : `-${tx.credits}`;
        const cashStr = isCredit ? `+₹${tx.cashValue.toFixed(2)}` : `-₹${tx.cashValue.toFixed(2)}`;

        tr.innerHTML = `
          <td><span class="muted">${dateStr}</span></td>
          <td>${typeBadge}</td>
          <td><strong>${creditStr}</strong></td>
          <td><span>${cashStr}</span></td>
          <td>${tx.note || "-"}</td>
          <td><strong class="text-gold">${tx.balanceAfter} pts</strong></td>
        `;
        el.historyTableBody.appendChild(tr);
      });
    }
  }

  el.memberHistoryModal.classList.remove("hidden");
  el.memberHistoryModal.setAttribute("aria-hidden", "false");
}

function closeHistoryModal() {
  if (!el.memberHistoryModal) return;
  el.memberHistoryModal.classList.add("hidden");
  el.memberHistoryModal.setAttribute("aria-hidden", "true");
  state.activeHistoryMemberId = null;
}

/* ==========================================================================
   Prize History Direct Credit Modal Controls
   ========================================================================== */

function openRecordCreditModal(recordId) {
  const record = getRecordById(recordId);
  if (!record || !el.recordCreditModal) return;

  state.activeRecordForCredit = record;
  const member = (record.customerNumber ? findMemberByPhone(record.customerNumber) : null) || findMemberByName(record.customerName);
  const phoneDisplay = formatDisplayPhone(record.customerNumber);

  if (el.recordCreditCustName) el.recordCreditCustName.textContent = record.customerName;
  if (el.recordCreditCustPhone) el.recordCreditCustPhone.textContent = phoneDisplay || "No phone number recorded";
  if (el.recordCreditRecId) el.recordCreditRecId.textContent = `#${record.recordId}`;
  if (el.recordCreditRecItem) el.recordCreditRecItem.textContent = record.purchasedItem || "Wheel Spin";
  if (el.recordCreditRecAmount) el.recordCreditRecAmount.textContent = `₹${formatAmount(record.amount)}`;
  if (el.recordCreditRecPrize) el.recordCreditRecPrize.textContent = record.prize;

  if (member) {
    if (el.recordCreditMemberStatusPill) {
      el.recordCreditMemberStatusPill.className = "status-pill status-member";
      el.recordCreditMemberStatusPill.textContent = "👑 Active Member";
    }
    if (el.recordCreditBalanceRow) el.recordCreditBalanceRow.classList.remove("hidden");
    if (el.recordCreditCurrentBal) el.recordCreditCurrentBal.textContent = `${member.credits} Credits`;
    if (el.recordCreditCashWorth) el.recordCreditCashWorth.textContent = `₹${creditsToCash(member.credits).toFixed(2)}`;
    if (el.recordCreditNotMemberNotice) el.recordCreditNotMemberNotice.classList.add("hidden");

    if (el.recordCreditAddTitle) el.recordCreditAddTitle.textContent = "Add Purchase Credits";
    if (el.recordCreditAddSub) el.recordCreditAddSub.textContent = `Award points for ₹${formatAmount(record.amount)} bill`;
    if (el.recordCreditDeductTitle) el.recordCreditDeductTitle.textContent = "Reduce / Redeem Credits";
    if (el.recordCreditDeductSub) el.recordCreditDeductSub.textContent = `Balance: ${member.credits} pts (Max ₹${creditsToCash(member.credits).toFixed(2)})`;
    if (el.recordCreditViewMemberBtn) el.recordCreditViewMemberBtn.classList.remove("hidden");
  } else {
    if (el.recordCreditMemberStatusPill) {
      el.recordCreditMemberStatusPill.className = "status-pill status-not-member";
      el.recordCreditMemberStatusPill.textContent = "✨ Not a Member";
    }
    if (el.recordCreditBalanceRow) el.recordCreditBalanceRow.classList.add("hidden");
    if (el.recordCreditNotMemberNotice) el.recordCreditNotMemberNotice.classList.remove("hidden");

    if (el.recordCreditAddTitle) el.recordCreditAddTitle.textContent = "Register Member & Add Credits";
    if (el.recordCreditAddSub) el.recordCreditAddSub.textContent = "Auto-enrolls customer + records purchase points";
    if (el.recordCreditDeductTitle) el.recordCreditDeductTitle.textContent = "Register Member & Reduce Credits";
    if (el.recordCreditDeductSub) el.recordCreditDeductSub.textContent = "Auto-enrolls customer with 0 starting points";
    if (el.recordCreditViewMemberBtn) el.recordCreditViewMemberBtn.classList.add("hidden");
  }

  el.recordCreditModal.classList.remove("hidden");
  el.recordCreditModal.setAttribute("aria-hidden", "false");
}

function closeRecordCreditModal() {
  if (!el.recordCreditModal) return;
  el.recordCreditModal.classList.add("hidden");
  el.recordCreditModal.setAttribute("aria-hidden", "true");
  state.activeRecordForCredit = null;
}

function onRecordCreditAddClick() {
  const record = state.activeRecordForCredit;
  if (!record) return;

  let member = (record.customerNumber ? findMemberByPhone(record.customerNumber) : null) || findMemberByName(record.customerName);
  if (!member) {
    // Customer not added yet -> Auto-register with data from record!
    member = addOrUpdateMember({
      name: record.customerName,
      phone: record.customerNumber,
      memberType: "Spin Winner",
      initialCredits: 0,
      notes: `Auto-registered from Prize Record #${record.recordId} (${record.purchasedItem || "Wheel"}, Prize: ${record.prize})`
    });
    if (!member) return;
    showToast(`Registered "${member.name}" in Members Club!`, "success");
  }

  closeRecordCreditModal();
  // Open Credit Modal with purchase amount prefilled!
  openCreditModal(
    member.id,
    record.amount,
    `Purchase: ${record.purchasedItem || "Wheel Spin"} (Record #${record.recordId})`
  );
}

function onRecordCreditDeductClick() {
  const record = state.activeRecordForCredit;
  if (!record) return;

  let member = (record.customerNumber ? findMemberByPhone(record.customerNumber) : null) || findMemberByName(record.customerName);
  if (!member) {
    // Customer not added yet -> Auto-register with data from record!
    member = addOrUpdateMember({
      name: record.customerName,
      phone: record.customerNumber,
      memberType: "Spin Winner",
      initialCredits: 0,
      notes: `Auto-registered from Prize Record #${record.recordId} (${record.purchasedItem || "Wheel"}, Prize: ${record.prize})`
    });
    if (!member) return;
    showToast(`Registered "${member.name}" in Members Club with 0 Credits.`, "info");
  }

  closeRecordCreditModal();
  // Open Deduct Modal with purchase amount prefilled as bill amount!
  openDeductModal(
    member.id,
    record.amount,
    `Redemption for Record #${record.recordId}`
  );
}

function onRecordCreditViewMemberClick() {
  const record = state.activeRecordForCredit;
  if (!record) return;
  const member = (record.customerNumber ? findMemberByPhone(record.customerNumber) : null) || findMemberByName(record.customerName);
  closeRecordCreditModal();
  if (member) {
    setActiveTab("members");
    if (el.memberSearchInput) {
      el.memberSearchInput.value = member.phone || member.name;
      el.memberSearchInput.dispatchEvent(new Event("input", { bubbles: true }));
    }
  }
}

function closeAllMemberModals() {
  closeMemberModal();
  closeCreditModal();
  closeDeductModal();
  closeHistoryModal();
  closeRecordCreditModal();
}

/* ==========================================================================
   Members Event Listeners
   ========================================================================== */

function bindMembersEvents() {
  // Section Header Action Buttons
  if (el.openNewMemberModalBtn) {
    el.openNewMemberModalBtn.addEventListener("click", () => openAddMemberModal());
  }
  if (el.emptyAddMemberBtn) {
    el.emptyAddMemberBtn.addEventListener("click", () => openAddMemberModal());
  }
  if (el.openAddCreditModalBtn) {
    el.openAddCreditModalBtn.addEventListener("click", () => openCreditModal());
  }
  if (el.openDeductModalBtn) {
    el.openDeductModalBtn.addEventListener("click", () => openDeductModal());
  }
  if (el.exportMembersBtn) {
    el.exportMembersBtn.addEventListener("click", exportMembersToCSV);
  }
  if (el.pushMembersBtn) {
    el.pushMembersBtn.addEventListener("click", () => void pushAllMembersToSheets(true));
  }
  if (el.refreshMembersBtn) {
    el.refreshMembersBtn.addEventListener("click", () => void loadMembersFromSheets(true));
  }

  // Toolbar Search & Filter
  if (el.memberSearchInput) {
    el.memberSearchInput.addEventListener("input", () => {
      state.memberSearchQuery = el.memberSearchInput.value;
      if (el.clearMemberSearchBtn) {
        el.clearMemberSearchBtn.classList.toggle("hidden", !state.memberSearchQuery);
      }
      state.membersPage = 1;
      renderMembersTable();
    });
  }

  if (el.clearMemberSearchBtn) {
    el.clearMemberSearchBtn.addEventListener("click", () => {
      state.memberSearchQuery = "";
      if (el.memberSearchInput) el.memberSearchInput.value = "";
      el.clearMemberSearchBtn.classList.add("hidden");
      state.membersPage = 1;
      renderMembersTable();
    });
  }

  if (el.memberFilterSelect) {
    el.memberFilterSelect.addEventListener("change", () => {
      state.memberFilter = el.memberFilterSelect.value;
      state.membersPage = 1;
      renderMembersTable();
    });
  }

  // Members Bulk Selection (like Prize History)
  if (el.membersTableBody) {
    el.membersTableBody.addEventListener("change", onMemberSelectionChange);
  }
  if (el.selectAllMembers) {
    el.selectAllMembers.addEventListener("change", onSelectAllMembersChange);
  }
  if (el.applyMemberBulkActionBtn) {
    el.applyMemberBulkActionBtn.addEventListener("click", () => void onApplyMemberBulkAction());
  }
  if (el.clearMemberSelectionBtn) {
    el.clearMemberSelectionBtn.addEventListener("click", onClearMemberSelection);
  }

  // Members Pagination Controls
  if (el.membersPageSizeSelect) {
    el.membersPageSizeSelect.addEventListener("change", () => {
      const size = parseInt(el.membersPageSizeSelect.value, 10);
      state.membersPageSize = size > 0 ? size : 10;
      state.membersPage = 1;
      renderMembersTable();
    });
  }

  if (el.membersFirstPageBtn) {
    el.membersFirstPageBtn.addEventListener("click", () => goToMembersPage(1));
  }
  if (el.membersPrevPageBtn) {
    el.membersPrevPageBtn.addEventListener("click", () => goToMembersPage(state.membersPage - 1));
  }
  if (el.membersNextPageBtn) {
    el.membersNextPageBtn.addEventListener("click", () => goToMembersPage(state.membersPage + 1));
  }
  if (el.membersLastPageBtn) {
    el.membersLastPageBtn.addEventListener("click", () => goToMembersPage(getLastMembersPage()));
  }

  // Member Modal
  if (el.memberForm) {
    el.memberForm.addEventListener("submit", (e) => {
      e.preventDefault();
      const id = el.memberFormId.value.trim();
      const name = el.memberFormName.value.trim();
      const phone = el.memberFormPhone.value.trim();
      const memberType = el.memberFormType.value;
      const initialCredits = el.memberFormInitialCredits ? el.memberFormInitialCredits.value : "0";
      const notes = el.memberFormNotes ? el.memberFormNotes.value.trim() : "";

      if (!name) {
        showToast("Customer name is required.", "error");
        return;
      }
      if (!phone) {
        showToast("Phone number is required.", "error");
        return;
      }

      const res = addOrUpdateMember({ id, name, phone, memberType, initialCredits, notes });
      if (res) {
        closeMemberModal();
      }
    });
  }

  if (el.memberFormInitialCredits) {
    el.memberFormInitialCredits.addEventListener("input", updateInitialCreditsHint);
  }

  if (el.memberModal) {
    el.memberModal.addEventListener("click", (e) => {
      if (e.target && ((e.target.closest && e.target.closest("[data-close-member-modal='true']")) || e.target.classList.contains("modal-backdrop") || e.target === el.memberModal)) {
        closeMemberModal();
      }
    });
  }

  // Credit Modal
  if (el.creditForm) {
    el.creditForm.addEventListener("submit", (e) => {
      e.preventDefault();
      const memberId = el.creditMemberSelect.value;
      if (!memberId) {
        showToast("Please select a member.", "error");
        return;
      }
      const credits = parseInt(el.creditAmountInput.value, 10);
      if (!credits || credits <= 0) {
        showToast("Please enter a valid credit amount.", "error");
        return;
      }
      const purchaseAmount = el.creditPurchaseAmount ? el.creditPurchaseAmount.value : 0;
      const note = el.creditNoteInput ? el.creditNoteInput.value.trim() : "";

      const ok = addMemberCredits(memberId, credits, purchaseAmount, note);
      if (ok) {
        closeCreditModal();
      }
    });
  }

  if (el.creditMemberSelect) {
    el.creditMemberSelect.addEventListener("change", updateCreditModalPreview);
  }

  if (el.creditAmountInput) {
    el.creditAmountInput.addEventListener("input", updateCreditModalPreview);
  }

  if (el.creditQuickAddMemberBtn) {
    el.creditQuickAddMemberBtn.addEventListener("click", () => {
      closeCreditModal();
      openAddMemberModal();
    });
  }

  if (el.creditModal) {
    el.creditModal.addEventListener("click", (e) => {
      if (e.target && ((e.target.closest && e.target.closest("[data-close-credit-modal='true']")) || e.target.classList.contains("modal-backdrop") || e.target === el.creditModal)) {
        closeCreditModal();
      }
    });
  }

  // Deduct Modal
  if (el.modePurchaseDiscountBtn) {
    el.modePurchaseDiscountBtn.addEventListener("click", () => setRedeemMode("purchase_discount"));
  }
  if (el.modeCashPayoutBtn) {
    el.modeCashPayoutBtn.addEventListener("click", () => setRedeemMode("cash_payout"));
  }
  if (el.deductBillAmountInput) {
    el.deductBillAmountInput.addEventListener("input", () => updateDeductModalPreview("credits"));
  }

  if (el.deductForm) {
    el.deductForm.addEventListener("submit", (e) => {
      e.preventDefault();
      const memberId = el.deductMemberSelect.value;
      if (!memberId) {
        showToast("Please select a member.", "error");
        return;
      }
      const credits = parseInt(el.deductCreditsInput.value, 10);
      if (!credits || credits <= 0) {
        showToast("Please enter credits to deduct.", "error");
        return;
      }
      const redeemMode = el.deductRedeemMode ? el.deductRedeemMode.value : "purchase_discount";
      const billAmount = el.deductBillAmountInput ? num(el.deductBillAmountInput.value) : 0;
      const note = el.deductNoteInput ? el.deductNoteInput.value.trim() : "";
      const sendReceipt = el.deductSendWhatsAppCheck ? el.deductSendWhatsAppCheck.checked : false;

      const res = deductMemberCredits(memberId, credits, note, sendReceipt, redeemMode, billAmount);
      if (res.success) {
        closeDeductModal();
      } else {
        showToast(res.error, "error");
      }
    });
  }

  if (el.deductMemberSelect) {
    el.deductMemberSelect.addEventListener("change", () => updateDeductModalPreview("credits"));
  }

  if (el.deductCreditsInput) {
    el.deductCreditsInput.addEventListener("input", () => updateDeductModalPreview("credits"));
  }

  if (el.deductCashInput) {
    el.deductCashInput.addEventListener("input", () => updateDeductModalPreview("cash"));
  }

  if (el.deductModal) {
    el.deductModal.addEventListener("click", (e) => {
      if (e.target && ((e.target.closest && e.target.closest("[data-close-deduct-modal='true']")) || e.target.classList.contains("modal-backdrop") || e.target === el.deductModal)) {
        closeDeductModal();
      }
    });
  }

  // History Modal
  if (el.memberHistoryModal) {
    el.memberHistoryModal.addEventListener("click", (e) => {
      if (e.target && ((e.target.closest && e.target.closest("[data-close-history-modal='true']")) || e.target.classList.contains("modal-backdrop") || e.target === el.memberHistoryModal)) {
        closeHistoryModal();
      }
    });
  }

  // Prize Record Direct Credit Modal Listeners
  if (el.recordCreditAddBtn) {
    el.recordCreditAddBtn.addEventListener("click", onRecordCreditAddClick);
  }
  if (el.recordCreditDeductBtn) {
    el.recordCreditDeductBtn.addEventListener("click", onRecordCreditDeductClick);
  }
  if (el.recordCreditViewMemberBtn) {
    el.recordCreditViewMemberBtn.addEventListener("click", onRecordCreditViewMemberClick);
  }
  if (el.recordCreditModal) {
    el.recordCreditModal.addEventListener("click", (e) => {
      if (e.target && ((e.target.closest && e.target.closest("[data-close-record-credit='true']")) || e.target.classList.contains("modal-backdrop") || e.target === el.recordCreditModal)) {
        closeRecordCreditModal();
      }
    });
  }
}


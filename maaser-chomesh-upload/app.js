const STORAGE_KEY = "maaser-chomesh-data-v2";
const PROFILES_KEY = "maaser-chomesh-import-profiles-v1";
const PROFILE_SETTINGS_KEY = "maaser-chomesh-import-profile-settings-v1";

/** @type {{ entries: Array<any>, version: string, date: string }} */
let state = {
  entries: [],
  version: "4.0",
  date: new Date().toISOString()
};

let activeTab = "all";
let activeTopTab = "main";
let reportChart = null;
let excelRows = [];
let excelWorkbook = null;
let excelFileName = "";
let noticeTimer = null;
const MAX_HISTORY = 100;
let undoStack = [];
let redoStack = [];
let manualSelectedRows = new Set();

const els = {
  form: document.getElementById("entry-form"),
  editingId: document.getElementById("editing-id"),
  saveBtn: document.getElementById("save-btn"),
  cancelEditBtn: document.getElementById("cancel-edit-btn"),
  type: document.getElementById("type"),
  date: document.getElementById("date"),
  hebrewDate: document.getElementById("hebrew-date"),
  description: document.getElementById("description"),
  amount: document.getElementById("amount"),
  recipient: document.getElementById("recipient"),
  recipientWrap: document.getElementById("recipient-wrap"),
  notes: document.getElementById("notes"),
  entriesBodyAll: document.getElementById("entries-body-all"),
  entriesBodyIncome: document.getElementById("entries-body-income"),
  entriesBodyDonation: document.getElementById("entries-body-donation"),
  tabPanels: {
    all: document.getElementById("panel-all"),
    income: document.getElementById("panel-income"),
    donation: document.getElementById("panel-donation")
  },
  rowTemplate: document.getElementById("row-template"),
  totalIncome: document.getElementById("total-income"),
  maaserTarget: document.getElementById("maaser-target"),
  chomeshTarget: document.getElementById("chomesh-target"),
  totalDonations: document.getElementById("total-donations"),
  remainingMaaser: document.getElementById("remaining-maaser"),
  remainingChomesh: document.getElementById("remaining-chomesh"),
  search: document.getElementById("search"),
  filterYear: document.getElementById("filter-year"),
  fromDate: document.getElementById("from-date"),
  toDate: document.getElementById("to-date"),
  tabBtns: Array.from(document.querySelectorAll(".tab-btn")),
  topTabMain: document.getElementById("top-tab-main"),
  topTabImport: document.getElementById("top-tab-import"),
  topPanelMain: document.getElementById("top-panel-main"),
  topPanelImport: document.getElementById("top-panel-import"),
  appNotice: document.getElementById("app-notice"),
  importSteps: Array.from(document.querySelectorAll("#import-steps .import-step")),
  undoBtn: document.getElementById("undo-btn"),
  redoBtn: document.getElementById("redo-btn"),
  exportBtn: document.getElementById("export-btn"),
  exportCsvBtn: document.getElementById("export-csv-btn"),
  exportXlsxBtn: document.getElementById("export-xlsx-btn"),
  importInput: document.getElementById("import-input"),
  clearBtn: document.getElementById("clear-btn"),
  reportYear: document.getElementById("report-year"),
  reportMode: document.getElementById("report-mode"),
  reportChart: document.getElementById("report-chart"),
  excelInput: document.getElementById("excel-input"),
  excelSheet: document.getElementById("excel-sheet"),
  excelType: document.getElementById("excel-type"),
  excelAmountMode: document.getElementById("excel-amount-mode"),
  excelHasHeader: document.getElementById("excel-has-header"),
  excelFixedDate: document.getElementById("excel-fixed-date"),
  excelMapper: document.getElementById("excel-mapper"),
  profileName: document.getElementById("profile-name"),
  saveProfileBtn: document.getElementById("save-profile-btn"),
  profileSelect: document.getElementById("profile-select"),
  loadProfileBtn: document.getElementById("load-profile-btn"),
  deleteProfileBtn: document.getElementById("delete-profile-btn"),
  setDefaultProfileBtn: document.getElementById("set-default-profile-btn"),
  clearDefaultProfileBtn: document.getElementById("clear-default-profile-btn"),
  autoProfileMode: document.getElementById("auto-profile-mode"),
  exportProfilesBtn: document.getElementById("export-profiles-btn"),
  importProfilesInput: document.getElementById("import-profiles-input"),
  mapDescription: document.getElementById("map-description"),
  mapAmount: document.getElementById("map-amount"),
  mapDate: document.getElementById("map-date"),
  mapNotes: document.getElementById("map-notes"),
  mapRecipient: document.getElementById("map-recipient"),
  excelRowMode: document.getElementById("excel-row-mode"),
  excelImportSearch: document.getElementById("excel-import-search"),
  excelStartRow: document.getElementById("excel-start-row"),
  selectAllRowsBtn: document.getElementById("select-all-rows-btn"),
  clearAllRowsBtn: document.getElementById("clear-all-rows-btn"),
  selectedRowsCounter: document.getElementById("selected-rows-counter"),
  autoMapBtn: document.getElementById("auto-map-btn"),
  excelRawPreview: document.getElementById("excel-raw-preview"),
  excelParsedPreview: document.getElementById("excel-parsed-preview"),
  excelLegacyPreview: document.getElementById("excel-legacy-preview"),
  importExcelBtn: document.getElementById("import-excel-btn"),
  quickImportBtn: document.getElementById("quick-import-btn")
};

/** @type {Record<string, any>} */
let importProfiles = {};
let profileSettings = {
  defaultProfile: "",
  autoProfileMode: "on"
};

function formatCurrency(num) {
  return new Intl.NumberFormat("he-IL", { style: "currency", currency: "ILS" }).format(num || 0);
}

function toNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function toIsoDate(value) {
  if (!value) return "";
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return "";
    const month = String(parsed.m).padStart(2, "0");
    const day = String(parsed.d).padStart(2, "0");
    return `${parsed.y}-${month}-${day}`;
  }
  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
  if (typeof value === "string") {
    const s = value.trim();
    const m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
    if (m) {
      const day = Number(m[1]);
      const month = Number(m[2]);
      let year = Number(m[3]);
      if (year < 100) year += year >= 70 ? 1900 : 2000;
      if (day >= 1 && day <= 31 && month >= 1 && month <= 12 && year >= 1900 && year <= 2200) {
        return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
      }
    }
  }
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return "";
  return d.toISOString().slice(0, 10);
}

function toHebrewLetters(num) {
  const ones = ["", "א", "ב", "ג", "ד", "ה", "ו", "ז", "ח", "ט"];
  const tens = ["", "י", "כ", "ל", "מ", "נ", "ס", "ע", "פ", "צ"];
  const hundreds = ["", "ק", "ר", "ש", "ת", "תק", "תר", "תש", "תת", "תתק"];

  let n = Number(num);
  if (!Number.isInteger(n) || n <= 0) return "";

  let result = "";
  const h = Math.floor(n / 100);
  if (h > 0) {
    result += hundreds[h] || "";
  }

  n = n % 100;
  if (n === 15) return result + "טו";
  if (n === 16) return result + "טז";

  const t = Math.floor(n / 10);
  const o = n % 10;
  result += tens[t] || "";
  result += ones[o] || "";

  return result;
}

function addGereshGershayim(hebrewText) {
  if (!hebrewText) return "";
  if (hebrewText.length === 1) return `${hebrewText}׳`;
  return `${hebrewText.slice(0, -1)}״${hebrewText.slice(-1)}`;
}

function toHebrewDate(gregorianDate) {
  if (!gregorianDate) return "";
  const parsed = new Date(gregorianDate);
  if (Number.isNaN(parsed.getTime())) return "";

  try {
    const parts = new Intl.DateTimeFormat("he-IL-u-ca-hebrew", {
      day: "numeric",
      month: "long",
      year: "numeric"
    }).formatToParts(parsed);

    const dayPart = parts.find((p) => p.type === "day");
    const monthPart = parts.find((p) => p.type === "month");
    const yearPart = parts.find((p) => p.type === "year");

    if (!dayPart || !monthPart || !yearPart) return "";

    const dayNum = Number(dayPart.value);
    const yearNum = Number(yearPart.value) % 1000;
    const dayHebrew = addGereshGershayim(toHebrewLetters(dayNum));
    const yearHebrew = addGereshGershayim(toHebrewLetters(yearNum));
    const monthHebrew = monthPart.value.trim();

    if (!dayHebrew || !monthHebrew || !yearHebrew) return "";
    return `${dayHebrew} ${monthHebrew} ${yearHebrew}`;
  } catch (e) {
    console.error("Hebrew date parsing error:", e);
    return "";
  }
}

function saveState() {
  state.date = new Date().toISOString();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function cloneEntries(entries) {
  return entries.map((e) => ({ ...e }));
}

function updateUndoRedoButtons() {
  if (els.undoBtn) els.undoBtn.disabled = undoStack.length === 0;
  if (els.redoBtn) els.redoBtn.disabled = redoStack.length === 0;
}

function showNotice(message, type = "info", timeoutMs = 3200) {
  if (!els.appNotice) return;
  els.appNotice.textContent = message;
  els.appNotice.classList.remove("hidden", "success", "error");
  if (type === "success") els.appNotice.classList.add("success");
  if (type === "error") els.appNotice.classList.add("error");

  if (noticeTimer) {
    clearTimeout(noticeTimer);
    noticeTimer = null;
  }
  if (timeoutMs > 0) {
    noticeTimer = setTimeout(() => {
      els.appNotice.classList.add("hidden");
    }, timeoutMs);
  }
}

function setImportStep(step) {
  if (!els.importSteps || !els.importSteps.length) return;
  const activeStep = Math.max(1, Math.min(4, Number(step) || 1));
  els.importSteps.forEach((el) => {
    const n = Number(el.dataset.step || 0);
    el.classList.toggle("active", n === activeStep);
  });
}

function pushHistorySnapshot() {
  undoStack.push(cloneEntries(state.entries));
  if (undoStack.length > MAX_HISTORY) {
    undoStack.shift();
  }
  redoStack = [];
  updateUndoRedoButtons();
}

function undoLastAction() {
  if (!undoStack.length) return;
  redoStack.push(cloneEntries(state.entries));
  state.entries = undoStack.pop();
  saveState();
  rerender();
  updateUndoRedoButtons();
}

function redoLastAction() {
  if (!redoStack.length) return;
  undoStack.push(cloneEntries(state.entries));
  state.entries = redoStack.pop();
  saveState();
  rerender();
  updateUndoRedoButtons();
}

function loadProfiles() {
  const raw = localStorage.getItem(PROFILES_KEY);
  if (!raw) return;
  try {
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === "object") {
      importProfiles = parsed;
    }
  } catch (_err) {
    importProfiles = {};
  }
}

function loadProfileSettings() {
  const raw = localStorage.getItem(PROFILE_SETTINGS_KEY);
  if (!raw) return;
  try {
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === "object") {
      profileSettings.defaultProfile = parsed.defaultProfile || "";
      profileSettings.autoProfileMode = parsed.autoProfileMode === "off" ? "off" : "on";
    }
  } catch (_err) {
    profileSettings = { defaultProfile: "", autoProfileMode: "on" };
  }
}

function saveProfileSettings() {
  localStorage.setItem(PROFILE_SETTINGS_KEY, JSON.stringify(profileSettings));
}

function saveProfiles() {
  localStorage.setItem(PROFILES_KEY, JSON.stringify(importProfiles));
}

function renderProfileOptions() {
  const names = Object.keys(importProfiles).sort((a, b) => a.localeCompare(b, "he"));
  const options = ['<option value="">בחר תבנית...</option>']
    .concat(names.map((n) => `<option value="${n}">${n}</option>`))
    .join("");
  els.profileSelect.innerHTML = options;
  els.autoProfileMode.value = profileSettings.autoProfileMode;
}

function updateDefaultProfileUiHint() {
  const current = profileSettings.defaultProfile;
  if (!current || !importProfiles[current]) return;
  if (!els.profileSelect.value) {
    els.profileSelect.value = current;
  }
}

function getCurrentMappingModel() {
  return {
    excelType: els.excelType.value,
    excelAmountMode: els.excelAmountMode.value,
    excelHasHeader: els.excelHasHeader.value,
    mapDescription: els.mapDescription.value,
    mapAmount: els.mapAmount.value,
    mapDate: els.mapDate.value,
    mapNotes: els.mapNotes.value,
    mapRecipient: els.mapRecipient.value
  };
}

function applyMappingModel(model) {
  if (!model) return;
  els.excelType.value = model.excelType || els.excelType.value;
  els.excelAmountMode.value = model.excelAmountMode || els.excelAmountMode.value;
  els.excelHasHeader.value = model.excelHasHeader || els.excelHasHeader.value;
  els.mapDescription.value = model.mapDescription ?? els.mapDescription.value;
  els.mapAmount.value = model.mapAmount ?? els.mapAmount.value;
  els.mapDate.value = model.mapDate ?? els.mapDate.value;
  els.mapNotes.value = model.mapNotes ?? els.mapNotes.value;
  els.mapRecipient.value = model.mapRecipient ?? els.mapRecipient.value;
  renderExcelPreview();
  renderParsedExcelPreview();
}

function saveCurrentProfile() {
  const name = (els.profileName.value || "").trim();
  if (!name) {
    alert("יש להזין שם תבנית");
    return;
  }
  importProfiles[name] = getCurrentMappingModel();
  saveProfiles();
  renderProfileOptions();
  els.profileSelect.value = name;
  updateDefaultProfileUiHint();
  alert("התבנית נשמרה");
}

function loadSelectedProfile() {
  const name = els.profileSelect.value;
  if (!name || !importProfiles[name]) {
    alert("לא נבחרה תבנית");
    return;
  }
  applyMappingModel(importProfiles[name]);
  els.profileName.value = name;
  alert("התבנית נטענה");
}

function deleteSelectedProfile() {
  const name = els.profileSelect.value;
  if (!name || !importProfiles[name]) {
    alert("לא נבחרה תבנית למחיקה");
    return;
  }
  delete importProfiles[name];
  if (profileSettings.defaultProfile === name) {
    profileSettings.defaultProfile = "";
    saveProfileSettings();
  }
  saveProfiles();
  renderProfileOptions();
  els.profileName.value = "";
}

function setDefaultProfile() {
  const name = els.profileSelect.value;
  if (!name || !importProfiles[name]) {
    alert("בחר תבנית לפני קביעה כברירת מחדל");
    return;
  }
  profileSettings.defaultProfile = name;
  saveProfileSettings();
  alert(`התבנית '${name}' נקבעה כברירת מחדל`);
}

function clearDefaultProfile() {
  profileSettings.defaultProfile = "";
  saveProfileSettings();
  alert("ברירת המחדל נוקתה");
}

function onAutoProfileModeChange() {
  profileSettings.autoProfileMode = els.autoProfileMode.value === "off" ? "off" : "on";
  saveProfileSettings();
}

function exportProfilesJson() {
  const payload = {
    version: "1.0",
    exportedAt: new Date().toISOString(),
    settings: profileSettings,
    profiles: importProfiles
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `import_profiles_${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

async function importProfilesJson(file) {
  const text = await file.text();
  const parsed = JSON.parse(text);
  if (!parsed || typeof parsed !== "object" || typeof parsed.profiles !== "object") {
    throw new Error("קובץ תבניות לא תקין");
  }

  importProfiles = { ...importProfiles, ...parsed.profiles };
  if (parsed.settings && typeof parsed.settings === "object") {
    profileSettings.defaultProfile = parsed.settings.defaultProfile || profileSettings.defaultProfile;
    profileSettings.autoProfileMode = parsed.settings.autoProfileMode === "off" ? "off" : profileSettings.autoProfileMode;
  }

  saveProfiles();
  saveProfileSettings();
  renderProfileOptions();
  updateDefaultProfileUiHint();
}

function findProfileByFileName(fileName) {
  const lower = (fileName || "").toLowerCase();
  const names = Object.keys(importProfiles);
  for (const name of names) {
    const tokens = name
      .toLowerCase()
      .split(/[\s\-_]+/)
      .filter((t) => t.length >= 3);
    if (tokens.some((t) => lower.includes(t))) {
      return name;
    }
  }
  return "";
}

function applyBestProfileForCurrentFile() {
  if (profileSettings.autoProfileMode === "off") return;

  let name = findProfileByFileName(excelFileName);
  if (!name && profileSettings.defaultProfile && importProfiles[profileSettings.defaultProfile]) {
    name = profileSettings.defaultProfile;
  }
  if (!name) return;

  applyMappingModel(importProfiles[name]);
  els.profileSelect.value = name;
  els.profileName.value = name;
}

function loadState() {
  const current = localStorage.getItem(STORAGE_KEY);
  const legacy = localStorage.getItem("maaser-chomesh-data-v1");
  const raw = current || legacy;
  if (!raw) return;

  try {
    const parsed = JSON.parse(raw);
    if (!parsed || !Array.isArray(parsed.entries)) return;

    state = {
      entries: parsed.entries.map((e) => ({
        id: e.id || `${Date.now()}-${Math.random()}`,
        type: e.type === "donation" ? "donation" : "income",
        date: toIsoDate(e.date) || new Date().toISOString().slice(0, 10),
        description: String(e.description || ""),
        amount: toNumber(e.amount),
        recipient: String(e.recipient || ""),
        notes: String(e.notes || ""),
        hebrewDate: String(e.hebrewDate || toHebrewDate(toIsoDate(e.date)))
      })),
      version: parsed.version || "4.0",
      date: parsed.date || new Date().toISOString()
    };

    saveState();
  } catch (_err) {
    console.warn("Failed to parse local state");
  }
}

function calcSummary(entries) {
  const income = entries.filter((e) => e.type === "income").reduce((sum, e) => sum + toNumber(e.amount), 0);
  const donations = entries
    .filter((e) => e.type === "donation")
    .reduce((sum, e) => sum + Math.max(0, toNumber(e.amount)), 0);
  const maaser = Math.max(0, income * 0.1);
  const chomesh = Math.max(0, income * 0.2);

  return {
    income,
    donations,
    maaser,
    chomesh,
    remainingMaaser: Math.max(0, maaser - donations),
    remainingChomesh: Math.max(0, chomesh - donations)
  };
}

function rowTypeLabel(type) {
  return type === "donation" ? "תרומה" : "הכנסה";
}

function getFilteredEntries() {
  const q = (els.search.value || "").trim().toLowerCase();
  const filterYear = (els.filterYear && els.filterYear.value) || "";
  const from = els.fromDate.value;
  const to = els.toDate.value;

  return state.entries.filter((entry) => {
    if (activeTab !== "all" && entry.type !== activeTab) return false;
    if (filterYear && !String(entry.date || "").startsWith(`${filterYear}-`)) return false;
    if (from && entry.date < from) return false;
    if (to && entry.date > to) return false;

    if (!q) return true;
    const haystack = [entry.description, entry.notes, entry.recipient].filter(Boolean).join(" ").toLowerCase();
    return haystack.includes(q);
  });
}

function renderFilterYearOptions() {
  if (!els.filterYear) return;
  const prev = els.filterYear.value;
  const years = Array.from(new Set(
    state.entries
      .map((e) => Number(String(e.date || "").slice(0, 4)))
      .filter((n) => Number.isFinite(n) && n > 1900)
  )).sort((a, b) => b - a);

  const options = ['<option value="">כל השנים</option>']
    .concat(years.map((y) => `<option value="${y}">${y}</option>`))
    .join("");
  els.filterYear.innerHTML = options;
  if (prev && years.includes(Number(prev))) {
    els.filterYear.value = prev;
  }
}

function renderSummary() {
  const s = calcSummary(state.entries);
  els.totalIncome.textContent = formatCurrency(s.income);
  els.maaserTarget.textContent = formatCurrency(s.maaser);
  els.chomeshTarget.textContent = formatCurrency(s.chomesh);
  els.totalDonations.textContent = formatCurrency(s.donations);
  els.remainingMaaser.textContent = formatCurrency(s.remainingMaaser);
  els.remainingChomesh.textContent = formatCurrency(s.remainingChomesh);
}

function renderTable() {
  const sortedAll = state.entries.slice().sort((a, b) => (a.date < b.date ? 1 : -1));
  const filteredAll = getFilteredEntries().slice().sort((a, b) => (a.date < b.date ? 1 : -1));
  const incomes = sortedAll.filter((x) => x.type === "income");
  const donations = sortedAll.filter((x) => x.type === "donation");

  fillTableBody(els.entriesBodyAll, filteredAll);
  fillTableBody(els.entriesBodyIncome, incomes);
  fillTableBody(els.entriesBodyDonation, donations);
}

function fillTableBody(bodyEl, list) {
  bodyEl.innerHTML = "";
  for (const item of list) {
    const frag = els.rowTemplate.content.cloneNode(true);
    const row = frag.querySelector("tr");
    row.querySelector('[data-k="type"]').textContent = rowTypeLabel(item.type);
    row.querySelector('[data-k="date"]').textContent = item.date;
    row.querySelector('[data-k="hebrewDate"]').textContent = toHebrewDate(item.date) || item.hebrewDate || "-";
    row.querySelector('[data-k="description"]').textContent = item.description || "";
    row.querySelector('[data-k="amount"]').textContent = formatCurrency(toNumber(item.amount));
    row.querySelector('[data-k="recipient"]').textContent = item.recipient || "-";
    row.querySelector('[data-k="notes"]').textContent = item.notes || "-";
    row.querySelector('[data-action="edit"]').dataset.id = String(item.id);
    row.querySelector('[data-action="delete"]').dataset.id = String(item.id);
    bodyEl.appendChild(frag);
  }
}

function renderTabPanels() {
  Object.entries(els.tabPanels).forEach(([name, panel]) => {
    panel.classList.toggle("hidden", name !== activeTab);
  });
}

function renderTopPanels() {
  if (els.topPanelMain) {
    els.topPanelMain.classList.toggle("hidden", activeTopTab !== "main");
  }
  if (els.topPanelImport) {
    els.topPanelImport.classList.toggle("hidden", activeTopTab !== "import");
  }
  if (els.topTabMain) {
    els.topTabMain.classList.toggle("active", activeTopTab === "main");
  }
  if (els.topTabImport) {
    els.topTabImport.classList.toggle("active", activeTopTab === "import");
  }
  if (activeTopTab === "import" && !excelWorkbook) {
    setImportStep(1);
  }
}

function resetFormToCreateMode() {
  els.editingId.value = "";
  els.saveBtn.textContent = "שמור פעולה";
  els.cancelEditBtn.classList.add("hidden");
  els.form.reset();
  els.date.value = new Date().toISOString().slice(0, 10);
  updateHebrewDatePreview();
  els.type.value = "income";
  toggleRecipient();
}

function enterEditMode(id) {
  const item = state.entries.find((x) => String(x.id) === String(id));
  if (!item) return;

  els.editingId.value = String(item.id);
  els.type.value = item.type;
  els.date.value = item.date;
  els.hebrewDate.value = toHebrewDate(item.date) || item.hebrewDate || "";
  els.description.value = item.description;
  els.amount.value = String(item.amount);
  els.recipient.value = item.recipient || "";
  els.notes.value = item.notes || "";
  toggleRecipient();
  els.saveBtn.textContent = "עדכן פעולה";
  els.cancelEditBtn.classList.remove("hidden");
}

function upsertEntry(entry) {
  const idx = state.entries.findIndex((x) => String(x.id) === String(entry.id));
  if (idx === -1) {
    state.entries.push(entry);
  } else {
    state.entries[idx] = entry;
  }
}

function onSubmit(e) {
  e.preventDefault();
  const type = els.type.value;
  const amount = toNumber(els.amount.value);

  if (!els.date.value || !els.description.value.trim()) return;
  if (type === "donation" && amount <= 0) {
    alert("בתרומה יש להזין סכום חיובי גדול מאפס");
    return;
  }
  if (type === "income" && amount === 0) {
    alert("בהכנסה יש להזין סכום שונה מאפס (אפשר גם שלילי)");
    return;
  }

  const existingId = els.editingId.value;
  const computedHebrewDate = toHebrewDate(els.date.value);
  const entry = {
    id: existingId || `${Date.now()}-${Math.random()}`,
    type,
    date: els.date.value,
    description: els.description.value.trim(),
    amount,
    recipient: type === "donation" ? (els.recipient.value || "").trim() : "",
    notes: (els.notes.value || "").trim(),
    hebrewDate: computedHebrewDate
  };

  pushHistorySnapshot();
  upsertEntry(entry);
  saveState();
  resetFormToCreateMode();
  rerender();
}

function toggleRecipient() {
  const donation = els.type.value === "donation";
  els.recipientWrap.classList.toggle("hidden", !donation);
  els.recipient.required = false;
}

function onRowActions(e) {
  const target = e.target;
  if (!(target instanceof HTMLElement)) return;
  const action = target.dataset.action;
  const id = target.dataset.id;
  if (!action || !id) return;

  if (action === "delete") {
    pushHistorySnapshot();
    state.entries = state.entries.filter((x) => String(x.id) !== String(id));
    saveState();
    rerender();
    return;
  }

  if (action === "edit") {
    enterEditMode(id);
  }
}

function exportBackup() {
  const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `backup_${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

function exportBackupBeforeClear() {
  const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const stamp = new Date().toISOString().replace(/[:]/g, "-").slice(0, 19);
  a.download = `backup_before_clear_${stamp}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

function exportCsv() {
  const header = ["type", "date", "hebrewDate", "description", "amount", "recipient", "notes"];
  const lines = [header.join(",")];

  for (const e of state.entries) {
    const row = [e.type, e.date, toHebrewDate(e.date) || e.hebrewDate || "", e.description, String(e.amount), e.recipient || "", e.notes || ""].map((v) => {
      const value = String(v).replaceAll('"', '""');
      return `"${value}"`;
    });
    lines.push(row.join(","));
  }

  const blob = new Blob(["\ufeff" + lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `entries_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

function exportXlsx() {
  const rows = state.entries.map((e) => ({
    type: e.type,
    date: e.date,
    hebrewDate: toHebrewDate(e.date) || e.hebrewDate || "",
    description: e.description,
    amount: e.amount,
    recipient: e.recipient || "",
    notes: e.notes || ""
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "entries");
  XLSX.writeFile(wb, `entries_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function normalizeBackup(raw) {
  if (!raw || !Array.isArray(raw.entries)) {
    throw new Error("קובץ גיבוי לא תקין");
  }

  return {
    entries: raw.entries.map((e) => ({
      id: e.id || `${Date.now()}-${Math.random()}`,
      type: e.type === "donation" ? "donation" : "income",
      date: toIsoDate(e.date) || new Date().toISOString().slice(0, 10),
      description: String(e.description || ""),
      amount: toNumber(e.amount),
      recipient: String(e.recipient || ""),
      notes: String(e.notes || ""),
      hebrewDate: String(e.hebrewDate || toHebrewDate(toIsoDate(e.date)))
    })),
    version: raw.version || "4.0",
    date: raw.date || new Date().toISOString()
  };
}

async function importBackup(file) {
  const text = await file.text();
  const parsed = JSON.parse(text);
  const normalized = normalizeBackup(parsed);

  const byId = new Map(state.entries.map((e) => [String(e.id), e]));
  pushHistorySnapshot();
  for (const item of normalized.entries) {
    byId.set(String(item.id), item);
  }

  state.entries = Array.from(byId.values());
  state.version = normalized.version;
  state.date = new Date().toISOString();
  saveState();
  rerender();
}

function renderReportYearOptions() {
  const yearSet = new Set(state.entries.map((e) => Number((e.date || "").slice(0, 4))).filter((n) => Number.isFinite(n)));
  if (yearSet.size === 0) yearSet.add(new Date().getFullYear());
  const years = Array.from(yearSet).sort((a, b) => b - a);

  const current = els.reportYear.value;
  els.reportYear.innerHTML = years.map((y) => `<option value="${y}">${y}</option>`).join("");
  els.reportYear.value = current && years.includes(Number(current)) ? current : String(years[0]);
}

function getReportData() {
  const mode = els.reportMode.value;
  const year = Number(els.reportYear.value);

  if (mode === "yearly") {
    const byYear = new Map();
    for (const e of state.entries) {
      const y = Number((e.date || "").slice(0, 4));
      if (!Number.isFinite(y)) continue;
      const bucket = byYear.get(y) || { income: 0, donation: 0 };
      if (e.type === "income") bucket.income += toNumber(e.amount);
      if (e.type === "donation") bucket.donation += Math.max(0, toNumber(e.amount));
      byYear.set(y, bucket);
    }
    const labels = Array.from(byYear.keys()).sort((a, b) => a - b).map(String);
    const income = labels.map((y) => byYear.get(Number(y)).income);
    const donation = labels.map((y) => byYear.get(Number(y)).donation);
    return { labels, income, donation };
  }

  const labels = ["ינו", "פבר", "מרץ", "אפר", "מאי", "יונ", "יול", "אוג", "ספט", "אוק", "נוב", "דצמ"];
  const income = new Array(12).fill(0);
  const donation = new Array(12).fill(0);

  for (const e of state.entries) {
    const d = e.date || "";
    const y = Number(d.slice(0, 4));
    const m = Number(d.slice(5, 7)) - 1;
    if (y !== year || m < 0 || m > 11) continue;
    if (e.type === "income") income[m] += toNumber(e.amount);
    if (e.type === "donation") donation[m] += Math.max(0, toNumber(e.amount));
  }

  return { labels, income, donation };
}

function renderReportChart() {
  const data = getReportData();
  if (reportChart) reportChart.destroy();

  reportChart = new Chart(els.reportChart, {
    type: "bar",
    data: {
      labels: data.labels,
      datasets: [
        { label: "הכנסות", data: data.income, backgroundColor: "#1a73e8" },
        { label: "תרומות", data: data.donation, backgroundColor: "#188038" }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      aspectRatio: 3.8,
      layout: {
        padding: { top: 4, right: 8, bottom: 4, left: 8 }
      },
      plugins: { legend: { position: "top", rtl: true } },
      scales: { y: { beginAtZero: true } }
    }
  });
}

function detectColumnsByHeaders(headers) {
  // Patterns for common bank/financial Excel column names
  const patterns = {
    date: ['תאריך שעה', 'תאריך', 'date', 'יום', 'תאריך עסקה', 'ממועד', 'ממועד ערך', 'ערך'],
    description: ['תיאור הרשומה', 'תיאור', 'description', 'פרטים', 'פירוט', 'מהות הפעולה', 'שם פעולה', 'שם בית עסק'],
    amount: ['סכום', 'amount', 'זכות', 'חובה', 'חיוב', 'זיכוי', 'יתרה פעולה', 'סכום פעולה', 'value'],
    notes: ['הערות', 'notes', 'הוספות', 'הערה', 'אסמכתא', 'פרטי עסקה'],
    recipient: ['שם התאגיד', 'מקבל', 'recipient', 'beneficiary', 'למי', 'שם', 'בעל החשבון', 'מוטב']
  };
  
  const result = {
    date: null,
    description: null,
    amount: null,
    notes: null,
    recipient: null
  };
  
  // Iterate through headers and match patterns
  headers.forEach((header, index) => {
    if (!header) return;
    const normalized = String(header).trim().toLowerCase();
    
    for (const [field, keywords] of Object.entries(patterns)) {
      if (result[field] !== null) continue; // Already found
      
      for (const keyword of keywords) {
        if (normalized.includes(keyword.toLowerCase())) {
          result[field] = index;
          break;
        }
      }
    }
  });
  
  return result;
}

function detectBestHeaderRow(maxScanRows = 30) {
  if (!excelRows.length) return 0;
  const scanLimit = Math.min(maxScanRows, excelRows.length);

  let bestRow = 0;
  let bestScore = -1;

  for (let i = 0; i < scanLimit; i += 1) {
    const row = excelRows[i] || [];
    const headers = row.map((h, idx) => (h == null || h === "" ? `טור ${idx + 1}` : String(h)));
    const detected = detectColumnsByHeaders(headers);
    let score = 0;
    if (detected.date !== null) score += 2;
    if (detected.amount !== null) score += 3;
    if (detected.description !== null) score += 3;
    if (detected.notes !== null) score += 1;
    if (detected.recipient !== null) score += 1;

    if (score > bestScore) {
      bestScore = score;
      bestRow = i;
    }
  }

  return bestScore >= 5 ? bestRow : 0;
}

function autoMapColumns() {
  if (!excelRows.length) return;
  
  const hasHeader = els.excelHasHeader.value === "yes";
  if (!hasHeader) return; // Can't auto-detect without headers
  
  const startRow = Math.max(0, Number(els.excelStartRow.value) - 1) || 0;
  const headerRow = excelRows[startRow] || [];
  const headers = headerRow.map((h, i) => h || `טור ${i + 1}`);
  
  const detected = detectColumnsByHeaders(headers);
  
  // Apply detected mappings
  if (detected.description !== null) els.mapDescription.value = detected.description;
  if (detected.amount !== null) els.mapAmount.value = detected.amount;
  if (detected.date !== null) els.mapDate.value = detected.date;
  if (detected.notes !== null) els.mapNotes.value = detected.notes;
  if (detected.recipient !== null) els.mapRecipient.value = detected.recipient;
  
  renderParsedExcelPreview();
}

function getAmountByMode(rawAmount, type, mode) {
  const amount = toNumber(rawAmount);
  if (mode === "as-is") return amount;
  if (mode === "abs") return Math.abs(amount);
  if (mode === "flip") return amount * -1;

  // auto mode: bank files are often +/- mixed. Keep income as-is,
  // and make donations positive to simplify credit-card donation imports.
  if (type === "donation") return Math.abs(amount);
  return amount;
}

function renderExcelPreview() {
  if (!excelRows.length) {
    els.excelRawPreview.innerHTML = "";
    els.excelRawPreview.style.display = "none";
    return;
  }

  const hasHeader = els.excelHasHeader.value === "yes";
  const startRow = Math.max(0, Number(els.excelStartRow.value) - 1) || 0;
  const headerRow = hasHeader ? excelRows[startRow] : excelRows[startRow].map((_x, i) => `טור ${i + 1}`);
  const headers = headerRow.map((h, i) => h || `טור ${i + 1}`);
  const bodyStart = hasHeader ? startRow + 1 : startRow;
  const bodyRows = excelRows.slice(bodyStart);
  const q = (els.excelImportSearch.value || "").trim().toLowerCase();
  const indexedRows = bodyRows.map((row, idx) => ({ row, index1: idx + 1 }));
  const filteredRows = q
    ? indexedRows.filter(({ row }) => row.some((c) => String(c == null ? "" : c).toLowerCase().includes(q)))
    : indexedRows;

  const showCheckboxes = els.excelRowMode.value === "selected";
  
  const rowsHtml = filteredRows
    .map(({ row, index1 }) => {
      const isChecked = manualSelectedRows.has(index1);
      const checkbox = showCheckboxes ? `<td><input class="row-check" type="checkbox" data-row="${index1}" ${isChecked ? "checked" : ""} /></td>` : "";
      const tds = row.map((c) => `<td>${c == null ? "" : String(c)}</td>`).join("");
      return `<tr>${checkbox}<td style="text-align:center;color:#999;">${startRow + 1 + index1}</td>${tds}</tr>`;
    })
    .slice(0, 100)
    .join("");

  const visibleSelectedCount = filteredRows.filter(({ index1 }) => manualSelectedRows.has(index1)).length;
  const allVisibleSelected = filteredRows.length > 0 && visibleSelectedCount === filteredRows.length;
  const headerCheckbox = showCheckboxes
    ? `<th><input id="row-check-all" type="checkbox" ${allVisibleSelected ? "checked" : ""} title="בחר/בטל הכל" /></th>`
    : "";
  els.excelRawPreview.innerHTML = `
    <h4>תצוגה גולמית (${q ? "מסונן" : "כללי"}, עד 100 שורות)</h4>
    <table style="font-size:0.85rem;">
      <thead><tr>${headerCheckbox}<th>#</th>${headers.map((h) => `<th>${h}</th>`).join("")}</tr></thead>
      <tbody>${rowsHtml}</tbody>
    </table>
  `;
  els.excelRawPreview.style.display = "block";
  
  // Re-run parsed preview when checkboxes change
  if (showCheckboxes) {
    const checks = Array.from(els.excelRawPreview.querySelectorAll(".row-check"));
    checks.forEach((check) => {
      check.addEventListener("change", () => {
        const rowNum = Number(check.dataset.row);
        if (check.checked) manualSelectedRows.add(rowNum);
        else manualSelectedRows.delete(rowNum);
        renderExcelPreview();
        renderParsedExcelPreview();
      });
    });

    const master = els.excelRawPreview.querySelector("#row-check-all");
    if (master) {
      master.addEventListener("change", () => {
        filteredRows.forEach(({ index1 }) => {
          if (master.checked) manualSelectedRows.add(index1);
          else manualSelectedRows.delete(index1);
        });
        renderExcelPreview();
        renderParsedExcelPreview();
      });
    }
  }
}

function getExcelHeaders() {
  if (!excelRows.length) return [];
  const startRow = Math.max(0, Number(els.excelStartRow.value) - 1) || 0;
  const hasHeader = els.excelHasHeader.value === "yes";
  const firstRow = excelRows[startRow] || [];
  if (hasHeader) return firstRow.map((h, i) => h || `טור ${i + 1}`);
  return firstRow.map((_h, i) => `טור ${i + 1}`);
}

function updateMappingOptionsFromSelectedRow() {
  if (!excelRows.length) return;
  const startRow = Math.max(0, Number(els.excelStartRow.value) - 1) || 0;
  const selectedRow = excelRows[startRow] || [];

  const options = ['<option value="">לא נבחר</option>']
    .concat(selectedRow.map((cell, i) => {
      const text = String(cell == null ? "" : cell).trim();
      const label = text || `(ריק) טור ${i + 1}`;
      return `<option value="${i}">${label}</option>`;
    }))
    .join("");

  [els.mapDescription, els.mapAmount, els.mapDate, els.mapNotes, els.mapRecipient].forEach((sel) => {
    const prev = sel.value;
    sel.innerHTML = options;
    if (prev !== "" && selectedRow[Number(prev)] !== undefined) {
      sel.value = prev;
    }
  });
}

async function onExcelFileChosen(file) {
  excelFileName = file.name || "";
  const buf = await file.arrayBuffer();
  excelWorkbook = XLSX.read(buf, { type: "array" });
  els.excelSheet.innerHTML = excelWorkbook.SheetNames.map((name) => `<option value="${name}">${name}</option>`).join("");
  setImportStep(2);
  loadSelectedSheetRows();
  applyBestProfileForCurrentFile();
  showNotice(`הקובץ נטען: ${excelFileName}`, "success");
}

function loadSelectedSheetRows() {
  if (!excelWorkbook) return;
  const name = els.excelSheet.value || excelWorkbook.SheetNames[0];
  const ws = excelWorkbook.Sheets[name];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
  if (!rows.length) throw new Error("empty");

  excelRows = rows;

  const bestHeaderRow = detectBestHeaderRow(40);
  els.excelHasHeader.value = "yes";
  els.excelStartRow.value = String(bestHeaderRow + 1);
  manualSelectedRows = new Set(excelRows.slice(bestHeaderRow + 1).map((_row, idx) => idx + 1));

  updateMappingOptionsFromSelectedRow();
  autoMapColumns();
  renderExcelPreview();
  renderParsedExcelPreview();
  updateSelectedRowsCounter();
  els.excelMapper.classList.remove("hidden");
  setImportStep(3);
}

function getCurrentFilteredBodyIndexSet() {
  const hasHeader = els.excelHasHeader.value === "yes";
  const startRow = Math.max(0, Number(els.excelStartRow.value) - 1) || 0;
  const bodyStart = hasHeader ? startRow + 1 : startRow;
  const bodyRows = excelRows.slice(bodyStart);
  const q = (els.excelImportSearch.value || "").trim().toLowerCase();
  const indexedRows = bodyRows.map((row, idx) => ({ row, index1: idx + 1 }));
  const filteredRows = q
    ? indexedRows.filter(({ row }) => row.some((c) => String(c == null ? "" : c).toLowerCase().includes(q)))
    : indexedRows;
  return new Set(filteredRows.map((x) => x.index1));
}

function collectRowsForImport() {
  const mode = els.excelRowMode.value;
  const hasHeader = els.excelHasHeader.value === "yes";
  const startRow = Math.max(0, Number(els.excelStartRow.value) - 1) || 0;
  const bodyStart = hasHeader ? startRow + 1 : startRow;
  const allRows = excelRows.slice(bodyStart);

  const visibleSet = getCurrentFilteredBodyIndexSet();
  if (mode === "all") {
    return allRows.filter((_row, idx) => visibleSet.has(idx + 1));
  }

  return allRows.filter((_row, idx) => {
    const index1 = idx + 1;
    return visibleSet.has(index1) && manualSelectedRows.has(index1);
  });
}

function setAllRowChecks(checked) {
  const visibleSet = getCurrentFilteredBodyIndexSet();
  visibleSet.forEach((idx) => {
    if (checked) manualSelectedRows.add(idx);
    else manualSelectedRows.delete(idx);
  });
  renderExcelPreview();
  renderParsedExcelPreview();
}

function updateSelectedRowsCounter() {
  if (!els.selectedRowsCounter) return;
  const hasHeader = els.excelHasHeader.value === "yes";
  const startRow = Math.max(0, Number(els.excelStartRow.value) - 1) || 0;
  const bodyStart = hasHeader ? startRow + 1 : startRow;
  const totalRows = Math.max(0, excelRows.slice(bodyStart).length);
  const visibleSet = getCurrentFilteredBodyIndexSet();

  let selectedCount;
  if (els.excelRowMode.value === "all") {
    selectedCount = visibleSet.size;
  } else {
    selectedCount = Array.from(visibleSet).filter((idx) => manualSelectedRows.has(idx)).length;
  }

  els.selectedRowsCounter.textContent = `נבחרו ${selectedCount} מתוך ${totalRows} שורות`;
}

function parseImportedEntries() {
  const rows = collectRowsForImport();
  const type = els.excelType.value;
  const fixedDate = els.excelFixedDate.value;
  const amountMode = els.excelAmountMode.value;
  const idx = {
    description: Number(els.mapDescription.value),
    amount: Number(els.mapAmount.value),
    date: els.mapDate.value === "" ? null : Number(els.mapDate.value),
    notes: els.mapNotes.value === "" ? null : Number(els.mapNotes.value),
    recipient: els.mapRecipient.value === "" ? null : Number(els.mapRecipient.value)
  };

  if (!Number.isFinite(idx.description) || !Number.isFinite(idx.amount)) {
    throw new Error("יש לבחור לפחות טור תיאור וטור סכום");
  }

  const imported = [];
  for (const row of rows) {
    const amount = getAmountByMode(row[idx.amount], type, amountMode);
    const description = String(row[idx.description] || "").trim();
    if (!description) continue;

    if (type === "donation" && amount <= 0) continue;
    if (type === "income" && amount === 0) continue;

    const resolvedDate = fixedDate || toIsoDate(idx.date == null ? "" : row[idx.date]) || new Date().toISOString().slice(0, 10);

    imported.push({
      id: `${Date.now()}-${Math.random()}`,
      type,
      date: resolvedDate,
      description,
      amount,
      recipient: type === "donation" ? String(idx.recipient == null ? "" : row[idx.recipient] || "").trim() : "",
      notes: String(idx.notes == null ? "" : row[idx.notes] || "").trim(),
      hebrewDate: toHebrewDate(resolvedDate)
    });
  }

  return imported;
}

function renderParsedExcelPreview() {
  if (!excelRows.length) {
    els.excelParsedPreview.innerHTML = "";
    return;
  }

  try {
    const imported = parseImportedEntries();
    if (!imported.length) {
      els.excelParsedPreview.innerHTML = "<p style='color:#d93025;'>לא נמצאו שורות מתאימות לייבוא (בדוק את הבחירה והמיפוי)</p>";
      return;
    }
    setImportStep(4);

    const rowsHtml = imported.slice(0, 50)
      .map((entry) => {
        return `<tr>
          <td>${entry.type === "donation" ? "תרומה" : "הכנסה"}</td>
          <td>${entry.date}</td>
          <td>${entry.hebrewDate}</td>
          <td>${entry.description}</td>
          <td>${formatCurrency(entry.amount)}</td>
          <td>${entry.recipient || "-"}</td>
          <td>${entry.notes || "-"}</td>
        </tr>`;
      })
      .join("");

    els.excelParsedPreview.innerHTML = `
      <h4>תצוגה מקדימה של הנתונים שיובאו (${imported.length} שורות בסה"כ)</h4>
      <table style="font-size:0.85rem;">
        <thead><tr>
          <th>סוג</th>
          <th>תאריך</th>
          <th>תאריך עברי</th>
          <th>תיאור</th>
          <th>סכום</th>
          <th>מקבל</th>
          <th>הערות</th>
        </tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    `;
  } catch (err) {
    els.excelParsedPreview.innerHTML = `<p style='color:#d93025;'>שגיאה בעיבוד: ${err.message}</p>`;
  }
}

function updateHebrewDatePreview() {
  els.hebrewDate.value = toHebrewDate(els.date.value);
}

function onImportExcel() {
  try {
    const imported = parseImportedEntries();
    if (!imported.length) {
      showNotice("לא נמצאו שורות מתאימות לייבוא", "error");
      return;
    }

    pushHistorySnapshot();
    for (const entry of imported) {
      state.entries.push(entry);
    }
    saveState();
    rerender();
    showNotice(`יובאו ${imported.length} שורות בהצלחה`, "success");
  } catch (err) {
    showNotice(`ייבוא אקסל נכשל: ${err.message || "שגיאה לא ידועה"}`, "error", 5200);
  }
}

function onQuickImport() {
  if (!excelRows.length) {
    showNotice("בחר קובץ אקסל לפני ייבוא מהיר", "error");
    return;
  }

  const fallback = profileSettings.defaultProfile;
  const detected = profileSettings.autoProfileMode === "on" ? findProfileByFileName(excelFileName) : "";
  const selected = detected || fallback;

  if (!selected || !importProfiles[selected]) {
    showNotice("אין תבנית מזוהה או ברירת מחדל לייבוא מהיר", "error");
    return;
  }

  applyMappingModel(importProfiles[selected]);
  onImportExcel();
}

function rerender() {
  renderSummary();
  renderTable();
  renderFilterYearOptions();
  renderReportYearOptions();
  renderReportChart();
}

function bindEvents() {
  els.form.addEventListener("submit", onSubmit);
  els.type.addEventListener("change", toggleRecipient);
  els.cancelEditBtn.addEventListener("click", resetFormToCreateMode);
  els.entriesBodyAll.addEventListener("click", onRowActions);
  els.entriesBodyIncome.addEventListener("click", onRowActions);
  els.entriesBodyDonation.addEventListener("click", onRowActions);

  [els.search, els.filterYear, els.fromDate, els.toDate].forEach((el) => {
    if (!el) return;
    el.addEventListener("input", renderTable);
    el.addEventListener("change", renderTable);
  });

  els.date.addEventListener("change", updateHebrewDatePreview);
  els.date.addEventListener("input", updateHebrewDatePreview);

  els.tabBtns.forEach((btn) => {
    btn.addEventListener("click", () => {
      activeTab = btn.dataset.tab || "all";
      els.tabBtns.forEach((x) => x.classList.toggle("active", x === btn));
      renderTabPanels();
      renderTable();
    });
  });

  if (els.topTabMain) {
    els.topTabMain.addEventListener("click", () => {
      activeTopTab = "main";
      renderTopPanels();
    });
  }
  if (els.topTabImport) {
    els.topTabImport.addEventListener("click", () => {
      activeTopTab = "import";
      renderTopPanels();
    });
  }

  els.exportBtn.addEventListener("click", exportBackup);
  els.exportCsvBtn.addEventListener("click", exportCsv);
  els.exportXlsxBtn.addEventListener("click", exportXlsx);

  if (els.undoBtn) {
    els.undoBtn.addEventListener("click", undoLastAction);
  }
  if (els.redoBtn) {
    els.redoBtn.addEventListener("click", redoLastAction);
  }

  els.importInput.addEventListener("change", async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    try {
      await importBackup(file);
      showNotice("הגיבוי יובא בהצלחה", "success");
    } catch (_err) {
      showNotice("ייבוא נכשל: הקובץ לא בפורמט תקין", "error", 5000);
    } finally {
      e.target.value = "";
    }
  });

  els.clearBtn.addEventListener("click", () => {
    const ok = confirm("למחוק את כל הנתונים? פעולה זו לא ניתנת לביטול.");
    if (!ok) return;
    exportBackupBeforeClear();
    pushHistorySnapshot();
    state.entries = [];
    saveState();
    resetFormToCreateMode();
    rerender();
    showNotice("כל הנתונים נמחקו (נוצר גיבוי אוטומטי לפני המחיקה)", "success", 4800);
  });

  els.reportYear.addEventListener("change", renderReportChart);
  els.reportMode.addEventListener("change", renderReportChart);

  els.excelInput.addEventListener("change", async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    try {
      await onExcelFileChosen(file);
    } catch (err) {
      const details = err && err.message ? `: ${err.message}` : "";
      showNotice(`קריאת קובץ אקסל נכשלה${details}`, "error", 5200);
    } finally {
      e.target.value = "";
    }
  });

  els.excelSheet.addEventListener("change", () => {
    try {
      loadSelectedSheetRows();
      renderParsedExcelPreview();
    } catch (_err) {
      showNotice("טעינת הגיליון נכשלה", "error", 5000);
    }
  });

  els.excelHasHeader.addEventListener("change", () => {
    if (!excelRows.length) return;
    updateMappingOptionsFromSelectedRow();
    
    renderExcelPreview();
    renderParsedExcelPreview();
    updateSelectedRowsCounter();
  });

  els.excelStartRow.addEventListener("change", () => {
    if (!excelRows.length) return;
    updateMappingOptionsFromSelectedRow();
    
    renderExcelPreview();
    renderParsedExcelPreview();
    updateSelectedRowsCounter();
  });

  els.excelType.addEventListener("change", () => renderParsedExcelPreview());
  els.excelAmountMode.addEventListener("change", () => renderParsedExcelPreview());
  els.excelFixedDate.addEventListener("change", () => renderParsedExcelPreview());
  els.excelImportSearch.addEventListener("input", () => {
    renderExcelPreview();
    renderParsedExcelPreview();
    updateSelectedRowsCounter();
  });

  [els.mapDescription, els.mapAmount, els.mapDate, els.mapNotes, els.mapRecipient].forEach((sel) => {
    sel.addEventListener("change", () => renderParsedExcelPreview());
  });

  els.saveProfileBtn.addEventListener("click", saveCurrentProfile);
  els.loadProfileBtn.addEventListener("click", loadSelectedProfile);
  els.deleteProfileBtn.addEventListener("click", deleteSelectedProfile);
  els.setDefaultProfileBtn.addEventListener("click", setDefaultProfile);
  els.clearDefaultProfileBtn.addEventListener("click", clearDefaultProfile);
  els.autoProfileMode.addEventListener("change", onAutoProfileModeChange);
  els.exportProfilesBtn.addEventListener("click", exportProfilesJson);
  els.autoMapBtn.addEventListener("click", autoMapColumns);

  els.importProfilesInput.addEventListener("change", async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    try {
      await importProfilesJson(file);
      showNotice("תבניות יובאו בהצלחה", "success");
    } catch (_err) {
      showNotice("ייבוא תבניות נכשל", "error", 5000);
    } finally {
      e.target.value = "";
    }
  });

  els.excelRowMode.addEventListener("change", () => {
    renderExcelPreview();
    renderParsedExcelPreview();
    updateSelectedRowsCounter();
  });

  els.selectAllRowsBtn.addEventListener("click", () => {
    if (els.excelRowMode.value !== "selected") {
      els.excelRowMode.value = "selected";
      renderExcelPreview();
    }
    setAllRowChecks(true);
    updateSelectedRowsCounter();
  });

  els.clearAllRowsBtn.addEventListener("click", () => {
    if (els.excelRowMode.value !== "selected") {
      els.excelRowMode.value = "selected";
      renderExcelPreview();
    }
    setAllRowChecks(false);
    updateSelectedRowsCounter();
  });

  els.importExcelBtn.addEventListener("click", onImportExcel);
  els.quickImportBtn.addEventListener("click", onQuickImport);
}

function init() {
  loadState();
  state.entries = state.entries.map((e) => ({
    ...e,
    hebrewDate: toHebrewDate(e.date) || e.hebrewDate || ""
  }));
  loadProfiles();
  loadProfileSettings();
  bindEvents();
  renderProfileOptions();
  updateDefaultProfileUiHint();
  resetFormToCreateMode();
  setImportStep(1);
  renderTopPanels();
  renderTabPanels();
  rerender();
  updateSelectedRowsCounter();
  updateUndoRedoButtons();
}

init();

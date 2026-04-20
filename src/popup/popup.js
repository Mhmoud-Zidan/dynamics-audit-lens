/**
 * popup.js — Entry point for the Dynamics Audit Lens popup UI.
 *
 * Two modes:
 *   Tab 1 — Record Audit: export audit history for selected grid records.
 *   Tab 2 — User Audit:   export audit history for a specific user on any
 *                          entity, with optional date range filter.
 *
 * User tab state (entity, user, dates) is persisted to chrome.storage.local
 * so it survives popup close/reopen.
 *
 * All persistence goes through chrome.storage.local — strictly local,
 * no external requests are ever made from this extension.
 */

import * as XLSX from "xlsx";

// ── Constants ─────────────────────────────────────────────────────────────────

const MAX_EXPORT_RECORDS = 250;
const MAX_EXPORT_ROWS = 100_000;
const SEARCH_DEBOUNCE_MS = 300;
const MIN_SEARCH_LENGTH = 2;
const STATE_STORAGE_KEY = "userAuditState";

// ── DOM references (Records tab) ──────────────────────────────────────────────

const statusEl = document.getElementById("status-msg");
const recordInfoEl = document.getElementById("record-info");
const recordCountEl = document.getElementById("record-count");
const entityNameEl = document.getElementById("entity-name");
const exportBtn = document.getElementById("export-btn");
const progressSection = document.getElementById("progress-section");
const progressFill = document.getElementById("progress-fill");
const progressText = document.getElementById("progress-text");

// ── DOM references (User tab) ─────────────────────────────────────────────────

const userStatusEl = document.getElementById("user-status-msg");
const entitySearchInput = document.getElementById("entity-search-input");
const entitySearchDropdown = document.getElementById("entity-search-dropdown");
const userSearchInput = document.getElementById("user-search-input");
const userSearchDropdown = document.getElementById("user-search-dropdown");
const selectedUserEl = document.getElementById("selected-user");
const selectedUserNameEl = document.getElementById("selected-user-name");
const clearUserBtn = document.getElementById("clear-user-btn");
const dateFromInput = document.getElementById("date-from");
const dateToInput = document.getElementById("date-to");
const userExportBtn = document.getElementById("user-export-btn");
const userProgressSection = document.getElementById("user-progress-section");
const userProgressFill = document.getElementById("user-progress-fill");
const userProgressText = document.getElementById("user-progress-text");

// ── State ─────────────────────────────────────────────────────────────────────

let currentContext = null;
let currentTabId = null;
let exporting = false;
let userExporting = false;
let selectedUser = null;

// ── Version badge ────────────────────────────────────────────────────────────

(function setVersionBadge() {
  const v = `v${chrome.runtime.getManifest().version}`;
  const header = document.getElementById("app-version");
  if (header) header.textContent = v;
  const modal = document.getElementById("modal-version");
  if (modal) modal.textContent = v;
})();

// ── Tab switching ─────────────────────────────────────────────────────────────

const tabBtns = document.querySelectorAll(".tab");
const tabPanels = document.querySelectorAll(".tab-panel");

tabBtns.forEach((btn) => {
  btn.addEventListener("click", () => {
    const target = btn.dataset.tab;
    tabBtns.forEach((b) => b.classList.toggle("tab--active", b.dataset.tab === target));
    tabPanels.forEach((p) => {
      const isActive = p.id === `tab-${target}`;
      p.classList.toggle("tab-panel--active", isActive);
      p.hidden = !isActive;
    });
  });
});

// ── UI helpers ────────────────────────────────────────────────────────────────

function makeStatusSetter(el) {
  return (text, type = "idle") => {
    el.textContent = text;
    el.className = `status status--${type}`;
  };
}

const setStatus = makeStatusSetter(statusEl);
const setUserStatus = makeStatusSetter(userStatusEl);

function makeProgressUpdater(fillEl, textEl) {
  return (processed, total) => {
    const pct = total > 0 ? Math.round((processed / total) * 100) : 0;
    fillEl.style.width = `${pct}%`;
    textEl.textContent = `Processed ${processed} of ${total} records\u2026`;
    textEl.className = "progress-text";
  };
}

const updateProgress = makeProgressUpdater(progressFill, progressText);
const updateUserProgress = makeProgressUpdater(userProgressFill, userProgressText);

function formatDateStamp() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, "0")}${String(d.getDate()).padStart(2, "0")}`;
}

function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

// ── Excel export ──────────────────────────────────────────────────────────────

function generateExcel(rows, entityName, filenameSuffix) {
  let ws = XLSX.utils.json_to_sheet(rows);

  const range = XLSX.utils.decode_range(ws["!ref"]);
  let dateCol = -1;
  for (let C = range.s.c; C <= range.e.c; ++C) {
    const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: C });
    const cell = ws[cellAddress];
    if (cell && cell.v === "ChangedDate") {
      dateCol = C;
      break;
    }
  }
  if (dateCol >= 0) {
    for (let R = range.s.r + 1; R <= range.e.r; ++R) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: dateCol });
      const cell = ws[cellAddress];
      if (cell && cell.v != null && cell.v !== "") {
        if (cell.v instanceof Date || typeof cell.v === "number") {
          cell.t = "d";
          cell.z = "yyyy-mm-dd hh:mm:ss";
        }
      }
    }
  }

  let wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Audit History");

  if (rows.length > 0) {
    const keys = Object.keys(rows[0]);
    const sample = rows.slice(0, 200);
    ws["!cols"] = keys.map((key) => {
      const maxLen = Math.max(
        key.length,
        ...sample.map((r) => String(r[key] ?? "").length),
      );
      return { wch: Math.min(maxLen + 2, 50) };
    });

    ws["!autofilter"] = { ref: ws["!ref"] };
  }

  const safeName = String(entityName ?? "Unknown").replace(
    /[^a-zA-Z0-9_-]/g,
    "_",
  );
  const suffix = filenameSuffix ? `_${filenameSuffix}` : "";
  const filename = `AuditExport_${safeName}${suffix}_${formatDateStamp()}.xlsx`;

  const wbOut = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  wb = null;
  ws = null;

  const blob = new Blob([wbOut], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// ── Messaging ─────────────────────────────────────────────────────────────────

function sendToTab(tabId, message) {
  return chrome.tabs.sendMessage(tabId, message);
}

// ── State persistence ─────────────────────────────────────────────────────────

async function saveUserAuditState() {
  const state = {
    entityName: entitySearchInput.value.trim() || null,
    selectedUser,
    dateFrom: dateFromInput.value || null,
    dateTo: dateToInput.value || null,
  };
  try {
    await chrome.storage.local.set({ [STATE_STORAGE_KEY]: state });
  } catch { /* storage unavailable */ }
}

async function loadUserAuditState() {
  try {
    const result = await chrome.storage.local.get(STATE_STORAGE_KEY);
    const state = result?.[STATE_STORAGE_KEY];
    if (!state) return;

    if (state.entityName) {
      entitySearchInput.value = state.entityName;
    }
    if (state.selectedUser) {
      selectedUser = state.selectedUser;
      selectedUserNameEl.textContent =
        selectedUser.fullname || selectedUser.email || (selectedUser.id ?? "").slice(0, 8);
      selectedUserEl.hidden = false;
    }
    if (state.dateFrom) {
      dateFromInput.value = state.dateFrom;
    }
    if (state.dateTo) {
      dateToInput.value = state.dateTo;
    }
  } catch { /* storage unavailable */ }
}

// ── Context detection ─────────────────────────────────────────────────────────

const DYNAMICS_PATTERN =
  /^https?:\/\/[^/]+\.(crm\d*\.dynamics\.com|crm\.microsoftdynamics\.us|crm\.appsplatform\.us|crm\.dynamics\.cn)\//;

async function fetchContext() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  if (!tab?.id) {
    setStatus("Cannot access current tab.", "error");
    setUserStatus("Cannot access current tab.", "error");
    return;
  }

  currentTabId = tab.id;

  const isDynamics = DYNAMICS_PATTERN.test(tab.url ?? "");

  if (!isDynamics) {
    setStatus("Not a Dynamics / Dataverse page.", "idle");
    setUserStatus("Not a Dynamics / Dataverse page.", "idle");
    return;
  }

  try {
    const response = await sendToTab(tab.id, { type: "GET_CONTEXT" });

    if (!response?.ok || !response.context) {
      setStatus("Could not read page context.", "error");
      setUserStatus("Could not read page context.", "error");
      return;
    }

    currentContext = response.context;
    const count = currentContext.selectedIds?.length ?? 0;
    const entity = currentContext.entityName;
    const hostname = new URL(tab.url).hostname;

    // ── Records tab ──
    setStatus(`Active on: ${hostname}`, "active");
    recordCountEl.textContent = `${count} record${count !== 1 ? "s" : ""} selected`;
    if (entity) entityNameEl.textContent = entity;
    recordInfoEl.hidden = false;

    if (currentContext.selectionUnavailable) {
      recordCountEl.textContent = "0 records selected";
    }

    if (count > MAX_EXPORT_RECORDS) {
      setStatus(
        `Too many records selected (max ${MAX_EXPORT_RECORDS}). Narrow your selection.`,
        "error",
      );
      exportBtn.disabled = true;
    } else {
      exportBtn.disabled = count === 0;
    }

    // ── User tab ──
    setUserStatus(`Active on: ${hostname}`, "active");

    // Pre-fill entity from page context if the input is empty (first visit).
    // If persisted state already has a value, keep it.
    if (entity && !entitySearchInput.value) {
      entitySearchInput.value = entity;
    }
  } catch (err) {
    console.warn("[Audit Lens] fetchContext failed:", err);
    setStatus("Content script not ready. Reload the page.", "error");
    setUserStatus("Content script not ready. Reload the page.", "error");
  }

  // Load persisted state (entity, user, dates) and apply defaults.
  await loadUserAuditState();

  // Default "To" date to today if not already set. "From" stays empty (open-ended).
  if (!dateToInput.value) {
    dateToInput.value = todayISO();
  }

  updateUserExportBtnState();
}

// ── Search wiring ─────────────────────────────────────────────────────────────

function wireSearchInput({ input, dropdown, timeoutRef, searchFn, renderFn, onChange }) {
  let timeout;

  input.addEventListener("input", () => {
    clearTimeout(timeout);
    const query = input.value.trim();

    if (query.length < MIN_SEARCH_LENGTH) {
      dropdown.hidden = true;
      if (onChange) onChange();
      return;
    }

    if (onChange) onChange();
    timeout = setTimeout(() => searchFn(query), SEARCH_DEBOUNCE_MS);
  });

  input.addEventListener("blur", () => {
    setTimeout(() => { dropdown.hidden = true; }, 200);
  });

  input.addEventListener("focus", () => {
    if (dropdown.children.length > 0 && input.value.trim().length >= MIN_SEARCH_LENGTH) {
      dropdown.hidden = false;
    }
  });
}

function renderDropdown(dropdown, items, emptyText, onSelect) {
  dropdown.replaceChildren();

  if (items.length === 0) {
    const empty = document.createElement("div");
    empty.className = "search-dropdown__empty";
    empty.textContent = emptyText;
    dropdown.appendChild(empty);
    dropdown.hidden = false;
    return;
  }

  for (const item of items) {
    const el = document.createElement("div");
    el.className = "search-dropdown__item";

    const nameSpan = document.createElement("span");
    nameSpan.className = "search-dropdown__item-name";
    nameSpan.textContent = item.name;

    const subSpan = document.createElement("span");
    subSpan.className = "search-dropdown__item-email";
    subSpan.textContent = item.sub || "";

    el.appendChild(nameSpan);
    el.appendChild(subSpan);

    el.addEventListener("mousedown", (e) => {
      e.preventDefault();
      onSelect(item.raw);
    });

    dropdown.appendChild(el);
  }

  dropdown.hidden = false;
}

// ── Entity search ─────────────────────────────────────────────────────────────

async function performEntitySearch(query) {
  if (!currentTabId) return;

  try {
    const response = await sendToTab(currentTabId, {
      type: "SEARCH_ENTITIES",
      query,
    });

    if (!response?.ok || !response.entities?.length) {
      renderEntityDropdown([]);
      return;
    }
    renderEntityDropdown(response.entities);
  } catch {
    entitySearchDropdown.hidden = true;
  }
}

function renderEntityDropdown(entities) {
  renderDropdown(
    entitySearchDropdown,
    entities.map((ent) => ({
      name: ent.displayName || ent.logicalName,
      sub: ent.displayName !== ent.logicalName ? `(${ent.logicalName})` : "",
      raw: ent,
    })),
    "No entities found.",
    (ent) => {
      entitySearchInput.value = ent.logicalName;
      entitySearchDropdown.hidden = true;
      updateUserExportBtnState();
      saveUserAuditState();
    },
  );
}

wireSearchInput({
  input: entitySearchInput,
  dropdown: entitySearchDropdown,
  searchFn: performEntitySearch,
  renderFn: renderEntityDropdown,
  onChange: () => { updateUserExportBtnState(); saveUserAuditState(); },
});

// ── User search ───────────────────────────────────────────────────────────────

async function performUserSearch(query) {
  if (!currentTabId) return;

  try {
    const response = await sendToTab(currentTabId, {
      type: "SEARCH_USERS",
      query,
    });

    if (!response?.ok || !response.users?.length) {
      renderUserDropdown([]);
      return;
    }
    renderUserDropdown(response.users);
  } catch {
    userSearchDropdown.hidden = true;
  }
}

function renderUserDropdown(users) {
  renderDropdown(
    userSearchDropdown,
    users.map((user) => ({
      name: user.fullname || "(unnamed)",
      sub: user.email ? `(${user.email})` : "",
      raw: user,
    })),
    "No users found.",
    (user) => selectUser(user),
  );
}

wireSearchInput({
  input: userSearchInput,
  dropdown: userSearchDropdown,
  searchFn: performUserSearch,
  renderFn: renderUserDropdown,
});

function selectUser(user) {
  selectedUser = user;
  userSearchInput.value = "";
  userSearchDropdown.hidden = true;

  selectedUserNameEl.textContent =
    user.fullname || user.email || (user.id ?? "").slice(0, 8);
  selectedUserEl.hidden = false;

  updateUserExportBtnState();
  saveUserAuditState();
}

clearUserBtn.addEventListener("click", () => {
  selectedUser = null;
  selectedUserEl.hidden = true;
  updateUserExportBtnState();
  saveUserAuditState();
});

// ── Date change handlers ─────────────────────────────────────────────────────

dateFromInput.addEventListener("change", () => saveUserAuditState());
dateToInput.addEventListener("change", () => saveUserAuditState());

// ── Export button state ───────────────────────────────────────────────────────

function updateUserExportBtnState() {
  const entityOk = entitySearchInput.value.trim().length >= MIN_SEARCH_LENGTH;
  const userOk = !!selectedUser;
  userExportBtn.disabled = !entityOk || !userOk || userExporting;
}

// ── Record export orchestration ───────────────────────────────────────────────

async function startExport() {
  if (exporting || !currentContext || !currentTabId) return;

  const { entityName, selectedIds } = currentContext;
  if (!entityName || !selectedIds?.length) return;

  if (selectedIds.length > MAX_EXPORT_RECORDS) {
    setStatus(`Too many records (max ${MAX_EXPORT_RECORDS}).`, "error");
    return;
  }

  exporting = true;
  exportBtn.disabled = true;
  progressSection.hidden = false;
  updateProgress(0, selectedIds.length);

  const port = chrome.tabs.connect(currentTabId, { name: "audit-export" });

  port.onMessage.addListener((msg) => {
    if (msg.type === "progress") {
      updateProgress(msg.done, msg.total);
    }

    if (msg.type === "done") {
      const rows = (msg.rows ?? []).slice(0, MAX_EXPORT_ROWS);

      if (rows.length === 0) {
        progressText.textContent = "No audit records found.";
        progressText.className = "progress-text progress-text--empty";
      } else {
        if (msg.rows.length > MAX_EXPORT_ROWS) {
          progressText.textContent = `Capped at ${MAX_EXPORT_ROWS.toLocaleString()} rows. Generating file\u2026`;
        }
        generateExcel(rows, entityName);
        progressText.textContent = `Export complete \u2014 ${rows.length} row${rows.length !== 1 ? "s" : ""}.`;
        progressText.className = "progress-text progress-text--success";
      }

      exporting = false;
      exportBtn.disabled = false;
    }

    if (msg.type === "error") {
      progressText.textContent = `Error: ${msg.error}`;
      progressText.className = "progress-text progress-text--error";
      setStatus("Export failed.", "error");
      exporting = false;
      exportBtn.disabled = false;
    }
  });

  port.onDisconnect.addListener(() => {
    if (exporting) {
      progressText.textContent = "Connection lost. Reload the page and retry.";
      progressText.className = "progress-text progress-text--error";
      setStatus("Export failed.", "error");
      exporting = false;
      exportBtn.disabled = false;
    }
  });

  port.postMessage({
    entityLogicalName: entityName,
    guids: selectedIds,
  });
}

// ── User audit export orchestration ───────────────────────────────────────────

async function startUserExport() {
  if (userExporting || !currentTabId || !selectedUser) return;

  const entityLogicalName = entitySearchInput.value.trim().toLowerCase();
  if (!entityLogicalName) return;

  const dateFrom = dateFromInput.value || null;
  const dateTo = dateToInput.value || null;

  userExporting = true;
  userExportBtn.disabled = true;
  userProgressSection.hidden = false;
  userProgressFill.style.width = "0%";
  userProgressText.textContent = "Querying audit records\u2026";
  userProgressText.className = "progress-text";

  const port = chrome.tabs.connect(currentTabId, { name: "user-audit-export" });

  port.onMessage.addListener((msg) => {
    if (msg.type === "phase") {
      userProgressText.textContent = msg.text;
      userProgressText.className = "progress-text";
    }

    if (msg.type === "progress") {
      updateUserProgress(msg.done, msg.total);
    }

    if (msg.type === "done") {
      const rows = (msg.rows ?? []).slice(0, MAX_EXPORT_ROWS);

      if (rows.length === 0) {
        userProgressText.textContent = "No audit records found for this user.";
        userProgressText.className = "progress-text progress-text--empty";
      } else {
        const suffix = selectedUser.fullname
          ? selectedUser.fullname.replace(/[^a-zA-Z0-9_-]/g, "_")
          : (selectedUser.id ?? "unknown").slice(0, 8);
        generateExcel(rows, entityLogicalName, suffix);
        userProgressText.textContent =
          `Export complete \u2014 ${rows.length} row${rows.length !== 1 ? "s" : ""}.`;
        userProgressText.className = "progress-text progress-text--success";
      }

      userExporting = false;
      updateUserExportBtnState();
    }

    if (msg.type === "error") {
      userProgressText.textContent = `Error: ${msg.error}`;
      userProgressText.className = "progress-text progress-text--error";
      setUserStatus("Export failed.", "error");
      userExporting = false;
      updateUserExportBtnState();
    }
  });

  port.onDisconnect.addListener(() => {
    if (userExporting) {
      userProgressText.textContent = "Connection lost. Reload the page and retry.";
      userProgressText.className = "progress-text progress-text--error";
      setUserStatus("Export failed.", "error");
      userExporting = false;
      updateUserExportBtnState();
    }
  });

  port.postMessage({
    entityLogicalName,
    userGuid: selectedUser.id,
    dateFrom,
    dateTo,
  });
}

// ── Event wiring ──────────────────────────────────────────────────────────────

exportBtn.addEventListener("click", startExport);
userExportBtn.addEventListener("click", startUserExport);

// ── Fill Data ─────────────────────────────────────────────────────────────────

const fillDataBtn = document.getElementById("fill-data-btn");
const fillStatusEl = document.getElementById("fill-status");

let fillStatusTimer = null;

function showFillStatus(text, type) {
  clearTimeout(fillStatusTimer);
  fillStatusEl.textContent = text;
  fillStatusEl.className = "fill-status fill-status--" + type;
  fillStatusEl.hidden = false;
  if (type === "ok" || type === "err") {
    fillStatusTimer = setTimeout(() => { fillStatusEl.hidden = true; }, 4000);
  }
}

fillDataBtn.addEventListener("click", async () => {
  if (!currentTabId) return;

  fillDataBtn.disabled = true;
  showFillStatus("Filling form fields\u2026", "loading");

  try {
    const response = await sendToTab(currentTabId, { type: "FILL_DATA" });

    if (response?.ok) {
      const lookupErrs = response.lookupErrors || [];
      if (response.filled > 0) {
        let msg =
          "Filled " + response.filled + " of " + response.total +
          " fields (" + response.skipped + " skipped).";
        if (lookupErrs.length > 0) {
          msg += " Lookup issues: " + lookupErrs.slice(0, 2).join("; ");
        }
        showFillStatus(msg, "ok");
        fillDataBtn.classList.add("fill-data-flash");
        setTimeout(() => fillDataBtn.classList.remove("fill-data-flash"), 700);
      } else {
        let msg =
          "0 fields filled (" + response.skipped + " skipped). " +
          "All fields may already have values or be read-only.";
        if (lookupErrs.length > 0) {
          msg += " Lookup errors: " + lookupErrs.join("; ");
        }
        showFillStatus(msg, "err");
      }
      if (lookupErrs.length > 0 || (response.errors || []).length > 0) {
        console.log("[Audit Lens] Fill debug — formType:", response.formType,
          "sample:", (response.sample || []).join(", "),
          "errors:", (response.errors || []).join("; "),
          "lookupErrors:", lookupErrs.join("; "));
      }
    } else {
      showFillStatus(response?.error || "Failed to fill form data.", "err");
    }
  } catch {
    showFillStatus("Could not reach content script. Reload the page.", "err");
  }

  fillDataBtn.disabled = false;
});

// ── Settings / Theme / About ──────────────────────────────────────────────────

const settingsBtn      = document.getElementById("settings-btn");
const settingsMenu     = document.getElementById("settings-menu");
const themeToggleBtn   = document.getElementById("theme-toggle-btn");
const themeLabel       = document.getElementById("theme-label");
const aboutBtn         = document.getElementById("about-btn");
const aboutModal       = document.getElementById("about-modal");
const aboutCloseBtn    = document.getElementById("about-close-btn");

const THEME_STORAGE_KEY = "theme";

function applyTheme(theme) {
  if (theme === "light") {
    document.documentElement.setAttribute("data-theme", "light");
    themeLabel.textContent = "Dark Mode";
  } else {
    document.documentElement.removeAttribute("data-theme");
    themeLabel.textContent = "Light Mode";
  }
}

async function loadTheme() {
  try {
    const result = await chrome.storage.local.get(THEME_STORAGE_KEY);
    applyTheme(result?.[THEME_STORAGE_KEY] ?? "light");
  } catch {
    applyTheme("light");
  }
}

async function toggleTheme() {
  const isLight = document.documentElement.getAttribute("data-theme") === "light";
  const next = isLight ? "dark" : "light";
  applyTheme(next);
  try {
    await chrome.storage.local.set({ [THEME_STORAGE_KEY]: next });
  } catch { /* storage unavailable */ }
}

// Open / close settings dropdown
settingsBtn.addEventListener("click", (e) => {
  e.stopPropagation();
  settingsMenu.hidden = !settingsMenu.hidden;
});

// Close on outside click
document.addEventListener("click", () => {
  settingsMenu.hidden = true;
});

// Keep clicks inside the menu from bubbling to the document handler
settingsMenu.addEventListener("click", (e) => {
  e.stopPropagation();
});

themeToggleBtn.addEventListener("click", () => {
  settingsMenu.hidden = true;
  toggleTheme();
});

aboutBtn.addEventListener("click", () => {
  settingsMenu.hidden = true;
  aboutModal.hidden = false;
});

// ── Audit Settings ─────────────────────────────────────────────────────────────

const auditSettingsBtn = document.getElementById("audit-settings-btn");

async function resolveAppId(tab) {
  try {
    const appId = new URL(tab.url).searchParams.get("appid");
    if (appId) return appId;
  } catch (_) {}

  try {
    const resp = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: () =>
        fetch("/api/data/v9.2/appmodules?$select=appmoduleid&$top=1", {
          headers: { Accept: "application/json" },
        }).then((r) => r.json()),
    });
    const apps = resp?.[0]?.result?.value;
    if (apps?.length) return apps[0].appmoduleid;
  } catch (_) {}

  return null;
}

auditSettingsBtn.addEventListener("click", async () => {
  settingsMenu.hidden = true;
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  const origin =
    tab?.url && DYNAMICS_PATTERN.test(tab.url)
      ? new URL(tab.url).origin
      : null;

  if (!origin) {
    chrome.tabs.create({ url: "https://admin.powerplatform.microsoft.com/" });
    return;
  }

  const appId = tab ? await resolveAppId(tab) : null;
  const encodedData = encodeURIComponent('{"area":"nav_audit"}');
  const url = `${origin}/main.aspx?appid=${appId}&pagetype=control&controlName=PowerAdmin.EnvironmentSettings.NavigatorPage&data=${encodedData}`;
  chrome.tabs.create({ url });
});

aboutCloseBtn.addEventListener("click", () => {
  aboutModal.hidden = true;
});

// Close modal on backdrop click
aboutModal.addEventListener("click", (e) => {
  if (e.target === aboutModal) {
    aboutModal.hidden = true;
  }
});

// External links inside About modal
document.getElementById("linkedin-btn").addEventListener("click", () => {
  chrome.tabs.create({ url: "https://www.linkedin.com/in/mahmoudzidan55" });
});

document.getElementById("github-repo-btn").addEventListener("click", () => {
  chrome.tabs.create({ url: "https://github.com/Mhmoud-Zidan/dynamics-audit-lens/releases" });
});

document.addEventListener("DOMContentLoaded", () => {
  loadTheme();
  fetchContext();
});

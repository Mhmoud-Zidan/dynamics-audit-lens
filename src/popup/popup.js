/**
 * popup.js — Entry point for the Dynamics Audit Lens popup UI.
 *
 * Handles context detection, audit extraction with progress tracking,
 * and Excel export via SheetJS (xlsx).
 * All persistence goes through chrome.storage.local — strictly local,
 * no external requests are ever made from this extension.
 */

import * as XLSX from "xlsx";

// ── Constants ─────────────────────────────────────────────────────────────────

/** Hard cap on records to export — protects against Dataverse 429 rate limits. */
const MAX_EXPORT_RECORDS = 250;

/** Safety cap on total formatted rows to prevent OOM in the popup context. */
const MAX_EXPORT_ROWS = 100_000;

// ── DOM references ────────────────────────────────────────────────────────────

const statusEl = document.getElementById("status-msg");
const recordInfoEl = document.getElementById("record-info");
const recordCountEl = document.getElementById("record-count");
const entityNameEl = document.getElementById("entity-name");
const exportBtn = document.getElementById("export-btn");
const progressSection = document.getElementById("progress-section");
const progressFill = document.getElementById("progress-fill");
const progressText = document.getElementById("progress-text");

// ── State ─────────────────────────────────────────────────────────────────────

let currentContext = null;
let currentTabId = null;
let exporting = false;

// ── UI helpers ────────────────────────────────────────────────────────────────

/** Update the status banner in the popup. */
function setStatus(text, type = "idle") {
  statusEl.textContent = text;
  statusEl.className = `status status--${type}`; // idle | active | error
}

/** Update the progress bar and text. */
function updateProgress(processed, total) {
  const pct = total > 0 ? Math.round((processed / total) * 100) : 0;
  progressFill.style.width = `${pct}%`;
  progressText.textContent = `Processed ${processed} of ${total} records\u2026`;
}

/** Return a YYYYMMDD date stamp for filenames. */
function formatDateStamp() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, "0")}${String(d.getDate()).padStart(2, "0")}`;
}

// ── Excel export ──────────────────────────────────────────────────────────────

/**
 * Generate an .xlsx file from an array of row objects and trigger a download.
 *
 * @param {FormattedAuditRow[]} rows        Flat audit data rows.
 * @param {string|null}         entityName  Entity logical name for the filename.
 */
function generateExcel(rows, entityName) {
  let ws = XLSX.utils.json_to_sheet(rows);
  let wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Audit History");

  // Auto-size columns (sample first 200 rows to avoid perf hit on large sets).
  if (rows.length > 0) {
    const keys = Object.keys(rows[0]);
    ws["!cols"] = keys.map((key) => {
      const sample = rows.slice(0, 200);
      const maxLen = Math.max(
        key.length,
        ...sample.map((r) => String(r[key] ?? "").length),
      );
      return { wch: Math.min(maxLen + 2, 50) };
    });
  }

  const safeName = String(entityName ?? "Unknown").replace(
    /[^a-zA-Z0-9_-]/g,
    "_",
  );
  const filename = `AuditExport_${safeName}_${formatDateStamp()}.xlsx`;

  const wbOut = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  // Release intermediate objects so GC can reclaim memory before blob creation.
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

/** Send a message to the content script running in the given tab. */
function sendToTab(tabId, message) {
  return chrome.tabs.sendMessage(tabId, message);
}

// ── Context detection ─────────────────────────────────────────────────────────

const DYNAMICS_PATTERN =
  /^https?:\/\/[^/]+\.(crm\d*\.dynamics\.com|crm\.microsoftdynamics\.us|crm\.appsplatform\.us|crm\.dynamics\.cn)\//;

/** Query the active tab, verify it's a Dynamics page, and fetch Xrm context. */
async function fetchContext() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  if (!tab?.id) {
    setStatus("Cannot access current tab.", "error");
    return;
  }

  currentTabId = tab.id;

  if (!DYNAMICS_PATTERN.test(tab.url ?? "")) {
    setStatus("Not a Dynamics / Dataverse page.", "idle");
    return;
  }

  try {
    const response = await sendToTab(tab.id, { type: "GET_CONTEXT" });

    if (!response?.ok || !response.context) {
      setStatus("Could not read page context.", "error");
      return;
    }

    currentContext = response.context;
    const count = currentContext.selectedIds?.length ?? 0;
    const entity = currentContext.entityName;

    setStatus(`Active on: ${new URL(tab.url).hostname}`, "active");

    recordCountEl.textContent = `${count} record${count !== 1 ? "s" : ""} selected`;
    if (entity) entityNameEl.textContent = entity;
    recordInfoEl.hidden = false;

    if (currentContext.selectionUnavailable) {
      setStatus(
        "Grid selection could not be read. Try selecting records and re-opening.",
        "error",
      );
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
  } catch (err) {
    console.warn("[Audit Lens] fetchContext failed:", err);
    setStatus("Content script not ready. Reload the page.", "error");
  }
}

// ── Export orchestration ──────────────────────────────────────────────────────

/**
 * Kick off the audit extraction + Excel export pipeline.
 *
 * Opens a named port to the content script and sends all GUIDs at once.
 * The content script runs them through its concurrency pool (MAX_CONCURRENT = 5)
 * and streams incremental progress messages back through the port.
 */
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
      } else {
        if (msg.rows.length > MAX_EXPORT_ROWS) {
          progressText.textContent = `Capped at ${MAX_EXPORT_ROWS.toLocaleString()} rows. Generating file\u2026`;
        }
        generateExcel(rows, entityName);
        progressText.textContent = `Export complete \u2014 ${rows.length} row${rows.length !== 1 ? "s" : ""}.`;
      }

      exporting = false;
      exportBtn.disabled = false;
    }

    if (msg.type === "error") {
      progressText.textContent = `Error: ${msg.error}`;
      setStatus("Export failed.", "error");
      exporting = false;
      exportBtn.disabled = false;
    }
  });

  port.onDisconnect.addListener(() => {
    if (exporting) {
      progressText.textContent = "Connection lost. Reload the page and retry.";
      setStatus("Export failed.", "error");
      exporting = false;
      exportBtn.disabled = false;
    }
  });

  // Send all GUIDs in one shot — the content script's pool handles concurrency.
  port.postMessage({
    entityLogicalName: entityName,
    guids: selectedIds,
  });
}

// ── Event wiring ──────────────────────────────────────────────────────────────

exportBtn.addEventListener("click", startExport);

document.addEventListener("DOMContentLoaded", () => {
  fetchContext();
});

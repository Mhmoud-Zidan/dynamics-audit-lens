/**
 * service-worker.js — MV3 background service worker.
 *
 * Responsibilities:
 *  - Receive messages from the content script and popup.
 *  - Keep the popup badge in sync with page activity.
 *
 * Security notes:
 *  - No fetch() calls to external origins.
 *  - No data persisted to storage — badge state is ephemeral.
 *  - Message origins are validated against the extension's own ID.
 */

"use strict";

// ── Lifecycle ───────────────────────────────────────────────────────────────

chrome.runtime.onInstalled.addListener(({ reason, previousVersion }) => {
  if (reason === chrome.runtime.OnInstalledReason.INSTALL) {
    console.debug("[Audit Lens] Extension installed.");
  }

  if (reason === chrome.runtime.OnInstalledReason.UPDATE) {
    console.debug(`[Audit Lens] Updated from ${previousVersion}.`);
  }
});

// ── Message router ──────────────────────────────────────────────────────────

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (sender.id !== chrome.runtime.id) return false;

  switch (message.type) {
    case "DYNAMICS_PAGE_ACTIVE":
      handleDynamicsPageActive(message.payload, sender.tab);
      sendResponse({ ok: true });
      break;

    case "DYNAMICS_CONTEXT_UPDATE":
      handleContextUpdate(message.payload, sender.tab);
      sendResponse({ ok: true });
      break;

    default:
      console.warn("[Audit Lens] Unknown message type:", message.type);
  }

  return false;
});

// ── Handlers ────────────────────────────────────────────────────────────────

function handleDynamicsPageActive(payload, tab) {
  if (!payload?.hostname) return;

  if (tab?.id != null) {
    chrome.action.setBadgeText({ text: "●", tabId: tab.id });
    chrome.action.setBadgeBackgroundColor({ color: "#4caf7d", tabId: tab.id });
  }
}

function handleContextUpdate(payload, tab) {
  if (!payload?.hostname) return;

  const ctx = payload.context ?? {};
  const pageType = ctx.pageType ? String(ctx.pageType).slice(0, 64) : null;

  if (tab?.id != null) {
    const colour = pageType === "entityrecord" ? "#0078d4" : "#4caf7d";
    chrome.action.setBadgeText({ text: "●", tabId: tab.id });
    chrome.action.setBadgeBackgroundColor({ color: colour, tabId: tab.id });
  }
}

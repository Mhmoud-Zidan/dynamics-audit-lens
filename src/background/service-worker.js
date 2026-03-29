/**
 * service-worker.js — MV3 background service worker.
 *
 * Responsibilities:
 *  - Receive messages from the content script and popup.
 *  - Persist audit data to chrome.storage.local (never to remote servers).
 *  - Keep the popup badge in sync with page activity.
 *
 * Security notes:
 *  - No fetch() calls to external origins.
 *  - All storage operations use chrome.storage.local — data never leaves the device.
 *  - Message origins are validated against the extension's own ID.
 */

'use strict';

// ── Lifecycle ───────────────────────────────────────────────────────────────

chrome.runtime.onInstalled.addListener(({ reason, previousVersion }) => {
  if (reason === chrome.runtime.OnInstalledReason.INSTALL) {
    console.debug('[Audit Lens] Extension installed.');
    // Initialise default storage schema (no external calls).
    chrome.storage.local.set({ sessions: [], settings: { enabled: true } });
  }

  if (reason === chrome.runtime.OnInstalledReason.UPDATE) {
    console.debug(`[Audit Lens] Updated from ${previousVersion}.`);
  }
});

// ── Message router ──────────────────────────────────────────────────────────

/**
 * All messages must come from within the extension itself.
 * External page scripts cannot pass the origin check because they run in an
 * isolated world and must go through chrome.runtime.sendMessage.
 */
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  // Reject messages that do not originate from this extension.
  if (sender.id !== chrome.runtime.id) return false;

  switch (message.type) {
    case 'DYNAMICS_PAGE_ACTIVE':
      handleDynamicsPageActive(message.payload, sender.tab);
      sendResponse({ ok: true });
      break;

    case 'DYNAMICS_CONTEXT_UPDATE':
      handleContextUpdate(message.payload, sender.tab);
      sendResponse({ ok: true });
      break;

    case 'GET_SESSIONS':
      chrome.storage.local.get('sessions', ({ sessions }) => {
        sendResponse({ sessions: sessions ?? [] });
      });
      return true; // Keep message channel open for async response.

    default:
      console.warn('[Audit Lens] Unknown message type:', message.type);
  }

  return false;
});

// ── Handlers ────────────────────────────────────────────────────────────────

/**
 * Persist a page-active event and update the action badge.
 *
 * @param {{ hostname: string, pathname: string, title: string }} payload
 * @param {chrome.tabs.Tab | undefined} tab
 */
function handleDynamicsPageActive(payload, tab) {
  if (!payload?.hostname) return;

  const record = {
    hostname:  String(payload.hostname).slice(0, 253),
    pathname:  String(payload.pathname).slice(0, 2000),
    title:     String(payload.title).slice(0, 200),
    timestamp: Date.now(),
  };

  chrome.storage.local.get('sessions', ({ sessions }) => {
    const updated = [record, ...(sessions ?? [])].slice(0, 500);
    chrome.storage.local.set({ sessions: updated });
  });

  if (tab?.id != null) {
    chrome.action.setBadgeText({ text: '●', tabId: tab.id });
    chrome.action.setBadgeBackgroundColor({ color: '#4caf7d', tabId: tab.id });
  }
}

/**
 * Store a rich Xrm context snapshot (entityName, entityId, selectedIds, …)
 * alongside the session record for this tab.
 *
 * @param {{ hostname: string, pathname: string, title: string, context: object }} payload
 * @param {chrome.tabs.Tab | undefined} tab
 */
function handleContextUpdate(payload, tab) {
  if (!payload?.hostname) return;

  const ctx = payload.context ?? {};

  // Sanitise: only store known scalar fields from the Xrm context.
  const record = {
    hostname:    String(payload.hostname).slice(0, 253),
    pathname:    String(payload.pathname).slice(0, 2000),
    title:       String(payload.title ?? '').slice(0, 200),
    pageType:    ctx.pageType    ? String(ctx.pageType).slice(0, 64)    : null,
    entityName:  ctx.entityName  ? String(ctx.entityName).slice(0, 128) : null,
    entityId:    ctx.entityId    ? String(ctx.entityId).slice(0, 36)    : null,
    selectedIds: Array.isArray(ctx.selectedIds)
      ? ctx.selectedIds.slice(0, 250).map(id => String(id).slice(0, 36))
      : [],
    timestamp:   Date.now(),
  };

  chrome.storage.local.get('sessions', ({ sessions }) => {
    const updated = [record, ...(sessions ?? [])].slice(0, 500);
    chrome.storage.local.set({ sessions: updated });
  });

  // Update badge colour: blue for a form (single record), green for a list.
  if (tab?.id != null) {
    const colour = record.pageType === 'entityrecord' ? '#0078d4' : '#4caf7d';
    chrome.action.setBadgeText({ text: '●', tabId: tab.id });
    chrome.action.setBadgeBackgroundColor({ color: colour, tabId: tab.id });
  }
}

/**
 * popup.js — Entry point for the Dynamics Audit Lens popup UI.
 *
 * All persistence goes through chrome.storage.local — strictly local,
 * no external requests are ever made from this extension.
 */

const statusEl = document.getElementById('status-msg');

/** Update the status banner in the popup. */
function setStatus(text, type = 'idle') {
  statusEl.textContent = text;
  statusEl.className = `status status--${type}`; // idle | active | error
}

/** Query the active tab and verify we are on a supported Dynamics page. */
async function checkActiveTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  if (!tab?.url) {
    setStatus('Cannot read the current tab URL.', 'error');
    return;
  }

  const DYNAMICS_PATTERN = /^https?:\/\/[^/]+\.crm\d*\.dynamics\.com\//;

  if (DYNAMICS_PATTERN.test(tab.url)) {
    setStatus(`Active on: ${new URL(tab.url).hostname}`, 'active');
  } else {
    setStatus('Not a Dynamics / Dataverse page.', 'idle');
  }
}

document.addEventListener('DOMContentLoaded', () => {
  checkActiveTab();
});

/**
 * page-bridge.js — Runs in the PAGE context (main / MAIN world).
 *
 * WHY this file exists:
 *   Chrome content scripts run in an "isolated world" — they share the DOM but NOT
 *   the page's JavaScript globals. window.Xrm is on the page's JS heap, invisible
 *   to a normal content script. The only reliable way to reach it is to inject a
 *   <script> tag that runs in the page's own world.
 *
 * HOW it communicates back:
 *   Uses window.postMessage(). The content script (isolated world) listens for
 *   these messages and validates origin before consuming them.
 *
 * MESSAGE PROTOCOL (all messages are scoped with __DAL__ prefix):
 *   PAGE → CONTENT  __DAL__BRIDGE_READY    { payload: initialContext }
 *   CONTENT → PAGE  __DAL__CONTEXT_REQUEST {}
 *   PAGE → CONTENT  __DAL__CONTEXT_RESPONSE { payload: context }
 *
 * ContextPayload shape:
 * {
 *   available:   boolean,        // whether window.Xrm was found
 *   pageType:    string | null,  // 'entityrecord' | 'entitylist' | 'dashboard' | …
 *   entityName:  string | null,  // logical name, e.g. 'account'
 *   entityId:    string | null,  // bare GUID, lowercase, no braces
 *   selectedIds: string[],       // GUIDs of selected records
 * }
 */

(function dynamicsAuditLensBridge() {
  "use strict";

  // ── Message type constants ───────────────────────────────────────────────
  const T_READY = "__DAL__BRIDGE_READY";
  const T_REQUEST = "__DAL__CONTEXT_REQUEST";
  const T_RESPONSE = "__DAL__CONTEXT_RESPONSE";

  // Target origin for outbound postMessage — always the page's own origin.
  // Dynamics 365 is always HTTPS so this is well-defined.
  const TARGET_ORIGIN = window.location.origin;

  // ── Helpers ──────────────────────────────────────────────────────────────

  /** Strip curly braces and lowercase a raw Dynamics GUID string. */
  function normaliseGuid(raw) {
    if (!raw) return null;
    return raw.replace(/[{}]/g, "").toLowerCase();
  }

  /** Validate a string is a well-formed GUID. Rejects attacker-controlled junk. */
  const GUID_RE =
    /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  function isGuid(s) {
    return typeof s === "string" && GUID_RE.test(s);
  }

  // ── Xrm context readers ──────────────────────────────────────────────────

  /**
   * UCI (Unified Client Interface) path.
   * Xrm.Utility.getPageContext() is the officially documented API.
   * Returns null if the API is unavailable.
   */
  function readUciContext() {
    try {
      const input = window.Xrm?.Utility?.getPageContext?.()?.input;
      if (!input) return null;

      const entityId = normaliseGuid(input.entityId);
      return {
        pageType: String(input.pageType ?? "unknown").slice(0, 64),
        entityName: input.entityName
          ? String(input.entityName).slice(0, 128)
          : null,
        entityId: isGuid(entityId) ? entityId : null,
      };
    } catch {
      return null;
    }
  }

  /**
   * Legacy (web client / early UCI) path via Xrm.Page.
   * Still present in many on-premise / older online orgs.
   */
  function readLegacyFormContext() {
    try {
      const entity = window.Xrm?.Page?.data?.entity;
      if (!entity) return null;

      const rawId = entity.getId?.();
      const rawName = entity.getEntityName?.();
      if (!rawId && !rawName) return null;

      const entityId = normaliseGuid(rawId);
      return {
        pageType: "entityrecord",
        entityName: rawName ? String(rawName).slice(0, 128) : null,
        entityId: isGuid(entityId) ? entityId : null,
      };
    } catch {
      return null;
    }
  }

  /**
   * Read selected row GUIDs from the UCI list view.
   *
   * The official Xrm API does not expose selected main-grid rows outside of a
   * ribbon command context. We therefore read the `data-id` / `row-id` ARIA
   * attributes on rows marked `aria-selected="true"`. These attributes are part
   * of Dynamics' accessibility contract and are more stable than internal CSS
   * class names (which are what the user warned against).
   *
   * Every extracted value is validated as a well-formed GUID before it is
   * returned, so attacker-controlled DOM content cannot inject arbitrary strings.
   *
   * @returns {string[]}  Array of normalised GUIDs.
   */
  function readSelectedGridIds() {
    try {
      // UCI renders the main grid as an ag-grid; rows carry row-id or data-id.
      const candidates = document.querySelectorAll(
        '[aria-selected="true"][data-id], [aria-selected="true"][row-id]',
      );
      const ids = [];
      candidates.forEach((el) => {
        const raw =
          el.getAttribute("data-id") || el.getAttribute("row-id") || "";
        const guid = normaliseGuid(raw);
        if (isGuid(guid) && !ids.includes(guid)) {
          ids.push(guid);
        }
      });
      return ids;
    } catch {
      return [];
    }
  }

  /**
   * Try Xrm subgrid controls for selected rows (subgrids on a form).
   * @returns {string[]}
   */
  function readSubgridSelectedIds() {
    try {
      const controls = window.Xrm?.Page?.controls?.get?.() ?? [];
      const ids = [];
      controls.forEach((ctrl) => {
        const type = ctrl.getControlType?.();
        if (type !== "subgrid") return;
        const grid = ctrl.getGrid?.();
        const rows = grid?.getSelectedRows?.();
        rows?.getAll?.().forEach((row) => {
          const raw = row.getData?.()?.entity?.getId?.();
          const guid = normaliseGuid(raw);
          if (isGuid(guid) && !ids.includes(guid)) {
            ids.push(guid);
          }
        });
      });
      return ids;
    } catch {
      return [];
    }
  }

  // ── Main context collector ───────────────────────────────────────────────

  function collectContext() {
    if (!window.Xrm) {
      return {
        available: false,
        pageType: null,
        entityName: null,
        entityId: null,
        selectedIds: [],
      };
    }

    // Prefer UCI API, fall back to legacy
    const base = readUciContext() ?? readLegacyFormContext();

    if (!base) {
      return {
        available: true,
        pageType: "unknown",
        entityName: null,
        entityId: null,
        selectedIds: [],
      };
    }

    let selectedIds = [];

    if (base.pageType === "entityrecord") {
      // Form — selected "record" is the open record itself, plus any subgrid selections
      if (base.entityId) selectedIds = [base.entityId];
      const sub = readSubgridSelectedIds();
      sub.forEach((id) => {
        if (!selectedIds.includes(id)) selectedIds.push(id);
      });
    } else if (base.pageType === "entitylist") {
      // List / grid view — collect ARIA-selected row IDs
      selectedIds = readSelectedGridIds();
    }

    // Flag when we're on a list page but couldn't detect any selection method.
    // This helps the popup warn the user instead of silently showing "0 selected".
    const selectionUnavailable =
      base.pageType === "entitylist" &&
      selectedIds.length === 0 &&
      document.querySelectorAll('[aria-selected="true"]').length > 0;

    return {
      available: true,
      pageType: base.pageType,
      entityName: base.entityName,
      entityId: base.entityId,
      selectedIds,
      selectionUnavailable,
    };
  }

  // ── Message handler ──────────────────────────────────────────────────────

  window.addEventListener("message", function handleBridgeRequest(event) {
    // Only accept messages originating from the same window.
    // (Cross-origin iframes will have a different event.source.)
    if (event.source !== window) return;
    if (event.data?.type !== T_REQUEST) return;

    const payload = collectContext();
    window.postMessage({ type: T_RESPONSE, payload }, TARGET_ORIGIN);
  });

  // ── Announce readiness ───────────────────────────────────────────────────
  // Send an initial snapshot immediately so the content script can cache it
  // without waiting for an explicit request.
  window.postMessage(
    { type: T_READY, payload: collectContext() },
    TARGET_ORIGIN,
  );
})();


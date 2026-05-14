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

  /**
   * Detect whether Select All is active on the UCI grid.
   *
   * Why this is tricky: Dynamics' ag-grid virtualises rows — only the rows
   * inside the viewport (~25-30) exist in the DOM at any time. Comparing the
   * count of aria-selected="true" rows against the count of rendered rows is
   * unreliable because rows can be in transient render states (and the user
   * scrolling does not re-mark off-screen rows). We therefore detect via
   * multiple independent heuristics and treat ANY of them as proof of
   * Select-All:
   *
   *   1. The header column's checkbox is checked (most authoritative).
   *   2. A page banner explicitly says "N records selected" or
   *      "N of M selected" with N > number of currently-selected DOM rows.
   *   3. All rendered data rows are aria-selected="true" (the original signal).
   *
   * When active, we also try to read the total record count from pagination
   * text ("1-23 of 90") or selected-items banner text so the popup can show
   * the right number even before the API fallback resolves the full GUID set.
   *
   * @returns {{ active: boolean, totalRecords: number|null }}
   */
  function readGridSelectAllInfo() {
    try {
      var selected = document.querySelectorAll(
        '[aria-selected="true"][data-id], [aria-selected="true"][row-id]'
      );
      var allSelectable = document.querySelectorAll(
        "[aria-selected][data-id], [aria-selected][row-id]"
      );

      // ── Heuristic 1: header checkbox is checked ──────────────────────────
      // The grid header has a checkbox the user clicks to select-all.
      // Match by role+aria-checked, by aria-label, and by the legacy
      // data-id="header" row containing a checked checkbox.
      var headerChecked = false;
      try {
        var checkBoxes = document.querySelectorAll(
          '[role="checkbox"][aria-checked="true"]'
        );
        for (var c = 0; c < checkBoxes.length; c++) {
          var cb = checkBoxes[c];
          var label = (cb.getAttribute("aria-label") || "").toLowerCase();
          if (
            label.indexOf("select all") !== -1 ||
            label.indexOf("all rows") !== -1 ||
            label.indexOf("toggle selection of all") !== -1
          ) {
            headerChecked = true;
            break;
          }
          // Also: a checked checkbox sitting inside the header row.
          var hostRow = cb.closest(
            '[role="row"][data-id="header"], [role="row"][aria-rowindex="1"], [role="columnheader"]'
          );
          if (hostRow) {
            headerChecked = true;
            break;
          }
        }
        if (!headerChecked) {
          var input = document.querySelector(
            'input[type="checkbox"][aria-label*="elect all" i]:checked'
          );
          if (input) headerChecked = true;
        }
      } catch {
        // ignore — fall through to other heuristics
      }

      // ── Heuristic 2: "N records selected" / "N of M selected" banner ─────
      var bannerTotal = null;
      var bannerSelectedCount = null;
      var spans = document.querySelectorAll("span, div");
      for (var s = 0; s < spans.length; s++) {
        var txt = (spans[s].textContent ?? "").trim();
        if (!txt || txt.length > 80) continue;
        // "N of M selected" form — picks both numbers.
        var mOf = txt.match(
          /(\d+)\s+of\s+(\d+)\s+(?:records?|items?|rows?)?\s*(?:are\s+)?selected/i
        );
        if (mOf) {
          var nSel = parseInt(mOf[1], 10);
          var nTot = parseInt(mOf[2], 10);
          if (Number.isFinite(nSel) && Number.isFinite(nTot)) {
            bannerSelectedCount = Math.max(bannerSelectedCount ?? 0, nSel);
            bannerTotal = Math.max(bannerTotal ?? 0, nTot);
          }
          continue;
        }
        // "N records selected" form.
        var mN = txt.match(
          /(\d+)\s+(?:records?|items?|rows?)\s+(?:are\s+)?selected/i
        );
        if (mN) {
          var n = parseInt(mN[1], 10);
          if (Number.isFinite(n) && n > 0) {
            bannerSelectedCount = Math.max(bannerSelectedCount ?? 0, n);
          }
        }
      }
      var bannerSaysSelectAll =
        bannerSelectedCount !== null &&
        bannerSelectedCount > selected.length;

      // ── Heuristic 3: every rendered row aria-selected="true" (legacy) ────
      var allRenderedSelected =
        selected.length > 0 &&
        allSelectable.length > 0 &&
        selected.length >= allSelectable.length;

      var active = headerChecked || bannerSaysSelectAll || allRenderedSelected;

      if (!active) {
        return { active: false, totalRecords: null };
      }

      // ── Resolve totalRecords ─────────────────────────────────────────────
      var totalRecords = bannerTotal;

      // "1-23 of 90" style pagination text.
      if (totalRecords === null) {
        for (var i = 0; i < spans.length; i++) {
          var t = (spans[i].textContent ?? "").trim();
          if (!t || t.length > 80) continue;
          var m = t.match(/(?:\d+\s*[-–]\s*\d+\s+)?of\s+(\d+)/i);
          if (m) {
            var nn = parseInt(m[1], 10);
            if (Number.isFinite(nn) && nn > selected.length) {
              totalRecords = nn;
              break;
            }
          }
        }
      }

      // Banner-only "N selected" — use it as the floor.
      if (totalRecords === null && bannerSelectedCount !== null) {
        totalRecords = bannerSelectedCount;
      }

      return { active: true, totalRecords: totalRecords };
    } catch {
      return { active: false, totalRecords: null };
    }
  }

  /**
   * Extract the current view ID and type.
   *
   * Primary source: Xrm.Utility.getPageContext().input (most reliable).
   * Fallback:      URL query-string / hash parameters.
   *
   * @returns {{ viewId: string, viewType: string }|null}
   */
  function getViewInfo() {
    try {
      // --- Try Xrm API first (works even with hash-based routing) ---
      var input = window.Xrm?.Utility?.getPageContext?.()?.input;
      if (input?.viewId) {
        var xrmId = normaliseGuid(String(input.viewId));
        if (isGuid(xrmId)) {
          var xrmVt = String(input.viewType ?? "");
          return {
            viewId: xrmId,
            viewType: xrmVt === "4230" ? "userquery" : "savedquery",
          };
        }
      }

      // --- Fallback: parse URL parameters (case-insensitive on key names) ---
      // Different Dynamics versions/links emit viewid / viewId / ViewId etc.
      var paramMap = {};
      function ingestParams(usp) {
        usp.forEach(function (v, k) {
          var lk = k.toLowerCase();
          if (paramMap[lk] === undefined) paramMap[lk] = v;
        });
      }
      ingestParams(new URLSearchParams(window.location.search));
      var hash = window.location.hash;
      if (hash) {
        var h = hash.replace(/^#\/?/, "");
        var hashQuery = h.indexOf("?") !== -1 ? h : "?" + h;
        ingestParams(new URLSearchParams(hashQuery));
      }
      var rawId = paramMap["viewid"];
      if (!rawId) return null;
      var id = normaliseGuid(rawId);
      if (!isGuid(id)) return null;
      var vt = paramMap["viewtype"];
      return { viewId: id, viewType: vt === "4230" ? "userquery" : "savedquery" };
    } catch {
      return null;
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
    let selectAllActive = false;
    let totalRecordCount = null;
    let viewId = null;
    let viewType = null;

    if (base.pageType === "entityrecord") {
      // Form — selected "record" is the open record itself, plus any subgrid selections
      if (base.entityId) selectedIds = [base.entityId];
      const sub = readSubgridSelectedIds();
      sub.forEach((id) => {
        if (!selectedIds.includes(id)) selectedIds.push(id);
      });
    } else if (base.pageType === "entitylist") {
      selectedIds = readSelectedGridIds();

      var selectAllInfo = readGridSelectAllInfo();
      selectAllActive = selectAllInfo.active;
      totalRecordCount = selectAllInfo.active ? selectAllInfo.totalRecords : null;

      var viewInfo = getViewInfo();
      if (viewInfo) {
        viewId = viewInfo.viewId;
        viewType = viewInfo.viewType;
      }

      console.log("[Audit Lens] entitylist context:", {
        selectedIds: selectedIds.length,
        selectAllActive: selectAllActive,
        totalRecordCount: totalRecordCount,
        viewId: viewId,
        viewType: viewType,
        allSelectableRows: document.querySelectorAll("[aria-selected][data-id], [aria-selected][row-id]").length,
        selectedRows: document.querySelectorAll('[aria-selected="true"][data-id], [aria-selected="true"][row-id]').length,
        xrmInput: window.Xrm?.Utility?.getPageContext?.()?.input,
      });
    }

    // Flag when we're on a list page but couldn't detect any selection method.
    // This helps the popup warn the user instead of silently showing "0 selected".
    const selectionUnavailable =
      base.pageType === "entitylist" &&
      selectedIds.length === 0 &&
      !selectAllActive &&
      document.querySelectorAll('[aria-selected="true"]').length > 0;

    return {
      available: true,
      pageType: base.pageType,
      entityName: base.entityName,
      entityId: base.entityId,
      selectedIds,
      selectionUnavailable,
      selectAllActive,
      totalRecordCount,
      viewId,
      viewType,
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


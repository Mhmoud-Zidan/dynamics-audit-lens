/**
 * content.js — Isolated-world content script.
 *
 * ARCHITECTURE OVERVIEW
 * ─────────────────────
 *
 *   ┌──────────────────────────────────────────────────────────┐
 *   │  Chrome Extension                                        │
 *   │                                                          │
 *   │   popup.js ──chrome.tabs.sendMessage──► content.js       │
 *   │                       ◄──sendResponse──                  │
 *   │                                  │  ▲                    │
 *   │                    postMessage   │  │ postMessage        │
 *   │              ┌───────────────────┘  │                    │
 *   │              ▼                      │                    │
 *   │   ┌──────── PAGE CONTEXT ─────────────────────────────┐  │
 *   │   │  page-bridge.js  (window.Xrm accessible here) ✓   │  │
 *   │   └───────────────────────────────────────────────────┘  │
 *   └──────────────────────────────────────────────────────────┘
 *
 * SECURITY MODEL
 * ──────────────
 *  • The bridge script is loaded via chrome.runtime.getURL() — Chrome owns
 *    the URL; the host page cannot forge or replace it.
 *  • All postMessages from the bridge are validated:
 *      - event.source === window  (same window, rules out cross-origin iframes)
 *      - event.origin === EXPECTED_ORIGIN  (same HTTP origin, rules out spoofing)
 *  • GUID values are treated as opaque display strings — never eval'd or used
 *    in innerHTML.
 *  • No data is sent to any external server.
 */

"use strict";

// ── Guard against double injection ────────────────────────────────────────────
// Uses a module-scoped closure variable instead of a window property to prevent
// a hostile page script from pre-setting the flag and silently disabling us.
{
  const key = Symbol.for("__dalContentV1");
  if (globalThis[key]) {
    throw new Error("DAL: content script already active");
  }
  Object.defineProperty(globalThis, key, {
    value: true,
    writable: false,
    configurable: false,
    enumerable: false,
  });
}

// ── Message type constants ────────────────────────────────────────────────────
const T_READY = "__DAL__BRIDGE_READY";
const T_REQUEST = "__DAL__CONTEXT_REQUEST";
const T_RESPONSE = "__DAL__CONTEXT_RESPONSE";
const T_FILL_DATA_RESPONSE = "__DAL__FILL_DATA_RESPONSE";

/**
 * The origin we expect on all postMessages from the bridge.
 * Dynamics 365 is always HTTPS, so origin is always well-defined.
 * If somehow null (sandboxed frame), we reject all messages.
 */
const EXPECTED_ORIGIN = (() => {
  const o = window.location.origin;
  return o && o !== "null" ? o : null;
})();

// ── State ─────────────────────────────────────────────────────────────────────
/** Latest context snapshot received from the bridge. */
let cachedContext = null;

/** True once the <script> element has been appended. */
let bridgeInjected = false;

/** Pending fill-data bridge response. { resolve, timer } */
let pendingFillRequest = null;

// ── Bridge injection ──────────────────────────────────────────────────────────

/**
 * Inject page-bridge.js into the page's main world by appending a
 * <script src="chrome.runtime.getURL(...)"> element.
 *
 * Because the URL is a chrome-extension:// URL listed in
 * web_accessible_resources, Chrome permits it regardless of the page's own CSP.
 * The element is removed from the DOM immediately after load — the script
 * has already executed by that point.
 */
function injectBridge() {
  if (bridgeInjected) return;
  bridgeInjected = true;

  try {
    const script = document.createElement("script");
    script.src = chrome.runtime.getURL("src/inject/page-bridge.js");
    script.type = "text/javascript";

    const parent = document.head ?? document.documentElement;
    parent.appendChild(script);

    script.addEventListener("load", () => script.remove(), { once: true });
    script.addEventListener(
      "error",
      () => {
        console.warn(
          "[Audit Lens] Bridge script blocked by page CSP. " +
            "Xrm context detection is unavailable on this page.",
        );
        script.remove();
      },
      { once: true },
    );
  } catch (err) {
    console.error("[Audit Lens] Bridge injection failed:", err);
  }
}

// ── Pending request queue ────────────────────────────────────────────────────
/**
 * Each entry represents one popup waiting for a fresh GET_CONTEXT reply.
 * Shape: { resolve: Function, timer: number }
 */
const pendingRequests = [];

// ── postMessage listener ──────────────────────────────────────────────────────

window.addEventListener("message", function onBridgeMessage(event) {
  // Reject messages from other windows or origins up-front.
  if (event.source !== window) return;
  if (!EXPECTED_ORIGIN || event.origin !== EXPECTED_ORIGIN) return;

  const { type, payload } = event.data ?? {};

  if (type === T_READY && payload) {
    // Bridge has loaded and sent an initial snapshot.
    cachedContext = payload;
    notifyBackground(payload);
    return;
  }

  if (type === T_RESPONSE && payload) {
    // Bridge replied to an explicit request — update cache and unblock callers.
    cachedContext = payload;
    while (pendingRequests.length) {
      const { resolve, timer } = pendingRequests.shift();
      clearTimeout(timer);
      resolve(payload);
    }
  }

  if (type === T_FILL_DATA_RESPONSE && payload) {
    if (pendingFillRequest) {
      clearTimeout(pendingFillRequest.timer);
      pendingFillRequest.resolve(payload);
      pendingFillRequest = null;
    }
  }
});

// ── Background notification ───────────────────────────────────────────────────

function notifyBackground(context) {
  chrome.runtime
    .sendMessage({
      type: "DYNAMICS_CONTEXT_UPDATE",
      payload: {
        hostname: window.location.hostname,
        pathname: window.location.pathname,
        title: document.title.slice(0, 200),
        context,
      },
    })
    .catch(() => {
      /* extension context may be invalidated; ignore */
    });
}

// ── Context request helper ────────────────────────────────────────────────────

/**
 * Ask the bridge for a fresh context reading.
 * Resolves with the bridge's response, or after 2 s falls back to cachedContext.
 *
 * @returns {Promise<object>}
 */
function requestFreshContext() {
  return new Promise((resolve) => {
    const FALLBACK_CONTEXT = {
      available: false,
      pageType: null,
      entityName: null,
      entityId: null,
      selectedIds: [],
    };

    const timer = setTimeout(() => {
      const idx = pendingRequests.findIndex((r) => r.resolve === resolve);
      if (idx !== -1) pendingRequests.splice(idx, 1);
      resolve(cachedContext ?? FALLBACK_CONTEXT);
    }, 2000);

    pendingRequests.push({ resolve, timer });
    window.postMessage({ type: T_REQUEST }, EXPECTED_ORIGIN);
  });
}

// ═════════════════════════════════════════════════════════════════════════════
// DATA EXTRACTION ENGINE
// ═════════════════════════════════════════════════════════════════════════════
//
// Calls the Dataverse Web API from within the content script.
// Because the script runs in the same browser context as the Dynamics page,
// the user's active MSCRM session cookies are sent automatically by the browser
// on same-origin requests — no Authorization header is needed or added.
//
// API reference:
//   GET [orgUri]/api/data/v9.2/[entitySetName]([guid])
//        /Microsoft.Dynamics.CRM.RetrieveRecordChangeHistory()
//
// Service Protection API limits (per user per server):
//   • 6 000 requests / 5 min
//   • 20 min combined execution time / 5 min
//   • 52 concurrent requests
// We stay well inside those limits with a hard cap of MAX_CONCURRENT = 5.
// ═════════════════════════════════════════════════════════════════════════════

// ── Constants ─────────────────────────────────────────────────────────────────

const API_VERSION = "9.2";
const MAX_CONCURRENT = 5;
const MAX_EXPORT_ROWS = 100_000;
const MAX_USER_AUDIT_RECORDS = 500;
const MAX_AUDIT_QUERY_PAGES = 20;

/**
 * Standard OData headers required by the Dataverse REST endpoint.
 * Prefer `application/json` over `application/atom+xml`.
 * `OData-MaxVersion` pins the protocol so future server upgrades don't break parsing.
 */
const ODATA_HEADERS = Object.freeze({
  Accept: "application/json; odata.metadata=minimal",
  "OData-MaxVersion": "4.0",
  "OData-Version": "4.0",
  "Content-Type": "application/json; charset=utf-8",
  Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
});

// ── GUID validation ───────────────────────────────────────────────────────────

const GUID_PATTERN =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * Throw if the supplied value is not a well-formed GUID string.
 * Used to sanitise caller-supplied record IDs before they are interpolated
 * into URL paths, eliminating path-traversal / injection risk.
 *
 * @param {string} guid
 * @throws {TypeError}
 */
function assertGuid(guid) {
  if (typeof guid !== "string" || !GUID_PATTERN.test(guid)) {
    throw new TypeError(`[Audit Lens] Invalid GUID: "${guid}"`);
  }
}

/**
 * Validate an entity set name.
 * Dataverse entity set names are alphanumeric + underscore, 1–128 chars.
 * Rejecting anything outside this set prevents URL injection.
 *
 * @param {string} name
 * @throws {TypeError}
 */
function assertEntitySetName(name) {
  if (typeof name !== "string" || !/^[A-Za-z][A-Za-z0-9_]{0,127}$/.test(name)) {
    throw new TypeError(`[Audit Lens] Invalid entity set name: "${name}"`);
  }
}

// ── Promise concurrency pool ──────────────────────────────────────────────────

/**
 * Run an array of async task-factories with at most `limit` running in parallel.
 *
 * Unlike Promise.all() (which fires everything at once) this pool keeps exactly
 * `limit` tasks in-flight at any moment, picking the next task as soon as a
 * slot becomes free. This is the standard "promise pool" / "worker pool" pattern.
 *
 * Each factory in `tasks` is a zero-argument function that returns a Promise.
 * Results are returned in the same order as `tasks` (stable ordering).
 *
 * @template T
 * @param {Array<() => Promise<T>>} tasks  Array of task factories.
 * @param {number}                  limit  Max concurrency (default MAX_CONCURRENT).
 * @returns {Promise<T[]>}
 */
async function runPool(tasks, limit = MAX_CONCURRENT) {
  const results = new Array(tasks.length);
  let nextIdx = 0;

  async function worker() {
    while (nextIdx < tasks.length) {
      const idx = nextIdx++; // claim a task slot atomically within the microtask
      results[idx] = await tasks[idx](); // may throw; propagated to caller via Promise.all
    }
  }

  // Spin up exactly `limit` worker coroutines (or fewer if there aren't enough tasks).
  const workers = Array.from({ length: Math.min(limit, tasks.length) }, () =>
    worker(),
  );

  await Promise.all(workers);
  return results;
}

// ── Retry / back-off helper ───────────────────────────────────────────────────

/**
 * Retry a failing async operation with exponential back-off.
 *
 * Handles HTTP 429 (Service Protection limit) and 503 (transient server error).
 * On a 429 response the server MAY supply a `Retry-After` header (seconds);
 * we honour it when present, otherwise fall back to the computed back-off.
 *
 * @param {() => Promise<Response>} fetchFn   Zero-arg fetch factory.
 * @param {number}                  maxRetries Max number of re-attempts (default 3).
 * @returns {Promise<Response>}
 * @throws  When all retries are exhausted.
 */
async function fetchWithRetry(fetchFn, maxRetries = 3) {
  let lastError;
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    let response;
    try {
      response = await fetchFn();
    } catch (networkErr) {
      // Network-level error (offline, DNS failure, CORS block, etc.)
      lastError = networkErr;
      if (attempt === maxRetries) break;
      await sleep(backoffMs(attempt));
      continue;
    }

    // Success path
    if (response.ok) return response;

    // Retryable server-side errors
    if (response.status === 429 || response.status === 503) {
      if (attempt === maxRetries) {
        lastError = new ApiError(
          response.status,
          "Rate-limited or service unavailable",
          response,
        );
        break;
      }
      const retryAfter = Number(response.headers.get("Retry-After") ?? 0);
      const delay = retryAfter > 0 ? retryAfter * 1000 : backoffMs(attempt);
      console.warn(
        `[Audit Lens] HTTP ${response.status} — retrying in ${delay}ms (attempt ${attempt + 1}/${maxRetries})`,
      );
      await sleep(delay);
      continue;
    }

    // Non-retryable HTTP error (400, 401, 403, 404, 500, …)
    lastError = new ApiError(response.status, response.statusText, response);
    break;
  }
  throw lastError;
}

/** Exponential back-off: 1 s, 2 s, 4 s, … capped at 30 s. */
function backoffMs(attempt) {
  return Math.min(1000 * 2 ** attempt, 30_000);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// ── Custom error type ─────────────────────────────────────────────────────────

class ApiError extends Error {
  /**
   * @param {number}   status      HTTP status code
   * @param {string}   message
   * @param {Response} response    The raw fetch Response (may be consumed)
   */
  constructor(status, message, response) {
    super(`[Audit Lens] API ${status}: ${message}`);
    this.name = "ApiError";
    this.status = status;
    this.response = response;
  }
}

// ── Organisation URI helper ───────────────────────────────────────────────────

/**
 * Derive the Dataverse Organisation Root URI from the current page URL.
 *
 * Dynamics 365 Online:  https://contoso.crm.dynamics.com
 * On-premise / custom:  Same assumption — the Web API lives at the origin root.
 *
 * Returns a string with NO trailing slash.
 *
 * @returns {string}
 */
function getOrgUri() {
  return window.location.origin; // e.g. "https://contoso.crm.dynamics.com"
}

function pushRowsSafe(allRows, rows, maxRows, onCap) {
  if (allRows.length + rows.length <= maxRows) {
    for (let i = 0; i < rows.length; i++) allRows.push(rows[i]);
  } else {
    const remaining = maxRows - allRows.length;
    for (let i = 0; i < remaining; i++) allRows.push(rows[i]);
    onCap();
  }
}

// ── Single-record audit fetch ─────────────────────────────────────────────────

/**
 * Fetch the full audit history for one record.
 *
 * Endpoint (bound function):
 *   GET [orgUri]/api/data/v9.2/[entitySetName]([guid])
 *        /Microsoft.Dynamics.CRM.RetrieveRecordChangeHistory()
 *
 * The `@ChangedRecords` / `AuditDetailCollection` response structure varies
 * slightly by Dataverse version; we return the raw parsed JSON so the caller
 * can handle schema differences without re-fetching.
 *
 * @param {string} entitySetName  Plural entity set name, e.g. "accounts"
 * @param {string} guid           Record ID (bare, lowercase, no braces)
 * @returns {Promise<{ guid: string, entitySetName: string, data: object }>}
 * @throws  {ApiError | TypeError}
 */
async function fetchAuditHistoryForRecord(entitySetName, guid) {
  // Validate inputs before they reach the network layer.
  assertEntitySetName(entitySetName);
  assertGuid(guid);

  // Use the UNBOUND RetrieveRecordChangeHistory function with a Target parameter
  // alias. The bound syntax (/{entitySetName}({guid})/Microsoft.Dynamics.CRM.RetrieveRecordChangeHistory())
  // returns 404 on many Dataverse versions / configurations. The unbound form
  // works consistently across all versions.
  const target = `${entitySetName}(${guid})`;
  const baseUrl =
    `${getOrgUri()}/api/data/v${API_VERSION}` +
    `/RetrieveRecordChangeHistory(Target=@target)` +
    `?@target={'@odata.id':'${target}'}`;

  // Follow @odata.nextLink pagination — RetrieveRecordChangeHistory can return
  // paginated results when audit history exceeds the server page size (~5 000).
  let allDetails = [];
  let url = baseUrl;
  const MAX_PAGES = 50; // safety cap to prevent infinite loops

  for (let page = 0; page < MAX_PAGES; page++) {
    const response = await fetchWithRetry(() =>
      fetch(url, {
        method: "GET",
        credentials: "include", // send the MSCRM session cookie
        headers: ODATA_HEADERS,
      }),
    );

    const json = await response.json();
    const details = json?.AuditDetailCollection?.AuditDetails ?? [];
    allDetails.push(...details);

    const nextLink = json["@odata.nextLink"];
    if (!nextLink) break;

    // Validate nextLink is same-origin to prevent SSRF via tampered response.
    try {
      const parsedNext = new URL(nextLink);
      if (parsedNext.origin !== window.location.origin) break;
      url = nextLink;
    } catch {
      break;
    }
  }

  // Reconstruct the expected response shape with all pages merged.
  const data = { AuditDetailCollection: { AuditDetails: allDetails } };
  return { guid, entitySetName, data };
}

// ═════════════════════════════════════════════════════════════════════════════
// METADATA RESOLUTION ENGINE
// ═════════════════════════════════════════════════════════════════════════════
//
// Fetches entity attribute metadata from the Dataverse Metadata API and caches
// it in-memory per entity logical name for the lifetime of the tab.
//
// Responsibilities:
//   • Map attribute logical names → human-readable Display Names.
//   • Map integer OptionSet / State / Status / Boolean values → label strings.
//   • Decode the `attributemask` field on AuditRecord into a list of changed
//     attribute logical names (cross-referenced against ColumnNumber ordering).
//   • Format raw RetrieveRecordChangeHistory JSON into clean row objects.
// ═════════════════════════════════════════════════════════════════════════════

// ── Types (JSDoc only) ────────────────────────────────────────────────────────

/**
 * @typedef {object} AttrMeta
 * @property {string}                displayName  Human-readable field label.
 * @property {string}                type         Dataverse AttributeType string.
 * @property {Map<number,string>|null} options    Integer → label for enum types.
 */

/**
 * @typedef {object} EntityMeta
 * @property {string}              primaryId   Logical name of the PK attribute.
 * @property {string|null}          primaryName Logical name of the primary name attribute.
 * @property {string|null}          entitySetName Plural entity set name for API URLs.
 * @property {number|null}          objectTypeCode Integer type code used in audit filters.
 * @property {Map<string,AttrMeta>} attributes logicalName → AttrMeta.
 * @property {Map<number,string>}   byColumn   ColumnNumber → logicalName (for attributemask).
 */

/**
 * @typedef {object} FormattedAuditRow
 * @property {string} RecordID    GUID of the audited record.
 * @property {string} RecordName  Primary name of the audited record.
 * @property {string} ChangedBy   Display name of the user who made the change.
 * @property {Date}   ChangedDate Date of the change.
 * @property {string} Operation   Human-readable operation (e.g. "Update").
 * @property {string} FieldName   Display name of the changed attribute.
 * @property {string} OldValue    Resolved human-readable old value.
 * @property {string} NewValue    Resolved human-readable new value.
 */

// ── In-memory cache ───────────────────────────────────────────────────────────

/** entityLogicalName → EntityMeta */
const metadataCache = new Map();

// ── Validation ────────────────────────────────────────────────────────────────

/**
 * Validate a Dataverse entity logical name.
 * These are always lowercase alphanumeric + underscore (publisher prefix + name).
 *
 * @param {string} name
 * @throws {TypeError}
 */
function assertEntityLogicalName(name) {
  if (typeof name !== "string" || !/^[a-z][a-z0-9_]{0,127}$/.test(name)) {
    throw new TypeError(`[Audit Lens] Invalid entity logical name: "${name}"`);
  }
}

// ── Metadata fetcher ──────────────────────────────────────────────────────────

/**
 * Fetch (and cache) entity attribute metadata from the Dataverse Metadata API.
 *
 * Dataverse forbids multi-level $expand (expanding OptionSet inside Attributes
 * causes 0x80060888 or "Multiple levels of expansion aren't supported").
 * The workaround is two parallel request groups:
 *
 *   Request 1 — base attribute list (no nested expand):
 *     GET EntityDefinitions(LogicalName='<entity>')
 *       ?$select=LogicalName,PrimaryIdAttribute,EntitySetName
 *       &$expand=Attributes($select=LogicalName,DisplayName,AttributeType,ColumnNumber)
 *
 *   Requests 2-5 (parallel) — typed path-cast requests; each supports one level
 *   of $expand because OptionSet IS a property of the cast type's own schema:
 *     GET EntityDefinitions(LogicalName='<entity>')/Attributes
 *           /Microsoft.Dynamics.CRM.PicklistAttributeMetadata
 *           ?$select=LogicalName&$expand=OptionSet($select=Options)
 *     (…same pattern for MultiSelectPicklist, Status, State, Boolean)
 *
 * @param {string} entityLogicalName  e.g. "account"
 * @returns {Promise<EntityMeta>}
 */
async function fetchEntityMetadata(entityLogicalName) {
  assertEntityLogicalName(entityLogicalName);

  if (metadataCache.has(entityLogicalName)) {
    return metadataCache.get(entityLogicalName);
  }

  const baseUrl = `${getOrgUri()}/api/data/v${API_VERSION}`;
  const entityBase =
    `${baseUrl}/EntityDefinitions(LogicalName='${entityLogicalName}')`;

  // Helper: fire a GET and return parsed JSON, falling back to {} on error
  // so a missing cast type (e.g. entity has no picklist attrs) doesn't abort.
  const getJson = (url) =>
    fetchWithRetry(() =>
      fetch(url, { method: "GET", credentials: "include", headers: ODATA_HEADERS }),
    )
      .then((r) => r.json())
      .catch(() => ({}));

  // Typed attribute requests — each is a flat, single-level expand that
  // Dataverse accepts without error.
  const OPTION_TYPES = [
    {
      cast: "Microsoft.Dynamics.CRM.PicklistAttributeMetadata",
      expandSelect: "Options",
      isBool: false,
    },
    {
      cast: "Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata",
      expandSelect: "Options",
      isBool: false,
    },
    {
      cast: "Microsoft.Dynamics.CRM.StatusAttributeMetadata",
      expandSelect: "Options",
      isBool: false,
    },
    {
      cast: "Microsoft.Dynamics.CRM.StateAttributeMetadata",
      expandSelect: "Options",
      isBool: false,
    },
    {
      cast: "Microsoft.Dynamics.CRM.BooleanAttributeMetadata",
      expandSelect: "TrueOption,FalseOption",
      isBool: true,
    },
  ];

  const attrUrl =
    `${entityBase}` +
    `?$select=LogicalName,PrimaryIdAttribute,PrimaryNameAttribute,EntitySetName,ObjectTypeCode` +
    `&$expand=Attributes($select=LogicalName,DisplayName,AttributeType,ColumnNumber)`;

  // Fire all requests in parallel.
  const [entityJson, ...optionResults] = await Promise.all([
    getJson(attrUrl),
    ...OPTION_TYPES.map(({ cast, expandSelect }) =>
      getJson(
        `${entityBase}/Attributes/${cast}` +
          `?$select=LogicalName&$expand=OptionSet($select=${expandSelect})`,
      ),
    ),
  ]);

  // Build logicalName → { optionSet, isBool } from the typed results.
  /** @type {Map<string, { optionSet: object, isBool: boolean }>} */
  const optionSetByName = new Map();
  optionResults.forEach((result, i) => {
    const { isBool } = OPTION_TYPES[i];
    for (const attr of result.value ?? []) {
      if (typeof attr.LogicalName === "string" && attr.OptionSet) {
        optionSetByName.set(attr.LogicalName, { optionSet: attr.OptionSet, isBool });
      }
    }
  });

  const attributes = new Map(); // logicalName → AttrMeta
  const byColumn = new Map(); // ColumnNumber → logicalName

  for (const attr of entityJson.Attributes ?? []) {
    const logicalName = attr.LogicalName;
    if (typeof logicalName !== "string") continue;

    const displayName =
      attr.DisplayName?.UserLocalizedLabel?.Label ?? logicalName;
    const type = String(attr.AttributeType ?? "Unknown");

    // ── Build integer → label map for option-set backed types ──────────────
    let options = null;
    const optData = optionSetByName.get(logicalName);
    if (optData) {
      if (optData.isBool) {
        const trueLabel =
          optData.optionSet.TrueOption?.Label?.UserLocalizedLabel?.Label ?? "Yes";
        const falseLabel =
          optData.optionSet.FalseOption?.Label?.UserLocalizedLabel?.Label ?? "No";
        options = new Map([
          [1, trueLabel],
          [0, falseLabel],
        ]);
      } else if (Array.isArray(optData.optionSet.Options)) {
        options = new Map();
        for (const opt of optData.optionSet.Options) {
          const label = opt.Label?.UserLocalizedLabel?.Label;
          if (label !== undefined && opt.Value !== undefined) {
            options.set(Number(opt.Value), String(label));
          }
        }
      }
    }

    /** @type {AttrMeta} */
    const meta = { displayName, type, options };
    attributes.set(logicalName, meta);

    if (typeof attr.ColumnNumber === "number") {
      byColumn.set(attr.ColumnNumber, logicalName);
    }
  }

  /** @type {EntityMeta} */
  const entityMeta = {
    primaryId: entityJson.PrimaryIdAttribute ?? `${entityLogicalName}id`,
    primaryName: entityJson.PrimaryNameAttribute ?? null,
    entitySetName: entityJson.EntitySetName ?? null,
    objectTypeCode: typeof entityJson.ObjectTypeCode === "number"
      ? entityJson.ObjectTypeCode
      : null,
    attributes,
    byColumn,
  };

  metadataCache.set(entityLogicalName, entityMeta);
  return entityMeta;
}

// ── Entity set name resolver ──────────────────────────────────────────────────

/**
 * Resolve the plural entity set name from an entity logical name.
 * Checks the metadata cache first; falls back to fetching metadata.
 *
 * @param {string} entityLogicalName  e.g. "account"
 * @returns {Promise<string>}         e.g. "accounts"
 */
async function resolveEntitySetName(entityLogicalName) {
  assertEntityLogicalName(entityLogicalName);
  const cached = metadataCache.get(entityLogicalName);
  if (cached?.entitySetName) return cached.entitySetName;
  const meta = await fetchEntityMetadata(entityLogicalName);
  if (!meta.entitySetName) {
    throw new Error(
      `Could not resolve EntitySetName for "${entityLogicalName}". Ensure the entity exists and you have read access.`,
    );
  }
  return meta.entitySetName;
}

// ── Record name fetcher ───────────────────────────────────────────────────────

/**
 * Fetch the primary name attribute value for a single record.
 *
 * @param {string} entitySetName       e.g. "accounts"
 * @param {string} guid                Record GUID (bare, lowercase)
 * @param {string|null} primaryNameAttr Logical name of the primary name attribute
 * @returns {Promise<string>}          The display name, or "(unknown)" on failure
 */
async function fetchRecordName(entitySetName, guid, primaryNameAttr) {
  if (!primaryNameAttr) return "(unknown)";
  assertEntitySetName(entitySetName);
  assertGuid(guid);

  try {
    const url =
      `${getOrgUri()}/api/data/v${API_VERSION}` +
      `/${entitySetName}(${guid})?$select=${primaryNameAttr}`;
    const response = await fetchWithRetry(() =>
      fetch(url, { method: "GET", credentials: "include", headers: ODATA_HEADERS }),
    );
    const json = await response.json();
    const name = json[primaryNameAttr];
    return typeof name === "string" && name ? name : "(unnamed)";
  } catch {
    return "(unknown)";
  }
}

/**
 * Batch-fetch primary name values for multiple GUIDs.
 * Returns a Map<guid, name>.
 *
 * @param {string} entitySetName
 * @param {string[]} guids
 * @param {string|null} primaryNameAttr
 * @returns {Promise<Map<string,string>>}
 */
async function fetchRecordNames(entitySetName, guids, primaryNameAttr) {
  /** @type {Map<string,string>} */
  const nameMap = new Map();
  if (!primaryNameAttr || !guids.length) return nameMap;

  const tasks = guids.map((guid) => async () => {
    const name = await fetchRecordName(entitySetName, guid, primaryNameAttr);
    nameMap.set(guid, name);
  });
  await runPool(tasks, MAX_CONCURRENT);
  return nameMap;
}

// ── AttributeMask decoder ─────────────────────────────────────────────────────

/**
 * Decode the `attributemask` field from an AuditRecord into an array of
 * attribute logical names.
 *
 * Dataverse stores `attributemask` on the `audit` entity as a comma-separated
 * string of ColumnNumber integers, e.g. "1,7,12".  Each integer maps to an
 * attribute via the ColumnNumber field in EntityDefinitions.Attributes (captured
 * in `byColumn` during the metadata fetch).
 *
 * Returns an empty array when the mask is absent or cannot be decoded.
 * The caller should then fall back to diffing OldValue / NewValue keys.
 *
 * @param {string|number|null|undefined} attributemask  Raw value from AuditRecord.
 * @param {Map<number,string>}           byColumn       ColumnNumber → logicalName.
 * @returns {string[]}
 */
function parseAttributeMask(attributemask, byColumn) {
  if (attributemask === null || attributemask === undefined) return [];

  const raw = String(attributemask).trim();
  if (!raw) return [];

  const logicalNames = [];
  for (const part of raw.split(",")) {
    const colNum = Number(part.trim());
    if (!Number.isFinite(colNum)) continue;
    const name = byColumn.get(colNum);
    if (name) logicalNames.push(name);
    // If colNum is not in byColumn, the attribute was deleted from the schema
    // after the audit entry was created — silently skip it.
  }
  return logicalNames;
}

// ── Value resolver ────────────────────────────────────────────────────────────

/** OData formatted-value annotation suffix added by Dataverse. */
const FORMATTED_SUFFIX = "@OData.Community.Display.V1.FormattedValue";

/**
 * Dataverse Audit entity OperationType option set values → human-readable labels.
 * Used as a fallback when the server does not return formatted annotations.
 */
const OPERATION_MAP = new Map([
  [1, "Create"],
  [2, "Update"],
  [3, "Delete"],
  [4, "Access"],
  [5, "Upsert"],
]);

/**
 * Resolve a single attribute value to a human-readable string.
 *
 * Resolution priority:
 *   1. `key@OData.Community.Display.V1.FormattedValue` present in container
 *      → the server already provides a formatted label; use it directly.
 *   2. null / undefined → '(empty)'
 *   3. Enum option label (Picklist / State / Status / Boolean with options map)
 *   4. Boolean with no options map → 'Yes' / 'No'
 *   5. ISO-8601 date/datetime string → locale-formatted date+time
 *   6. Default: String(value)
 *
 * @param {string}        logicalName  Attribute logical name.
 * @param {unknown}       value        Raw value from OldValue / NewValue object.
 * @param {object}        container    Full OldValue or NewValue object (for annotations).
 * @param {AttrMeta|null} attrMeta     Cached metadata for this attribute, or null.
 * @returns {string}
 */
function resolveFieldValue(logicalName, value, container, attrMeta) {
  // 1. Server-provided formatted annotation (most reliable path).
  const annotationKey = `${logicalName}${FORMATTED_SUFFIX}`;
  if (Object.prototype.hasOwnProperty.call(container, annotationKey)) {
    const fv = container[annotationKey];
    return fv !== null && fv !== undefined ? String(fv) : "(empty)";
  }

  // 2. Null / undefined sentinel.
  if (value === null || value === undefined) return "(empty)";

  // 3. Enum option resolution (Picklist, State, Status, Boolean).
  if (attrMeta?.options && typeof value === "number") {
    const label = attrMeta.options.get(value);
    if (label !== undefined) return label;
  }

  // 4. Raw boolean with no options map.
  if (typeof value === "boolean") {
    if (attrMeta?.options) {
      return attrMeta.options.get(value ? 1 : 0) ?? (value ? "Yes" : "No");
    }
    return value ? "Yes" : "No";
  }

  // 5. ISO-8601 datetime string.
  if (
    typeof value === "string" &&
    /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2})?/.test(value)
  ) {
    const d = new Date(value);
    if (!Number.isNaN(d.getTime())) {
      return d.toLocaleString(undefined, {
        dateStyle: "short",
        timeStyle: "short",
      });
    }
  }

  // 6. Default.
  return String(value);
}

// ── User name resolution helpers ──────────────────────────────────────────────

/**
 * Collect unique user GUIDs from an array of AuditDetail objects.
 *
 * @param {object[]} auditDetails  Array of AuditDetail objects.
 * @returns {Set<string>}  Set of lowercase user GUIDs.
 */
function collectUserGuids(auditDetails) {
  const guids = new Set();
  for (const detail of auditDetails) {
    const uid = detail.AuditRecord?.["_userid_value"];
    if (typeof uid === "string" && GUID_PATTERN.test(uid)) {
      guids.add(uid.toLowerCase());
    }
  }
  return guids;
}

/**
 * Batch-fetch user display names for a set of user GUIDs.
 * Merges results into `targetMap` (existing entries are not overwritten).
 *
 * @param {Set<string>|string[]} userGuids  User GUIDs to resolve.
 * @param {Map<string,string>}   [targetMap]  Map to merge results into.
 * @returns {Promise<Map<string,string>>}  guid → fullname map.
 */
async function fetchUserDisplayNames(userGuids, targetMap) {
  const map = targetMap ?? new Map();
  const toFetch = [...userGuids].filter((uid) => !map.has(uid));
  if (toFetch.length === 0) return map;

  const tasks = toFetch.map((uid) => async () => {
    try {
      const url =
        `${getOrgUri()}/api/data/v${API_VERSION}` +
        `/systemusers(${uid})?$select=fullname`;
      const resp = await fetchWithRetry(() =>
        fetch(url, { method: "GET", credentials: "include", headers: ODATA_HEADERS }),
      );
      const json = await resp.json();
      if (json.fullname) map.set(uid, json.fullname);
    } catch { /* best-effort */ }
  });
  await runPool(tasks, MAX_CONCURRENT);
  return map;
}

/**
 * Fetch user display names from AuditDetail entries (per-record fallback).
 * Used when no shared map is provided to formatAuditResults.
 *
 * @param {object[]} auditDetails
 * @returns {Promise<Map<string,string>>}
 */
async function fetchUserNames(auditDetails) {
  const guids = collectUserGuids(auditDetails);
  if (guids.size === 0) return new Map();
  return fetchUserDisplayNames(guids);
}

// ── Audit result formatter ────────────────────────────────────────────────────

/**
 * Transform the raw JSON from `RetrieveRecordChangeHistory` into an array of
 * flat, human-readable FormattedAuditRow objects.
 *
 * Algorithm per AuditDetail entry:
 *   1. Extract provenance: ChangedBy (_userid_value formatted), ChangedDate
 *      (createdon), Operation (formatted value).
 *   2. Decode `attributemask` using ColumnNumber → logicalName mapping.
 *      Fall back to diffing OldValue / NewValue keys when mask is absent.
 *   3. For each changed field:
 *      a. Resolve display name from entity metadata (falls back to logicalName).
 *      b. Resolve OldValue and NewValue to human-readable strings (formatted
 *         annotations → option labels → date formatting → String()).
 *   4. Skip fields where both OldValue and NewValue are absent from the payload.
 *
 * @param {string} guid               The GUID of the audited record.
 * @param {string} entityLogicalName  Entity logical name, e.g. "account".
 * @param {object} rawAuditData       Parsed JSON body from the API response.
 * @param {string} recordName         Display name of the record.
 * @param {Map<string,string>} [sharedUserNameMap]  Pre-fetched user name map
 *   (shared across records to avoid redundant API calls). If omitted, user
 *   names are fetched per-record (backwards-compatible fallback).
 * @returns {Promise<FormattedAuditRow[]>}
 */
async function formatAuditResults(guid, entityLogicalName, rawAuditData, recordName, sharedUserNameMap) {
  assertGuid(guid);
  assertEntityLogicalName(entityLogicalName);

  // Gracefully degrade if metadata is unavailable (403, deleted entity, etc.).
  let entityMeta;
  try {
    entityMeta = await fetchEntityMetadata(entityLogicalName);
  } catch {
    entityMeta = {
      primaryId: `${entityLogicalName}id`,
      primaryName: null,
      entitySetName: null,
      attributes: new Map(),
      byColumn: new Map(),
    };
  }
  const auditDetails = rawAuditData?.AuditDetailCollection?.AuditDetails ?? [];

  // Use the shared map if provided (pre-fetched at the batch level to avoid
  // O(records × users) API calls). Otherwise fall back to per-record fetching.
  const userNameMap = sharedUserNameMap ?? await fetchUserNames(auditDetails);

  /** @type {FormattedAuditRow[]} */
  const rows = [];

  for (const detail of auditDetails) {
    const auditRecord = detail.AuditRecord ?? {};
    const oldValue = detail.OldValue ?? {};
    const newValue = detail.NewValue ?? {};

    // ── Provenance ──────────────────────────────────────────────────────────
    const userId = auditRecord["_userid_value"];
    const changedBy = String(
      auditRecord[`_userid_value${FORMATTED_SUFFIX}`] ??
        (typeof userId === "string" ? userNameMap.get(userId.toLowerCase()) : undefined) ??
        userId ??
        "(unknown user)",
    );

    const rawDate = auditRecord.createdon ?? "";
    let changedDate = "";
    if (rawDate && rawDate.trim()) {
      const d = new Date(rawDate);
      if (!isNaN(d.getTime())) changedDate = d;
    }

    const rawOp =
      auditRecord[`operation${FORMATTED_SUFFIX}`] ??
      auditRecord.operation ??
      "";
    const operation = OPERATION_MAP.get(Number(rawOp)) ?? String(rawOp);

    // ── Determine changed fields ────────────────────────────────────────────
    // Primary: decode attributemask (ColumnNumber-based, most authoritative).
    const maskNames = parseAttributeMask(
      auditRecord.attributemask,
      entityMeta.byColumn,
    );

    // Fallback / supplement: diff the object keys (excludes '@'-annotated keys).
    const diffKeys = new Set([
      ...Object.keys(oldValue).filter((k) => !k.includes("@")),
      ...Object.keys(newValue).filter((k) => !k.includes("@")),
    ]);

    // Use mask-decoded names if available; otherwise fall back to key diffing.
    // Either way, also include diffKeys so we never miss fields.
    const changedFields = new Set(diffKeys);
    if (maskNames.length > 0) {
      for (const n of maskNames) changedFields.add(n);
    }

    // ── One row per changed field ───────────────────────────────────────────
    for (const logicalName of changedFields) {
      const attrMeta = entityMeta.attributes.get(logicalName) ?? null;
      // When metadata is unavailable (deleted field or inaccessible schema),
      // tag the name so users understand why the display name is raw.
      const displayName = attrMeta?.displayName ?? `${logicalName} (deleted)`;

      const hasOld = Object.prototype.hasOwnProperty.call(
        oldValue,
        logicalName,
      );
      const hasNew = Object.prototype.hasOwnProperty.call(
        newValue,
        logicalName,
      );

      // Skip entries where neither side carries a value for this field.
      if (!hasOld && !hasNew) continue;

      rows.push({
        RecordID: guid,
        RecordName: recordName ?? "",
        ChangedBy: changedBy,
        ChangedDate: changedDate,
        Operation: operation,
        FieldName: displayName,
        OldValue: resolveFieldValue(
          logicalName,
          hasOld ? oldValue[logicalName] : undefined,
          oldValue,
          attrMeta,
        ),
        NewValue: resolveFieldValue(
          logicalName,
          hasNew ? newValue[logicalName] : undefined,
          newValue,
          attrMeta,
        ),
      });
    }
  }

  return rows;
}

// ── User search ───────────────────────────────────────────────────────────────

async function searchUsers(query) {
  if (typeof query !== "string" || query.trim().length < 2) return [];

  const sanitized = query.trim().replace(/'/g, "''");
  const filterStr =
    `(contains(fullname,'${sanitized}') or contains(internalemailaddress,'${sanitized}')) and isdisabled eq false`;
  const url =
    `${getOrgUri()}/api/data/v${API_VERSION}/systemusers` +
    `?$filter=${encodeURIComponent(filterStr)}` +
    `&$select=systemuserid,fullname,internalemailaddress&$top=15`;

  try {
    const response = await fetchWithRetry(() =>
      fetch(url, { method: "GET", credentials: "include", headers: ODATA_HEADERS }),
    );
    const json = await response.json();
    return (json.value ?? []).map((u) => ({
      id: u.systemuserid ?? "",
      fullname: u.fullname ?? "",
      email: u.internalemailaddress ?? "",
    }));
  } catch {
    return [];
  }
}

// ── Entity search ─────────────────────────────────────────────────────────────

/** Full entity list cached after first fetch. */
let entityListCache = null;
/** Origin at the time entityListCache was populated. */
let entityListCacheOrigin = null;

/**
 * Search entities by display name OR logical name.
 *
 * Dataverse EntityDefinitions does not support $filter on DisplayName (complex
 * type), so we fetch the full entity list once, cache it, and filter locally
 * — the same pattern used by user search with contains(fullname,...).
 */
async function searchEntities(query) {
  if (typeof query !== "string" || query.trim().length < 2) return [];

  // Invalidate cache if the org origin has changed (e.g. user navigated to a different org).
  const currentOrigin = getOrgUri();
  if (entityListCache && entityListCacheOrigin !== currentOrigin) {
    entityListCache = null;
    entityListCacheOrigin = null;
  }

  // Fetch and cache the full entity list on first call.
  if (!entityListCache) {
    try {
      const url =
        `${getOrgUri()}/api/data/v${API_VERSION}/EntityDefinitions` +
        `?$select=LogicalName,DisplayName,EntitySetName`;
      const response = await fetchWithRetry(() =>
        fetch(url, { method: "GET", credentials: "include", headers: ODATA_HEADERS }),
      );
      const json = await response.json();
      entityListCache = (json.value ?? []).map((e) => ({
        logicalName: e.LogicalName ?? "",
        displayName:
          e.DisplayName?.UserLocalizedLabel?.Label ?? e.LogicalName ?? "",
        entitySetName: e.EntitySetName ?? null,
      }));
      entityListCacheOrigin = currentOrigin;
    } catch {
      return [];
    }
  }

  const lowerQuery = query.trim().toLowerCase();

  const matched = entityListCache.filter((e) => {
    const ln = e.logicalName.toLowerCase();
    const dn = e.displayName.toLowerCase();
    return ln.includes(lowerQuery) || dn.includes(lowerQuery);
  });

  matched.sort((a, b) => {
    const aStartsD = a.displayName.toLowerCase().startsWith(lowerQuery) ? 0
      : a.displayName.toLowerCase().includes(lowerQuery) ? 1 : 2;
    const bStartsD = b.displayName.toLowerCase().startsWith(lowerQuery) ? 0
      : b.displayName.toLowerCase().includes(lowerQuery) ? 1 : 2;
    const aStartsL = a.logicalName.toLowerCase().startsWith(lowerQuery) ? 0 : 1;
    const bStartsL = b.logicalName.toLowerCase().startsWith(lowerQuery) ? 0 : 1;

    const aScore = Math.min(aStartsD, aStartsL);
    const bScore = Math.min(bStartsD, bStartsL);
    if (aScore !== bScore) return aScore - bScore;
    return a.displayName.localeCompare(b.displayName);
  });

  return matched.slice(0, 20);
}

// ── User audit record discovery ───────────────────────────────────────────────

/**
 * Discover which record GUIDs a user has touched for a given entity, within an
 * optional date range. Uses two complementary data sources:
 *
 * Source 1 — Entity table query (always reliable):
 *   GET /{entitySetName}?$filter=_modifiedby_value eq {userGuid} [and modifiedon ...]
 *   Works for any user with entity-read permission — does NOT require direct
 *   audit-entity access. Finds records where the user was the last modifier.
 *
 * Source 2 — Audit entity direct query (supplementary):
 *   GET /audits?$filter=_userid_value eq {userGuid} and objecttypecode eq '...'
 *   Catches records where the user made intermediate changes (not the last editor).
 *   Requires System Admin / direct audit-entity read access. Silently skipped
 *   when the user lacks the required privilege (403 or empty result).
 *
 * Results from both sources are merged and de-duplicated.
 *
 * @param {string}      entityLogicalName  e.g. "account"
 * @param {string}      userGuid           Bare GUID of the target user.
 * @param {string|null} dateFrom           ISO date string "YYYY-MM-DD" or null.
 * @param {string|null} dateTo             ISO date string "YYYY-MM-DD" or null.
 * @returns {Promise<string[]>}            Array of unique record GUIDs.
 */
async function fetchUserAuditRecordGuids(entityLogicalName, userGuid, dateFrom, dateTo) {
  assertEntityLogicalName(entityLogicalName);
  assertGuid(userGuid);

  const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
  if (dateFrom && !DATE_RE.test(dateFrom)) throw new Error("Invalid dateFrom format");
  if (dateTo && !DATE_RE.test(dateTo)) throw new Error("Invalid dateTo format");

  const recordGuids = new Set();

  // ── Source 1: Entity table query ─────────────────────────────────────────────
  // Find records where _modifiedby_value matches the user, filtered by modifiedon.
  // This is the primary discovery path and works with standard entity-read access.
  try {
    const entitySetName = await resolveEntitySetName(entityLogicalName);
    const meta = await fetchEntityMetadata(entityLogicalName);

    let entityFilter = `_modifiedby_value eq ${userGuid}`;
    if (dateFrom) {
      entityFilter += ` and modifiedon ge ${new Date(`${dateFrom}T00:00:00`).toISOString()}`;
    }
    if (dateTo) {
      entityFilter += ` and modifiedon le ${new Date(`${dateTo}T23:59:59.999`).toISOString()}`;
    }

    let nextUrl =
      `${getOrgUri()}/api/data/v${API_VERSION}/${entitySetName}` +
      `?$filter=${entityFilter}&$select=${meta.primaryId}`;

    for (let page = 0; page < MAX_AUDIT_QUERY_PAGES; page++) {
      const resp = await fetchWithRetry(() =>
        fetch(nextUrl, { method: "GET", credentials: "include", headers: ODATA_HEADERS }),
      );
      const json = await resp.json();
      for (const record of json.value ?? []) {
        const id = record[meta.primaryId];
        if (typeof id === "string" && GUID_PATTERN.test(id)) {
          recordGuids.add(id.toLowerCase());
        }
      }

      if (recordGuids.size >= MAX_USER_AUDIT_RECORDS) break;

      const nextLink = json["@odata.nextLink"];
      if (!nextLink) break;
      try {
        const parsed = new URL(nextLink);
        if (parsed.origin !== window.location.origin) break;
        nextUrl = nextLink;
      } catch {
        break;
      }
    }
  } catch { /* entity table query failed — will rely on Source 2 */ }

  // ── Source 2: Audit entity direct query ──────────────────────────────────────
  // Catches records where the user made an intermediate edit (and was NOT the
  // final modifier). Requires direct read access to the audit entity table
  // (typically System Admin or prvReadAudit). Silently skipped when denied.
  if (recordGuids.size < MAX_USER_AUDIT_RECORDS) {
    try {
      // objecttypecode is an integer in Dataverse — use the numeric value from
      // entity metadata, NOT the string logical name (which never matches).
      const auditMeta = await fetchEntityMetadata(entityLogicalName);
      const objectTypeCode = auditMeta.objectTypeCode;
      if (objectTypeCode === null) throw new Error("objectTypeCode unavailable");

      let auditFilter =
        `_userid_value eq ${userGuid} and objecttypecode eq ${objectTypeCode}`;
      if (dateFrom) {
        auditFilter += ` and createdon ge ${new Date(`${dateFrom}T00:00:00`).toISOString()}`;
      }
      if (dateTo) {
        auditFilter += ` and createdon le ${new Date(`${dateTo}T23:59:59.999`).toISOString()}`;
      }

      let nextUrl =
        `${getOrgUri()}/api/data/v${API_VERSION}/audits` +
        `?$filter=${auditFilter}&$select=_objectid_value`;

      for (let page = 0; page < MAX_AUDIT_QUERY_PAGES; page++) {
        const resp = await fetchWithRetry(() =>
          fetch(nextUrl, { method: "GET", credentials: "include", headers: ODATA_HEADERS }),
        );

        if (!resp.ok) break; // access denied or unsupported — stop silently

        const json = await resp.json();
        for (const record of json.value ?? []) {
          const objId = record._objectid_value;
          if (typeof objId === "string" && GUID_PATTERN.test(objId)) {
            recordGuids.add(objId.toLowerCase());
          }
        }

        if (recordGuids.size >= MAX_USER_AUDIT_RECORDS) break;

        const nextLink = json["@odata.nextLink"];
        if (!nextLink) break;
        try {
          const parsed = new URL(nextLink);
          if (parsed.origin !== window.location.origin) break;
          nextUrl = nextLink;
        } catch {
          break;
        }
      }
    } catch { /* audit entity access unavailable — silently ignored */ }
  }

  return [...recordGuids].slice(0, MAX_USER_AUDIT_RECORDS);
}

// ── Audit detail filtering by user ────────────────────────────────────────────

/**
 * Filter raw AuditDetails to only those matching a specific user and date range.
 * Applied after RetrieveRecordChangeHistory returns full history for a record.
 *
 * @param {object[]}    auditDetails  Array of AuditDetail objects.
 * @param {string}      userGuid      Target user GUID (lowercase).
 * @param {string|null} dateFrom      ISO date string or null.
 * @param {string|null} dateTo        ISO date string or null.
 * @returns {object[]}  Filtered AuditDetail subset.
 */
function filterAuditDetailsByUser(auditDetails, userGuid, dateFrom, dateTo) {
  // Parse dates as LOCAL midnight / end-of-day (no Z suffix → local timezone).
  const fromMs = dateFrom ? new Date(`${dateFrom}T00:00:00`).getTime() : -Infinity;
  const toMs = dateTo ? new Date(`${dateTo}T23:59:59.999`).getTime() : Infinity;

  return auditDetails.filter((detail) => {
    const auditRecord = detail.AuditRecord ?? {};
    const userId = auditRecord["_userid_value"];

    if (typeof userId === "string") {
      if (userId.toLowerCase() !== userGuid.toLowerCase()) return false;
    } else {
      return false;
    }

    const rawDate = auditRecord.createdon;
    if (rawDate) {
      const ts = new Date(rawDate).getTime();
      if (Number.isFinite(ts) && (ts < fromMs || ts > toMs)) return false;
    }

    return true;
  });
}

// ── Extension message router ──────────────────────────────────────────────────

/**
 * Messages accepted from the popup or service worker:
 *
 *   GET_CONTEXT — Requests the current Dynamics page context.
 *     Response: { ok: true, context: ContextPayload }
 *
 *   FETCH_AUDIT_HISTORY — Fetch raw audit change history for one or more records.
 *     Payload: { entitySetName: string, guids: string[] }
 *     Response (success): { ok: true,  results: AuditResult[] }
 *     Response (error):   { ok: false, error: string }
 *
 *   FETCH_AND_FORMAT_AUDIT — Fetch + resolve metadata + format into human-readable rows.
 *     Payload: { entityLogicalName: string, entitySetName: string, guids: string[] }
 *     Response (success): { ok: true,  rows: FormattedAuditRow[] }
 *     Response (error):   { ok: false, error: string }
 *
 *   PING — Liveness check.
 *     Response: { ok: true, alive: true }
 */
chrome.runtime.onMessage.addListener(
  function onExtensionMessage(message, sender, sendResponse) {
    // Reject messages not originating from this extension.
    if (sender.id !== chrome.runtime.id) return false;
    if (message.type === "GET_CONTEXT") {
      requestFreshContext()
        .then((context) => sendResponse({ ok: true, context }))
        .catch(() =>
          sendResponse({ ok: false, context: cachedContext, error: "timeout" }),
        );
      return true;
    }

    if (message.type === "FILL_DATA") {
      if (!EXPECTED_ORIGIN) {
        sendResponse({ ok: false, error: "Cannot reach page bridge." });
        return false;
      }
      var fillScript = document.createElement("script");
      fillScript.src = chrome.runtime.getURL("src/inject/fill-data.js");
      fillScript.type = "text/javascript";

      var fillTimer = setTimeout(function () {
        if (pendingFillRequest) {
          pendingFillRequest = null;
          sendResponse({ ok: false, error: "Fill script timed out." });
        }
        if (fillScript.parentNode) fillScript.remove();
      }, 15000);

      pendingFillRequest = {
        resolve: function (result) {
          clearTimeout(fillTimer);
          sendResponse(result);
        },
        timer: fillTimer,
      };

      (document.head || document.documentElement).appendChild(fillScript);
      fillScript.addEventListener("load", function () { fillScript.remove(); });
      fillScript.addEventListener("error", function () {
        fillScript.remove();
        if (pendingFillRequest) {
          clearTimeout(fillTimer);
          pendingFillRequest = null;
          sendResponse({ ok: false, error: "Failed to load fill script. Check extension permissions." });
        }
      });

      return true;
    }

    if (message.type === "SEARCH_USERS") {
      const { query } = message;
      if (typeof query !== "string" || query.trim().length < 2) {
        sendResponse({ ok: false, error: "Query must be at least 2 characters." });
        return false;
      }
      searchUsers(query)
        .then((users) => sendResponse({ ok: true, users }))
        .catch((err) =>
          sendResponse({ ok: false, error: err.message ?? String(err) }),
        );
      return true;
    }

    if (message.type === "SEARCH_ENTITIES") {
      const { query } = message;
      if (typeof query !== "string" || query.trim().length < 2) {
        sendResponse({ ok: false, error: "Query must be at least 2 characters." });
        return false;
      }
      searchEntities(query)
        .then((entities) => sendResponse({ ok: true, entities }))
        .catch((err) =>
          sendResponse({ ok: false, error: err.message ?? String(err) }),
        );
      return true;
    }
  },
);

// ── Port-based audit export with streaming progress ───────────────────────────
//
// The popup opens a named port ("audit-export") and sends all GUIDs at once.
// This handler runs them through the concurrency pool while posting incremental
// progress messages back through the port so the popup can update a progress bar.

chrome.runtime.onConnect.addListener(function onPortConnect(port) {
  if (port.name === "audit-export") {
    handleAuditExportPort(port);
    return;
  }

  if (port.name === "user-audit-export") {
    handleUserAuditExportPort(port);
    return;
  }
});

// ── Record audit export port handler ──────────────────────────────────────────

function makeErrorRow(guid, nameMap, operation, message) {
  return {
    RecordID: guid,
    RecordName: nameMap.get(guid) ?? "",
    ChangedBy: "",
    ChangedDate: "",
    Operation: operation,
    FieldName: "",
    OldValue: "",
    NewValue: message,
  };
}

function classifyFetchError(err) {
  const status = err instanceof ApiError ? err.status : undefined;
  let errorMsg = err.message ?? String(err);
  if (status === 401) {
    errorMsg = "Session expired \u2014 please reload the page and re-authenticate.";
  } else if (status === 403) {
    errorMsg =
      'Access denied \u2014 you need the "Audit Summary View" (prvReadAuditSummary) privilege.';
  } else if (status === 404) {
    errorMsg = "Record not found \u2014 it may have been deleted.";
  }
  return { errorMsg, status };
}

async function processAuditRecords({ port, portAlive, entitySetName, entityLogicalName, guids, nameMap, sharedUserNameMap, dataTransform }) {
  const total = guids.length;
  let done = 0;
  const allRows = [];
  let rowCapHit = false;

  const tasks = guids.map((guid) => async () => {
    if (!portAlive()) return { guid, entitySetName, error: "cancelled" };

    let result;
    try {
      result = await fetchAuditHistoryForRecord(entitySetName, guid);
    } catch (err) {
      const { errorMsg, status } = classifyFetchError(err);
      result = { guid, entitySetName, error: errorMsg, status };
    }

    if (result.error) {
      allRows.push(makeErrorRow(result.guid, nameMap, "FETCH_ERROR", result.error));
    } else {
      try {
        let data = result.data;
        if (dataTransform) data = dataTransform(result, data);

        const auditDetails = data?.AuditDetailCollection?.AuditDetails ?? [];
        const newGuids = collectUserGuids(auditDetails);
        if (newGuids.size > 0) {
          await fetchUserDisplayNames(newGuids, sharedUserNameMap);
        }

        const rows = await formatAuditResults(
          result.guid, entityLogicalName, data,
          nameMap.get(result.guid) ?? "", sharedUserNameMap,
        );
        pushRowsSafe(allRows, rows, MAX_EXPORT_ROWS, () => { rowCapHit = true; });
      } catch (fmtErr) {
        allRows.push(makeErrorRow(result.guid, nameMap, "FORMAT_ERROR", fmtErr.message ?? String(fmtErr)));
      }
    }

    done++;
    if (portAlive()) {
      try { port.postMessage({ type: "progress", done, total }); } catch { /* port closed */ }
    }
    return result;
  });

  await runPool(tasks, MAX_CONCURRENT);
  return { allRows, rowCapHit };
}

function handleAuditExportPort(port) {

  let alive = true;
  port.onDisconnect.addListener(() => { alive = false; });
  const portAlive = () => alive;

  port.onMessage.addListener(async function onPortMessage(msg) {
    const { entityLogicalName, guids } = msg ?? {};
    if (
      typeof entityLogicalName !== "string" ||
      !Array.isArray(guids) ||
      guids.length === 0
    ) {
      port.postMessage({ type: "error", error: "Invalid payload." });
      return;
    }

    try {
      let entitySetName;
      try {
        entitySetName = await resolveEntitySetName(entityLogicalName);
      } catch (err) {
        if (portAlive()) port.postMessage({ type: "error", error: err.message });
        return;
      }

      const meta = await fetchEntityMetadata(entityLogicalName);
      const nameMap = await fetchRecordNames(entitySetName, guids, meta.primaryName);
      const sharedUserNameMap = new Map();

      const { allRows, rowCapHit } = await processAuditRecords({
        port, portAlive, entitySetName, entityLogicalName,
        guids, nameMap, sharedUserNameMap,
      });

      if (portAlive()) port.postMessage({ type: "done", rows: allRows, capped: rowCapHit });
    } catch (err) {
      if (portAlive()) {
        try { port.postMessage({ type: "error", error: err.message ?? String(err) }); } catch { /* port closed */ }
      }
    }
  });
}

// ── User audit export port handler ────────────────────────────────────────────

function handleUserAuditExportPort(port) {
  let alive = true;
  port.onDisconnect.addListener(() => { alive = false; });
  const portAlive = () => alive;

  port.onMessage.addListener(async function onUserPortMessage(msg) {
    const { entityLogicalName, userGuid, dateFrom, dateTo } = msg ?? {};
    if (
      typeof entityLogicalName !== "string" ||
      typeof userGuid !== "string" ||
      !GUID_PATTERN.test(userGuid)
    ) {
      port.postMessage({ type: "error", error: "Invalid payload." });
      return;
    }

    try {
      if (portAlive()) {
        port.postMessage({ type: "phase", text: "Discovering records touched by user\u2026" });
      }

      const recordGuids = await fetchUserAuditRecordGuids(
        entityLogicalName, userGuid, dateFrom || null, dateTo || null,
      );

      if (recordGuids.length === 0) {
        if (portAlive()) port.postMessage({ type: "done", rows: [] });
        return;
      }

      if (portAlive()) {
        port.postMessage({
          type: "phase",
          text: `Found ${recordGuids.length} record${recordGuids.length !== 1 ? "s" : ""}. Fetching audit history\u2026`,
        });
      }

      let entitySetName;
      try {
        entitySetName = await resolveEntitySetName(entityLogicalName);
      } catch (err) {
        if (portAlive()) port.postMessage({ type: "error", error: err.message });
        return;
      }

      const meta = await fetchEntityMetadata(entityLogicalName);
      const nameMap = await fetchRecordNames(entitySetName, recordGuids, meta.primaryName);
      const sharedUserNameMap = new Map();

      const dataTransform = (result, data) => {
        const filtered = { ...data };
        if (filtered.AuditDetailCollection) {
          filtered.AuditDetailCollection = {
            AuditDetails: filterAuditDetailsByUser(
              filtered.AuditDetailCollection.AuditDetails ?? [],
              userGuid, dateFrom || null, dateTo || null,
            ),
          };
        }
        return filtered;
      };

      const { allRows, rowCapHit } = await processAuditRecords({
        port, portAlive, entitySetName, entityLogicalName,
        guids: recordGuids, nameMap, sharedUserNameMap, dataTransform,
      });

      if (portAlive()) port.postMessage({ type: "done", rows: allRows, capped: rowCapHit });
    } catch (err) {
      if (portAlive()) {
        try { port.postMessage({ type: "error", error: err.message ?? String(err) }); } catch { /* port closed */ }
      }
    }
  });
}

// ── Initialisation ────────────────────────────────────────────────────────────

function init() {
  injectBridge();

  // Send basic page metadata to the service worker immediately (before the
  // bridge has read Xrm); the richer DYNAMICS_CONTEXT_UPDATE follows shortly.
  chrome.runtime
    .sendMessage({
      type: "DYNAMICS_PAGE_ACTIVE",
      payload: {
        hostname: window.location.hostname,
        pathname: window.location.pathname,
        title: document.title.slice(0, 200),
      },
    })
    .catch(() => {});
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", init, { once: true });
} else {
  init();
}

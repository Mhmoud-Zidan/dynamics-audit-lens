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

'use strict';

// ── Guard against double injection ────────────────────────────────────────────
if (window.__dalContentV1) {
  throw new Error('DAL: content script already active');
}
Object.defineProperty(window, '__dalContentV1', {
  value: true, writable: false, configurable: false,
});

// ── Message type constants ────────────────────────────────────────────────────
const T_READY    = '__DAL__BRIDGE_READY';
const T_REQUEST  = '__DAL__CONTEXT_REQUEST';
const T_RESPONSE = '__DAL__CONTEXT_RESPONSE';

/**
 * The origin we expect on all postMessages from the bridge.
 * Dynamics 365 is always HTTPS, so origin is always well-defined.
 * If somehow null (sandboxed frame), we reject all messages.
 */
const EXPECTED_ORIGIN = (() => {
  const o = window.location.origin;
  return o && o !== 'null' ? o : null;
})();

// ── State ─────────────────────────────────────────────────────────────────────
/** Latest context snapshot received from the bridge. */
let cachedContext = null;

/** True once the <script> element has been appended. */
let bridgeInjected = false;

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
    const script       = document.createElement('script');
    script.src        = chrome.runtime.getURL('src/inject/page-bridge.js');
    script.type       = 'text/javascript';

    const parent = document.head ?? document.documentElement;
    parent.appendChild(script);

    script.addEventListener('load',  () => script.remove(), { once: true });
    script.addEventListener('error', () => {
      console.warn(
        '[Audit Lens] Bridge script blocked by page CSP. ' +
        'Xrm context detection is unavailable on this page.'
      );
      script.remove();
    }, { once: true });
  } catch (err) {
    console.error('[Audit Lens] Bridge injection failed:', err);
  }
}

// ── postMessage validation ────────────────────────────────────────────────────

function isTrustedBridgeEvent(event, expectedType) {
  if (event.source !== window)                      return false;
  if (!EXPECTED_ORIGIN || event.origin !== EXPECTED_ORIGIN) return false;
  if (!event.data || event.data.type !== expectedType)      return false;
  return true;
}

// ── Pending request queue ────────────────────────────────────────────────────
/**
 * Each entry represents one popup waiting for a fresh GET_CONTEXT reply.
 * Shape: { resolve: Function, timer: number }
 */
const pendingRequests = [];

// ── postMessage listener ──────────────────────────────────────────────────────

window.addEventListener('message', function onBridgeMessage(event) {
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
});

// ── Background notification ───────────────────────────────────────────────────

function notifyBackground(context) {
  chrome.runtime.sendMessage({
    type: 'DYNAMICS_CONTEXT_UPDATE',
    payload: {
      hostname:  window.location.hostname,
      pathname:  window.location.pathname,
      title:     document.title.slice(0, 200),
      context,
    },
  }).catch(() => { /* extension context may be invalidated; ignore */ });
}

// ── Context request helper ────────────────────────────────────────────────────

/**
 * Ask the bridge for a fresh context reading.
 * Resolves with the bridge's response, or after 2 s falls back to cachedContext.
 *
 * @returns {Promise<object>}
 */
function requestFreshContext() {
  return new Promise(resolve => {
    const FALLBACK_CONTEXT = {
      available: false, pageType: null, entityName: null,
      entityId: null, selectedIds: [],
    };

    const timer = setTimeout(() => {
      const idx = pendingRequests.findIndex(r => r.resolve === resolve);
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

const API_VERSION    = '9.2';
const MAX_CONCURRENT = 5;

/**
 * Standard OData headers required by the Dataverse REST endpoint.
 * Prefer `application/json` over `application/atom+xml`.
 * `OData-MaxVersion` pins the protocol so future server upgrades don't break parsing.
 */
const ODATA_HEADERS = Object.freeze({
  'Accept':          'application/json; odata.metadata=minimal',
  'OData-MaxVersion': '4.0',
  'OData-Version':   '4.0',
  'Content-Type':    'application/json; charset=utf-8',
});

// ── GUID validation ───────────────────────────────────────────────────────────

const GUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * Throw if the supplied value is not a well-formed GUID string.
 * Used to sanitise caller-supplied record IDs before they are interpolated
 * into URL paths, eliminating path-traversal / injection risk.
 *
 * @param {string} guid
 * @throws {TypeError}
 */
function assertGuid(guid) {
  if (typeof guid !== 'string' || !GUID_PATTERN.test(guid)) {
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
  if (typeof name !== 'string' || !/^[A-Za-z][A-Za-z0-9_]{0,127}$/.test(name)) {
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
  let   nextIdx = 0;

  async function worker() {
    while (nextIdx < tasks.length) {
      const idx  = nextIdx++;          // claim a task slot atomically within the microtask
      results[idx] = await tasks[idx](); // may throw; propagated to caller via Promise.all
    }
  }

  // Spin up exactly `limit` worker coroutines (or fewer if there aren't enough tasks).
  const workers = Array.from(
    { length: Math.min(limit, tasks.length) },
    () => worker()
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
        lastError = new ApiError(response.status, 'Rate-limited or service unavailable', response);
        break;
      }
      const retryAfter = Number(response.headers.get('Retry-After') ?? 0);
      const delay      = retryAfter > 0 ? retryAfter * 1000 : backoffMs(attempt);
      console.warn(`[Audit Lens] HTTP ${response.status} — retrying in ${delay}ms (attempt ${attempt + 1}/${maxRetries})`);
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
  return new Promise(resolve => setTimeout(resolve, ms));
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
    this.name     = 'ApiError';
    this.status   = status;
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

  const url = `${getOrgUri()}/api/data/v${API_VERSION}`
            + `/${entitySetName}(${encodeURIComponent(guid)})`
            + `/Microsoft.Dynamics.CRM.RetrieveRecordChangeHistory()`;

  const response = await fetchWithRetry(() =>
    fetch(url, {
      method:      'GET',
      credentials: 'include',      // send the MSCRM session cookie
      headers:     ODATA_HEADERS,
    })
  );

  const data = await response.json();
  return { guid, entitySetName, data };
}

// ── Batch fetch with concurrency limiter ──────────────────────────────────────

/**
 * Fetch audit history for multiple record GUIDs, with at most MAX_CONCURRENT
 * simultaneous HTTP requests in-flight.
 *
 * Records that fail (network error, 403, 404, …) are captured as
 * `{ guid, entitySetName, error: string }` objects so one bad record does NOT
 * abort the whole batch.
 *
 * @param {string}   entitySetName  Plural entity set name, e.g. "accounts"
 * @param {string[]} guids          Array of bare, lowercase GUIDs
 * @returns {Promise<Array<AuditResult>>}
 *
 * @typedef {{ guid: string, entitySetName: string, data: object }
 *           | { guid: string, entitySetName: string, error: string, status?: number }} AuditResult
 */
async function fetchAuditHistoryBatch(entitySetName, guids) {
  if (!Array.isArray(guids) || guids.length === 0) return [];

  // Build one task-factory per GUID.
  const tasks = guids.map(guid => async () => {
    try {
      return await fetchAuditHistoryForRecord(entitySetName, guid);
    } catch (err) {
      // Capture per-record failures; don't let one record block others.
      return {
        guid,
        entitySetName,
        error:  err.message ?? String(err),
        status: err instanceof ApiError ? err.status : undefined,
      };
    }
  });

  return runPool(tasks, MAX_CONCURRENT);
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
 * @property {Map<string,AttrMeta>} attributes logicalName → AttrMeta.
 * @property {Map<number,string>}   byColumn   ColumnNumber → logicalName (for attributemask).
 */

/**
 * @typedef {object} FormattedAuditRow
 * @property {string} RecordID    GUID of the audited record.
 * @property {string} ChangedBy   Display name of the user who made the change.
 * @property {string} ChangedDate Localised date/time string.
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
  if (typeof name !== 'string' || !/^[a-z][a-z0-9_]{0,127}$/.test(name)) {
    throw new TypeError(`[Audit Lens] Invalid entity logical name: "${name}"`);
  }
}

// ── Metadata fetcher ──────────────────────────────────────────────────────────

/**
 * Fetch (and cache) entity attribute metadata from the Dataverse Metadata API.
 *
 * Endpoint:
 *   GET [orgUri]/api/data/v9.2/EntityDefinitions(LogicalName='<entity>')
 *     ?$select=LogicalName,PrimaryIdAttribute
 *     &$expand=Attributes(
 *         $select=LogicalName,DisplayName,AttributeType,ColumnNumber,GlobalOptionSetName;
 *         $expand=OptionSet($select=IsGlobal,Options,TrueOption,FalseOption)
 *       )
 *
 * The nested $expand on OptionSet retrieves inline option labels for
 * Picklist / Status / State / Boolean attributes without extra round-trips.
 * For attributes using a GlobalOptionSet the `OptionSet` property is still
 * populated with the resolved options by the server.
 *
 * @param {string} entityLogicalName  e.g. "account"
 * @returns {Promise<EntityMeta>}
 */
async function fetchEntityMetadata(entityLogicalName) {
  assertEntityLogicalName(entityLogicalName);

  if (metadataCache.has(entityLogicalName)) {
    return metadataCache.get(entityLogicalName);
  }

  // Nested OData query options inside $expand use semicolons as separators.
  const nestedOpts = [
    '$select=LogicalName,DisplayName,AttributeType,ColumnNumber,GlobalOptionSetName',
    '$expand=OptionSet($select=IsGlobal,Options,TrueOption,FalseOption)',
  ].join(';');

  const url = `${getOrgUri()}/api/data/v${API_VERSION}`
            + `/EntityDefinitions(LogicalName='${entityLogicalName}')`
            + `?$select=LogicalName,PrimaryIdAttribute`
            + `&$expand=Attributes(${nestedOpts})`;

  const response = await fetchWithRetry(() =>
    fetch(url, {
      method:      'GET',
      credentials: 'include',
      headers:     ODATA_HEADERS,
    })
  );

  const json = await response.json();

  const attributes = new Map(); // logicalName → AttrMeta
  const byColumn   = new Map(); // ColumnNumber → logicalName

  for (const attr of (json.Attributes ?? [])) {
    const logicalName = attr.LogicalName;
    if (typeof logicalName !== 'string') continue;

    const displayName = attr.DisplayName?.UserLocalizedLabel?.Label ?? logicalName;
    const type        = String(attr.AttributeType ?? 'Unknown');

    // ── Build integer → label map for option-set backed types ──────────────
    let options = null;

    if (type === 'Boolean' && attr.OptionSet) {
      // Boolean attributes use TrueOption / FalseOption, not an Options array.
      const trueLabel  = attr.OptionSet.TrueOption?.Label?.UserLocalizedLabel?.Label  ?? 'Yes';
      const falseLabel = attr.OptionSet.FalseOption?.Label?.UserLocalizedLabel?.Label ?? 'No';
      options = new Map([[1, trueLabel], [0, falseLabel]]);

    } else if (Array.isArray(attr.OptionSet?.Options)) {
      // Picklist, State, Status — and GlobalOptionSets resolved inline by server.
      options = new Map();
      for (const opt of attr.OptionSet.Options) {
        const label = opt.Label?.UserLocalizedLabel?.Label;
        if (label !== undefined && opt.Value !== undefined) {
          options.set(Number(opt.Value), String(label));
        }
      }
    }

    /** @type {AttrMeta} */
    const meta = { displayName, type, options };
    attributes.set(logicalName, meta);

    if (typeof attr.ColumnNumber === 'number') {
      byColumn.set(attr.ColumnNumber, logicalName);
    }
  }

  /** @type {EntityMeta} */
  const entityMeta = {
    primaryId:  json.PrimaryIdAttribute ?? `${entityLogicalName}id`,
    attributes,
    byColumn,
  };

  metadataCache.set(entityLogicalName, entityMeta);
  return entityMeta;
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
  for (const part of raw.split(',')) {
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
const FORMATTED_SUFFIX = '@OData.Community.Display.V1.FormattedValue';

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
    return fv !== null && fv !== undefined ? String(fv) : '(empty)';
  }

  // 2. Null / undefined sentinel.
  if (value === null || value === undefined) return '(empty)';

  // 3. Enum option resolution (Picklist, State, Status, Boolean).
  if (attrMeta?.options && typeof value === 'number') {
    const label = attrMeta.options.get(value);
    if (label !== undefined) return label;
  }

  // 4. Raw boolean with no options map.
  if (typeof value === 'boolean') {
    if (attrMeta?.options) {
      return attrMeta.options.get(value ? 1 : 0) ?? (value ? 'Yes' : 'No');
    }
    return value ? 'Yes' : 'No';
  }

  // 5. ISO-8601 datetime string.
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2})?/.test(value)) {
    const d = new Date(value);
    if (!Number.isNaN(d.getTime())) {
      return d.toLocaleString(undefined, { dateStyle: 'short', timeStyle: 'short' });
    }
  }

  // 6. Default.
  return String(value);
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
 * @returns {Promise<FormattedAuditRow[]>}
 */
async function formatAuditResults(guid, entityLogicalName, rawAuditData) {
  assertGuid(guid);
  assertEntityLogicalName(entityLogicalName);

  const entityMeta   = await fetchEntityMetadata(entityLogicalName);
  const auditDetails = rawAuditData?.AuditDetailCollection?.AuditDetails ?? [];

  /** @type {FormattedAuditRow[]} */
  const rows = [];

  for (const detail of auditDetails) {
    const auditRecord = detail.AuditRecord ?? {};
    const oldValue    = detail.OldValue    ?? {};
    const newValue    = detail.NewValue    ?? {};

    // ── Provenance ──────────────────────────────────────────────────────────
    const changedBy = String(
      auditRecord[`_userid_value${FORMATTED_SUFFIX}`]
      ?? auditRecord['_userid_value']
      ?? '(unknown user)'
    );

    const rawDate    = auditRecord.createdon ?? '';
    const changedDate = rawDate
      ? new Date(rawDate).toLocaleString(undefined, { dateStyle: 'short', timeStyle: 'short' })
      : '(unknown date)';

    const operation = String(
      auditRecord[`operation${FORMATTED_SUFFIX}`]
      ?? auditRecord.operation
      ?? ''
    );

    // ── Determine changed fields ────────────────────────────────────────────
    // Primary: decode attributemask (ColumnNumber-based, most authoritative).
    const maskNames = parseAttributeMask(auditRecord.attributemask, entityMeta.byColumn);

    // Fallback / supplement: diff the object keys (excludes '@'-annotated keys).
    const diffKeys = new Set([
      ...Object.keys(oldValue).filter(k => !k.includes('@')),
      ...Object.keys(newValue).filter(k => !k.includes('@')),
    ]);

    // Use mask-decoded names if available; otherwise fall back to key diffing.
    // Either way, also include diffKeys so we never miss fields.
    const changedFields = maskNames.length > 0
      ? [...new Set([...maskNames, ...diffKeys])]
      : [...diffKeys];

    // ── One row per changed field ───────────────────────────────────────────
    for (const logicalName of changedFields) {
      const attrMeta    = entityMeta.attributes.get(logicalName) ?? null;
      const displayName = attrMeta?.displayName ?? logicalName;

      const hasOld = Object.prototype.hasOwnProperty.call(oldValue, logicalName);
      const hasNew = Object.prototype.hasOwnProperty.call(newValue, logicalName);

      // Skip entries where neither side carries a value for this field.
      if (!hasOld && !hasNew) continue;

      rows.push({
        RecordID:    guid,
        ChangedBy:   changedBy,
        ChangedDate: changedDate,
        Operation:   operation,
        FieldName:   displayName,
        OldValue:    resolveFieldValue(logicalName, hasOld ? oldValue[logicalName] : undefined, oldValue, attrMeta),
        NewValue:    resolveFieldValue(logicalName, hasNew ? newValue[logicalName] : undefined, newValue, attrMeta),
      });
    }
  }

  return rows;
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
chrome.runtime.onMessage.addListener(function onExtensionMessage(message, _sender, sendResponse) {
  if (message.type === 'GET_CONTEXT') {
    requestFreshContext()
      .then(context => sendResponse({ ok: true, context }))
      .catch(()      => sendResponse({ ok: false, context: cachedContext, error: 'timeout' }));
    return true;
  }

  if (message.type === 'FETCH_AUDIT_HISTORY') {
    const { entitySetName, guids } = message.payload ?? {};
    if (typeof entitySetName !== 'string' || !Array.isArray(guids) || guids.length === 0) {
      sendResponse({ ok: false, error: 'Invalid payload: entitySetName (string) and guids (array) are required.' });
      return false;
    }
    fetchAuditHistoryBatch(entitySetName, guids)
      .then(results => sendResponse({ ok: true, results }))
      .catch(err    => sendResponse({ ok: false, error: err.message ?? String(err) }));
    return true;
  }

  if (message.type === 'FETCH_AND_FORMAT_AUDIT') {
    const { entityLogicalName, entitySetName, guids } = message.payload ?? {};
    if (
      typeof entityLogicalName !== 'string' ||
      typeof entitySetName     !== 'string' ||
      !Array.isArray(guids) || guids.length === 0
    ) {
      sendResponse({ ok: false, error: 'Invalid payload: entityLogicalName, entitySetName (strings) and guids (array) are required.' });
      return false;
    }

    (async () => {
      const rawResults = await fetchAuditHistoryBatch(entitySetName, guids);
      const allRows    = [];

      for (const result of rawResults) {
        if (result.error) {
          // Surface per-record fetch failures as sentinel rows.
          allRows.push({
            RecordID:    result.guid,
            ChangedBy:   '',
            ChangedDate: '',
            Operation:   'FETCH_ERROR',
            FieldName:   '',
            OldValue:    '',
            NewValue:    result.error,
          });
          continue;
        }
        const rows = await formatAuditResults(result.guid, entityLogicalName, result.data);
        allRows.push(...rows);
      }
      return allRows;
    })()
      .then(rows => sendResponse({ ok: true,  rows }))
      .catch(err => sendResponse({ ok: false, error: err.message ?? String(err) }));
    return true;
  }

  if (message.type === 'PING') {
    sendResponse({ ok: true, alive: true });
    return false;
  }
});

// ── Initialisation ────────────────────────────────────────────────────────────

function init() {
  injectBridge();

  // Send basic page metadata to the service worker immediately (before the
  // bridge has read Xrm); the richer DYNAMICS_CONTEXT_UPDATE follows shortly.
  chrome.runtime.sendMessage({
    type: 'DYNAMICS_PAGE_ACTIVE',
    payload: {
      hostname: window.location.hostname,
      pathname: window.location.pathname,
      title:    document.title.slice(0, 200),
    },
  }).catch(() => {});
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init, { once: true });
} else {
  init();
}


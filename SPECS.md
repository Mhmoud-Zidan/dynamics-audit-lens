# Dynamics Audit Lens — Project Plan & Retrospective Spec

> **Type:** Chrome Extension (Manifest V3)
> **Codename:** `dynamics-audit-lens`
> **Version:** 1.0.2
> **Stack:** Vite 8 + @crxjs/vite-plugin + Vanilla ES6 + SheetJS (xlsx)
> **Target Platform:** Microsoft Dynamics 365 / Dataverse (all CRM regions)

---

## 1. Retrospective: Initiation

### 1.1 Problem Statement

Microsoft Dynamics 365 / Dataverse provides audit logging at the platform level, but **native audit export is cumbersome** — administrators must navigate the Audit Summary Viewer in the web client, export is limited, and there is no one-click mechanism to pull audit change history for multiple selected records into a structured Excel file.

### 1.2 Objectives

| # | Objective | Success Metric |
|---|-----------|----------------|
| O1 | Provide one-click audit export from any Dynamics 365 grid or form | User clicks "Export to Excel" and receives an `.xlsx` file |
| O2 | Zero data exfiltration — all processing stays in-browser | CSP `connect-src 'self'`; no external `fetch`/XHR anywhere |
| O3 | Handle Dataverse API rate limits gracefully | Concurrency cap of 5; exponential backoff on 429/503 |
| O4 | Resolve raw GUIDs/integer codes into human-readable labels | Metadata engine resolves OptionSets, display names, user names |
| O5 | Work across all Dynamics 365 Online regions and sovereign clouds | Manifest covers 18+ CRM host patterns including GCC, China |
| O6 | Provide user-centric audit export filtered by actor and date range | User selects entity + user + date range → exports only that user's changes |

### 1.3 Constraints

- **Manifest V3 only** — no persistent background pages; service worker lifecycle managed by Chrome.
- **No external dependencies at runtime** — SheetJS is the only bundled library; no CDN, no remote code.
- **Content script isolation** — `window.Xrm` is invisible to isolated-world scripts; a page-bridge injection strategy is mandatory.
- **Dataverse Service Protection API limits** — 6,000 requests/5 min, 52 concurrent — the extension must stay well within these bounds.
- **Single-purpose, no analytics** — no telemetry, no tracking, no third-party services.

### 1.4 Technical Decisions Log

| Decision | Rationale |
|----------|-----------|
| Vite + @crxjs/vite-plugin | HMR-like dev experience (`npm run dev --watch`), automatic manifest handling, tree-shaking |
| Vanilla ES6 (no React/Vue) | Popup is a single-view UI (~60 LOC HTML); framework overhead unjustified |
| SheetJS (`xlsx`) over ExcelJS | Smaller bundle, synchronous API suits the popup's single-thread context |
| Page-bridge via `<script src>` injection | Only reliable way to access `window.Xrm` from isolated world; Chrome owns the `chrome-extension://` URL |
| Port-based messaging for export | Allows streaming progress updates from content script → popup without polling |
| Symbol-based double-injection guard | `Symbol.for("__dalContentV1")` with `Object.defineProperty` — hostile page scripts cannot pre-set or tamper with it |
| Unbound `RetrieveRecordChangeHistory` over bound | Bound syntax (`/{entitySet}({guid})/...`) returns 404 on many Dataverse versions; unbound form with `@target` alias works universally |
| Parallel typed metadata requests | Dataverse forbids multi-level `$expand`; path-cast requests (`PicklistAttributeMetadata`, etc.) enable single-level expand per type |
| Manifest CSP `connect-src 'self'` applies only to extension pages | Content scripts inherit the host page's origin for `fetch()` calls — so Dataverse API requests from `content.js` use the CRM tenant's origin, not the extension's CSP. The manifest CSP (`connect-src 'self'`) governs only extension pages (popup, service worker), preventing accidental external calls from those contexts |

---

## 2. Architecture Overview

```
┌─────────────────────────────────────────────────────────────────┐
│  Chrome Extension (MV3)                                         │
│                                                                 │
│  ┌──────────────┐    chrome.tabs.connect     ┌───────────────┐ │
│  │  popup.js    │◄──────── port ────────────►│  content.js   │ │
│  │  (popup UI)  │    sendMessage / sendResp  │  (isolated)   │ │
│  └──────────────┘                            └───────┬───────┘ │
│                                                      │         │
│                                              postMessage       │
│                                              (bridge protocol) │
│                                                      │         │
│                                              ┌───────▼───────┐ │
│                                              │ page-bridge.js│ │
│                                              │ (main world)  │ │
│                                              │ window.Xrm ✓  │ │
│                                              └───────────────┘ │
│                                                                 │
│  ┌──────────────────┐                                          │
│  │ service-worker.js│◄── DYNAMICS_PAGE_ACTIVE / CONTEXT_UPDATE │
│  │ (background)     │    badge + chrome.storage.local          │
│  └──────────────────┘                                          │
└─────────────────────────────────────────────────────────────────┘
```

### Component Inventory

| File | Role | LOC | Key Responsibilities |
|------|------|-----|---------------------|
| `manifest.json` | Extension manifest (source of truth for @crxjs) | 113 | Permissions, content script matching, CSP, web-accessible resources |
| `src/popup/popup.html` | Popup UI shell | ~120 | Two-tab UI (Records / Users), settings gear, about modal, progress sections |
| `src/popup/popup.css` | Dark + light theme styles | ~900 | CSS custom properties for both themes, modal system, search dropdowns, link rows |
| `src/popup/popup.js` | Popup controller | ~750 | Dual-tab logic, entity/user search, date range, state persistence, theme toggle, about modal |
| `src/content/content.js` | Content script (main engine) | ~1700 | Bridge injection, Dataverse API, metadata, formatting, record export, user audit export. **Internal sections** (maintain clear separator comments): (1) Constants & Guards, (2) Bridge Layer, (3) API Engine & Concurrency, (4) Metadata Resolution, (5) Formatting Pipeline, (6) Search Handlers, (7) Record Audit Export, (8) User Audit Export |
| `src/inject/page-bridge.js` | Page-context bridge | 243 | Reads `window.Xrm` (UCI + legacy), collects selected IDs, postMessage protocol |
| `src/background/service-worker.js` | Background service worker | 132 | Badge updates, session persistence, message routing |
| `vite.config.js` | Build config | 19 | @crxjs plugin, oxc minification, no sourcemaps |

---

## 3. Phase Breakdown

### Phase 0 — Scaffolding & Build Pipeline

**Goal:** Establish a working MV3 extension skeleton that loads in Chrome.

**Tasks:**

| # | Task | Details |
|---|------|---------|
| 0.1 | Initialize npm project | `package.json` with `"type": "module"`, `vite`, `@crxjs/vite-plugin`, `esbuild`, `rollup` as dev dependencies; `xlsx` as production dependency |
| 0.2 | Configure Vite | `vite.config.js` — import `crx` from `@crxjs/vite-plugin`, feed `manifest.json`, set `outDir: "dist"`, `minify: "oxc"`, `sourcemap: false` |
| 0.3 | Author manifest.json | MV3 manifest: 18 CRM region host patterns for `content_scripts.matches`, `host_permissions`, `web_accessible_resources`; permissions `["storage", "activeTab"]`; strict CSP (`script-src 'self'`, `connect-src 'self'`) |
| 0.4 | Create directory structure | `src/popup/`, `src/content/`, `src/inject/`, `src/background/`, `public/icons/` |
| 0.5 | Add icons | `icon16.png`, `icon48.png`, `icon128.png` in `public/icons/` |
| 0.6 | Configure `.gitignore` | Exclude `node_modules/`, `dist/`, `.env`, OS files |
| 0.7 | Define npm scripts | `"dev": "vite build --watch"`, `"build": "vite build"`, `"clean": "rimraf dist"` |

**Acceptance Criteria:**
- `npm run build` produces a `dist/` folder
- Extension loads via `chrome://extensions` → "Load unpacked" → `dist/`
- No console errors on install

---

### Phase 1 — Page Context Detection (page-bridge.js + content.js bridge layer)

**Goal:** Detect when the user is on a Dynamics 365 page and extract Xrm context (entity name, record ID, selected grid rows).

**Tasks:**

| # | Task | File(s) | Details |
|---|------|---------|---------|
| 1.1 | Write page-bridge.js | `src/inject/page-bridge.js` | IIFE running in main world. Reads `window.Xrm.Utility.getPageContext()` (UCI path) with fallback to `Xrm.Page.data.entity` (legacy path). Collects `pageType`, `entityName`, `entityId`. Normalizes GUIDs (strip braces, lowercase). Validates with regex. |
| 1.2 | Grid selection reader | `src/inject/page-bridge.js` | On `entitylist` pages, queries `document.querySelectorAll('[aria-selected="true"][data-id], [aria-selected="true"][row-id]')` to extract selected row GUIDs from the UCI ag-grid. Validates each as well-formed GUID. **Fallback:** if DOM selectors yield zero results on an `entitylist` page, attempts `Xrm.App.grid.getGrid().getSelectedRows()` (when available). Logs a warning if both methods fail — selectors are internal UCI implementation details that may change across Dynamics updates. |
| 1.3 | Subgrid selection reader | `src/inject/page-bridge.js` | On `entityrecord` pages, iterates `Xrm.Page.controls` for subgrid controls, calls `getGrid().getSelectedRows()` to gather subgrid selections. |
| 1.4 | postMessage protocol | Both files | Define 3 message types: `__DAL__BRIDGE_READY` (page→content on load), `__DAL__CONTEXT_REQUEST` (content→page), `__DAL__CONTEXT_RESPONSE` (page→content). All scoped with `__DAL__` prefix. |
| 1.5 | Bridge injection in content.js | `src/content/content.js` | Create `<script src="chrome.runtime.getURL('src/inject/page-bridge.js')">`, append to `document.head`, remove after load event. Guard against double-injection with `Symbol.for("__dalContentV1")` + `Object.defineProperty`. |
| 1.6 | postMessage listener | `src/content/content.js` | Listen for `message` events; validate `event.source === window` and `event.origin === EXPECTED_ORIGIN`. Cache context from `T_READY` and resolve pending requests on `T_RESPONSE`. |
| 1.7 | Context request with timeout | `src/content/content.js` | `requestFreshContext()` posts `T_REQUEST`, resolves from `T_RESPONSE` or falls back to cached context after 2s timeout. |
| 1.8 | Background notification | `src/content/content.js` | On `T_READY`, forward context to service worker via `DYNAMICS_CONTEXT_UPDATE` message. Also send `DYNAMICS_PAGE_ACTIVE` on init. |

**ContextPayload Schema:**
```typescript
interface ContextPayload {
  available: boolean;
  pageType: "entityrecord" | "entitylist" | "dashboard" | "unknown" | null;
  entityName: string | null;     // e.g. "account"
  entityId: string | null;       // bare GUID, lowercase
  selectedIds: string[];         // array of bare GUIDs
  selectionUnavailable?: boolean; // warning flag for list pages
}
```

**Acceptance Criteria:**
- Navigating to `*.crm.dynamics.com` triggers bridge injection
- Opening popup shows "Active on: <hostname>" with correct entity name
- Selecting rows in a grid populates `selectedIds` with valid GUIDs
- Badge turns green on list pages, blue on form pages
- Double-injection guard prevents duplicate content script execution

---

### Phase 2 — Service Worker & Badge Management

**Goal:** Background coordination — badge updates and session persistence.

**Tasks:**

| # | Task | File(s) | Details |
|---|------|---------|---------|
| 2.1 | Install handler | `src/background/service-worker.js` | On `chrome.runtime.OnInstalledReason.INSTALL`, initialize `chrome.storage.local` with `{ sessions: [], settings: { enabled: true } }`. |
| 2.2 | Message validation | `src/background/service-worker.js` | Verify `sender.id === chrome.runtime.id` on all incoming messages. Reject external messages. |
| 2.3 | DYNAMICS_PAGE_ACTIVE handler | `src/background/service-worker.js` | Persist `{ hostname, pathname, title, timestamp }` to `sessions` array (capped at 500). Set badge to green dot. |
| 2.4 | DYNAMICS_CONTEXT_UPDATE handler | `src/background/service-worker.js` | Persist enriched context `{ hostname, pathname, title, pageType, entityName, entityId, selectedIds, timestamp }`. Sanitize all string lengths. Badge color: blue for `entityrecord`, green for `entitylist`. |
| 2.5 | GET_SESSIONS handler | `src/background/service-worker.js` | Return stored sessions from `chrome.storage.local`. Keep message channel open (`return true`). |

**Acceptance Criteria:**
- Badge appears on Dynamics pages with correct color coding
- Sessions accumulate in `chrome.storage.local`
- Unknown message types logged as warnings, not errors

---

### Phase 3 — Popup UI

**Goal:** Build the popup interface with context display and export trigger.

**Tasks:**

| # | Task | File(s) | Details |
|---|------|---------|---------|
| 3.1 | HTML structure | `src/popup/popup.html` | Header with logo + title, status message, record info bar, export button with SVG icon, progress section (bar + text), footer. Inline CSP meta tag matching manifest CSP. |
| 3.2 | Dark theme CSS | `src/popup/popup.css` | CSS custom properties for colors (`--color-bg: #0d0d1a`, `--color-accent: #0078d4`, `--color-accent-2: #00b4d8`). Animated gradient accent bar. Status variants: `--idle`, `--active`, `--error`. Progress bar with gradient fill and glow. Body width: 340px. |
| 3.3 | Context detection on open | `src/popup/popup.js` | On `DOMContentLoaded`, query active tab, validate URL against `DYNAMICS_PATTERN` regex, send `GET_CONTEXT` message to content script. Update status banner and record info. |
| 3.4 | Export button state | `src/popup/popup.js` | Disable button when 0 records selected or >250 records. Show entity name badge. |
| 3.5 | Export trigger via port | `src/popup/popup.js` | On click, open `chrome.tabs.connect(tabId, { name: "audit-export" })`. Send `{ entityLogicalName, guids }`. Listen for `progress`, `done`, `error` messages. Update progress bar and text in real-time. |
| 3.6 | Port disconnect handling | `src/popup/popup.js` | If port disconnects while exporting, show "Connection lost" error. |

**Acceptance Criteria:**
- Popup opens with dark theme and animated accent bar
- Shows correct status for Dynamics vs non-Dynamics pages
- Record count and entity name update when context is available
- Export button enables/disables correctly
- Progress bar animates during export

---

### Phase 4 — Dataverse API Engine (content.js data layer)

**Goal:** Fetch audit history from the Dataverse Web API with proper concurrency, retry, and pagination.

**Tasks:**

| # | Task | File(s) | Details |
|---|------|---------|---------|
| 4.1 | Constants & headers | `src/content/content.js` | `API_VERSION = "9.2"`, `MAX_CONCURRENT = 5`, `MAX_EXPORT_ROWS = 100_000`. OData headers: `Accept: application/json; odata.metadata=minimal`, `OData-MaxVersion: 4.0`, `Prefer: odata.include-annotations="OData.Community.Display.V1.FormattedValue"`. |
| 4.2 | Input validation | `src/content/content.js` | `assertGuid()` — regex `/^[0-9a-f]{8}-...-[0-9a-f]{12}$/i`. `assertEntitySetName()` — regex `/^[a-zA-Z_][a-zA-Z0-9_]*$/` (alphanumeric + underscore, must start with letter or underscore; Dataverse entity set names use PascalCase). `assertEntityLogicalName()` — regex `/^[a-z][a-z0-9_]*$/` (lowercase alphanumeric + underscore). All throw `TypeError` on invalid input. |
| 4.3 | Promise concurrency pool | `src/content/content.js` | `runPool(tasks, limit)` — creates `min(limit, tasks.length)` worker coroutines. Each worker claims tasks via `nextIdx++` and stores results in a pre-allocated array. Returns results in stable order. |
| 4.4 | Retry with exponential backoff | `src/content/content.js` | `fetchWithRetry(fetchFn, maxRetries=3)` — retries on 429 and 503. Honors `Retry-After` header. Backoff: `min(1000 * 2^attempt, 30000)`. Custom `ApiError` class with status code and raw response. |
| 4.5 | Single-record audit fetch | `src/content/content.js` | `fetchAuditHistoryForRecord(entitySetName, guid)` — calls unbound `RetrieveRecordChangeHistory(Target=@target)?@target={'@odata.id':'entitySet(guid)'}`. Follows `@odata.nextLink` pagination up to 50 pages. Validates nextLink is same-origin (SSRF prevention). |
| 4.6 | Batch fetch | `src/content/content.js` | `fetchAuditHistoryBatch(entitySetName, guids)` — wraps each GUID in a task factory, runs through `runPool`. Captures per-record errors as `{ guid, error, status }` objects — one bad record doesn't abort the batch. |
| 4.7 | Message handlers | `src/content/content.js` | `GET_CONTEXT` → requestFreshContext + sendResponse. `FETCH_AUDIT_HISTORY` → batch fetch raw data. `FETCH_AND_FORMAT_AUDIT` → batch fetch + metadata + format. `PING` → liveness check. All async handlers return `true` to keep channel open. |

**API Endpoint (Unbound form):**
```
GET {orgUri}/api/data/v9.2/RetrieveRecordChangeHistory(Target=@target)
    ?@target={'@odata.id':'accounts(00000000-0000-0000-0000-000000000000)'}
```

**Acceptance Criteria:**
- Audit history fetches succeed for valid GUIDs on a Dynamics page
- 429 responses trigger retry with correct backoff
- Failed records appear as error rows, not as silent skips
- Pagination correctly follows `@odata.nextLink`
- Concurrency never exceeds 5 simultaneous requests

---

### Phase 5 — Metadata Resolution Engine

**Goal:** Resolve raw attribute names, integer option-set values, and user GUIDs into human-readable labels.

**Tasks:**

| # | Task | File(s) | Details |
|---|------|---------|---------|
| 5.1 | Entity metadata fetch | `src/content/content.js` | `fetchEntityMetadata(entityLogicalName)` — parallel requests: (1) base attributes with `$expand=Attributes($select=...)`, (2-6) typed path-cast requests for `PicklistAttributeMetadata`, `MultiSelectPicklistAttributeMetadata`, `StatusAttributeMetadata`, `StateAttributeMetadata`, `BooleanAttributeMetadata` each with `$expand=OptionSet(...)`. Results merged by `LogicalName`. |
| 5.2 | Metadata cache | `src/content/content.js` | In-memory `Map<entityLogicalName, EntityMeta>` — cached for the lifetime of the tab. Avoids redundant API calls across multiple exports. |
| 5.3 | EntityMeta structure | `src/content/content.js` | `{ primaryId, primaryName, entitySetName, attributes: Map<logicalName, AttrMeta>, byColumn: Map<ColumnNumber, logicalName> }`. `AttrMeta` = `{ displayName, type, options: Map<int, string> \| null }`. |
| 5.4 | Entity set name resolver | `src/content/content.js` | `resolveEntitySetName(entityLogicalName)` — checks cache first, falls back to metadata fetch. Used to convert `entityName` → `entitySetName` for API URLs. |
| 5.5 | Record name batch fetch | `src/content/content.js` | `fetchRecordNames(entitySetName, guids, primaryNameAttr)` — batch-fetches primary name values via concurrent pool. Returns `Map<guid, name>`. |
| 5.6 | AttributeMask decoder | `src/content/content.js` | `parseAttributeMask(attributemask, byColumn)` — splits comma-separated ColumnNumber string, maps each through `byColumn` to logicalName. Falls back to empty array. |
| 5.7 | Value resolver | `src/content/content.js` | `resolveFieldValue(logicalName, value, container, attrMeta)` — priority: (1) OData formatted annotation, (2) null → "(empty)", (3) option-set label, (4) boolean → Yes/No, (5) ISO datetime → locale string, (6) String() fallback. |
| 5.8 | User name resolution | `src/content/content.js` | Inside `formatAuditResults()`, collect unique `_userid_value` GUIDs, batch-fetch from `/systemusers({guid})?$select=fullname` via pool, cache in `userNameMap`. |

**Dataverse Metadata API Endpoints:**
```
GET {orgUri}/api/data/v9.2/EntityDefinitions(LogicalName='account')
    ?$select=LogicalName,PrimaryIdAttribute,PrimaryNameAttribute,EntitySetName
    &$expand=Attributes($select=LogicalName,DisplayName,AttributeType,ColumnNumber)

GET {orgUri}/api/data/v9.2/EntityDefinitions(LogicalName='account')/Attributes
    /Microsoft.Dynamics.CRM.PicklistAttributeMetadata
    ?$select=LogicalName&$expand=OptionSet($select=Options)
```

**Acceptance Criteria:**
- Option-set integer values (e.g., statuscode 1) resolve to labels (e.g., "Active")
- Boolean fields show "Yes"/"No" or custom labels
- DateTime fields render in locale format
- Deleted/missing attributes tagged with "(deleted)"
- User GUIDs resolved to display names
- Metadata cached per entity per tab session

---

### Phase 6 — Audit Formatting Pipeline

**Goal:** Transform raw API responses into flat `FormattedAuditRow[]` for Excel export.

**Tasks:**

| # | Task | File(s) | Details |
|---|------|---------|---------|
| 6.1 | formatAuditResults() | `src/content/content.js` | For each `AuditDetail`: extract `ChangedBy`, `ChangedDate`, `Operation` from `AuditRecord`. Decode `attributemask` → field list. Supplement with OldValue/NewValue key diffing. For each changed field: resolve display name, resolve old/new values. Emit one row per changed field. |
| 6.2 | Changed field detection | `src/content/content.js` | Primary: `parseAttributeMask(auditRecord.attributemask, entityMeta.byColumn)`. Fallback: diff `Object.keys(OldValue)` ∪ `Object.keys(NewValue)`, excluding `@`-annotated keys. Union of both sets. |
| 6.3 | Operation type mapping | `src/content/content.js` | `OPERATION_MAP`: `{ 1: "Create", 2: "Update", 3: "Delete", 4: "Access", 5: "Upsert" }`. Falls back to raw value. |
| 6.4 | Error row generation | `src/content/content.js` | Failed fetch → `{ Operation: "FETCH_ERROR", NewValue: errorMsg }`. Failed format → `{ Operation: "FORMAT_ERROR", NewValue: fmtErr }`. Ensures user sees per-record errors in the Excel output. |

**FormattedAuditRow Schema:**
```typescript
interface FormattedAuditRow {
  RecordID: string;      // GUID
  RecordName: string;    // Primary name
  ChangedBy: string;     // User display name
  ChangedDate: Date;     // JS Date object
  Operation: string;     // "Create" | "Update" | "Delete" | ...
  FieldName: string;     // Human-readable attribute label
  OldValue: string;      // Resolved old value
  NewValue: string;      // Resolved new value
}
```

**Acceptance Criteria:**
- Each changed field produces exactly one row
- Fields with neither OldValue nor NewValue are skipped
- Error records surface as rows with `FETCH_ERROR`/`FORMAT_ERROR` operation
- All values are human-readable strings (no raw GUIDs or integers)

---

### Phase 7 — Port-Based Export with Streaming Progress

**Goal:** Wire the popup to the content script's export engine via a long-lived port for real-time progress updates.

**Tasks:**

| # | Task | File(s) | Details |
|---|------|---------|---------|
| 7.1 | Port connection | `src/popup/popup.js` | `chrome.tabs.connect(currentTabId, { name: "audit-export" })`. Send all GUIDs in one `postMessage`. |
| 7.2 | Port listener in content.js | `src/content/content.js` | `chrome.runtime.onConnect` — listen for port name `"audit-export"`. Parse `{ entityLogicalName, guids }` from first message. |
| 7.3 | Port alive tracking | `src/content/content.js` | `portAlive` flag set to `false` on `port.onDisconnect`. **All `port.postMessage()` calls must be wrapped in `try/catch`** — if the port disconnects between the `portAlive` check and the `postMessage()` call (TOCTOU race), the caught error is treated as an implicit disconnect (`portAlive = false`). Workers check flag before posting progress — prevents errors when popup closes mid-export. |
| 7.4 | Streaming progress | `src/content/content.js` | After each record completes (success or error), post `{ type: "progress", done, total }` through the port. **For memory efficiency, partial row arrays are posted in `progress` messages** as `{ type: "progress", done, total, rows: partialRows }`. The popup accumulates these chunks so that the final `done` message only signals completion rather than carrying the entire dataset. This prevents large result sets (up to 100k rows) from being held entirely in the content script's heap. |
| 7.5 | Row cap enforcement | `src/content/content.js` | `MAX_EXPORT_ROWS = 100_000`. If cumulative rows exceed cap, truncate and set `rowCapHit` flag. Post `{ type: "done", rows, capped }`. |
| 7.6 | Completion signaling | `src/content/content.js` | After all tasks complete, post `{ type: "done", totalRows }`. **All rows have already been transferred via `progress` chunks** — the `done` message signals completion and triggers Excel generation from the popup's accumulated rows. On unhandled error, post `{ type: "error", error: message }`. |
| 7.7 | Popup progress handler | `src/popup/popup.js` | On `progress` message → update bar width and text **and append `message.rows` to accumulated array**. On `done` → call `generateExcel()` with accumulated rows, then show success. On `error` → show error in progress text. On disconnect → show "Connection lost". |

**Port Message Protocol:**
```
popup ──connect("audit-export")──► content.js
popup ──{ entityLogicalName, guids }──► content.js

content.js ──{ type: "progress", done: N, total: M, rows: [...] }──► popup  (N times)
content.js ──{ type: "done", totalRows: N }──► popup
  OR
content.js ──{ type: "error", error: "..." }──► popup
```

**Acceptance Criteria:**
- Progress bar updates in real-time as records are processed
- Closing popup mid-export doesn't crash the content script
- Row cap prevents tab OOM on very large audits
- Final download triggers automatically

---

### Phase 8 — Excel Generation & Download

**Goal:** Convert `FormattedAuditRow[]` into a properly formatted `.xlsx` file.

**Tasks:**

| # | Task | File(s) | Details |
|---|------|---------|---------|
| 8.1 | SheetJS import | `src/popup/popup.js` | `import * as XLSX from "xlsx"` — bundled by Vite at build time. |
| 8.2 | Sheet creation | `src/popup/popup.js` | `XLSX.utils.json_to_sheet(rows)` — auto-generates headers from object keys. Sheet name: "Audit History". |
| 8.3 | Date column typing | `src/popup/popup.js` | Scan for `ChangedDate` column header. Convert cells to Excel date type (`t: 'd'`) with format `yyyy-mm-dd hh:mm:ss` for native Excel sorting/filtering. |
| 8.4 | Auto-column sizing | `src/popup/popup.js` | Sample first 200 rows, compute `max(key.length, max(value.length))` per column, cap at 50 chars. Set `ws["!cols"]`. |
| 8.5 | Auto-filter | `src/popup/popup.js` | `ws["!autofilter"] = { ref: ws["!ref"] }` — enables Excel filter dropdowns on all columns. |
| 8.6 | File naming | `src/popup/popup.js` | `AuditExport_{entityName}_{YYYYMMDD}.xlsx`. Entity name sanitized: non-alphanumeric → `_`. |
| 8.7 | Blob download | `src/popup/popup.js` | `XLSX.write(wb, { bookType: "xlsx", type: "array" })` → `Blob` → `URL.createObjectURL` → programmatic `<a>` click. Release URL after download. Null out `wb` and `ws` for GC before blob creation. |
| 8.8 | Safety caps | `src/popup/popup.js` | `MAX_EXPORT_RECORDS = 250` (GUIDs), `MAX_EXPORT_ROWS = 100_000` (formatted rows). Show warning when caps are hit. |

**Acceptance Criteria:**
- Excel file opens in Excel/Google Sheets without errors
- `ChangedDate` column is recognized as date type for sorting
- Auto-filter dropdowns work on all columns
- Column widths are readable (no truncated headers)
- Filename includes entity name and date

---

### Phase 9 — Security Hardening

**Goal:** Ensure the extension meets Chrome Web Store security requirements and cannot be exploited.

**Tasks:**

| # | Task | Details |
|---|------|---------|
| 9.1 | Content Security Policy | Manifest CSP: `default-src 'self'; script-src 'self'; style-src 'self'; connect-src 'self'; object-src 'none'; base-uri 'self'; form-action 'self'`. Popup inline meta tag mirrors this. |
| 9.2 | No innerHTML | All DOM manipulation uses `.textContent` and CSSOM — never `.innerHTML`. Prevents XSS via page content. |
| 9.3 | Input sanitization | GUID validation via regex before URL interpolation. Entity set name validation. Entity logical name validation. Prevents path traversal and injection. |
| 9.4 | Same-origin validation | Bridge postMessages validated: `event.source === window` + `event.origin === EXPECTED_ORIGIN`. `@odata.nextLink` validated as same-origin before following. |
| 9.5 | Message origin verification | Service worker checks `sender.id === chrome.runtime.id`. Rejects external messages. |
| 9.6 | Double-injection prevention | `Symbol.for("__dalContentV1")` with `Object.defineProperty` (non-writable, non-configurable). Throws on re-injection. |
| 9.7 | No eval / no remote code | `script-src 'self'` — no `eval()`, no `new Function()`, no remote script loading. |
| 9.8 | Storage overflow protection | Session list capped at 500 entries. Selected IDs capped at 250. Row output capped at 100,000. All string fields truncated to safe lengths. |
| 9.9 | No external network access | `connect-src 'self'` + no `fetch`/XHR to external origins anywhere in codebase. |
| 9.10 | Privacy policy | `PRIVACY.md` documenting no data collection, local-only processing, no third-party services. |

**Acceptance Criteria:**
- Extension passes Chrome Web Store review
- No CSP violations in console
- No data transmitted to external servers (verified via Network tab)

---

### Phase 10 — Build, Publish & Documentation

**Goal:** Production build, packaging, and user-facing documentation.

**Tasks:**

| # | Task | Details |
|---|------|---------|
| 10.1 | Production build | `npm run build` — Vite + @crxjs produces `dist/` with minified JS (`oxc`), no sourcemaps, bundled SheetJS. |
| 10.2 | Verify dist contents | `manifest.json`, icon PNGs, `page-bridge.js`, compiled `popup.html`, CSS assets, JS chunks, `service-worker-loader.js`. |
| 10.3 | Package for CWS | `publish/dynamics_audit_lens.zip` — zip of `dist/` contents for Chrome Web Store upload. |
| 10.4 | README.md | Project structure, quick start (install → dev → build → load unpacked), custom domain instructions, security design table, icon generation. |
| 10.5 | PRIVACY.md | Privacy policy with data collection, permissions justification, third-party services, contact info. |
| 10.6 | Version management | `package.json` and `manifest.json` version synchronized (currently `1.0.2`). |

**Acceptance Criteria:**
- `npm run build` succeeds without errors
- Extension loads and functions correctly from `dist/` folder
- Published zip contains all required files
- README accurately reflects project structure and usage

---

### Phase 11 — User Audit Tab

**Goal:** Add a second popup tab that lets administrators export all audit changes made by a specific user across a chosen entity, optionally filtered by date range.

**Tasks:**

| # | Task | File(s) | Details |
|---|------|---------|---------|
| 11.1 | Add second tab to HTML | `src/popup/popup.html` | Two tab buttons (`data-tab="records"`, `data-tab="users"`). Tab panels styled with `tabpanel` role. Users panel contains: entity search container, user search container, selected-user chip (clearable), date range (from/to), export button, user progress section. |
| 11.2 | Tab switching logic | `src/popup/popup.js` | On tab button click: toggle `tab-btn--active` class, show/hide panels, persist nothing (tabs are session-only). |
| 11.3 | Entity search with debounce | `src/popup/popup.js` | `entitySearchInput` fires `performEntitySearch(query)` after `SEARCH_DEBOUNCE_MS = 300ms` and `MIN_SEARCH_LENGTH = 2`. Sends `{ type: "SEARCH_ENTITIES", query }` to content script. Renders `renderEntityDropdown()`. On select: writes `ent.logicalName` to input field. |
| 11.4 | User search with debounce | `src/popup/popup.js` | `userSearchInput` fires `performUserSearch(query)` with same debounce/length constraints. Sends `{ type: "SEARCH_USERS", query }` to content script. Renders `renderUserDropdown()` showing fullname + email. On select: calls `selectUser(user)`, renders a chip showing the chosen name, clears search input. |
| 11.5 | User chip & clear | `src/popup/popup.js` | `selectedUserEl` chip (hidden by default). Shows `user.fullname` or email or first 8 chars of GUID. Clear button (`clearUserBtn`) nulls `selectedUser`, hides chip, re-evaluates export button state, persists. |
| 11.6 | Date range inputs | `src/popup/popup.js` | `dateFromInput` and `dateToInput` (`type="date"`). On `DOMContentLoaded`, `dateToInput` defaults to today via `todayISO()` if not in persisted state. |
| 11.7 | User export button state | `src/popup/popup.js` | `updateUserExportBtnState()` — disabled unless entity field has ≥ 2 chars **and** a user is selected **and** not already exporting. |
| 11.8 | State persistence | `src/popup/popup.js` | `saveUserAuditState()` / `loadUserAuditState()` using `chrome.storage.local` key `"userAuditState"`. Persisted fields: `{ entityName, selectedUser, dateFrom, dateTo }`. Loaded on `DOMContentLoaded` in `fetchContext()`. |
| 11.9 | Auto-fill entity from page context | `src/popup/popup.js` | In `fetchContext()`, if the Dynamics page has an entity name and the entity search field is empty (first visit), pre-populate `entitySearchInput` with the page's entity name. |
| 11.10 | SEARCH_ENTITIES handler | `src/content/content.js` | `searchEntities(query)` — **sanitizes query via `sanitizeOData(query)`** (escapes single quotes: `'` → `''`) to prevent OData injection. Then calls `EntityDefinitions?$filter=contains(LogicalName,'{sanitized}') or contains(...DisplayName,'{sanitized}')&$select=LogicalName,DisplayName,EntitySetName`. Returns `[{ logicalName, displayName }]`. Validates query ≥ 2 chars in message handler before delegating. |
| 11.11 | SEARCH_USERS handler | `src/content/content.js` | `searchUsers(query)` — **sanitizes query via `sanitizeOData(query)`** (escapes single quotes: `'` → `''`) to prevent OData injection. Then calls `/systemusers?$filter=contains(fullname,'{sanitized}') or contains(internalemailaddress,'{sanitized}')&$select=systemuserid,fullname,internalemailaddress`. Returns `[{ id, fullname, email }]`. Same validation gate. |
| 11.12 | User audit port handler | `src/content/content.js` | `chrome.runtime.onConnect` listens for port name `"user-audit-export"`. Delegates to `handleUserAuditExportPort(port)`. Validates `{ entityLogicalName, userGuid, dateFrom, dateTo }` payload; validates `userGuid` with `GUID_PATTERN`. **Validates date range:** if both `dateFrom` and `dateTo` are provided, asserts `new Date(dateFrom) <= new Date(dateTo)` — posts `{ type: "error", error: "Invalid date range: 'From' must be on or before 'To'." }` and returns early on failure. Also validates each date parses to a valid `Date` (not `NaN`). |
| 11.13 | Discover records by user | `src/content/content.js` | `fetchUserAuditRecordGuids(entityLogicalName, userGuid, dateFrom, dateTo)` — paginates the `/audits` OData feed filtered by `_userid_value eq {userGuid}` and `objectid_...` matching the entity, within the date range. Returns de-duplicated array of record GUIDs. Capped at `MAX_USER_AUDIT_RECORDS = 500`. Pages capped at `MAX_AUDIT_QUERY_PAGES = 20`. |
| 11.14 | Two-phase export progress | `src/content/content.js` | Phase 1: post `{ type: "phase", text: "Discovering records touched by user…" }`. Phase 2: post `{ type: "phase", text: "Found N records. Fetching audit history…" }`. Then stream `{ type: "progress", done, total }` as records complete. |
| 11.15 | Client-side user filter | `src/content/content.js` | `filterAuditDetailsByUser(auditDetails, userGuid, dateFrom, dateTo)` — filters raw `AuditDetails` array to entries where `AuditRecord._userid_value === userGuid` and `createdon` falls within the date range. Applied after fetching the full per-record audit history so only the target user's changes appear in the export. |
| 11.16 | User audit Excel generation | `src/popup/popup.js` | Calls `generateExcel(rows, entityLogicalName, suffix)` where `suffix` is `selectedUser.fullname` sanitized (`/[^a-zA-Z0-9_-]/g → "_"`). Filename: `AuditExport_{entity}_{YYYYMMDD}_{UserName}.xlsx`. |
| 11.17 | User status banner | `src/popup/popup.js` | Separate `setUserStatus(text, type)` for the Users tab. Shares the same status variants (idle / active / error) as the Records tab. |

**User Audit Port Message Protocol:**
```
popup ──connect("user-audit-export")──► content.js
popup ──{ entityLogicalName, userGuid, dateFrom, dateTo }──► content.js

content.js ──{ type: "phase",    text: "Discovering records…" }──► popup
content.js ──{ type: "phase",    text: "Found N records. Fetching…" }──► popup
content.js ──{ type: "progress", done: N, total: M, rows: [...] }──► popup  (N times)
content.js ──{ type: "done",     totalRows: N }──► popup
  OR
content.js ──{ type: "error",    error: "…" }──► popup
```

**SEARCH_ENTITIES Response Schema:**
```typescript
interface EntityResult {
  logicalName: string;   // e.g. "account"
  displayName: string;   // e.g. "Account"
}
```

**SEARCH_USERS Response Schema:**
```typescript
interface UserResult {
  id: string;       // systemuserid GUID
  fullname: string; // e.g. "Jane Doe"
  email: string;    // internalemailaddress
}
```

**Acceptance Criteria:**
- Typing ≥ 2 chars in entity search shows a dropdown of matching entities within 300ms
- Typing ≥ 2 chars in user search shows a dropdown of matching users within 300ms
- Selecting a user from the dropdown renders a chip; clicking "×" clears it
- Export button is disabled until both entity and user are set
- Date range defaults "To" to today on first open; "From" is optional
- State (entity name, selected user, dates) survives popup close/reopen
- Entity field is pre-filled from page context on first visit to a Dynamics page
- Progress shows two phases: discovery then per-record fetching
- Output Excel filename includes the user's display name as a suffix
- Closing popup mid-export does not crash content script (portAlive guard)

---

### Phase 12 — Settings, Theme & About Modal

**Goal:** Add a settings gear menu to the popup header with light/dark theme toggle and an About modal containing author contact and license information.

**Tasks:**

| # | Task | File(s) | Details |
|---|------|---------|---------|
| 12.1 | CSS custom properties (dark theme) | `src/popup/popup.css` | Define full design token set on `:root`: `--color-bg`, `--color-surface`, `--color-surface-2`, `--color-border`, `--color-text`, `--color-text-2`, `--color-muted`, `--color-accent`, `--color-accent-2`, `--color-danger`, `--color-success`. |
| 12.2 | Light theme overrides | `src/popup/popup.css` | `[data-theme="light"]` selector overrides all color tokens with light-mode equivalents. Toggled by setting `document.documentElement.setAttribute("data-theme", "light")` or removing the attribute. |
| 12.3 | Settings button in header | `src/popup/popup.html` | Gear `⚙` icon button (`id="settings-btn"`) in popup header, right-aligned. Clicking toggles `.settings-menu--open` class on the dropdown. Outside clicks close it via document listener (`stopPropagation` on button itself). |
| 12.4 | Settings dropdown | `src/popup/popup.html` | `.settings-menu` positioned absolute below gear button. Contains: (1) Theme toggle button — label reads "Light Mode" or "Dark Mode" to indicate what switching *to* yields; (2) About button that opens the About modal. |
| 12.5 | Theme toggle logic | `src/popup/popup.js` | `applyTheme(theme)` — sets or removes `data-theme` attribute, updates label text. `toggleTheme()` — reads current state, flips it, saves to `chrome.storage.local`. `loadTheme()` — reads `chrome.storage.local` key `"theme"` on `DOMContentLoaded`, falls back to "dark". |
| 12.6 | Theme persistence | `src/popup/popup.js` | `THEME_STORAGE_KEY = "theme"`. Persisted to `chrome.storage.local` (not `localStorage`) for reliability across extension lifecycle events. |
| 12.7 | About modal structure | `src/popup/popup.html` | `.modal-overlay` wrapping `.modal-card`. Sections: (1) Hero — extension icon + "Dynamics Audit Lens" title + version badge; (2) Description — one-line purpose statement; (3) Contact — two full-row `<button>` tap targets (LinkedIn, GitHub) with branded icon pill + display name + monospace handle + chevron; (4) License — MIT badge card; (5) Legal text card; (6) Copyright footer. |
| 12.8 | Modal scroll fix | `src/popup/popup.css` | `.modal-overlay`: `align-items: flex-start; overflow-y: auto; padding: 10px`. `.modal-card`: `margin: auto; flex-shrink: 0`. Prevents content clipping in the Chrome popup viewport. |
| 12.9 | Contact link rows | `src/popup/popup.css` | `.modal__link-row` — full-width `<button>`, flex row. `.modal__link-icon--linkedin` — `#0a66c2` background pill. `.modal__link-icon--github` — `#24292f` background pill. Contains label span + monospace handle span + chevron `›`. Entire row is the tap target (no nested `<a>` tags). |
| 12.10 | External links via tabs API | `src/popup/popup.js` | LinkedIn and GitHub buttons use `chrome.tabs.create({ url: "..." })` — compliant with CSP `connect-src 'self'`; no `<a target="_blank">` required. |
| 12.11 | Modal open/close | `src/popup/popup.js` | About button shows modal. Close button and overlay backdrop click hide it. Focus is not trapped (extension popup is a controlled surface). |

**Theme Token Table:**

| Token | Dark value | Light value | Usage |
|-------|-----------|-------------|-------|
| `--color-bg` | `#0d0d1a` | `#f4f6fb` | Page background |
| `--color-surface` | `#13132b` | `#ffffff` | Card / panel background |
| `--color-surface-2` | `#1a1a3e` | `#eef0f5` | Inner card / chip background |
| `--color-border` | `#2a2a5a` | `#dde1ea` | Borders and dividers |
| `--color-text` | `#e8e8ff` | `#1a1a2e` | Primary text |
| `--color-text-2` | `#9898c0` | `#4a4a6a` | Secondary / muted text |
| `--color-muted` | `#4a4a70` | `#6b6b8a` | Placeholder / disabled text |
| `--color-accent` | `#0078d4` | `#0078d4` | Primary accent (Fluent blue) |
| `--color-accent-2` | `#00b4d8` | `#0369a1` | Gradient second stop |

**Acceptance Criteria:**
- Clicking the gear button opens the settings dropdown; clicking elsewhere closes it
- Theme toggle switches between dark and dark visuals instantly, label updates correctly
- Theme choice persists across popup close/reopen and browser restart
- About modal opens from settings menu, scrolls correctly in the popup viewport
- LinkedIn button opens `linkedin.com/in/mahmoud-zidan` in a new tab
- GitHub button opens the extension's repository page in a new tab
- All text in the modal is readable in both dark and light themes

---

## 4. Data Flow Summary

```
User selects rows in Dynamics 365 grid
         │
         ▼
page-bridge.js reads ARIA-selected row IDs
         │ postMessage(__DAL__BRIDGE_READY)
         ▼
content.js caches context, notifies service worker
         │
         ▼
User opens popup → popup.js sends GET_CONTEXT
         │
         ▼
popup shows record count + entity name
         │
         ▼
User clicks "Export to Excel"
         │ chrome.tabs.connect("audit-export")
         ▼
content.js resolves entitySetName via metadata API
         │
         ▼
content.js fetches record names concurrently (pool of 5)
         │
         ▼
content.js fetches audit history per GUID (pool of 5)
  ├─ For each record: RetrieveRecordChangeHistory API call
  ├─ Follow @odata.nextLink pagination (max 50 pages)
  ├─ Retry on 429/503 (exponential backoff, 3 retries)
  ├─ Stream { type: "progress", done, total } per record
  └─ Capture per-record errors as sentinel rows
         │
         ▼
For each audit record:
  ├─ Decode attributemask → changed field names
  ├─ Resolve field display names from metadata
  ├─ Resolve OptionSet/Boolean/State/Status labels
  ├─ Resolve user GUIDs → display names
  ├─ Format dates to locale strings
  └─ Emit one FormattedAuditRow per changed field
         │
         ▼
content.js sends { type: "done", rows: [...] } via port
         │
         ▼
popup.js generates Excel via SheetJS:
  ├─ json_to_sheet with auto-columns
  ├─ ChangedDate typed as Excel date
  ├─ Auto-filter on all columns
  ├─ Filename: AuditExport_{entity}_{date}.xlsx
  └─ Blob download via programmatic <a> click
         │
         ▼
User opens .xlsx file in Excel
```

### User Audit Data Flow

```
User opens popup → clicks "Users" tab
         │
         ▼
popup.js pre-fills entity from page context (if first visit)
popup.js restores persisted state (entity, selected user, dates)
         │
         ▼
User types in entity search → debounced after 300ms
         │ chrome.tabs.sendMessage({ type: "SEARCH_ENTITIES", query })
         ▼
content.js → GET /api/data/v9.2/EntityDefinitions?$filter=contains(LogicalName,...)
         │ returns [{ logicalName, displayName }]
         ▼
popup renders entity dropdown → user selects entity
         │
         ▼
User types in user search → debounced after 300ms
         │ chrome.tabs.sendMessage({ type: "SEARCH_USERS", query })
         ▼
content.js → GET /api/data/v9.2/systemusers?$filter=contains(fullname,...)
         │ returns [{ id, fullname, email }]
         ▼
popup renders user dropdown → user selects user → chip rendered
         │
         ▼
User sets optional date range (From / To)
         │
         ▼
User clicks "Export Audit"
         │ chrome.tabs.connect(tabId, { name: "user-audit-export" })
         ▼
content.js Phase 1 — discover records:
  │ fetchUserAuditRecordGuids(entityLogicalName, userGuid, dateFrom, dateTo)
  ├─ GET /api/data/v9.2/audits?$filter=_userid_value eq {userGuid} and objectidtypecode eq ...
  ├─ Follow @odata.nextLink pagination (max 20 pages)
  ├─ De-duplicate record GUIDs (cap: 500)
  └─ Post { type: "phase", text: "Found N records. Fetching audit history…" }
         │
         ▼
content.js Phase 2 — fetch + filter per record (pool of 5):
  ├─ resolveEntitySetName → fetchEntityMetadata → fetchRecordNames
  ├─ For each discovered GUID: fetchAuditHistoryForRecord(entitySetName, guid)
  ├─ filterAuditDetailsByUser(auditDetails, userGuid, dateFrom, dateTo)
  ├─ formatAuditResults() on filtered data → FormattedAuditRow[]
  └─ Stream { type: "progress", done, total } per record
         │
         ▼
content.js sends { type: "done", rows: [...] } via port
         │
         ▼
popup.js generates Excel via SheetJS:
  ├─ Same column structure as Record Audit export
  ├─ Filename: AuditExport_{entity}_{date}_{UserName}.xlsx
  └─ Blob download via programmatic <a> click
         │
         ▼
User opens .xlsx file in Excel
```

---

## 5. File-by-File Specification

### 5.1 `manifest.json`

| Field | Value |
|-------|-------|
| `manifest_version` | 3 |
| `name` | "Dynamics Audit Lens" |
| `short_name` | "Audit Lens" |
| `version` | "1.0.2" |
| `permissions` | `["storage", "activeTab"]` |
| `host_permissions` | 18 CRM region patterns (`*.crm.dynamics.com`, `*.crm2.dynamics.com`, ..., GCC, China) |
| `content_scripts` | Single entry: `src/content/content.js`, `run_at: document_idle`, `all_frames: false` |
| `background` | `service_worker: src/background/service-worker.js`, `type: module` |
| `web_accessible_resources` | `src/inject/page-bridge.js` — same 18 host patterns, `use_dynamic_urls: false` |
| `content_security_policy.extension_pages` | `default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data: blob:; connect-src 'self'; object-src 'none'; base-uri 'self'; form-action 'self'` |

### 5.2 `src/inject/page-bridge.js`

| Function | Purpose |
|----------|---------|
| `normaliseGuid(raw)` | Strip `{}`, lowercase |
| `isGuid(s)` | Validate GUID format |
| `readUciContext()` | `Xrm.Utility.getPageContext().input` → `{ pageType, entityName, entityId }` |
| `readLegacyFormContext()` | `Xrm.Page.data.entity` → `{ pageType: "entityrecord", entityName, entityId }` |
| `readSelectedGridIds()` | Query `[aria-selected="true"][data-id]` / `[row-id]` → validate as GUIDs |
| `readSubgridSelectedIds()` | `Xrm.Page.controls` → filter subgrid → `getGrid().getSelectedRows()` |
| `collectContext()` | Aggregate all sources → `ContextPayload` |

### 5.3 `src/content/content.js`

| Function / Section | Purpose |
|---------------------|---------|
| Double-injection guard | `Symbol.for("__dalContentV1")` + `Object.defineProperty` |
| `injectBridge()` | Inject page-bridge.js into page's main world |
| `onBridgeMessage()` | Listen for T_READY / T_RESPONSE, cache context |
| `requestFreshContext()` | Post T_REQUEST, resolve from T_RESPONSE (2s timeout) |
| `runPool(tasks, limit)` | Promise concurrency pool (max 5) |
| `fetchWithRetry(fetchFn, maxRetries)` | Exponential backoff on 429/503 |
| `fetchAuditHistoryForRecord(entitySetName, guid)` | Single-record audit fetch with pagination |
| `fetchAuditHistoryBatch(entitySetName, guids)` | Batch fetch via pool |
| `fetchEntityMetadata(entityLogicalName)` | Parallel metadata fetch + cache |
| `resolveEntitySetName(entityLogicalName)` | Entity name → entity set name |
| `fetchRecordName()` / `fetchRecordNames()` | Primary name resolution |
| `parseAttributeMask(attributemask, byColumn)` | Decode ColumnNumber mask |
| `resolveFieldValue(logicalName, value, container, attrMeta)` | Value → human-readable string |
| `formatAuditResults(guid, entityLogicalName, rawAuditData, recordName)` | Raw API → FormattedAuditRow[] |
| `sanitizeOData(value)` | Escape single quotes (`'` → `''`) in user-supplied strings before OData `$filter` interpolation. Prevents OData injection in `searchEntities()` and `searchUsers()`. |
| Message handlers | GET_CONTEXT, FETCH_AUDIT_HISTORY, FETCH_AND_FORMAT_AUDIT, PING, SEARCH_USERS, SEARCH_ENTITIES |
| `searchUsers(query)` | OData query `/systemusers` → `[{ id, fullname, email }]` |
| `searchEntities(query)` | OData query `/EntityDefinitions` → `[{ logicalName, displayName }]` |
| `fetchUserAuditRecordGuids(entityLogicalName, userGuid, dateFrom, dateTo)` | Page `/audits` for entity+user combo → de-duped record GUID array |
| `filterAuditDetailsByUser(auditDetails, userGuid, dateFrom, dateTo)` | Filter AuditDetails to target user within date range |
| Record audit port handler (`"audit-export"`) | Streaming record-by-record audit export |
| User audit port handler (`"user-audit-export"`) | Two-phase streaming export: discover → fetch → filter |

### 5.4 `src/popup/popup.js`

| Function | Purpose |
|----------|---------|
| `setStatus(text, type)` | Update Records tab status banner (idle/active/error) |
| `setUserStatus(text, type)` | Update Users tab status banner |
| `updateProgress(processed, total)` | Update Records tab progress bar + text |
| `updateUserProgress(processed, total)` | Update Users tab progress bar + text |
| `formatDateStamp()` | YYYYMMDD for filenames |
| `todayISO()` | Today's date as `YYYY-MM-DD` for date input default |
| `generateExcel(rows, entityName, suffix?)` | SheetJS → .xlsx → Blob download; optional `suffix` appended to filename |
| `fetchContext()` | Query active tab, validate URL, GET_CONTEXT, pre-fill entity search, load persisted state |
| `startExport()` | Records tab port-based export orchestration |
| `startUserExport()` | Users tab port-based export orchestration |
| `performEntitySearch(query)` | Send SEARCH_ENTITIES to content script, render dropdown |
| `renderEntityDropdown(entities)` | Populate entity search dropdown items |
| `performUserSearch(query)` | Send SEARCH_USERS to content script, render dropdown |
| `renderUserDropdown(users)` | Populate user search dropdown items |
| `selectUser(user)` | Set selectedUser, render chip, persist, update button state |
| `updateUserExportBtnState()` | Enable/disable user export button |
| `saveUserAuditState()` / `loadUserAuditState()` | Persist/restore user tab form state |
| `applyTheme(theme)` | Set/remove `data-theme` attribute, update toggle label |
| `loadTheme()` | Read chrome.storage.local and apply saved theme |
| `toggleTheme()` | Flip current theme, save, apply |
| Settings menu wiring | Open/close dropdown, outside-click handler |
| About modal wiring | Open/close modal, LinkedIn/GitHub `chrome.tabs.create()` handlers |
| Event wiring | DOMContentLoaded → loadTheme() + fetchContext() |

### 5.5 `src/background/service-worker.js`

| Function | Purpose |
|----------|---------|
| `onInstalled` handler | Initialize storage schema |
| `handleDynamicsPageActive(payload, tab)` | Persist session + green badge |
| `handleContextUpdate(payload, tab)` | Persist enriched context + colored badge |
| Message router | Validate sender.id, dispatch to handlers |

---

## 6. Constants & Limits

| Constant | Value | Location | Purpose |
|----------|-------|----------|---------|
| `MAX_EXPORT_RECORDS` | 250 | popup.js | Max GUIDs in one record-audit export batch |
| `MAX_EXPORT_ROWS` | 100,000 | popup.js, content.js | Max formatted rows in output file |
| `MAX_CONCURRENT` | 5 | content.js | Max parallel HTTP requests |
| `MAX_PAGES` | 50 | content.js | Max pagination follows per record audit |
| `MAX_USER_AUDIT_RECORDS` | 500 | content.js | Max distinct record GUIDs in user audit discovery |
| `MAX_AUDIT_QUERY_PAGES` | 20 | content.js | Max pages when paginating the `/audits` feed |
| `API_VERSION` | "9.2" | content.js | Dataverse Web API version |
| `maxRetries` | 3 | content.js | Retry attempts on 429/503 |
| `backoffMs` | min(1000 × 2^attempt, 30000) | content.js | Exponential backoff delay |
| `SEARCH_DEBOUNCE_MS` | 300 | popup.js | Entity and user search debounce delay |
| `MIN_SEARCH_LENGTH` | 2 | popup.js | Minimum characters to fire a search |
| `STATE_STORAGE_KEY` | `"userAuditState"` | popup.js | chrome.storage.local key for user tab state |
| `THEME_STORAGE_KEY` | `"theme"` | popup.js | chrome.storage.local key for theme preference |
| Session cap | 500 | service-worker.js | Max stored session records |
| Selected IDs cap | 250 | service-worker.js | Max IDs stored per session |
| Context timeout | 2,000 ms | content.js | Fallback timeout for fresh context |

---

## 7. Error Handling Matrix

| Error | Where | Behavior |
|-------|-------|----------|
| Non-Dynamics page | popup.js | Status: "Not a Dynamics / Dataverse page." |
| Content script not ready | popup.js | Status: "Content script not ready. Reload the page." |
| Too many records (>250) | popup.js | Status: "Too many records selected (max 250)." |
| HTTP 429 (rate limit) | content.js | Retry with backoff, honor Retry-After header |
| HTTP 503 (unavailable) | content.js | Retry with backoff |
| HTTP 403 (forbidden) | content.js | Per-record error: "Access denied — prvReadAuditSummary privilege" |
| HTTP 404 (not found) | content.js | Per-record error: "Record not found — may have been deleted" |
| Network error | content.js | Retry with backoff |
| Port disconnect | popup.js | "Connection lost. Reload the page and retry." |
| Bridge injection blocked | content.js | Warning log, Xrm detection unavailable |
| Metadata fetch failure | content.js | Graceful degradation — raw names used, options unavailable |
| Empty audit result (records tab) | popup.js | "No audit records found." (styled empty state) |
| Empty audit result (users tab) | popup.js | "No audit records found for this user." |
| Row cap hit | content.js | Truncate at 100k rows, show "Capped" message |
| User GUID invalid | content.js | Port error: "Invalid payload." — rejected before any API call |
| Entity search query < 2 chars | content.js | `sendResponse({ ok: false, error: "Query must be at least 2 characters." })` |
| User search query < 2 chars | content.js | Same validation response |
| Zero records discovered (user audit) | content.js | Port done with empty rows array — no API calls for phase 2 |

---

## 8. Future Considerations (Out of Scope for v1.0.2)

- [ ] Support for custom/vanity Dynamics domains via options page
- [ ] Export to CSV in addition to XLSX
- [ ] Column selection toggle (choose which fields to include)
- [ ] Audit history for related entities (expand scope beyond selected records)
- [ ] Localization / multi-language support
- [ ] Edge and Firefox compatibility
- [ ] Automated e2e tests with Puppeteer against a Dataverse trial org
- [ ] Chrome Web Store listing optimization (screenshots, description)

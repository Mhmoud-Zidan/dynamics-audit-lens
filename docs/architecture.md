# Architecture

> Dynamics Audit Lens — System architecture, component map, and data flow.

---

## Extension Components

```
dynamics-audit-lens/
├── manifest.json                  MV3 manifest — permissions, CSP, content scripts
├── vite.config.js                 Vite + @crxjs/vite-plugin build configuration
├── package.json                   Dependencies and scripts
├── public/
│   └── icons/                     Extension icons (16, 48, 128px)
├── src/
│   ├── popup/
│   │   ├── popup.html             Popup shell — tabbed UI (Records / By User)
│   │   ├── popup.css              Dark theme styles
│   │   └── popup.js               Popup controller — context detection, export, Excel
│   ├── content/
│   │   └── content.js             Content script — API engine, metadata, formatting
│   ├── inject/
│   │   └── page-bridge.js         Page-context bridge — reads window.Xrm
│   └── background/
│       └── service-worker.js      Background service worker — badge, storage
└── docs/                          Documentation
```

---

## Runtime Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                        Chrome Extension (MV3)                        │
│                                                                      │
│  ┌─────────────────────────────────────────────────────────────────┐│
│  │  Popup (popup.js)                                               ││
│  │                                                                  ││
│  │   ┌────────────┐                         ┌──────────────────┐  ││
│  │   │ Records Tab│  audit-export port      │  By User Tab     │  ││
│  │   │ (selected  │────────────────────┐    │  (user + dates)  │  ││
│  │   │  records)  │                    │    │                  │  ││
│  │   └────────────┘                    │    │ user-audit-export│  ││
│  │                                     │    │ port             │  ││
│  │   Shared: context detection,        │    └────────┬─────────┘  ││
│  │   Excel generation (SheetJS)        │             │            ││
│  └─────────────────────────────────────│─────────────│────────────┘│
│                                        │             │             │
│           chrome.tabs.connect          │             │             │
│                    │                   │             │             │
│                    ▼                   ▼             ▼             │
│  ┌──────────────────────────────────────────────────────────────┐  │
│  │  Content Script (content.js) — isolated world                │  │
│  │                                                              │  │
│  │   ┌──────────────────┐    ┌──────────────────────────────┐  │  │
│  │   │ Record Export    │    │  User Audit Export            │  │  │
│  │   │ Port Handler     │    │  Port Handler                 │  │  │
│  │   │                  │    │                                │  │  │
│  │   │ 1. Resolve meta  │    │ 1. Search users (Dataverse)   │  │  │
│  │   │ 2. Fetch names   │    │ 2. Query audit entityset      │  │  │
│  │   │ 3. Fetch history │    │ 3. Fetch history per record   │  │  │
│  │   │ 4. Format rows   │    │ 4. Filter by user + dates     │  │  │
│  │   │ 5. Stream prog.  │    │ 5. Format + stream progress   │  │  │
│  │   └──────────────────┘    └──────────────────────────────┘  │  │
│  │                                                              │  │
│  │   Shared: runPool, fetchWithRetry, metadata cache,           │  │
│  │           formatAuditResults, resolveFieldValue              │  │
│  └──────────────────────┬───────────────────────────────────────┘  │
│                          │ postMessage                              │
│                          ▼                                         │
│  ┌──────────────────────────────────────────────────────────────┐  │
│  │  Page Bridge (page-bridge.js) — main world                   │  │
│  │                                                              │  │
│  │   Reads window.Xrm (UCI + legacy)                            │  │
│  │   Collects entity name, record ID, selected grid rows         │  │
│  │   Communicates via window.postMessage                        │  │
│  └──────────────────────────────────────────────────────────────┘  │
│                                                                    │
│  ┌──────────────────────────────────────────────────────────────┐  │
│  │  Service Worker (service-worker.js)                          │  │
│  │                                                              │  │
│  │   Badge updates (green = list, blue = form)                  │  │
│  │   Session persistence (chrome.storage.local, cap 500)         │  │
│  │   Message routing with sender.id validation                  │  │
│  └──────────────────────────────────────────────────────────────┘  │
└────────────────────────────────────────────────────────────────────┘
```

---

## World Isolation Model

Chrome content scripts run in an **isolated world** — they share the DOM but not JavaScript globals. `window.Xrm` lives on the page's JS heap and is invisible to the content script.

```
┌─────────────────────────────────────┐
│  Page Context (main world)          │
│                                     │
│  window.Xrm        ✓ accessible    │
│  window.postMessage ✓ available    │
│  DOM                ✓ shared       │
│                                     │
│  page-bridge.js runs HERE           │
└───────────────┬─────────────────────┘
                │ postMessage
                ▼
┌─────────────────────────────────────┐
│  Content Script (isolated world)    │
│                                     │
│  window.Xrm        ✗ invisible     │
│  chrome.runtime     ✓ available    │
│  DOM                ✓ shared       │
│                                     │
│  content.js runs HERE               │
└─────────────────────────────────────┘
```

The bridge script (`page-bridge.js`) is injected via a `<script>` element whose `src` is a `chrome-extension://` URL (from `chrome.runtime.getURL()`). The host page cannot forge or replace this URL.

---

## Messaging Protocol

### Message Types (chrome.runtime.sendMessage / chrome.tabs.sendMessage)

| Type | Direction | Payload | Response |
|------|-----------|---------|----------|
| `GET_CONTEXT` | popup → content | — | `{ ok, context: ContextPayload }` |
| `FETCH_AUDIT_HISTORY` | popup → content | `{ entitySetName, guids }` | `{ ok, results }` |
| `FETCH_AND_FORMAT_AUDIT` | popup → content | `{ entityLogicalName, entitySetName, guids }` | `{ ok, rows }` |
| `SEARCH_USERS` | popup → content | `{ query }` | `{ ok, users: [{ id, fullname, email }] }` |
| `PING` | popup → content | — | `{ ok: true, alive: true }` |
| `DYNAMICS_PAGE_ACTIVE` | content → background | `{ hostname, pathname, title }` | `{ ok: true }` |
| `DYNAMICS_CONTEXT_UPDATE` | content → background | `{ hostname, pathname, title, context }` | `{ ok: true }` |
| `GET_SESSIONS` | popup → background | — | `{ sessions }` |

### Port Protocol (long-lived connections)

#### Record Audit Export (`audit-export`)

```
popup ──connect("audit-export")──► content.js
popup ──{ entityLogicalName, guids }──►

content.js ──{ type: "progress", done: N, total: M }──► popup
content.js ──{ type: "done", rows: [...], capped: bool }──► popup
content.js ──{ type: "error", error: "..." }──► popup
```

#### User Audit Export (`user-audit-export`)

```
popup ──connect("user-audit-export")──► content.js
popup ──{ entityLogicalName, userGuid, dateFrom, dateTo }──►

content.js ──{ type: "phase", text: "..." }──► popup
content.js ──{ type: "progress", done: N, total: M }──► popup
content.js ──{ type: "done", rows: [...], capped: bool }──► popup
content.js ──{ type: "error", error: "..." }──► popup
```

### Bridge Protocol (window.postMessage)

| Type | Direction | Purpose |
|------|-----------|---------|
| `__DAL__BRIDGE_READY` | bridge → content | Initial context snapshot on load |
| `__DAL__CONTEXT_REQUEST` | content → bridge | Request fresh context |
| `__DAL__CONTEXT_RESPONSE` | bridge → content | Respond with current context |

All bridge messages are validated in the content script:
- `event.source === window` (same window, not cross-origin iframe)
- `event.origin === EXPECTED_ORIGIN` (same HTTP origin)

---

## Data Flow — Record Audit Export

```
User selects rows in Dynamics 365 grid
         │
         ▼
page-bridge.js reads ARIA-selected row GUIDs
         │ __DAL__BRIDGE_READY
         ▼
content.js caches context, notifies service worker (badge update)
         │
         ▼
User opens popup → Records tab shows "N records selected"
         │
         ▼
User clicks "Export to Excel"
         │ chrome.tabs.connect("audit-export")
         ▼
content.js:
  1. resolveEntitySetName(entityLogicalName) → "accounts"
  2. fetchEntityMetadata("account") → attributes, option sets, column map
  3. fetchRecordNames(entitySetName, guids) → Map<guid, name>
  4. For each GUID (pool of 5):
     a. fetchAuditHistoryForRecord(entitySetName, guid)
        - GET /api/data/v9.2/RetrieveRecordChangeHistory(Target=@target)
        - Follow @odata.nextLink pagination (max 50 pages)
        - Retry on 429/503 (exponential backoff)
     b. formatAuditResults(guid, entityName, data, recordName)
        - Decode attributemask → field list
        - Resolve display names, option labels, user names
        - Emit one FormattedAuditRow per changed field
     c. Post { type: "progress", done, total }
         │
         ▼
popup.js receives { type: "done", rows: [...] }
         │
         ▼
generateExcel(rows, entityName):
  - XLSX.utils.json_to_sheet(rows)
  - ChangedDate column → Excel date type
  - Auto-size columns (sample first 200 rows)
  - Auto-filter on all columns
  - Blob download: AuditExport_{entity}_{date}.xlsx
```

## Data Flow — User Audit Export

```
User opens popup → switches to "By User" tab
         │
         ▼
Entity auto-detected from page context
         │
         ▼
User types in search → debounced 300ms → SEARCH_USERS message
         │
         ▼
content.js queries /systemusers?$filter=contains(fullname,'...')
         │
         ▼
Dropdown shows matching users → user selects one
         │
         ▼
User optionally sets date range (From / To)
         │
         ▼
User clicks "Export User Audit"
         │ chrome.tabs.connect("user-audit-export")
         ▼
content.js:
  1. fetchUserAuditRecordGuids(entity, userGuid, dateFrom, dateTo)
     - GET /audit?$filter=_userid_value eq {guid} and objecttypecode eq '{entity}'
     - Paginate, collect unique _objectid_value GUIDs (cap 500)
  2. resolveEntitySetName + fetchEntityMetadata
  3. fetchRecordNames for all discovered GUIDs
  4. For each record GUID (pool of 5):
     a. fetchAuditHistoryForRecord → full audit history
     b. filterAuditDetailsByUser → keep only target user's changes
     c. formatAuditResults → human-readable rows
     d. Post { type: "progress", done, total }
         │
         ▼
popup.js generates Excel: AuditExport_{entity}_{username}_{date}.xlsx
```

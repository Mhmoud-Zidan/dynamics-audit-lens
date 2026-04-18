# Record Audit Mode

> How the "Records" tab works — export audit history for selected grid or form records.

---

## Overview

The Records tab exports audit change history for records the user has selected in the Dynamics 365 grid, or for the record currently open in a form.

---

## User Flow

```
1. User navigates to a Dynamics 365 list view or entity form
2. User selects one or more records in the grid (or the form record is auto-detected)
3. User opens the extension popup
4. Popup shows: entity name + number of selected records
5. User clicks "Export to Excel"
6. Progress bar shows real-time progress
7. Excel file downloads automatically
```

---

## How Record Selection Works

### List View (entitylist)

The page bridge reads selected row GUIDs from ARIA attributes on the ag-grid DOM:

```javascript
document.querySelectorAll(
  '[aria-selected="true"][data-id], [aria-selected="true"][row-id]'
)
```

Each element's `data-id` or `row-id` attribute contains the record GUID. Values are validated as well-formed GUIDs before being returned.

### Form View (entityrecord)

The form's record ID is read via the UCI API:

```javascript
window.Xrm.Utility.getPageContext().input.entityId
```

With legacy fallback:

```javascript
window.Xrm.Page.data.entity.getId()
```

### Subgrid (on a form)

Subgrid selections are read via the Xrm API:

```javascript
Xrm.Page.controls.get()
  → filter by controlType === "subgrid"
  → getGrid().getSelectedRows()
```

---

## Export Pipeline

### Step 1: Context Detection

```
popup.js → chrome.tabs.sendMessage({ type: "GET_CONTEXT" }) → content.js
content.js → window.postMessage({ type: "__DAL__CONTEXT_REQUEST" }) → page-bridge.js
page-bridge.js → collectContext() → postMessage response
content.js → sendResponse({ ok: true, context }) → popup.js
```

### Step 2: Export Trigger

```
popup.js → chrome.tabs.connect({ name: "audit-export" })
         → port.postMessage({ entityLogicalName, guids })
```

### Step 3: Metadata Resolution

The content script fetches entity metadata in parallel:

```
1. Base attributes (LogicalName, DisplayName, AttributeType, ColumnNumber)
2. Picklist option sets
3. Multi-select picklist option sets
4. Status option sets
5. State option sets
6. Boolean option sets (TrueOption / FalseOption)
```

This maps:
- Attribute logical names → human-readable display names
- Integer option values → label strings
- Column numbers → logical names (for `attributemask` decoding)

### Step 4: Audit History Fetch

For each selected GUID (max 5 concurrent):

```
GET /api/data/v9.2/RetrieveRecordChangeHistory(Target=@target)
    ?@target={'@odata.id':'accounts(00000000-...-000000000000)'}
```

- Follows `@odata.nextLink` pagination (max 50 pages)
- Retries on 429/503 with exponential backoff
- Failed records captured as error rows (don't abort the batch)

### Step 5: Filtering

For each audit detail entry:

1. **Decode `attributemask`** — comma-separated ColumnNumber string → logical names
2. **Fallback to key diff** — union of OldValue and NewValue object keys (excluding `@` annotations)
3. **One row per changed field** — skip fields with neither OldValue nor NewValue

### Step 6: Value Resolution

Each old/new value is resolved through this priority chain:

```
1. OData formatted annotation (server-provided label)
2. null / undefined → "(empty)"
3. Option set integer → label string (from metadata)
4. Boolean → "Yes"/"No" (or custom labels from metadata)
5. ISO datetime → locale-formatted date string
6. Default: String(value)
```

User GUIDs in `_userid_value` are resolved to display names by querying:

```
GET /api/data/v9.2/systemusers({guid})?$select=fullname
```

### Step 7: Excel Generation

In the popup, SheetJS generates the `.xlsx` file:

- `json_to_sheet(rows)` with auto-generated headers
- `ChangedDate` column typed as Excel date (`yyyy-mm-dd hh:mm:ss`)
- Auto-sized columns (sample first 200 rows)
- Auto-filter enabled on all columns
- Filename: `AuditExport_{entityName}_{YYYYMMDD}.xlsx`

---

## Limits

| Limit | Value | Purpose |
|-------|-------|---------|
| Max selected records | 250 | Prevents Dataverse 429 rate limits |
| Max formatted rows | 100,000 | Prevents popup OOM |
| Max concurrent API requests | 5 | Stays within Dataverse Service Protection limits |
| Max pagination pages | 50 per record | Prevents infinite loops |
| Max retry attempts | 3 | Exponential backoff on 429/503 |

---

## Error Handling

| Scenario | Behavior |
|----------|----------|
| No records selected | Export button disabled |
| >250 records selected | Error message shown |
| 403 on audit fetch | Error row: "Access denied — prvReadAuditSummary privilege" |
| 404 on audit fetch | Error row: "Record not found — may have been deleted" |
| Popup closed mid-export | Workers abort via `portAlive` flag |
| Content script not ready | "Content script not ready. Reload the page." |
| Empty audit result | "No audit records found." (amber) |

# User Audit Mode

> How the "By User" tab works — export audit history for a specific user on the current entity.

---

## Overview

The User Audit tab lets you search for a Dynamics 365 user and export all their audit actions on the current entity, with optional date range filtering.

This is useful for:
- Reviewing what changes a specific user made to accounts, contacts, opportunities, etc.
- Compliance audits requiring a full trail of a user's activity on an entity
- Investigating data changes within a specific time window

---

## User Flow

```
1. User navigates to a Dynamics 365 page (list or form)
2. User opens the extension popup → switches to "By User" tab
3. Entity is auto-detected from the page context
4. User types a name or email in the search box
5. Dropdown shows matching users → user selects one
6. User optionally sets a date range (From / To)
7. User clicks "Export User Audit"
8. Two-phase progress: discovery → fetch → Excel download
```

---

## Two-Phase Export Pipeline

### Phase 1: Record Discovery

Before fetching audit history, the extension needs to know which records the user has modified. It queries the Dataverse `audits` entity set:

```http
GET {orgUri}/api/data/v9.2/audits
    ?$filter=_userid_value eq {userGuid}
        and objecttypecode eq {ObjectTypeCode}
        and createdon ge {dateFrom}T00:00:00Z
        and createdon le {dateTo}T23:59:59Z
    &$select=_objectid_value
```

- `objecttypecode` uses the **integer** ObjectTypeCode (e.g., `1` for account), resolved from entity metadata
- Results are paginated (max 20 pages)
- Unique `_objectid_value` GUIDs are collected and deduplicated
- Capped at **500 unique records**

The popup shows: `"Found N records. Fetching audit history..."`

### Phase 2: Audit History Fetch + Filter

For each discovered record GUID (max 5 concurrent), the extension:

1. **Fetches full audit history** via `RetrieveRecordChangeHistory` (same as Record Audit mode)
2. **Filters by user** — only keeps `AuditDetail` entries where `_userid_value` matches the target user
3. **Filters by date** — only keeps entries within the specified date range
4. **Formats** — uses the same metadata resolution and value formatting as Record Audit mode

### Filtering Logic

```javascript
function filterAuditDetailsByUser(auditDetails, userGuid, dateFrom, dateTo) {
  // Convert date range to timestamps
  const fromMs = dateFrom ? new Date(`${dateFrom}T00:00:00Z`).getTime() : -Infinity;
  const toMs = dateTo ? new Date(`${dateTo}T23:59:59Z`).getTime() : Infinity;

  return auditDetails.filter((detail) => {
    const auditRecord = detail.AuditRecord ?? {};

    // Filter by user GUID (case-insensitive)
    const userId = auditRecord["_userid_value"];
    if (typeof userId === "string" && userId.toLowerCase() !== userGuid.toLowerCase()) {
      return false;
    }

    // Filter by date range
    const rawDate = auditRecord.createdon;
    if (rawDate) {
      const ts = new Date(rawDate).getTime();
      if (Number.isFinite(ts) && (ts < fromMs || ts > toMs)) return false;
    }

    return true;
  });
}
```

---

## User Search

The search input queries the Dataverse `systemusers` entity:

```http
GET {orgUri}/api/data/v9.2/systemusers
    ?$filter=contains(fullname,'{query}') or contains(internalemailaddress,'{query}')
    &$select=systemuserid,fullname,internalemailaddress
    &$top=15
```

- Debounced at 300ms to avoid excessive API calls
- Results shown in a dropdown with name and email
- User selected via mousedown (with blur prevention)
- Selected user shown as a chip with clear button

---

## Date Range

| Input | Format | Behavior |
|-------|--------|----------|
| Both empty | — | No date filter applied (all time) |
| From only | `YYYY-MM-DD` | Records from that date onward |
| To only | `YYYY-MM-DD` | Records up to and including that date |
| Both set | `YYYY-MM-DD` | Records within the inclusive range |

Dates are passed from the popup as `YYYY-MM-DD` strings (native HTML date input format). The content script appends time components: `T00:00:00Z` for "from" and `T23:59:59Z` for "to".

---

## Excel Output

Same format as Record Audit mode, with filename including the user's name:

```
AuditExport_{entityName}_{UserName}_{YYYYMMDD}.xlsx
```

Columns: RecordID, RecordName, ChangedBy, ChangedDate, Operation, FieldName, OldValue, NewValue

---

## Limits

| Limit | Value | Purpose |
|-------|-------|---------|
| Max discovered records | 500 | Caps the number of records to process |
| Max audit query pages | 20 | Limits pagination during discovery phase |
| Max formatted rows | 100,000 | Prevents OOM |
| Max concurrent requests | 5 | Dataverse Service Protection |
| Search results | 15 | Dropdown usability |

---

## Error Handling

| Scenario | Behavior |
|----------|----------|
| No entity on page | "No entity detected on this page." |
| No user selected | Export button disabled |
| User has no audit actions | "No audit records found for this user." (amber) |
| 403 on audit query | Error row with privilege name |
| 404 on record fetch | Error row: "Record not found" |
| Discovery returns 0 records | Immediate "done" with empty rows |
| Date format invalid | Ignored (date filter skipped) |
| Popup closed mid-export | Workers abort via `portAlive` flag |

---

## Port Message Protocol

```
popup ──connect("user-audit-export")──► content.js
popup ──{ entityLogicalName, userGuid, dateFrom, dateTo }──►

content.js ──{ type: "phase", text: "Discovering records..." }──► popup
content.js ──{ type: "phase", text: "Found N records. Fetching..." }──► popup
content.js ──{ type: "progress", done: N, total: M }──► popup     (per record)
content.js ──{ type: "done", rows: [...], capped: bool }──► popup
content.js ──{ type: "error", error: "..." }──► popup
```

The `phase` messages provide status text during the discovery phase before the progress bar starts counting records.

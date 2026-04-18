# API Reference

> Dynamics Audit Lens — Dataverse Web API endpoints used and their parameters.

---

## Authentication

All API calls use the user's existing Dynamics 365 session cookie. The extension sends `credentials: "include"` on every `fetch()` call. No Authorization header is set.

---

## Common Headers

Every API request includes these OData headers:

```http
Accept: application/json; odata.metadata=minimal
OData-MaxVersion: 4.0
OData-Version: 4.0
Content-Type: application/json; charset=utf-8
Prefer: odata.include-annotations="OData.Community.Display.V1.FormattedValue"
```

The `Prefer` header tells Dataverse to return formatted display values alongside raw values (e.g., option set labels, lookup names) as `@OData.Community.Display.V1.FormattedValue` annotations.

---

## Endpoints

### 1. Retrieve Audit History for a Record

Retrieves the full change history for a single record, including field-level old/new values.

```http
GET {orgUri}/api/data/v9.2/RetrieveRecordChangeHistory(Target=@target)
    ?@target={'@odata.id':'{entitySetName}({guid})'}
```

| Parameter | Type | Example |
|-----------|------|---------|
| `entitySetName` | string | `accounts` |
| `guid` | GUID | `00000000-0000-0000-0000-000000000000` |

**Response:**

```json
{
  "AuditDetailCollection": {
    "AuditDetails": [
      {
        "AuditRecord": {
          "auditid": "...",
          "_userid_value": "...",
          "createdon": "2025-01-15T10:30:00Z",
          "operation": 2,
          "attributemask": "1,7,12",
          "_userid_value@OData.Community.Display.V1.FormattedValue": "John Smith",
          "operation@OData.Community.Display.V1.FormattedValue": "Update"
        },
        "OldValue": {
          "name": "Old Name",
          "statuscode": 1,
          "statuscode@OData.Community.Display.V1.FormattedValue": "Active"
        },
        "NewValue": {
          "name": "New Name",
          "statuscode": 2,
          "statuscode@OData.Community.Display.V1.FormattedValue": "Inactive"
        }
      }
    ]
  },
  "@odata.nextLink": "https://org.crm.dynamics.com/api/data/v9.2/..."
}
```

**Pagination:** Follow `@odata.nextLink` (max 50 pages per record). Validate same-origin before following.

**Retry:** 429 (rate limit) and 503 (unavailable) retried with exponential backoff (1s, 2s, 4s, ... capped at 30s). Honors `Retry-After` header.

---

### 2. Query Audit Entityset (User Audit Discovery)

Discovers which records a specific user has modified, used by the "By User" export mode.

```http
GET {orgUri}/api/data/v9.2/audits
    ?$filter=_userid_value eq {userGuid} and objecttypecode eq {otc} [and date filters]
    &$select=_objectid_value
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `userGuid` | GUID | System user ID |
| `otc` | integer | Entity ObjectTypeCode from metadata (e.g., `1` for account) |
| Date filters | optional | `createdon ge YYYY-MM-DDT00:00:00Z` and/or `createdon le YYYY-MM-DDT23:59:59Z` |

**Response:**

```json
{
  "value": [
    { "_objectid_value": "aaaaaaaa-0000-0000-0000-000000000001" },
    { "_objectid_value": "bbbbbbbb-0000-0000-0000-000000000002" }
  ],
  "@odata.nextLink": "..."
}
```

Unique `_objectid_value` GUIDs are collected (deduplicated, capped at 500). These GUIDs are then passed to `RetrieveRecordChangeHistory` individually.

---

### 3. Fetch Entity Metadata

Resolves attribute display names, option set labels, and column number mappings for audit formatting.

#### Base attributes

```http
GET {orgUri}/api/data/v9.2/EntityDefinitions(LogicalName='{entity}')
    ?$select=LogicalName,PrimaryIdAttribute,PrimaryNameAttribute,EntitySetName,ObjectTypeCode
    &$expand=Attributes($select=LogicalName,DisplayName,AttributeType,ColumnNumber)
```

#### Typed option set requests (parallel, one per type)

```http
GET {orgUri}/api/data/v9.2/EntityDefinitions(LogicalName='{entity}')/Attributes
    /Microsoft.Dynamics.CRM.PicklistAttributeMetadata
    ?$select=LogicalName&$expand=OptionSet($select=Options)

GET .../Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata?...
GET .../Microsoft.Dynamics.CRM.StatusAttributeMetadata?...
GET .../Microsoft.Dynamics.CRM.StateAttributeMetadata?...
GET .../Microsoft.Dynamics.CRM.BooleanAttributeMetadata
    ?$select=LogicalName&$expand=OptionSet($select=TrueOption,FalseOption)
```

**Why parallel?** Dataverse forbids multi-level `$expand` (causes 0x80060888 error). Each typed path-cast supports one level of `$expand` because `OptionSet` is a direct property of the cast type.

**Caching:** Results cached in-memory per entity for the lifetime of the tab.

---

### 4. Search Users

Used by the "By User" tab's user search input.

```http
GET {orgUri}/api/data/v9.2/systemusers
    ?$filter=contains(fullname,'{query}') or contains(internalemailaddress,'{query}')
    &$select=systemuserid,fullname,internalemailaddress
    &$top=15
```

**Sanitization:** Single quotes in the query are doubled (`'` → `''`) to prevent OData injection.

---

### 5. Fetch Record Display Name

Resolves the primary name field for a single record (e.g., account name).

```http
GET {orgUri}/api/data/v9.2/{entitySetName}({guid})
    ?$select={primaryNameAttr}
```

---

### 6. Fetch User Display Name

Resolves the full name for a user GUID (used when `RetrieveRecordChangeHistory` doesn't include formatted annotations).

```http
GET {orgUri}/api/data/v9.2/systemusers({guid})
    ?$select=fullname
```

---

## Rate Limiting

Dataverse enforces Service Protection API limits per user per server:

| Limit | Value | Extension Mitigation |
|-------|-------|---------------------|
| Requests | 6,000 per 5 min | Concurrency pool of 5 keeps throughput well below limit |
| Execution time | 20 min combined per 5 min | Single-record fetches complete in seconds |
| Concurrent requests | 52 | Pool of 5 uses ~10% of allowed concurrency |

The extension retries 429 responses with exponential backoff and honors the `Retry-After` header.

---

## Error Codes

| HTTP Status | Meaning | Extension Behavior |
|-------------|---------|-------------------|
| 200 | Success | Process response |
| 400 | Bad request (malformed query) | Surface as error row in export |
| 403 | Forbidden (missing `prvReadAuditSummary` privilege) | Per-record error with privilege name |
| 404 | Record not found or endpoint doesn't exist | Per-record error for deleted records |
| 429 | Rate limited | Retry with backoff (up to 3 attempts) |
| 503 | Service unavailable | Retry with backoff (up to 3 attempts) |

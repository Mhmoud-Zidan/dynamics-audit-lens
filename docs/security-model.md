# Security Model

> Dynamics Audit Lens — Threat model, security architecture, and mitigations.

---

## Core Principle: Local-First, Zero Exfiltration

This extension processes all data entirely within the user's browser. No audit data, credentials, or CRM records ever leave the device.

---

## Threat Model

| Threat | Risk | Mitigation |
|--------|------|------------|
| Data exfiltration via network | CRM audit data sent to attacker server | CSP `connect-src 'self'` blocks all external requests. No `fetch`/XHR to external origins in any script. |
| XSS via page content | Malicious Dynamics page injects scripts into popup | All DOM manipulation uses `.textContent` and `document.createElement` — never `.innerHTML`. CSP `script-src 'self'` blocks inline scripts. |
| Script injection via URL parameters | Attacker crafts GUIDs/entity names with JS payloads | All inputs validated with strict regex before URL interpolation. No `eval()` or `new Function()`. |
| Message spoofing | Malicious extension or page sends fake messages | `sender.id` validated against `chrome.runtime.id` in both service worker and content script. Bridge `postMessage` validated with `event.source === window` and `event.origin` check. |
| Double script injection | Content script injected twice by Chrome bug or hostile page | `Symbol.for("__dalContentV1")` guard with `Object.defineProperty` (non-writable, non-configurable). |
| Path traversal via API URLs | Attacker-controlled strings interpolated into API paths | GUID validation (`/^[0-9a-f]{8}-...-[0-9a-f]{12}$/i`), entity name validation (`/^[a-z][a-z0-9_]{0,127}$/`), entity set name validation (`/^[A-Za-z][A-Za-z0-9_]{0,127}$/`). |
| SSRF via pagination links | Tampered `@odata.nextLink` points to attacker server | All `@odata.nextLink` URLs validated for same-origin before following. |
| Storage overflow | Excessive data stored in `chrome.storage.local` | Sessions capped at 500 entries. All string fields truncated. Row output capped at 100,000. |
| Port message flood | Popup receives too many progress messages | Progress messages are lightweight objects. Port disconnect handling aborts workers. |

---

## Content Security Policy

### Manifest CSP (enforced by Chrome for all extension pages)

```
default-src 'self';
script-src   'self';
style-src    'self';
img-src      'self' data: blob:;
connect-src  'self';
object-src   'none';
base-uri     'self';
form-action  'self';
```

### What this prevents

- **`script-src 'self'`** — No inline scripts, no `eval()`, no `new Function()`, no remote script loading
- **`connect-src 'self'`** — The popup and service worker cannot make network requests to any external server
- **`object-src 'none'`** — No Flash, Java, or other plugin content
- **`style-src 'self'`** — No inline styles (all styles in `popup.css`)

> Note: The popup `<meta>` CSP tag is present but Chrome ignores it for MV3 extension pages — the manifest CSP takes precedence.

---

## Input Validation Matrix

Every value that enters a URL path or API query is validated before use:

| Input | Validation | Regex / Rule | Location |
|-------|-----------|--------------|----------|
| Record GUID | `assertGuid()` | `/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i` | `content.js:237` |
| Entity set name | `assertEntitySetName()` | `/^[A-Za-z][A-Za-z0-9_]{0,127}$/` | `content.js:262` |
| Entity logical name | `assertEntityLogicalName()` | `/^[a-z][a-z0-9_]{0,127}$/` | `content.js:577` |
| User search query | Single-quote doubling | `'` → `''` (OData string escaping) | `content.js:1107` |
| Date inputs | Format from `<input type="date">` | `YYYY-MM-DD` (HTML5 guarantee) | `popup.js:443-444` |
| User GUID | `GUID_PATTERN.test()` | Same regex as record GUID | `content.js:1522` |

---

## Message Origin Validation

### Service Worker

```javascript
if (sender.id !== chrome.runtime.id) return false;
```

Rejects all messages not originating from this extension.

### Content Script (onMessage)

```javascript
if (sender.id !== chrome.runtime.id) return false;
```

Same check — defense-in-depth.

### Bridge postMessage (content → page)

```javascript
if (event.source !== window) return;
if (!EXPECTED_ORIGIN || event.origin !== EXPECTED_ORIGIN) return;
```

- `event.source === window` — ensures the message comes from the same window (not a cross-origin iframe)
- `event.origin === EXPECTED_ORIGIN` — ensures the message comes from the same HTTP origin

### Port Messages

Port names are checked (`"audit-export"` or `"user-audit-export"`). Payload fields are type-checked. GUIDs validated before use in API calls.

---

## Authentication Model

The extension does **not** handle authentication directly. It inherits the user's active Dynamics 365 session:

- All Dataverse API calls use `credentials: "include"` on `fetch()` — the browser automatically sends the MSCRM session cookie
- No `Authorization` header is set anywhere in the codebase
- No tokens, passwords, or credentials are stored or logged
- The extension only sees data the user already has permission to access via their Dynamics session

---

## Permissions Justification

| Permission | Why It's Needed | Scope |
|-----------|----------------|-------|
| `storage` | Save UI preferences and session metadata locally | `chrome.storage.local` only |
| `activeTab` | Access the active tab's URL and communicate with its content script | Only when popup is open |
| Host permissions (`*.crm.dynamics.com`, etc.) | Make same-origin Dataverse API calls from content scripts | 18 known Dynamics 365 CDN domains |

No `tabs`, `history`, `cookies`, `webRequest`, or `notifications` permissions are requested.

---

## Audit Trail

All `fetch()` calls in the extension target the same-origin Dataverse Web API. The browser's Network tab can confirm:

1. No requests to domains other than the current Dynamics 365 org
2. No requests to `chrome-extension://` URLs (except the bridge script load)
3. No WebSocket connections
4. No background sync or push notifications

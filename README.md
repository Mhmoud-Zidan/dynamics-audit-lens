# Dynamics Audit Lens

A **strictly local**, Manifest V3 Chrome Extension for auditing and inspecting Microsoft Dynamics 365 / Dataverse pages.  
Built with **Vite 4 + @crxjs/vite-plugin + Vanilla ES6**.

> **Privacy guarantee:** No data ever leaves your browser. Everything is stored in `chrome.storage.local` on your own machine.

---

## Project structure

```
dynamics-audit-lens/
├── manifest.json                  ← MV3 manifest (source of truth for @crxjs)
├── vite.config.js
├── package.json
├── public/
│   └── icons/                     ← Add icon16.png, icon48.png, icon128.png here
└── src/
    ├── popup/
    │   ├── popup.html
    │   ├── popup.js
    │   └── popup.css
    ├── content/
    │   └── content.js             ← Injected into *.crm*.dynamics.com pages
    └── background/
        └── service-worker.js      ← MV3 background service worker
```

---

## Quick start

```bash
# 1. Install dependencies
npm install

# 2. Development build (rebuilds on file changes)
npm run dev

# 3. Production build
npm run build
```

### Load the unpacked extension in Chrome

1. Open `chrome://extensions/`
2. Enable **Developer mode** (top-right toggle)
3. Click **Load unpacked** → select the `dist/` folder
4. Navigate to any `*.crm.dynamics.com` URL — the badge turns green when the content script activates

---

## Adding a custom Dataverse domain

If your organisation uses a vanity domain (e.g. `https://crm.contoso.com`), add it to **both** arrays in `manifest.json`:

```json
"content_scripts": [{ "matches": ["*://crm.contoso.com/*", ...] }],
"host_permissions":               ["*://crm.contoso.com/*", ...]
```

Then run `npm run build` again and reload the extension.

---

## Security design

| Concern              | Mitigation                                                                                   |
| -------------------- | -------------------------------------------------------------------------------------------- |
| Data exfiltration    | `connect-src 'self'` in CSP; no `fetch`/XHR to external origins in any script                |
| XSS via page content | Content script uses `.textContent` only, never `innerHTML`; data is sanitised before storage |
| Script injection     | `script-src 'self'` — no `eval`, no remote scripts                                           |
| Double-injection     | IIFE guard with `Object.defineProperty` (non-writable) on `window`                           |
| Message spoofing     | Service worker validates `sender.id === chrome.runtime.id` before processing messages        |
| Storage overflow     | Session list capped at 500 entries in the service worker                                     |

---

## Icons

Place PNG icons (transparent background recommended) at:

```
public/icons/icon16.png   (16×16)
public/icons/icon48.png   (48×48)
public/icons/icon128.png  (128×128)
```

A quick way to generate placeholder icons during development:

```bash
# Requires ImageMagick (https://imagemagick.org)
magick -size 16x16  xc:#0078d4 public/icons/icon16.png
magick -size 48x48  xc:#0078d4 public/icons/icon48.png
magick -size 128x128 xc:#0078d4 public/icons/icon128.png
```

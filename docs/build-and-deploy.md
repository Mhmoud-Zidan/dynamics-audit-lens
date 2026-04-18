# Build & Deploy Guide

> Dynamics Audit Lens — Development setup, build process, and Chrome Web Store publishing.

---

## Prerequisites

- **Node.js** 18+ (LTS recommended)
- **npm** 9+
- **Chrome** 116+ (Manifest V3 support)

---

## Quick Start

```bash
# 1. Install dependencies
npm install

# 2. Development build (watches for file changes)
npm run dev

# 3. Production build
npm run build
```

---

## npm Scripts

| Script | Command | Description |
|--------|---------|-------------|
| `dev` | `vite build --watch` | Builds to `dist/` and rebuilds on every file change |
| `build` | `vite build` | One-time production build to `dist/` |
| `clean` | `rimraf dist` | Removes the `dist/` folder |

---

## Loading the Extension in Chrome

1. Run `npm run build`
2. Open `chrome://extensions/`
3. Enable **Developer mode** (toggle in top-right)
4. Click **Load unpacked**
5. Select the `dist/` folder
6. The extension icon appears in the toolbar

> **Dev workflow:** After running `npm run dev`, go to `chrome://extensions/` and click the refresh icon on the extension card after each code change. Content script changes require reloading the Dynamics page.

---

## Build Configuration

### Vite (`vite.config.js`)

```javascript
import { defineConfig } from "vite";
import { crx } from "@crxjs/vite-plugin";
import manifest from "./manifest.json";

export default defineConfig({
  plugins: [crx({ manifest })],
  build: {
    outDir: "dist",
    emptyOutDir: true,
    sourcemap: false,     // No sourcemaps in production
    minify: "oxc",        // Fast minification via oxc
  },
  clearScreen: false,
});
```

### Key Build Details

- **@crxjs/vite-plugin** reads `manifest.json` as the source of truth and handles entry point resolution, HMR support, and asset bundling
- **SheetJS** (`xlsx`) is imported in `popup.js` and bundled by Vite at build time — no CDN loading
- **page-bridge.js** is listed in `web_accessible_resources` and copied as-is (no bundling, runs in page context)
- **Sourcemaps** are disabled to prevent source structure leakage in production

---

## Build Output

```
dist/
├── manifest.json                          Extension manifest
├── service-worker-loader.js               Service worker bootstrap
├── icons/
│   ├── icon16.png
│   ├── icon48.png
│   └── icon128.png
├── src/
│   ├── inject/
│   │   └── page-bridge.js                 Page bridge (unbundled)
│   └── popup/
│       └── popup.html                     Popup HTML
└── assets/
    ├── popup-*.css                        Compiled popup styles
    ├── popup-*.js                         Compiled popup + SheetJS
    ├── content-*.js                       Compiled content script
    └── service-worker-*.js                Compiled service worker
```

---

## Testing Checklist

Before publishing, verify on a real Dynamics 365 environment:

### Record Audit Export
- [ ] Extension loads without console errors
- [ ] Badge appears on Dynamics pages (green for lists, blue for forms)
- [ ] Opening popup shows "Active on: {hostname}" with correct entity name
- [ ] Selecting rows in a grid updates the record count
- [ ] Export button enables when records are selected
- [ ] Clicking export shows progress bar
- [ ] Excel file downloads with correct filename
- [ ] Excel opens with proper columns (RecordID, RecordName, ChangedBy, ChangedDate, Operation, FieldName, OldValue, NewValue)
- [ ] ChangedDate column is recognized as date type in Excel
- [ ] Auto-filter dropdowns work on all columns
- [ ] Option set values show labels (e.g., "Active") not integers (e.g., "1")
- [ ] User names are resolved (not raw GUIDs)
- [ ] Error records appear with FETCH_ERROR operation
- [ ] Closing popup mid-export doesn't crash the page

### User Audit Export
- [ ] "By User" tab shows entity name from page context
- [ ] User search returns matching users
- [ ] Selecting a user shows the chip with their name
- [ ] Clear button removes selected user
- [ ] Export button enables only when user is selected
- [ ] Date range filters work (optional)
- [ ] Export discovers records touched by the user
- [ ] Results only contain changes by the selected user
- [ ] Date range filters correctly exclude out-of-range entries
- [ ] Excel file downloads with username in filename
- [ ] "No audit records found" shown when user has no changes

### Security
- [ ] No requests to external domains in Network tab
- [ ] No CSP violations in console
- [ ] Extension works on different Dynamics regions (crm, crm2, crm3, etc.)

---

## Publishing to Chrome Web Store

### 1. Prepare the package

```bash
npm run build
```

### 2. Create the zip

Zip the contents of the `dist/` folder (not the folder itself):

```bash
cd dist && zip -r ../publish/dynamics_audit_lens.zip . && cd ..
```

### 3. Upload

1. Go to [Chrome Web Store Developer Dashboard](https://chrome.google.com/webstore/devconsole)
2. Pay the one-time $5 registration fee (if first time)
3. Click **New Item** → upload the zip
4. Fill in listing details:
   - Name: Dynamics Audit Lens
   - Short description: Local-only audit export tool for Microsoft Dynamics 365
   - Category: Developer Tools / Productivity
   - Language: English
5. Upload screenshots (1280x800 or 640x400)
6. Add privacy policy URL (link to `PRIVACY.md` in your repo)
7. Submit for review

---

## Version Management

When releasing a new version:

1. Update `version` in both `package.json` and `manifest.json`
2. Update `PRIVACY.md` date if privacy policy changes
3. Commit and tag: `git tag v1.x.x`
4. Build and re-zip
5. Upload to Chrome Web Store Developer Dashboard

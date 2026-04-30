# Code Review Report — RTL / Arabic Language Support

**Verdict:** Hard Fail — The extension has zero RTL/Arabic support. Every UI string is hardcoded English, all CSS layout is LTR-only, and no i18n infrastructure exists.

**Stats:** 12 Critical issues, 24 Minor issues

---

## Executive Summary

To add Arabic (RTL) support, the extension needs:

1. **i18n infrastructure** — Chrome's `chrome.i18n` API (`_locales/ar/messages.json`)
2. **RTL-aware CSS** — Replace all `left`/`right`/`margin-left`/`padding-left` with logical properties (`start`/`end`/`inline-start`/`inline-end`)
3. **Arabic font** — Add Dubai font to the font stack
4. **Dynamic `dir` attribute** — Set `<html dir="rtl" lang="ar">` when Arabic is active
5. **~80 hardcoded English strings** extracted to `_locales/ar/messages.json`

---

## Findings

### A. Architecture / Infrastructure (Critical)

| # | File | Line | Issue | Fix |
|---|------|------|-------|-----|
| A1 | `manifest.json` | — | No `default_locale` field. Chrome i18n requires it. | Add `"default_locale": "en"` |
| A2 | `_locales/` | — | Directory does not exist. | Create `_locales/en/messages.json` and `_locales/ar/messages.json` |
| A3 | `popup.html` | 2 | `<html lang="en">` is hardcoded, no `dir` attribute. | Add dynamic `dir` and `lang` based on user locale: `<html lang="ar" dir="rtl">` |
| A4 | `popup.js` | — | No language detection or toggle mechanism. | Add `LANG_STORAGE_KEY`, detect `chrome.i18n.getUILanguage()`, add language toggle to settings menu |
| A5 | `manifest.json` | 3-6 | `"name"` and `"description"` are hardcoded English. | Use `"__MSG_extName__"` and `"__MSG_extDescription__"` with entries in messages.json |

---

### B. CSS Layout — RTL-Unsafe Properties (Critical)

Every instance below will render broken in RTL. All must be converted to CSS **logical properties** or overridden with `[dir="rtl"]`.

| # | File | Line | Property | Current | Fix |
|---|------|------|----------|---------|-----|
| B1 | `popup.css` | 25 | Font stack | `"Segoe UI", system-ui, -apple-system, sans-serif` | Add Dubai font: `"Dubai", "Segoe UI", system-ui, -apple-system, sans-serif` |
| B2 | `popup.css` | 175 | `.status` padding | `padding: 9px 12px 9px 16px` | `padding: 9px 12px 9px 16px` → `padding-block: 9px; padding-inline: 12px; padding-inline-start: 16px;` |
| B3 | `popup.css` | 189-196 | `.status::before` left stripe | `left: 0; border-radius: 8px 0 0 8px` | `inset-inline-start: 0; border-start-start-radius: 8px; border-end-start-radius: 8px;` |
| B4 | `popup.css` | 504 | `.search-dropdown__item-email` spacing | `margin-left: 6px` | `margin-inline-start: 6px` |
| B5 | `popup.css` | 602 | `.header-actions` push-right | `margin-left: auto` | `margin-inline-start: auto` |
| B6 | `popup.css` | 688 | `.settings-menu` position | `right: 0` | `inset-inline-end: 0` |
| B7 | `popup.css` | 711 | `.settings-menu__item` alignment | `text-align: left` | `text-align: start` |
| B8 | `popup.css` | 875 | `.modal__section-label` padding | `padding-left: 2px` | `padding-inline-start: 2px` |
| B9 | `popup.css` | 912 | `.modal__link-row` alignment | `text-align: left` | `text-align: start` |
| B10 | `popup.css` | 94-99 | `.accent-bar` gradient | `background: linear-gradient(90deg, ...)` | Use `to right` (which auto-flips in RTL context) or `100deg` for RTL |
| B11 | `popup.css` | 294 | `.btn--primary` gradient | `linear-gradient(135deg, ...)` | Consider RTL-neutral angle or `[dir="rtl"]` override |

---

### C. All Hardcoded English Strings in `popup.html`

Every string below needs extraction to `_locales/ar/messages.json`.

| # | Line | Current English | Arabic Translation |
|---|------|----------------|--------------------|
| C1 | 16 | `Dynamics Audit Lens` | `عدسة تدقيق Dynamics` |
| C2 | 20 | `title="Fill form with sample data"` | `title="تعبئة النموذج ببيانات تجريبية"` |
| C3 | 26 | `title="Settings"` | `title="الإعدادات"` |
| C4 | 34 | `Light Mode` | `الوضع الفاتح` |
| C5 | 39 | `Audit Settings` | `إعدادات التدقيق` |
| C6 | 44 | `About` | `حول` |
| C7 | 52 | `Records` (tab) | `السجلات` |
| C8 | 53 | `Users` (tab) | `المستخدمون` |
| C9 | 60 | `Waiting for Dynamics page…` | `في انتظار صفحة Dynamics…` |
| C10 | 63 | `0 records selected` | `0 سجلات محددة` |
| C11 | 72 | `Export to Excel` | `تصدير إلى Excel` |
| C12 | 79 | `Preparing…` | `جارٍ التحضير…` |
| C13 | 84 | `Waiting for Dynamics page…` | `في انتظار صفحة Dynamics…` |
| C14 | 91 | `placeholder="Entity name (e.g. account, contact…)"` | `placeholder="اسم الكيان (مثال: account, contact…)"` |
| C15 | 103 | `placeholder="Search user by name or email…"` | `placeholder="بحث عن مستخدم بالاسم أو البريد…"` |
| C16 | 112 | `title="Remove"` | `title="إزالة"` |
| C17 | 117 | `From` (date label) | `من` |
| C18 | 121 | `To` (date label) | `إلى` |
| C19 | 131 | `Export User Audit` | `تصدير تدقيق المستخدم` |
| C20 | 138 | `Preparing…` | `جارٍ التحضير…` |
| C21 | 144 | `All data stays in your browser.` | `جميع البيانات تبقى في متصفحك.` |
| C22 | 151 | `About` (modal title) | `حول` |
| C23 | 152 | `title="Close"` | `title="إغلاق"` |
| C24 | 160 | `Dynamics Audit Lens` (hero) | `عدسة تدقيق Dynamics` |
| C25 | 167-169 | `Local-only audit export for Microsoft Dynamics 365 & Dataverse. Zero data exfiltration — all processing stays in your browser.` | `تصدير تدقيق محلي فقط لـ Microsoft Dynamics 365 و Dataverse. لا تسريب للبيانات — جميع المعالجة تبقى في متصفحك.` |
| C26 | 174 | `Contact` | `تواصل` |
| C27 | 183 | `LinkedIn` | `LinkedIn` |
| C28 | 199 | `GitHub` | `GitHub` |
| C29 | 216 | `MIT License` | `رخصة MIT` |
| C30 | 217 | `Free & Open Source — No warranty` | `مجاني ومفتوح المصدر — بدون ضمان` |
| C31 | 222 | `Legal` | `قانوني` |
| C32 | 225-233 | MIT legal text | (keep English — standard license text) |
| C33 | 237 | `© 2026 Mahmoud Zidan — Free & Open Source Software` | `© 2026 محمود زيدان — برنامج مجاني ومفتوح المصدر` |

---

### D. All Hardcoded English Strings in `popup.js`

| # | Line | Current English | Arabic Translation |
|---|------|----------------|--------------------|
| D1 | 105 | `Processed ${n} of ${m} records…` | `تمت معالجة ${n} من ${m} سجل…` |
| D2 | 246 | `Cannot access current tab.` | `لا يمكن الوصول للتبويب الحالي.` |
| D3 | 256 | `Not a Dynamics / Dataverse page.` | `ليست صفحة Dynamics / Dataverse.` |
| D4 | 265 | `Could not read page context.` | `تعذرت قراءة سياق الصفحة.` |
| D5 | 276 | `Active on: ${hostname}` | `نشط على: ${hostname}` |
| D6 | 277 | `${n} record(s) selected` | `${n} سجلات محددة` |
| D7 | 287 | `Too many records selected (max ${n}). Narrow your selection.` | `عدد السجلات المحددة كبير جداً (الحد الأقصى ${n}). قلّل التحديد.` |
| D8 | 305 | `Content script not ready. Reload the page.` | `برنامج المحتوى غير جاهز. أعد تحميل الصفحة.` |
| D9 | 356 | `No entities found.` | `لم يتم العثور على كيانات.` |
| D10 | 417 | `No entities found.` | `لم يتم العثور على كيانات.` |
| D11 | 461 | `(unnamed)` | `(بدون اسم)` |
| D12 | 539 | `No audit records found.` | `لم يتم العثور على سجلات تدقيق.` |
| D13 | 542 | `Capped at ${n} rows. Generating file…` | `محدود بـ ${n} صف. جارٍ إنشاء الملف…` |
| D14 | 545 | `Export complete — ${n} row(s).` | `اكتمل التصدير — ${n} صف.` |
| D15 | 554 | `Error: ${msg}` | `خطأ: ${msg}` |
| D16 | 556 | `Export failed.` | `فشل التصدير.` |
| D17 | 564 | `Connection lost. Reload the page and retry.` | `فقد الاتصال. أعد تحميل الصفحة وحاول مجدداً.` |
| D18 | 593 | `Querying audit records…` | `جارٍ الاستعلام عن سجلات التدقيق…` |
| D19 | 611 | `No audit records found for this user.` | `لم يتم العثور على سجلات تدقيق لهذا المستخدم.` |
| D20 | 619 | `Export complete — ${n} row(s).` | `اكتمل التصدير — ${n} صف.` |
| D21 | 631 | `Export failed.` | `فشل التصدير.` |
| D22 | 639 | `Connection lost. Reload the page and retry.` | `فقد الاتصال. أعد تحميل الصفحة وحاول مجدداً.` |
| D23 | 681 | `Filling form fields…` | `جارٍ تعبئة حقول النموذج…` |
| D24 | 689 | `Filled ${n} of ${m} fields (${x} skipped).` | `تم تعبئة ${n} من ${m} حقل (${x} تم تخطيه).` |
| D25 | 699 | `0 fields filled (${x} skipped). All fields may already have values or be read-only.` | `0 حقل تمت تعبئته (${x} تم تخطيه). قد تكون جميع الحقول ممتلئة أو للقراءة فقط.` |
| D26 | 714 | `Failed to fill form data.` | `فشلت تعبئة بيانات النموذج.` |
| D27 | 717 | `Could not reach content script. Reload the page.` | `تعذر الوصول لبرنامج المحتوى. أعد تحميل الصفحة.` |
| D28 | 738-741 | `Dark Mode` / `Light Mode` | `الوضع الداكن` / `الوضع الفاتح` |

---

### E. All Hardcoded English Strings in `content.js`

| # | Line | Current English | Arabic Translation |
|---|------|----------------|--------------------|
| E1 | 916 | `OPERATION_MAP`: `Create`, `Update`, `Delete`, `Access`, `Upsert` | `إنشاء`, `تحديث`, `حذف`, `وصول`, `إدراج/تحديث` |
| E2 | 1635 | `Session expired — please reload the page and re-authenticate.` | `انتهت الجلسة — أعد تحميل الصفحة وسجّل الدخول مجدداً.` |
| E3 | 1638 | `Access denied — you need the "Audit Summary View" (prvReadAuditSummary) privilege.` | `تم رفض الوصول — تحتاج صلاحية "عرض ملخص التدقيق" (prvReadAuditSummary).` |
| E4 | 1640 | `Record not found — it may have been deleted.` | `السجل غير موجود — قد يكون قد تم حذفه.` |
| E5 | 1709 | `Invalid payload.` | `حمولة غير صالحة.` |
| E6 | 1760 | `Discovering records touched by user…` | `جارٍ اكتشاف السجلات التي لمسها المستخدم…` |
| E7 | 1775 | `Found ${n} record(s). Fetching audit history…` | `تم العثور على ${n} سجل. جارٍ جلب سجل التدقيق…` |

---

### F. Design / UX Considerations for Arabic

| # | Area | Issue | Recommendation |
|---|------|-------|---------------|
| F1 | Font | Current font `"Segoe UI"` has poor Arabic rendering | Use **Dubai** as primary Arabic font: `"Dubai", "Segoe UI", system-ui, sans-serif`. Dubai WSB (or Dubai Regular) is freely available on Windows and renders Arabic beautifully. |
| F2 | Popup width | `body { width: 340px }` may be too narrow for Arabic text (which can be ~20% wider) | Consider `width: 360px` when `dir="rtl"`, or use `min-width: 340px; max-width: 380px` |
| F3 | SVG icons | Download arrow (↓) and chevron (›) are direction-sensitive | The chevron `›` on link rows should flip to `‹` in RTL. The download arrow is direction-neutral. |
| F4 | Settings gear | Positioned via `margin-left: auto` (pushes right) | Already fixed by B5 above. Gear stays right in LTR, moves left in RTL — correct behavior. |
| F5 | Date inputs | `<input type="date">` renders browser-native. Chrome shows it in the user's locale automatically. | No change needed — Chrome handles this. |
| F6 | Tab order | Tab text "Records" / "Users" — Arabic text is right-aligned naturally with `text-align: start` | Ensure `.tab` uses `text-align: center` (already implicit via flex centering). |
| F7 | Progress bar fill | Fills left-to-right via `width: N%` | Should fill right-to-left in RTL. Add `[dir="rtl"] .progress-bar__fill { transform: scaleX(-1); }` on the bar container, or set `direction: ltr` on the progress bar itself (numbers are universal). |
| F8 | Dropdown search | `.search-dropdown` positioned `left: 0; right: 0` | Already safe — stretches full width. |

---

### G. Recommended Implementation Plan

#### Phase 1 — i18n Infrastructure

1. Add `"default_locale": "en"` to `manifest.json`
2. Create `_locales/en/messages.json` with all English strings
3. Create `_locales/ar/messages.json` with all Arabic translations
4. Replace all hardcoded strings in `popup.html` with `data-i18n` attributes
5. Replace all hardcoded strings in `popup.js` and `content.js` with `chrome.i18n.getMessage("key")`

#### Phase 2 — RTL CSS

1. Convert all physical properties (`left`, `right`, `margin-left`, `padding-left`) to logical properties (`inset-inline-start`, `margin-inline-start`, `padding-inline-start`, `text-align: start`)
2. Add Dubai font to `:root { --font }`
3. Add RTL-specific overrides in `[dir="rtl"]` selectors for gradients and icons

#### Phase 3 — Language Detection & Toggle

1. Detect `chrome.i18n.getUILanguage()` on install — if `ar`, default to Arabic
2. Add language toggle to settings dropdown (Arabic / English)
3. Persist choice to `chrome.storage.local`
4. On popup load: read preference, set `document.documentElement.dir` and `document.documentElement.lang`

#### Phase 4 — `_locales/ar/messages.json` Structure

```json
{
  "extName": {
    "message": "عدسة تدقيق Dynamics"
  },
  "extDescription": {
    "message": "أداة تدقيق وفحص محلية فقط لـ Microsoft Dynamics 365 / Dataverse. لا تغادر البيانات متصفحك."
  },
  "tabRecords": {
    "message": "السجلات"
  },
  "tabUsers": {
    "message": "المستخدمون"
  },
  "exportToExcel": {
    "message": "تصدير إلى Excel"
  },
  "exportUserAudit": {
    "message": "تصدير تدقيق المستخدم"
  },
  "waitingForDynamics": {
    "message": "في انتظار صفحة Dynamics…"
  },
  "activeOn": {
    "message": "نشط على: $hostname$",
    "placeholders": { "hostname": { "content": "$1" } }
  },
  "recordsSelected": {
    "message": "$count$ سجلات محددة",
    "placeholders": { "count": { "content": "$1" } }
  },
  "preparing": {
    "message": "جارٍ التحضير…"
  },
  "exportComplete": {
    "message": "اكتمل التصدير — $count$ صف.",
    "placeholders": { "count": { "content": "$1" } }
  },
  "exportFailed": {
    "message": "فشل التصدير."
  },
  "noAuditRecords": {
    "message": "لم يتم العثور على سجلات تدقيق."
  },
  "connectionLost": {
    "message": "فقد الاتصال. أعد تحميل الصفحة وحاول مجدداً."
  },
  "settings": {
    "message": "الإعدادات"
  },
  "lightMode": {
    "message": "الوضع الفاتح"
  },
  "darkMode": {
    "message": "الوضع الداكن"
  },
  "about": {
    "message": "حول"
  },
  "auditSettings": {
    "message": "إعدادات التدقيق"
  },
  "from": {
    "message": "من"
  },
  "to": {
    "message": "إلى"
  },
  "allDataStaysLocal": {
    "message": "جميع البيانات تبقى في متصفحك."
  },
  "fillFormSample": {
    "message": "تعبئة النموذج ببيانات تجريبية"
  },
  "remove": {
    "message": "إزالة"
  },
  "close": {
    "message": "إغلاق"
  },
  "contact": {
    "message": "تواصل"
  },
  "legal": {
    "message": "قانوني"
  },
  "processedOf": {
    "message": "تمت معالجة $done$ من $total$ سجل…",
    "placeholders": { "done": { "content": "$1" }, "total": { "content": "$2" } }
  },
  "queryingAudit": {
    "message": "جارٍ الاستعلام عن سجلات التدقيق…"
  },
  "discoveringRecords": {
    "message": "جارٍ اكتشاف السجلات التي لمسها المستخدم…"
  },
  "foundRecordsFetching": {
    "message": "تم العثور على $count$ سجل. جارٍ جلب سجل التدقيق…",
    "placeholders": { "count": { "content": "$1" } }
  },
  "opCreate": { "message": "إنشاء" },
  "opUpdate": { "message": "تحديث" },
  "opDelete": { "message": "حذف" },
  "opAccess": { "message": "وصول" },
  "opUpsert": { "message": "إدراج/تحديث" }
}
```

---

### H. RTL-Specific CSS Patch (add to `popup.css`)

```css
/* ── RTL Support ────────────────────────────────────── */

:root {
  --font: "Dubai", "Segoe UI", system-ui, -apple-system, sans-serif;
}

[dir="rtl"] .status::before {
  left: auto;
  right: 0;
  border-radius: 0 8px 8px 0;
}

[dir="rtl"] .status {
  padding: 9px 16px 9px 12px;
}

[dir="rtl"] .search-dropdown__item-email {
  margin-left: 0;
  margin-right: 6px;
}

[dir="rtl"] .header-actions {
  margin-left: 0;
  margin-right: auto;
}

[dir="rtl"] .settings-menu {
  right: auto;
  left: 0;
}

[dir="rtl"] .settings-menu__item,
[dir="rtl"] .modal__link-row {
  text-align: right;
}

[dir="rtl"] .modal__section-label {
  padding-left: 0;
  padding-right: 2px;
}

[dir="rtl"] .modal__link-chevron {
  transform: scaleX(-1);
}

[dir="rtl"] .progress-bar__fill {
  /* Progress fills right-to-left in RTL */
  float: right;
}
```

Or preferably, convert to logical properties throughout and skip the overrides entirely.

---

### I. Files That Need No Changes

| File | Reason |
|------|--------|
| `src/background/service-worker.js` | No user-facing strings, no layout |
| `src/inject/page-bridge.js` | Internal bridge — no UI |
| `src/inject/fill-data.js` | Sample data generator — no UI strings |
| `vite.config.js` | Build config |

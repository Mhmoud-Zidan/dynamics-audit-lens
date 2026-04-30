# Dynamics Audit Lens — Pro Feature Roadmap & Business Plan

> **Product:** Dynamics Audit Lens
> **Current Version:** 2.0.0 (Free Tier)
> **Document Purpose:** Specification for premium features, tiered monetization strategy, and architectural implementation details for transitioning the extension into a commercial SaaS/B2B product.

---

## 1. Executive Summary

The base extension (v2.0.0) provides immense value by solving the fundamental problem of extracting and formatting Dataverse audit history into Excel. However, IT administrators, compliance officers, and Dynamics 365 consultants require advanced workflow automation, visual analytics, and data manipulation tools that justify a commercial license.

This document outlines the feature roadmap for transforming Dynamics Audit Lens into a **Freemium SaaS Product**. By gating advanced features behind a subscription layer—validated via a lightweight license key mechanism—we target a B2B audience with high willingness-to-pay.

### Proposed Pricing Structure

| Tier | Price (USD) | Target Audience | Core Value Proposition |
|------|-------------|-----------------|------------------------|
| **Free** | $0 | Junior Admins, single-use cases | Basic XLSX export, 250 record cap, standard metadata resolution. |
| **Pro** | $12/mo | Power Users, Consultants | Saved presets, visual HTML diffs, multi-layout exports, higher caps (1000+). |
| **Business** | $30/mo | IT Departments, Managers | Cross-entity auditing, local analytics dashboard, scheduled export reminders. |
| **Enterprise** | $50+/mo | Large Corporations, Compliance | Team sharing, actionable undo capabilities, PCF component embedding. |

---

## 2. Technical Implementation: Licensing Architecture

To maintain the strict Content Security Policy (`connect-src 'self'`) and zero-telemetry constraint, the licensing model must operate locally or validate directly against a Stripe-backed API without third-party analytics scripts.

*   **Gateway:** User purchases via Stripe Checkout on a marketing landing page.
*   **Delivery:** User receives a 25-character License Key.
*   **Validation:** User enters key in the Extension Settings tab. The extension sends a `fetch` request to the validation API. *(Note: This will require updating the CSP in `manifest.json` to allow `connect-src` to the specific license validation domain).*
*   **Caching:** Validated license status and expiration date are stored in `chrome.storage.local`.
*   **Feature Gating:** Functions in `popup.js` and `content.js` check `isProUser()` before executing advanced logic.

---

## 3. Phase 13 — Pro Tier Features (Efficiency & Formatting)

### 13.1 Saved Export Presets & Templates

**User Value:** Admins and consultants frequently run the exact same audit reports weekly (e.g., "Show me all Opportunity changes every Monday"). Manually entering entity names, dates, and filters is tedious.

**Technical Implementation:**
*   **Storage:** Utilize `chrome.storage.local` (or `sync` for cross-device roaming) to store an array of `PresetProfile` objects.
*   **UI:** Add a "Presets" dropdown at the top of the Records and Users tabs.
*   **Schema:**
    ```typescript
    interface PresetProfile {
      id: string; // UUID
      name: string; // "Weekly Opportunity Audit"
      tab: "records" | "users";
      entityLogicalName: string;
      userGuid?: string;
      dateOffset?: number; // e.g., -7 days from today
    }
    ```
*   **Logic:** When a preset is selected, `popup.js` auto-fills the entity, user, and dynamically calculates the date range based on the offset, instantly enabling the Export button.

### 13.2 Interactive HTML Audit Diff Viewer

**User Value:** Flat Excel rows are data-heavy. A visual, "Git-style" diff view allows users to instantly see the magnitude of changes without cross-referencing rows.

**Technical Implementation:**
*   **Generation:** `popup.js` will include a new function `generateHtmlDiff(rows)` instead of `generateExcel()`.
*   **Layout:** The generated HTML will use CSS flexbox to create a timeline.
*   **Visuals:** Old values are highlighted in soft red (`background: #ffd7d7`), new values in soft green (`background: #d7ffd7`).
*   **Delivery:** Instead of an `.xlsx` blob, the extension creates a `text/html` Blob and downloads it as `AuditDiff_{entity}_{date}.html`.
*   **Privacy:** This perfectly aligns with the offline-only architecture. No external CSS/JS is needed; the CSS is injected inline into the generated HTML file.

### 13.3 Custom Column Selection & Export Layouts

**User Value:** The current long-format (one row per changed field) is great for databases, but terrible for human readability.

**Technical Implementation:**
*   **UI:** A new "Layout" dropdown in the popup UI.
*   **Transformations:** Add a `pivotAuditData(rows)` function in `popup.js`.
*   **Wide Format (Pivot):** Pivots the data so one row = one audit event. Columns become the dynamic `FieldNames`.
    ```json
    // Before (Long - v2.0.0)
    { Record: "Contoso", Operation: "Update", Field: "Revenue", Old: "100", New: "200" }
    // After (Wide - Pro)
    { Record: "Contoso", Operation: "Update", Revenue_Old: "100", Revenue_New: "200" }
    ```
*   **Column Picker:** Allow users to select which of the default 8 columns (`RecordID`, `RecordName`, etc.) to actually include in the export, stripping unnecessary data.

### 13.4 Advanced Export Formats (CSV, JSON)

**User Value:** Power users and external BI tools (Power BI, Tableau) prefer structured JSON or CSV over binary XLSX.

**Technical Implementation:**
*   **UI:** Radio buttons or a dropdown next to the Export button to select format (XLSX, CSV, JSON).
*   **Logic:** SheetJS natively supports CSV export (`XLSX.utils.sheet_to_csv`). For JSON, simply `JSON.stringify(rows, null, 2)` is used.
*   **MIME Types:** Update the Blob generator (`application/json`, `text/csv`).

---

## 4. Phase 14 — Business Tier Features (Analytics & Scope)

### 14.1 Cross-Entity & Relationship Auditing

**User Value:** Currently, audits are siloed per entity. If an admin wants to see what happened to an Account *and* its child Contacts, they must run two exports. Relationship auditing allows querying the Dataverse Web API for related records.

**Technical Implementation:**
*   **Metadata Lookup:** Use the `EntityMetadata` cache to find 1:N and N:1 relationships (via `OneToManyRelationships` and `ManyToOneRelationships`).
*   **UI:** A toggle switch: "Include related entities?" with a dropdown to select the relationship path (e.g., `account -> contact`).
*   **API Flow:** `content.js` fetches the primary records, extracts the `accountid`, then automatically spawns a secondary batch fetch for `contacts` where `parentcustomerid = accountid`.

### 14.2 Local Audit Analytics Dashboard

**User Value:** Managers need high-level trends, not raw data. "Who is making the most changes?" "Which fields are changed most often?"

**Technical Implementation:**
*   **UI:** A new "Dashboard" tab in the popup (or opened in a new Chrome tab via `chrome.tabs.create` for more screen real estate).
*   **Library:** Bundle a lightweight, offline-first charting library (e.g., Chart.js or Primitives) into the extension build.
*   **Data Source:** Run the standard audit fetch, but instead of formatting into rows, `content.js` reduces the data into aggregate metrics:
    ```typescript
    interface AuditMetrics {
      changesByUser: Map<string, number>;
      changesByField: Map<string, number>;
      changesByHour: number[]; // Heatmap data
    }
    ```
*   **Privacy:** 100% local rendering. No data leaves the browser. The dashboard is generated and destroyed instantly.

### 14.3 Snapshot Comparison (Time Travel)

**User Value:** Allows an admin to say, "Show me the state of this record on March 1st vs. Today."

**Technical Implementation:**
*   **Logic:** `content.js` iterates backward through the `FormattedAuditRow[]` array (which is sorted chronologically).
*   **Algorithm:** Starting from the current known values, the engine applies the `OldValue` for each event timestamped *after* the target date, effectively "rewinding" the record state.
*   **Output:** Generates an HTML report showing the reconstructed state of the record at Date A vs. Date B.

---

## 5. Phase 15 — Enterprise Tier Features (Actionability & Teams)

### 15.1 Guarded One-Click Revert (Undo)

**User Value:** The ultimate admin tool. If a user accidentally mass-updates records, the admin can click "Revert" on the audit log.

**Technical Implementation:**
*   **CRITICAL SECURITY:** This transitions the extension from a `GET`/`POST` (query) tool to a `PATCH` (modify) tool.
*   **UI:** A "Revert" button next to each row in the HTML Diff Viewer.
*   **API Call:** `content.js` constructs a Dataverse Web API `PATCH` request to the specific entity GUID, sending *only* the `OldValue` of the targeted field.
*   **Safeguards:**
    *   Must be explicitly enabled in the Settings tab (off by default).
    *   Requires a secondary confirmation dialog built in `popup.js` ("WARNING: You are about to modify record...").
    *   Requires the user to possess `prvWriteEntity` permissions in Dynamics (the API will return 403 if they lack it, acting as a natural safeguard).

### 15.2 Dataverse Model-Driven App Embedding (PCF Control)

**User Value:** Enterprises prefer native UI over Chrome extensions. By wrapping the core auditing logic into a Power Apps Component Framework (PCF) control, it can be embedded directly inside a Dynamics 365 form.

**Technical Implementation:**
*   **Refactoring:** Extract the Dataverse API logic and formatting engine from `content.js` into a standalone `auditEngine.ts` module.
*   **PCF Wrapper:** Create a React/Fluent UI-based PCF project that imports `auditEngine`.
*   **Distribution:** Sell this as a managed solution uploaded to the Dataverse environment, bypassing the Chrome Web Store entirely.

### 15.3 Team Shared Presets & Centralized Billing

**User Value:** IT managers want to standardize tools. They buy 10 seats, assign them to the team, and push a standardized set of Audit Presets.

**Technical Implementation:**
*   **Backend:** A simple Node.js/Express or Azure Functions backend integrated with Stripe for volume licensing.
*   **Sync Mechanism:** The extension periodically checks the license API for team-level shared configurations (`sharedPresets: []`), merging them into the local UI and marking them with a "Team" badge.

---

## 6. Development Priority (Highest ROI First)

To validate the commercial viability of the product with minimal development overhead, the features should be implemented in the following order:

1.  **Advanced Export Formats (CSV, JSON):** Easiest to build (1-2 hours). Instant perceived value for data analysts.
2.  **Saved Export Presets:** Fast to build. Immediately creates user lock-in and habit formation.
3.  **Custom Column Selection & Wide Format:** High administrative value. Solves the biggest complaint about raw audit logs.
4.  **Interactive HTML Audit Diff Viewer:** The "Showpiece" feature. High wow-factor for marketing and sales demos.
5.  **Local Analytics Dashboard:** Justifies the recurring subscription cost.
6.  **One-Click Revert:** The ultimate enterprise hook. Requires careful UX design but commands premium pricing.
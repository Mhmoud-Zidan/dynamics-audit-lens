# Privacy Policy for Dynamics Audit Lens
**Last Updated: March 2026**

## 1. Overview
Dynamics Audit Lens ("the Extension") is a productivity tool for Microsoft Dynamics 365. We prioritize user privacy by utilizing a "Local-First" architecture.

## 2. Data Collection and Usage
* **No Server-Side Storage:** The Extension does not collect, store, or transmit any personal data or CRM records to any external servers or third parties.
* **Local Processing:** All data retrieved from the Microsoft Dataverse Web API is processed entirely within your browser's local memory. This data is used solely to generate the Excel (.xlsx) export and is cleared once the tab is closed.
* **Authentication:** The Extension inherits your existing, active Dynamics 365 session. It does not see, store, or request your login credentials.

## 3. Permissions Justification
* **Host Permissions (`*.dynamics.com`, `*.powerapps.com`):** Required to allow the Extension to communicate with your specific Dynamics 365 environment to fetch audit logs.
* **Storage:** Used only to save your UI preferences (e.g., column toggle settings) locally in your browser.

## 4. Third-Party Services
The Extension does not use any third-party analytics, tracking scripts, or remote code. All libraries (e.g., ExcelJS) are bundled locally within the extension package.

## 5. Contact
For questions regarding this policy, please open an issue on our GitHub repository: [INSERT_YOUR_REPO_URL_HERE]

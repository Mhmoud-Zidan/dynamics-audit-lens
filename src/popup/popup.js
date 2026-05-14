import * as XLSX from "xlsx";

const MAX_EXPORT_RECORDS = 250;
const MAX_EXPORT_ROWS = 100_000;
const SEARCH_DEBOUNCE_MS = 300;
const MIN_SEARCH_LENGTH = 2;
const STATE_STORAGE_KEY = "userAuditState";
const THEME_STORAGE_KEY = "theme";
const LANG_STORAGE_KEY = "lang";

const I18N = {
  en: {
    extName: "Dynamics Audit Lens",
    fillFormTitle: "Fill form with sample data",
    settingsTitle: "Settings",
    lightMode: "Light Mode",
    darkMode: "Dark Mode",
    auditSettings: "Audit Settings",
    about: "About",
    tabRecords: "Records",
    tabUsers: "Users",
    tabDeleted: "Deleted",
    deletedHint: "Find records that were deleted from the audit log for a specific entity and date range.",
    filterByUser: "Filter by user (optional)",
    exportDeleted: "Export Deleted Records",
    queryingDeletions: "Querying deletion audits…",
    fetchingDeletedNames: "Found $count$ deletion(s). Looking up record names…",
    noDeletionsFound: "No deletions found for the chosen filters.",
    deletedExportComplete: "Export complete — $count$ deleted record(s).",
    selectEntityAndDates: "Select an entity and a date range to enable export.",
    feedbackTitle: "A note from the developer",
    feedbackText: "Hi! I built this tool. Got feedback or a feature wish? I'd love to hear it.",
    feedbackEmail: "Email",
    feedbackLinkedIn: "LinkedIn",
    feedbackEmailSubject: "Dynamics Audit Lens — Feedback",
    feedbackEmailBody: "Hi Mahmoud,\n\n(Your feedback, wishlist, or bug report here.)\n\n— Sent from Dynamics Audit Lens v$version$",
    dismiss: "Dismiss",
    sendFeedback: "Send Feedback",
    feedbackModalDesc: "I read every message. Type below, then choose how to send it — or copy and paste anywhere.",
    feedbackKindLabel: "Type",
    feedbackKindFeature: "Feature wishlist",
    feedbackKindBug: "Bug report",
    feedbackKindGeneral: "General feedback",
    feedbackYourEmail: "Your email",
    optional: "(optional)",
    feedbackMessage: "Message",
    feedbackMessagePlaceholder: "What's on your mind? Wishlist, bugs, ideas…",
    feedbackMessageRequired: "Message can't be empty.",
    sendVia: "Send via",
    gmail: "Gmail",
    outlook: "Outlook",
    copy: "Copy",
    orEmailMeAt: "Or email me directly at",
    copyAddress: "Copy address",
    addressCopied: "Address copied to clipboard.",
    messageCopied: "Message copied to clipboard — paste it into your mail.",
    copyFailed: "Couldn't access clipboard. Select and copy manually.",
    openingGmail: "Opening Gmail compose…",
    openingOutlook: "Opening Outlook compose…",
    waitingForDynamics: "Waiting for Dynamics page\u2026",
    recordsSelected: "$count$ record(s) selected",
    exportToExcel: "Export to Excel",
    exportEntireView: "Export entire view",
    exportEntireViewWithCount: "Export entire view ($count$)",
    resolvingViewRecords: "Resolving records in view…",
    viewResolveFailed: "Could not load records from this view.",
    preparing: "Preparing\u2026",
    entityPlaceholder: "Entity name (e.g. account, contact\u2026)",
    userSearchPlaceholder: "Search user by name or email\u2026",
    remove: "Remove",
    from: "From",
    to: "To",
    exportUserAudit: "Export User Audit",
    allDataLocal: "All data stays in your browser.",
    close: "Close",
    modalDesc: "Local-only audit export for Microsoft Dynamics\u00a0365 & Dataverse. Zero data exfiltration \u2014 all processing stays in your browser.",
    contact: "Contact",
    license: "License",
    mitLicense: "MIT License",
    freeOpenSource: "Free & Open Source \u2014 No warranty",
    legal: "Legal",
    legalText: "THE SOFTWARE IS PROVIDED \u201cAS IS\u201d, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.",
    copyright: "\u00a9 2026 Mahmoud Zidan \u2014 Free & Open Source Software",
    cannotAccessTab: "Cannot access current tab.",
    notDynamicsPage: "Not a Dynamics / Dataverse page.",
    couldNotReadContext: "Could not read page context.",
    activeOn: "Active on: $hostname$",
    tooManyRecords: "Too many records selected (max $max$). Narrow your selection.",
    contentScriptNotReady: "Content script not ready. Reload the page.",
    noEntitiesFound: "No entities found.",
    unnamed: "(unnamed)",
    noAuditRecords: "No audit records found.",
    noAuditRecordsForUser: "No audit records found for this user.",
    cappedAtRows: "Capped at $count$ rows. Generating file\u2026",
    exportComplete: "Export complete \u2014 $count$ row(s).",
    exportFailed: "Export failed.",
    connectionLost: "Connection lost. Reload the page and retry.",
    queryingAudit: "Querying audit records\u2026",
    fillingFields: "Filling form fields\u2026",
    filledFields: "Filled $filled$ of $total$ fields ($skipped$ skipped).",
    zeroFieldsFilled: "0 fields filled ($skipped$ skipped). All fields may already have values or be read-only.",
    failedFill: "Failed to fill form data.",
    couldNotReachContent: "Could not reach content script. Reload the page.",
    processedOf: "Processed $done$ of $total$ records\u2026",
    lookupIssues: "Lookup issues: $issues$",
    lookupErrors: "Lookup errors: $errors$",
    error: "Error: $msg$",
    discoveringRecords: "Discovering records touched by user\u2026",
    foundRecordsFetching: "Found $count$ record(s). Fetching audit history\u2026",
    sessionExpired: "Session expired \u2014 please reload the page and re-authenticate.",
    accessDeniedAudit: "Access denied \u2014 you need the \"Audit Summary View\" (prvReadAuditSummary) privilege.",
    recordNotFound: "Record not found \u2014 it may have been deleted.",
    invalidPayload: "Invalid payload.",
    opCreate: "Create", opUpdate: "Update", opDelete: "Delete",
    opAccess: "Access", opUpsert: "Upsert",
    language: "Language",
    langEnglish: "English",
    langArabic: "\u0627\u0644\u0639\u0631\u0628\u064a\u0629",
    noUsersFound: "No users found.",
  },
  ar: {
    extName: "\u0639\u062f\u0633\u0629 \u062a\u062f\u0642\u064a\u0642 Dynamics",
    fillFormTitle: "\u062a\u0639\u0628\u0626\u0629 \u0627\u0644\u0646\u0645\u0648\u0630\u062c \u0628\u0628\u064a\u0627\u0646\u0627\u062a \u062a\u062c\u0631\u064a\u0628\u064a\u0629",
    settingsTitle: "\u0627\u0644\u0625\u0639\u062f\u0627\u062f\u0627\u062a",
    lightMode: "\u0627\u0644\u0648\u0636\u0639 \u0627\u0644\u0641\u0627\u062a\u062d",
    darkMode: "\u0627\u0644\u0648\u0636\u0639 \u0627\u0644\u062f\u0627\u0643\u0646",
    auditSettings: "\u0625\u0639\u062f\u0627\u062f\u0627\u062a \u0627\u0644\u062a\u062f\u0642\u064a\u0642",
    about: "\u062d\u0648\u0644 \u0627\u0644\u0628\u0631\u0646\u0627\u0645\u062c",
    tabRecords: "\u0627\u0644\u0633\u062c\u0644\u0627\u062a",
    tabUsers: "\u0627\u0644\u0645\u0633\u062a\u062e\u062f\u0645\u0648\u0646",
    tabDeleted: "\u0627\u0644\u0645\u062d\u0630\u0648\u0641\u0629",
    deletedHint: "\u0627\u0644\u0639\u062b\u0648\u0631 \u0639\u0644\u0649 \u0627\u0644\u0633\u062c\u0644\u0627\u062a \u0627\u0644\u062a\u064a \u062a\u0645 \u062d\u0630\u0641\u0647\u0627 \u0645\u0646 \u0633\u062c\u0644 \u0627\u0644\u062a\u062f\u0642\u064a\u0642 \u0644\u0643\u064a\u0627\u0646 \u0645\u0639\u064a\u0651\u0646 \u062e\u0644\u0627\u0644 \u0646\u0637\u0627\u0642 \u062a\u0627\u0631\u064a\u062e \u0645\u062d\u062f\u062f.",
    filterByUser: "\u062a\u0635\u0641\u064a\u0629 \u062d\u0633\u0628 \u0627\u0644\u0645\u0633\u062a\u062e\u062f\u0645 (\u0627\u062e\u062a\u064a\u0627\u0631\u064a)",
    exportDeleted: "\u062a\u0635\u062f\u064a\u0631 \u0627\u0644\u0633\u062c\u0644\u0627\u062a \u0627\u0644\u0645\u062d\u0630\u0648\u0641\u0629",
    queryingDeletions: "\u062c\u0627\u0631\u064d \u0627\u0644\u0627\u0633\u062a\u0639\u0644\u0627\u0645 \u0639\u0646 \u0633\u062c\u0644\u0627\u062a \u0627\u0644\u062d\u0630\u0641\u2026",
    fetchingDeletedNames: "\u062a\u0645 \u0627\u0644\u0639\u062b\u0648\u0631 \u0639\u0644\u0649 $count$ \u0639\u0645\u0644\u064a\u0629 \u062d\u0630\u0641. \u062c\u0627\u0631\u064d \u062c\u0644\u0628 \u0623\u0633\u0645\u0627\u0621 \u0627\u0644\u0633\u062c\u0644\u0627\u062a\u2026",
    noDeletionsFound: "\u0644\u0627 \u062a\u0648\u062c\u062f \u0639\u0645\u0644\u064a\u0627\u062a \u062d\u0630\u0641 \u0636\u0645\u0646 \u0627\u0644\u0645\u0639\u0627\u064a\u064a\u0631 \u0627\u0644\u0645\u062e\u062a\u0627\u0631\u0629.",
    deletedExportComplete: "\u0627\u0643\u062a\u0645\u0644 \u0627\u0644\u062a\u0635\u062f\u064a\u0631 \u2014 $count$ \u0633\u062c\u0644 \u0645\u062d\u0630\u0648\u0641.",
    selectEntityAndDates: "\u0627\u062e\u062a\u0631 \u0643\u064a\u0627\u0646\u0627\u064b \u0648\u0646\u0637\u0627\u0642\u0627\u064b \u0632\u0645\u0646\u064a\u0627\u064b \u0644\u062a\u0641\u0639\u064a\u0644 \u0627\u0644\u062a\u0635\u062f\u064a\u0631.",
    feedbackTitle: "\u0643\u0644\u0645\u0629 \u0645\u0646 \u0627\u0644\u0645\u0637\u0648\u0651\u0631",
    feedbackText: "\u0645\u0631\u062d\u0628\u0627\u064b! \u0623\u0646\u0627 \u0645\u0637\u0648\u0651\u0631 \u0647\u0630\u0647 \u0627\u0644\u0623\u062f\u0627\u0629. \u0647\u0644 \u0644\u062f\u064a\u0643 \u0645\u0644\u0627\u062d\u0638\u0629 \u0623\u0648 \u0637\u0644\u0628 \u0644\u0645\u064a\u0632\u0629 \u062c\u062f\u064a\u062f\u0629\u061f \u064a\u0633\u0639\u062f\u0646\u064a \u0633\u0645\u0627\u0639 \u0631\u0623\u064a\u0643.",
    feedbackEmail: "\u0628\u0631\u064a\u062f \u0625\u0644\u0643\u062a\u0631\u0648\u0646\u064a",
    feedbackLinkedIn: "\u0644\u064a\u0646\u0643\u062f\u0625\u0646",
    feedbackEmailSubject: "\u0645\u0644\u0627\u062d\u0638\u0627\u062a \u0639\u0646 \u0625\u0636\u0627\u0641\u0629 Dynamics Audit Lens",
    feedbackEmailBody: "\u0645\u0631\u062d\u0628\u0627\u064b \u0645\u062d\u0645\u0648\u062f\u060c\n\n(\u0627\u0643\u062a\u0628 \u0645\u0644\u0627\u062d\u0638\u0627\u062a\u0643 \u0623\u0648 \u0627\u0642\u062a\u0631\u0627\u062d\u0627\u062a\u0643 \u0623\u0648 \u0627\u0644\u0623\u062e\u0637\u0627\u0621 \u0627\u0644\u062a\u064a \u0648\u0627\u062c\u0647\u062a\u0647\u0627 \u0647\u0646\u0627.)\n\n\u2014 \u0645\u064f\u0631\u0633\u064e\u0644 \u0645\u0646 \u0625\u0636\u0627\u0641\u0629 Dynamics Audit Lens v$version$",
    dismiss: "\u0625\u063a\u0644\u0627\u0642",
    sendFeedback: "\u0625\u0631\u0633\u0627\u0644 \u0645\u0644\u0627\u062d\u0638\u0627\u062a",
    feedbackModalDesc: "\u0623\u0642\u0631\u0623 \u0643\u0644 \u0631\u0633\u0627\u0644\u0629. \u0627\u0643\u062a\u0628 \u0645\u0644\u0627\u062d\u0638\u0627\u062a\u0643 \u0623\u062f\u0646\u0627\u0647\u060c \u062b\u0645 \u0627\u062e\u062a\u0631 \u0643\u064a\u0641 \u062a\u0631\u0633\u0644\u0647\u0627 \u2014 \u0623\u0648 \u0627\u0646\u0633\u062e\u0647\u0627 \u0648\u0623\u0644\u0635\u0642\u0647\u0627 \u0623\u064a\u0646\u0645\u0627 \u062a\u0631\u064a\u062f.",
    feedbackKindLabel: "\u0627\u0644\u0646\u0648\u0639",
    feedbackKindFeature: "\u0627\u0642\u062a\u0631\u0627\u062d \u0645\u064a\u0632\u0629",
    feedbackKindBug: "\u0628\u0644\u0627\u063a \u062e\u0637\u0623",
    feedbackKindGeneral: "\u0645\u0644\u0627\u062d\u0638\u0627\u062a \u0639\u0627\u0645\u0629",
    feedbackYourEmail: "\u0628\u0631\u064a\u062f\u0643 \u0627\u0644\u0625\u0644\u0643\u062a\u0631\u0648\u0646\u064a",
    optional: "(\u0627\u062e\u062a\u064a\u0627\u0631\u064a)",
    feedbackMessage: "\u0627\u0644\u0631\u0633\u0627\u0644\u0629",
    feedbackMessagePlaceholder: "\u0645\u0627 \u0627\u0644\u0630\u064a \u064a\u062f\u0648\u0631 \u0641\u064a \u0630\u0647\u0646\u0643\u061f \u0627\u0642\u062a\u0631\u0627\u062d\u0627\u062a\u060c \u0623\u062e\u0637\u0627\u0621\u060c \u0623\u0641\u0643\u0627\u0631\u2026",
    feedbackMessageRequired: "\u0644\u0627 \u064a\u0645\u0643\u0646 \u0623\u0646 \u062a\u0643\u0648\u0646 \u0627\u0644\u0631\u0633\u0627\u0644\u0629 \u0641\u0627\u0631\u063a\u0629.",
    sendVia: "\u0625\u0631\u0633\u0627\u0644 \u0639\u0628\u0631",
    gmail: "Gmail",
    outlook: "Outlook",
    copy: "\u0646\u0633\u062e",
    orEmailMeAt: "\u0623\u0648 \u0631\u0627\u0633\u0644\u0646\u064a \u0645\u0628\u0627\u0634\u0631\u0629 \u0639\u0644\u0649",
    copyAddress: "\u0646\u0633\u062e \u0627\u0644\u0639\u0646\u0648\u0627\u0646",
    addressCopied: "\u062a\u0645 \u0646\u0633\u062e \u0627\u0644\u0639\u0646\u0648\u0627\u0646 \u0625\u0644\u0649 \u0627\u0644\u062d\u0627\u0641\u0638\u0629.",
    messageCopied: "\u062a\u0645 \u0646\u0633\u062e \u0627\u0644\u0631\u0633\u0627\u0644\u0629 \u0625\u0644\u0649 \u0627\u0644\u062d\u0627\u0641\u0638\u0629 \u2014 \u0623\u0644\u0635\u0642\u0647\u0627 \u0641\u064a \u0628\u0631\u064a\u062f\u0643.",
    copyFailed: "\u062a\u0639\u0630\u0651\u0631 \u0627\u0644\u0648\u0635\u0648\u0644 \u0625\u0644\u0649 \u0627\u0644\u062d\u0627\u0641\u0638\u0629. \u062d\u062f\u0651\u062f \u0627\u0644\u0631\u0633\u0627\u0644\u0629 \u0648\u0627\u0646\u0633\u062e\u0647\u0627 \u064a\u062f\u0648\u064a\u0627\u064b.",
    openingGmail: "\u062c\u0627\u0631\u064d \u0641\u062a\u062d Gmail\u2026",
    openingOutlook: "\u062c\u0627\u0631\u064d \u0641\u062a\u062d Outlook\u2026",
    waitingForDynamics: "\u0628\u0627\u0646\u062a\u0638\u0627\u0631 \u0635\u0641\u062d\u0629 Dynamics\u2026",
    recordsSelected: "\u062a\u0645 \u062a\u062d\u062f\u064a\u062f $count$ \u0633\u062c\u0644",
    exportToExcel: "\u062a\u0635\u062f\u064a\u0631 \u0625\u0644\u0649 Excel",
    exportEntireView: "\u062a\u0635\u062f\u064a\u0631 \u062c\u0645\u064a\u0639 \u0633\u062c\u0644\u0627\u062a \u0627\u0644\u0639\u0631\u0636",
    exportEntireViewWithCount: "\u062a\u0635\u062f\u064a\u0631 \u062c\u0645\u064a\u0639 \u0633\u062c\u0644\u0627\u062a \u0627\u0644\u0639\u0631\u0636 ($count$)",
    resolvingViewRecords: "\u062c\u0627\u0631\u064d \u062c\u0644\u0628 \u0633\u062c\u0644\u0627\u062a \u0627\u0644\u0639\u0631\u0636 \u0645\u0646 Dataverse\u2026",
    viewResolveFailed: "\u062a\u0639\u0630\u0651\u0631 \u062c\u0644\u0628 \u0633\u062c\u0644\u0627\u062a \u0647\u0630\u0627 \u0627\u0644\u0639\u0631\u0636.",
    preparing: "\u062c\u0627\u0631\u064d \u0627\u0644\u062a\u062d\u0636\u064a\u0631\u2026",
    entityPlaceholder: "\u0627\u0633\u0645 \u0627\u0644\u0643\u064a\u0627\u0646 (\u0645\u062b\u0627\u0644: account, contact\u2026)",
    userSearchPlaceholder: "\u0627\u0628\u062d\u062b \u0639\u0646 \u0645\u0633\u062a\u062e\u062f\u0645 \u0628\u0627\u0644\u0627\u0633\u0645 \u0623\u0648 \u0627\u0644\u0628\u0631\u064a\u062f \u0627\u0644\u0625\u0644\u0643\u062a\u0631\u0648\u0646\u064a\u2026",
    remove: "\u0625\u0632\u0627\u0644\u0629",
    from: "\u0645\u0646",
    to: "\u0625\u0644\u0649",
    exportUserAudit: "\u062a\u0635\u062f\u064a\u0631 \u062a\u062f\u0642\u064a\u0642 \u0646\u0634\u0627\u0637 \u0627\u0644\u0645\u0633\u062a\u062e\u062f\u0645",
    allDataLocal: "\u062a\u0628\u0642\u0649 \u062c\u0645\u064a\u0639 \u0627\u0644\u0628\u064a\u0627\u0646\u0627\u062a \u062f\u0627\u062e\u0644 \u0645\u062a\u0635\u0641\u062d\u0643.",
    close: "\u0625\u063a\u0644\u0627\u0642",
    modalDesc: "\u0623\u062f\u0627\u0629 \u0645\u062d\u0644\u064a\u0629 \u0628\u0627\u0644\u0643\u0627\u0645\u0644 \u0644\u062a\u0635\u062f\u064a\u0631 \u0633\u062c\u0644\u0627\u062a \u0627\u0644\u062a\u062f\u0642\u064a\u0642 \u0645\u0646 Microsoft Dynamics\u00a0365 \u0648 Dataverse. \u0644\u0627 \u064a\u062a\u0645 \u0625\u0631\u0633\u0627\u0644 \u0623\u064a \u0628\u064a\u0627\u0646\u0627\u062a \u0625\u0644\u0649 \u0623\u064a \u062e\u0627\u062f\u0645 \u062e\u0627\u0631\u062c\u064a \u2014 \u062c\u0645\u064a\u0639 \u0627\u0644\u0645\u0639\u0627\u0644\u062c\u0629 \u062a\u062a\u0645 \u062f\u0627\u062e\u0644 \u0645\u062a\u0635\u0641\u062d\u0643.",
    contact: "\u062a\u0648\u0627\u0635\u0644",
    license: "\u0627\u0644\u062a\u0631\u062e\u064a\u0635",
    mitLicense: "\u0631\u062e\u0635\u0629 MIT",
    freeOpenSource: "\u0645\u062c\u0627\u0646\u064a \u0648\u0645\u0641\u062a\u0648\u062d \u0627\u0644\u0645\u0635\u062f\u0631 \u2014 \u062f\u0648\u0646 \u0623\u064a \u0636\u0645\u0627\u0646",
    legal: "\u0625\u0634\u0639\u0627\u0631 \u0642\u0627\u0646\u0648\u0646\u064a",
    legalText: "\u064a\u064f\u0642\u062f\u064e\u0651\u0645 \u0647\u0630\u0627 \u0627\u0644\u0628\u0631\u0646\u0627\u0645\u062c \u201c\u0643\u0645\u0627 \u0647\u0648\u201d \u062f\u0648\u0646 \u0623\u064a \u0636\u0645\u0627\u0646 \u0645\u0646 \u0623\u064a \u0646\u0648\u0639\u060c \u0635\u0631\u064a\u062d\u0627\u064b \u0643\u0627\u0646 \u0623\u0645 \u0636\u0645\u0646\u064a\u0627\u064b\u060c \u0628\u0645\u0627 \u0641\u064a \u0630\u0644\u0643 \u0639\u0644\u0649 \u0633\u0628\u064a\u0644 \u0627\u0644\u0645\u062b\u0627\u0644 \u0644\u0627 \u0627\u0644\u062d\u0635\u0631 \u0636\u0645\u0627\u0646\u0627\u062a \u0627\u0644\u0642\u0627\u0628\u0644\u064a\u0629 \u0644\u0644\u062a\u0633\u0648\u064a\u0642 \u0648\u0627\u0644\u0645\u0644\u0627\u0621\u0645\u0629 \u0644\u063a\u0631\u0636 \u0645\u0639\u064a\u0651\u0646 \u0648\u0639\u062f\u0645 \u0627\u0646\u062a\u0647\u0627\u0643 \u062d\u0642\u0648\u0642 \u0627\u0644\u063a\u064a\u0631. \u0644\u0627 \u064a\u062a\u062d\u0645\u0651\u0644 \u0627\u0644\u0645\u0624\u0644\u0641\u0648\u0646 \u0623\u0648 \u0623\u0635\u062d\u0627\u0628 \u062d\u0642\u0648\u0642 \u0627\u0644\u0646\u0634\u0631 \u0628\u0623\u064a \u062d\u0627\u0644 \u0645\u0646 \u0627\u0644\u0623\u062d\u0648\u0627\u0644 \u0623\u064a \u0645\u0633\u0624\u0648\u0644\u064a\u0629 \u0639\u0646 \u0623\u064a \u0645\u0637\u0627\u0644\u0628\u0629 \u0623\u0648 \u0623\u0636\u0631\u0627\u0631 \u0623\u0648 \u0645\u0633\u0624\u0648\u0644\u064a\u0627\u062a \u0623\u062e\u0631\u0649\u060c \u0633\u0648\u0627\u0621 \u0623\u0643\u0627\u0646 \u0630\u0644\u0643 \u0641\u064a \u0625\u0637\u0627\u0631 \u0639\u0642\u062f \u0623\u0648 \u062a\u0642\u0635\u064a\u0631 \u0623\u0648 \u062e\u0644\u0627\u0641\u0647\u060c \u062a\u0646\u0634\u0623 \u0639\u0646 \u0627\u0644\u0628\u0631\u0646\u0627\u0645\u062c \u0623\u0648 \u0627\u0633\u062a\u062e\u062f\u0627\u0645\u0647 \u0623\u0648 \u0627\u0644\u062a\u0639\u0627\u0645\u0644 \u0645\u0639\u0647.",
    copyright: "\u00a9 2026 \u0645\u062d\u0645\u0648\u062f \u0632\u064a\u062f\u0627\u0646 \u2014 \u0628\u0631\u0646\u0627\u0645\u062c \u0645\u062c\u0627\u0646\u064a \u0648\u0645\u0641\u062a\u0648\u062d \u0627\u0644\u0645\u0635\u062f\u0631",
    cannotAccessTab: "\u062a\u0639\u0630\u0651\u0631 \u0627\u0644\u0648\u0635\u0648\u0644 \u0625\u0644\u0649 \u0639\u0644\u0627\u0645\u0629 \u0627\u0644\u062a\u0628\u0648\u064a\u0628 \u0627\u0644\u062d\u0627\u0644\u064a\u0629.",
    notDynamicsPage: "\u0647\u0630\u0647 \u0627\u0644\u0635\u0641\u062d\u0629 \u0644\u064a\u0633\u062a \u0635\u0641\u062d\u0629 Dynamics / Dataverse.",
    couldNotReadContext: "\u062a\u0639\u0630\u0651\u0631\u062a \u0642\u0631\u0627\u0621\u0629 \u0633\u064a\u0627\u0642 \u0627\u0644\u0635\u0641\u062d\u0629.",
    activeOn: "\u0646\u0634\u0637 \u0639\u0644\u0649: $hostname$",
    tooManyRecords: "\u0639\u062f\u062f \u0627\u0644\u0633\u062c\u0644\u0627\u062a \u0627\u0644\u0645\u062d\u062f\u062f\u0629 \u0643\u0628\u064a\u0631 \u062c\u062f\u0627\u064b (\u0627\u0644\u062d\u062f \u0627\u0644\u0623\u0642\u0635\u0649 $max$ \u0633\u062c\u0644). \u064a\u0631\u062c\u0649 \u062a\u0642\u0644\u064a\u0635 \u0646\u0637\u0627\u0642 \u0627\u0644\u062a\u062d\u062f\u064a\u062f.",
    contentScriptNotReady: "\u0627\u0644\u0625\u0636\u0627\u0641\u0629 \u0644\u0645 \u062a\u0643\u062a\u0645\u0644 \u062a\u0647\u064a\u0626\u062a\u0647\u0627 \u0628\u0639\u062f. \u064a\u0631\u062c\u0649 \u0625\u0639\u0627\u062f\u0629 \u062a\u062d\u0645\u064a\u0644 \u0627\u0644\u0635\u0641\u062d\u0629.",
    noEntitiesFound: "\u0644\u0627 \u062a\u0648\u062c\u062f \u0643\u064a\u0627\u0646\u0627\u062a \u0645\u0637\u0627\u0628\u0642\u0629.",
    unnamed: "(\u0628\u062f\u0648\u0646 \u0627\u0633\u0645)",
    noAuditRecords: "\u0644\u0627 \u062a\u0648\u062c\u062f \u0633\u062c\u0644\u0627\u062a \u062a\u062f\u0642\u064a\u0642 \u0644\u0647\u0630\u0647 \u0627\u0644\u0639\u0646\u0627\u0635\u0631.",
    noAuditRecordsForUser: "\u0644\u0627 \u062a\u0648\u062c\u062f \u0633\u062c\u0644\u0627\u062a \u062a\u062f\u0642\u064a\u0642 \u0644\u0647\u0630\u0627 \u0627\u0644\u0645\u0633\u062a\u062e\u062f\u0645.",
    cappedAtRows: "\u062a\u0645 \u0627\u0642\u062a\u0637\u0627\u0639 \u0627\u0644\u0646\u062a\u0627\u0626\u062c \u0639\u0646\u062f $count$ \u0635\u0641. \u062c\u0627\u0631\u064d \u0625\u0646\u0634\u0627\u0621 \u0627\u0644\u0645\u0644\u0641\u2026",
    exportComplete: "\u0627\u0643\u062a\u0645\u0644 \u0627\u0644\u062a\u0635\u062f\u064a\u0631 \u2014 $count$ \u0635\u0641.",
    exportFailed: "\u0641\u0634\u0644 \u0627\u0644\u062a\u0635\u062f\u064a\u0631.",
    connectionLost: "\u0627\u0646\u0642\u0637\u0639 \u0627\u0644\u0627\u062a\u0635\u0627\u0644. \u064a\u0631\u062c\u0649 \u0625\u0639\u0627\u062f\u0629 \u062a\u062d\u0645\u064a\u0644 \u0627\u0644\u0635\u0641\u062d\u0629 \u0648\u0627\u0644\u0645\u062d\u0627\u0648\u0644\u0629 \u0645\u062c\u062f\u062f\u0627\u064b.",
    queryingAudit: "\u062c\u0627\u0631\u064d \u0627\u0644\u0627\u0633\u062a\u0639\u0644\u0627\u0645 \u0639\u0646 \u0633\u062c\u0644\u0627\u062a \u0627\u0644\u062a\u062f\u0642\u064a\u0642\u2026",
    fillingFields: "\u062c\u0627\u0631\u064d \u062a\u0639\u0628\u0626\u0629 \u062d\u0642\u0648\u0644 \u0627\u0644\u0646\u0645\u0648\u0630\u062c\u2026",
    filledFields: "\u062a\u0645\u062a \u062a\u0639\u0628\u0626\u0629 $filled$ \u0645\u0646 \u0623\u0635\u0644 $total$ \u062d\u0642\u0644 (\u062a\u0645 \u062a\u062e\u0637\u064a $skipped$).",
    zeroFieldsFilled: "\u0644\u0645 \u062a\u062a\u0645 \u062a\u0639\u0628\u0626\u0629 \u0623\u064a \u062d\u0642\u0644 (\u062a\u0645 \u062a\u062e\u0637\u064a $skipped$). \u0642\u062f \u062a\u0643\u0648\u0646 \u062c\u0645\u064a\u0639 \u0627\u0644\u062d\u0642\u0648\u0644 \u0645\u0645\u062a\u0644\u0626\u0629 \u0628\u0627\u0644\u0641\u0639\u0644 \u0623\u0648 \u0644\u0644\u0642\u0631\u0627\u0621\u0629 \u0641\u0642\u0637.",
    failedFill: "\u062a\u0639\u0630\u0651\u0631\u062a \u062a\u0639\u0628\u0626\u0629 \u0628\u064a\u0627\u0646\u0627\u062a \u0627\u0644\u0646\u0645\u0648\u0630\u062c.",
    couldNotReachContent: "\u062a\u0639\u0630\u0651\u0631 \u0627\u0644\u0648\u0635\u0648\u0644 \u0625\u0644\u0649 \u0627\u0644\u0625\u0636\u0627\u0641\u0629. \u064a\u0631\u062c\u0649 \u0625\u0639\u0627\u062f\u0629 \u062a\u062d\u0645\u064a\u0644 \u0627\u0644\u0635\u0641\u062d\u0629.",
    processedOf: "\u062a\u0645\u062a \u0645\u0639\u0627\u0644\u062c\u0629 $done$ \u0645\u0646 \u0623\u0635\u0644 $total$ \u0633\u062c\u0644\u2026",
    lookupIssues: "\u062a\u0646\u0628\u064a\u0647\u0627\u062a \u0641\u064a \u062d\u0642\u0648\u0644 \u0627\u0644\u0628\u062d\u062b (Lookup): $issues$",
    lookupErrors: "\u0623\u062e\u0637\u0627\u0621 \u0641\u064a \u062d\u0642\u0648\u0644 \u0627\u0644\u0628\u062d\u062b (Lookup): $errors$",
    error: "\u062e\u0637\u0623: $msg$",
    discoveringRecords: "\u062c\u0627\u0631\u064d \u0627\u0644\u0628\u062d\u062b \u0639\u0646 \u0627\u0644\u0633\u062c\u0644\u0627\u062a \u0627\u0644\u062a\u064a \u0639\u062f\u0651\u0644\u0647\u0627 \u0647\u0630\u0627 \u0627\u0644\u0645\u0633\u062a\u062e\u062f\u0645\u2026",
    foundRecordsFetching: "\u062a\u0645 \u0627\u0644\u0639\u062b\u0648\u0631 \u0639\u0644\u0649 $count$ \u0633\u062c\u0644. \u062c\u0627\u0631\u064d \u062c\u0644\u0628 \u0633\u062c\u0644\u0627\u062a \u0627\u0644\u062a\u062f\u0642\u064a\u0642\u2026",
    sessionExpired: "\u0627\u0646\u062a\u0647\u062a \u0635\u0644\u0627\u062d\u064a\u0629 \u0627\u0644\u062c\u0644\u0633\u0629 \u2014 \u064a\u0631\u062c\u0649 \u0625\u0639\u0627\u062f\u0629 \u062a\u062d\u0645\u064a\u0644 \u0627\u0644\u0635\u0641\u062d\u0629 \u0648\u062a\u0633\u062c\u064a\u0644 \u0627\u0644\u062f\u062e\u0648\u0644 \u0645\u062c\u062f\u062f\u0627\u064b.",
    accessDeniedAudit: "\u062a\u0645 \u0631\u0641\u0636 \u0627\u0644\u0648\u0635\u0648\u0644 \u2014 \u062a\u062d\u062a\u0627\u062c \u0625\u0644\u0649 \u0635\u0644\u0627\u062d\u064a\u0629 \"\u0639\u0631\u0636 \u0645\u0644\u062e\u0635 \u0627\u0644\u062a\u062f\u0642\u064a\u0642\" (prvReadAuditSummary).",
    recordNotFound: "\u0627\u0644\u0633\u062c\u0644 \u063a\u064a\u0631 \u0645\u0648\u062c\u0648\u062f \u2014 \u0631\u0628\u0645\u0627 \u062a\u0645 \u062d\u0630\u0641\u0647.",
    invalidPayload: "\u0627\u0644\u0628\u064a\u0627\u0646\u0627\u062a \u0627\u0644\u0645\u064f\u0631\u0633\u064e\u0644\u0629 \u063a\u064a\u0631 \u0635\u0627\u0644\u062d\u0629.",
    opCreate: "\u0625\u0646\u0634\u0627\u0621", opUpdate: "\u062a\u062d\u062f\u064a\u062b", opDelete: "\u062d\u0630\u0641",
    opAccess: "\u0648\u0635\u0648\u0644", opUpsert: "\u0625\u062f\u0631\u0627\u062c/\u062a\u062d\u062f\u064a\u062b",
    language: "\u0627\u0644\u0644\u063a\u0629",
    langEnglish: "English",
    langArabic: "\u0627\u0644\u0639\u0631\u0628\u064a\u0629",
    noUsersFound: "\u0644\u0645 \u064a\u062a\u0645 \u0627\u0644\u0639\u062b\u0648\u0631 \u0639\u0644\u0649 \u0645\u0633\u062a\u062e\u062f\u0645\u064a\u0646.",
  },
};

let currentLang = "en";

function t(key, subs) {
  const dict = I18N[currentLang] || I18N.en;
  let msg = dict[key] || I18N.en[key] || key;
  if (subs !== undefined) {
    const arr = Array.isArray(subs) ? subs : [subs];
    arr.forEach((val, i) => {
      msg = msg.replace(/\$\w+\$/, val);
    });
  }
  return msg;
}

function updateLangIcon(lang) {
  const slot = document.getElementById("lang-icon-slot");
  if (!slot) return;
  if (lang === "en") {
    slot.innerHTML = `<svg class="lang-flag-icon" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 32" width="20" height="14"><rect width="48" height="32" fill="#007A3D"/><circle cx="24" cy="16" r="9.5" fill="none" stroke="#fff" stroke-width="1.2"/><path d="M24 6.5A9.5 9.5 0 1 0 24 25.5 9.5 9.5 0 0 0 24 6.5Zm0 2.2a7.3 7.3 0 1 1 0 14.6 7.3 7.3 0 0 1 0-14.6Z" fill="#fff"/><path d="M21.2 11.2a5.3 5.3 0 1 0 0 9.6 6.8 6.8 0 0 1-2.4-5 6.8 6.8 0 0 1 2.4-4.6z" fill="#fff"/><path d="M23.3 13.2l.5 1.5h1.6l-1.3.95.5 1.55-1.3-.95-1.3.95.5-1.55-1.3-.95h1.6z" fill="#007A3D"/><path d="M26.3 13.2l.5 1.5h1.6l-1.3.95.5 1.55-1.3-.95-1.3.95.5-1.55-1.3-.95h1.6z" fill="#007A3D"/><path d="M24.8 17l.5 1.5h1.6l-1.3.95.5 1.55-1.3-.95-1.3.95.5-1.55-1.3-.95h1.6z" fill="#007A3D"/></svg>`;
  } else {
    slot.innerHTML = `<svg class="lang-flag-icon" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 36 24" width="20" height="14"><rect width="36" height="24" fill="#B22234"/><rect y="1.85" width="36" height="1.85" fill="#fff"/><rect y="5.54" width="36" height="1.85" fill="#fff"/><rect y="9.23" width="36" height="1.85" fill="#fff"/><rect y="12.92" width="36" height="1.85" fill="#fff"/><rect y="16.62" width="36" height="1.85" fill="#fff"/><rect y="20.31" width="36" height="1.85" fill="#fff"/><rect width="14.4" height="12.92" fill="#3C3B6E"/><circle cx="2.4" cy="1.8" r="0.55" fill="#fff"/><circle cx="4.8" cy="1.8" r="0.55" fill="#fff"/><circle cx="7.2" cy="1.8" r="0.55" fill="#fff"/><circle cx="9.6" cy="1.8" r="0.55" fill="#fff"/><circle cx="12" cy="1.8" r="0.55" fill="#fff"/><circle cx="3.6" cy="3.6" r="0.55" fill="#fff"/><circle cx="6" cy="3.6" r="0.55" fill="#fff"/><circle cx="8.4" cy="3.6" r="0.55" fill="#fff"/><circle cx="10.8" cy="3.6" r="0.55" fill="#fff"/><circle cx="2.4" cy="5.4" r="0.55" fill="#fff"/><circle cx="4.8" cy="5.4" r="0.55" fill="#fff"/><circle cx="7.2" cy="5.4" r="0.55" fill="#fff"/><circle cx="9.6" cy="5.4" r="0.55" fill="#fff"/><circle cx="12" cy="5.4" r="0.55" fill="#fff"/><circle cx="3.6" cy="7.2" r="0.55" fill="#fff"/><circle cx="6" cy="7.2" r="0.55" fill="#fff"/><circle cx="8.4" cy="7.2" r="0.55" fill="#fff"/><circle cx="10.8" cy="7.2" r="0.55" fill="#fff"/><circle cx="2.4" cy="9" r="0.55" fill="#fff"/><circle cx="4.8" cy="9" r="0.55" fill="#fff"/><circle cx="7.2" cy="9" r="0.55" fill="#fff"/><circle cx="9.6" cy="9" r="0.55" fill="#fff"/><circle cx="12" cy="9" r="0.55" fill="#fff"/><circle cx="3.6" cy="10.8" r="0.55" fill="#fff"/><circle cx="6" cy="10.8" r="0.55" fill="#fff"/><circle cx="8.4" cy="10.8" r="0.55" fill="#fff"/><circle cx="10.8" cy="10.8" r="0.55" fill="#fff"/></svg>`;
  }
}

function applyLang(lang) {
  currentLang = lang;
  const html = document.documentElement;
  html.setAttribute("dir", lang === "ar" ? "rtl" : "ltr");
  html.setAttribute("lang", lang === "ar" ? "ar" : "en");

  document.querySelectorAll("[data-i18n]").forEach((el) => {
    const key = el.getAttribute("data-i18n");
    const msg = t(key);
    if (msg) el.textContent = msg;
  });

  document.querySelectorAll("[data-i18n-placeholder]").forEach((el) => {
    const key = el.getAttribute("data-i18n-placeholder");
    const msg = t(key);
    if (msg) el.placeholder = msg;
  });

  document.querySelectorAll("[data-i18n-title]").forEach((el) => {
    const key = el.getAttribute("data-i18n-title");
    const msg = t(key);
    if (msg) el.title = msg;
  });

  const langLabel = document.getElementById("lang-label");
  if (langLabel) {
    langLabel.textContent = lang === "ar" ? "English" : "\u0627\u0644\u0639\u0631\u0628\u064a\u0629";
  }

  updateLangIcon(lang);
}

async function loadLang() {
  let lang = null;
  try {
    const result = await chrome.storage.local.get(LANG_STORAGE_KEY);
    lang = result?.[LANG_STORAGE_KEY] ?? null;
  } catch { /* storage unavailable */ }

  if (!lang) {
    const uiLang = chrome.i18n.getUILanguage();
    lang = uiLang && uiLang.startsWith("ar") ? "ar" : "en";
  }

  applyLang(lang);
}

async function toggleLang() {
  const next = currentLang === "ar" ? "en" : "ar";
  applyLang(next);
  try {
    await chrome.storage.local.set({ [LANG_STORAGE_KEY]: next });
  } catch { /* storage unavailable */ }
}

const statusEl = document.getElementById("status-msg");
const recordInfoEl = document.getElementById("record-info");
const recordCountEl = document.getElementById("record-count");
const entityNameEl = document.getElementById("entity-name");
const exportBtn = document.getElementById("export-btn");
const exportViewBtn = document.getElementById("export-view-btn");
const exportViewLabel = document.getElementById("export-view-label");
const progressSection = document.getElementById("progress-section");
const progressFill = document.getElementById("progress-fill");
const progressText = document.getElementById("progress-text");

const userStatusEl = document.getElementById("user-status-msg");
const entitySearchInput = document.getElementById("entity-search-input");
const entitySearchDropdown = document.getElementById("entity-search-dropdown");
const userSearchInput = document.getElementById("user-search-input");
const userSearchDropdown = document.getElementById("user-search-dropdown");
const selectedUserEl = document.getElementById("selected-user");
const selectedUserNameEl = document.getElementById("selected-user-name");
const clearUserBtn = document.getElementById("clear-user-btn");
const dateFromInput = document.getElementById("date-from");
const dateToInput = document.getElementById("date-to");
const userExportBtn = document.getElementById("user-export-btn");
const userProgressSection = document.getElementById("user-progress-section");
const userProgressFill = document.getElementById("user-progress-fill");
const userProgressText = document.getElementById("user-progress-text");

let currentContext = null;
let currentTabId = null;
let exporting = false;
let userExporting = false;
let selectedUser = null;

(function setVersionBadge() {
  const v = `v${chrome.runtime.getManifest().version}`;
  const header = document.getElementById("app-version");
  if (header) header.textContent = v;
  const modal = document.getElementById("modal-version");
  if (modal) modal.textContent = v;
})();

const tabBtns = document.querySelectorAll(".tab");
const tabPanels = document.querySelectorAll(".tab-panel");

tabBtns.forEach((btn) => {
  btn.addEventListener("click", () => {
    const target = btn.dataset.tab;
    tabBtns.forEach((b) => b.classList.toggle("tab--active", b.dataset.tab === target));
    tabPanels.forEach((p) => {
      const isActive = p.id === `tab-${target}`;
      p.classList.toggle("tab-panel--active", isActive);
      p.hidden = !isActive;
    });
  });
});

function makeStatusSetter(el) {
  return (text, type = "idle") => {
    el.textContent = text;
    el.className = `status status--${type}`;
  };
}

const setStatus = makeStatusSetter(statusEl);
const setUserStatus = makeStatusSetter(userStatusEl);

function makeProgressUpdater(fillEl, textEl) {
  return (processed, total) => {
    const pct = total > 0 ? Math.round((processed / total) * 100) : 0;
    fillEl.style.width = `${pct}%`;
    textEl.textContent = t("processedOf", [String(processed), String(total)]);
    textEl.className = "progress-text";
  };
}

const updateProgress = makeProgressUpdater(progressFill, progressText);
const updateUserProgress = makeProgressUpdater(userProgressFill, userProgressText);

function formatDateStamp() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, "0")}${String(d.getDate()).padStart(2, "0")}`;
}

function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function generateExcel(rows, entityName, filenameSuffix, opts) {
  const sheetName = opts?.sheetName ?? "Audit History";
  const filenamePrefix = opts?.filenamePrefix ?? "AuditExport";
  const dateColumnNames = opts?.dateColumns ?? ["ChangedDate", "DeletedDate"];

  let ws = XLSX.utils.json_to_sheet(rows);

  const range = XLSX.utils.decode_range(ws["!ref"]);
  const dateCols = [];
  for (let C = range.s.c; C <= range.e.c; ++C) {
    const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: C });
    const cell = ws[cellAddress];
    if (cell && dateColumnNames.includes(cell.v)) {
      dateCols.push(C);
    }
  }
  for (const dateCol of dateCols) {
    for (let R = range.s.r + 1; R <= range.e.r; ++R) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: dateCol });
      const cell = ws[cellAddress];
      if (cell && cell.v != null && cell.v !== "") {
        if (cell.v instanceof Date || typeof cell.v === "number") {
          cell.t = "d";
          cell.z = "yyyy-mm-dd hh:mm:ss";
        }
      }
    }
  }

  let wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);

  if (rows.length > 0) {
    const keys = Object.keys(rows[0]);
    const sample = rows.slice(0, 200);
    ws["!cols"] = keys.map((key) => {
      const maxLen = Math.max(
        key.length,
        ...sample.map((r) => String(r[key] ?? "").length),
      );
      return { wch: Math.min(maxLen + 2, 50) };
    });

    ws["!autofilter"] = { ref: ws["!ref"] };
  }

  const safeName = String(entityName ?? "Unknown").replace(
    /[^a-zA-Z0-9_-]/g,
    "_",
  );
  const suffix = filenameSuffix ? `_${filenameSuffix}` : "";
  const filename = `${filenamePrefix}_${safeName}${suffix}_${formatDateStamp()}.xlsx`;

  const wbOut = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  wb = null;
  ws = null;

  const blob = new Blob([wbOut], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function sendToTab(tabId, message) {
  return chrome.tabs.sendMessage(tabId, message);
}

async function saveUserAuditState() {
  const state = {
    entityName: entitySearchInput.value.trim() || null,
    selectedUser,
    dateFrom: dateFromInput.value || null,
    dateTo: dateToInput.value || null,
  };
  try {
    await chrome.storage.local.set({ [STATE_STORAGE_KEY]: state });
  } catch { /* storage unavailable */ }
}

async function loadUserAuditState() {
  try {
    const result = await chrome.storage.local.get(STATE_STORAGE_KEY);
    const state = result?.[STATE_STORAGE_KEY];
    if (!state) return;

    if (state.entityName) {
      entitySearchInput.value = state.entityName;
    }
    if (state.selectedUser) {
      selectedUser = state.selectedUser;
      selectedUserNameEl.textContent =
        selectedUser.fullname || selectedUser.email || (selectedUser.id ?? "").slice(0, 8);
      selectedUserEl.hidden = false;
    }
    if (state.dateFrom) {
      dateFromInput.value = state.dateFrom;
    }
    if (state.dateTo) {
      dateToInput.value = state.dateTo;
    }
  } catch { /* storage unavailable */ }
}

const DYNAMICS_PATTERN =
  /^https?:\/\/[^/]+\.(crm\d*\.dynamics\.com|crm\.microsoftdynamics\.us|crm\.appsplatform\.us|crm\.dynamics\.cn)\//;

async function fetchContext() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  if (!tab?.id) {
    setStatus(t("cannotAccessTab"), "error");
    setUserStatus(t("cannotAccessTab"), "error");
    setDelStatus(t("cannotAccessTab"), "error");
    return;
  }

  currentTabId = tab.id;

  const isDynamics = DYNAMICS_PATTERN.test(tab.url ?? "");

  if (!isDynamics) {
    setStatus(t("notDynamicsPage"), "idle");
    setUserStatus(t("notDynamicsPage"), "idle");
    setDelStatus(t("notDynamicsPage"), "idle");
    return;
  }

  try {
    const response = await sendToTab(tab.id, { type: "GET_CONTEXT" });

    if (!response?.ok || !response.context) {
      setStatus(t("couldNotReadContext"), "error");
      setUserStatus(t("couldNotReadContext"), "error");
      setDelStatus(t("couldNotReadContext"), "error");
      return;
    }

    currentContext = response.context;
    const count = currentContext.totalRecordCount ?? currentContext.selectedIds?.length ?? 0;
    const entity = currentContext.entityName;
    const hostname = new URL(tab.url).hostname;

    setStatus(t("activeOn", [hostname]), "active");
    recordCountEl.textContent = t("recordsSelected", [String(count)]);
    if (entity) entityNameEl.textContent = entity;
    recordInfoEl.hidden = false;

    if (currentContext.selectionUnavailable) {
      recordCountEl.textContent = t("recordsSelected", ["0"]);
    }

    if (count > MAX_EXPORT_RECORDS) {
      setStatus(t("tooManyRecords", [String(MAX_EXPORT_RECORDS)]), "error");
      exportBtn.disabled = true;
    } else {
      exportBtn.disabled = count === 0;
    }

    // Show "Export entire view" only when we have a view to query against.
    const hasView = !!(currentContext.entityName && currentContext.viewId);
    exportViewBtn.hidden = !hasView;
    if (hasView) {
      const total = currentContext.totalRecordCount;
      exportViewLabel.textContent =
        total && total > 0
          ? t("exportEntireViewWithCount", [String(total)])
          : t("exportEntireView");
      exportViewBtn.disabled = false;
    }

    setUserStatus(t("activeOn", [hostname]), "active");
    setDelStatus(t("activeOn", [hostname]), "active");

    if (entity && !entitySearchInput.value) {
      entitySearchInput.value = entity;
    }
    if (entity && !delEntityInput.value) {
      delEntityInput.value = entity;
      updateDelExportBtnState();
    }
  } catch (err) {
    console.warn("[Audit Lens] fetchContext failed:", err);
    setStatus(t("contentScriptNotReady"), "error");
    setUserStatus(t("contentScriptNotReady"), "error");
    setDelStatus(t("contentScriptNotReady"), "error");
  }

  await loadUserAuditState();

  if (!dateToInput.value) {
    dateToInput.value = todayISO();
  }

  updateUserExportBtnState();
}

function wireSearchInput({ input, dropdown, timeoutRef, searchFn, renderFn, onChange }) {
  let timeout;

  input.addEventListener("input", () => {
    clearTimeout(timeout);
    const query = input.value.trim();

    if (query.length < MIN_SEARCH_LENGTH) {
      dropdown.hidden = true;
      if (onChange) onChange();
      return;
    }

    if (onChange) onChange();
    timeout = setTimeout(() => searchFn(query), SEARCH_DEBOUNCE_MS);
  });

  input.addEventListener("blur", () => {
    setTimeout(() => { dropdown.hidden = true; }, 200);
  });

  input.addEventListener("focus", () => {
    if (dropdown.children.length > 0 && input.value.trim().length >= MIN_SEARCH_LENGTH) {
      dropdown.hidden = false;
    }
  });
}

function renderDropdown(dropdown, items, emptyText, onSelect) {
  dropdown.replaceChildren();

  if (items.length === 0) {
    const empty = document.createElement("div");
    empty.className = "search-dropdown__empty";
    empty.textContent = emptyText;
    dropdown.appendChild(empty);
    dropdown.hidden = false;
    return;
  }

  for (const item of items) {
    const el = document.createElement("div");
    el.className = "search-dropdown__item";

    const nameSpan = document.createElement("span");
    nameSpan.className = "search-dropdown__item-name";
    nameSpan.textContent = item.name;

    const subSpan = document.createElement("span");
    subSpan.className = "search-dropdown__item-email";
    subSpan.textContent = item.sub || "";

    el.appendChild(nameSpan);
    el.appendChild(subSpan);

    el.addEventListener("mousedown", (e) => {
      e.preventDefault();
      onSelect(item.raw);
    });

    dropdown.appendChild(el);
  }

  dropdown.hidden = false;
}

async function performEntitySearch(query) {
  if (!currentTabId) return;

  try {
    const response = await sendToTab(currentTabId, {
      type: "SEARCH_ENTITIES",
      query,
    });

    if (!response?.ok || !response.entities?.length) {
      renderEntityDropdown([]);
      return;
    }
    renderEntityDropdown(response.entities);
  } catch {
    entitySearchDropdown.hidden = true;
  }
}

function renderEntityDropdown(entities) {
  renderDropdown(
    entitySearchDropdown,
    entities.map((ent) => ({
      name: ent.displayName || ent.logicalName,
      sub: ent.displayName !== ent.logicalName ? `(${ent.logicalName})` : "",
      raw: ent,
    })),
    t("noEntitiesFound"),
    (ent) => {
      entitySearchInput.value = ent.logicalName;
      entitySearchDropdown.hidden = true;
      updateUserExportBtnState();
      saveUserAuditState();
    },
  );
}

wireSearchInput({
  input: entitySearchInput,
  dropdown: entitySearchDropdown,
  searchFn: performEntitySearch,
  renderFn: renderEntityDropdown,
  onChange: () => { updateUserExportBtnState(); saveUserAuditState(); },
});

async function performUserSearch(query) {
  if (!currentTabId) return;

  try {
    const response = await sendToTab(currentTabId, {
      type: "SEARCH_USERS",
      query,
    });

    if (!response?.ok || !response.users?.length) {
      renderUserDropdown([]);
      return;
    }
    renderUserDropdown(response.users);
  } catch {
    userSearchDropdown.hidden = true;
  }
}

function renderUserDropdown(users) {
  renderDropdown(
    userSearchDropdown,
    users.map((user) => ({
      name: user.fullname || t("unnamed"),
      sub: user.email ? `(${user.email})` : "",
      raw: user,
    })),
    t("noUsersFound"),
    (user) => selectUser(user),
  );
}

wireSearchInput({
  input: userSearchInput,
  dropdown: userSearchDropdown,
  searchFn: performUserSearch,
  renderFn: renderUserDropdown,
});

function selectUser(user) {
  selectedUser = user;
  userSearchInput.value = "";
  userSearchDropdown.hidden = true;

  selectedUserNameEl.textContent =
    user.fullname || user.email || (user.id ?? "").slice(0, 8);
  selectedUserEl.hidden = false;

  updateUserExportBtnState();
  saveUserAuditState();
}

clearUserBtn.addEventListener("click", () => {
  selectedUser = null;
  selectedUserEl.hidden = true;
  updateUserExportBtnState();
  saveUserAuditState();
});

dateFromInput.addEventListener("change", () => saveUserAuditState());
dateToInput.addEventListener("change", () => saveUserAuditState());

function updateUserExportBtnState() {
  const entityOk = entitySearchInput.value.trim().length >= MIN_SEARCH_LENGTH;
  const userOk = !!selectedUser;
  userExportBtn.disabled = !entityOk || !userOk || userExporting;
}

async function startExport(guidsOverride) {
  if (exporting || !currentContext || !currentTabId) return;

  const { entityName } = currentContext;
  const guids = guidsOverride ?? currentContext.selectedIds;
  if (!entityName || !guids?.length) return;

  if (guids.length > MAX_EXPORT_RECORDS) {
    setStatus(t("tooManyRecords", [String(MAX_EXPORT_RECORDS)]), "error");
    return;
  }

  exporting = true;
  exportBtn.disabled = true;
  exportViewBtn.disabled = true;
  progressSection.hidden = false;
  updateProgress(0, guids.length);

  const port = chrome.tabs.connect(currentTabId, { name: "audit-export" });

  port.onMessage.addListener((msg) => {
    if (msg.type === "progress") {
      updateProgress(msg.done, msg.total);
    }

    if (msg.type === "done") {
      const rows = (msg.rows ?? []).slice(0, MAX_EXPORT_ROWS);

      if (rows.length === 0) {
        progressText.textContent = t("noAuditRecords");
        progressText.className = "progress-text progress-text--empty";
      } else {
        if (msg.rows.length > MAX_EXPORT_ROWS) {
          progressText.textContent = t("cappedAtRows", [MAX_EXPORT_ROWS.toLocaleString()]);
        }
        generateExcel(rows, entityName);
        progressText.textContent = t("exportComplete", [String(rows.length)]);
        progressText.className = "progress-text progress-text--success";
      }

      exporting = false;
      exportBtn.disabled = false;
      exportViewBtn.disabled = false;
    }

    if (msg.type === "error") {
      progressText.textContent = t("error", [msg.error]);
      progressText.className = "progress-text progress-text--error";
      setStatus(t("exportFailed"), "error");
      exporting = false;
      exportBtn.disabled = false;
      exportViewBtn.disabled = false;
    }
  });

  port.onDisconnect.addListener(() => {
    if (exporting) {
      progressText.textContent = t("connectionLost");
      progressText.className = "progress-text progress-text--error";
      setStatus(t("exportFailed"), "error");
      exporting = false;
      exportBtn.disabled = false;
      exportViewBtn.disabled = false;
    }
  });

  port.postMessage({
    entityLogicalName: entityName,
    guids,
  });
}

async function startExportEntireView() {
  if (exporting || !currentContext || !currentTabId) return;
  const { entityName, viewId, viewType } = currentContext;
  if (!entityName || !viewId) return;

  exportViewBtn.disabled = true;
  exportBtn.disabled = true;
  progressSection.hidden = false;
  progressText.textContent = t("resolvingViewRecords");
  progressText.className = "progress-text";
  updateProgress(0, 0);

  let resp;
  try {
    resp = await sendToTab(currentTabId, {
      type: "RESOLVE_VIEW_RECORDS",
      entityLogicalName: entityName,
      viewId,
      viewType: viewType || "savedquery",
    });
  } catch {
    resp = null;
  }

  if (!resp?.ok || !Array.isArray(resp.guids) || resp.guids.length === 0) {
    progressText.textContent = resp?.error
      ? t("error", [resp.error])
      : t("viewResolveFailed");
    progressText.className = "progress-text progress-text--error";
    exportViewBtn.disabled = false;
    exportBtn.disabled = (currentContext.selectedIds?.length ?? 0) === 0;
    return;
  }

  // Hand resolved GUIDs to the standard export pipeline.
  await startExport(resp.guids);
}

async function startUserExport() {
  if (userExporting || !currentTabId || !selectedUser) return;

  const entityLogicalName = entitySearchInput.value.trim().toLowerCase();
  if (!entityLogicalName) return;

  const dateFrom = dateFromInput.value || null;
  const dateTo = dateToInput.value || null;

  userExporting = true;
  userExportBtn.disabled = true;
  userProgressSection.hidden = false;
  userProgressFill.style.width = "0%";
  userProgressText.textContent = t("queryingAudit");
  userProgressText.className = "progress-text";

  const port = chrome.tabs.connect(currentTabId, { name: "user-audit-export" });

  port.onMessage.addListener((msg) => {
    if (msg.type === "phase") {
      userProgressText.textContent = msg.text;
      userProgressText.className = "progress-text";
    }

    if (msg.type === "progress") {
      updateUserProgress(msg.done, msg.total);
    }

    if (msg.type === "done") {
      const rows = (msg.rows ?? []).slice(0, MAX_EXPORT_ROWS);

      if (rows.length === 0) {
        userProgressText.textContent = t("noAuditRecordsForUser");
        userProgressText.className = "progress-text progress-text--empty";
      } else {
        const suffix = selectedUser.fullname
          ? selectedUser.fullname.replace(/[^a-zA-Z0-9_-]/g, "_")
          : (selectedUser.id ?? "unknown").slice(0, 8);
        generateExcel(rows, entityLogicalName, suffix);
        userProgressText.textContent = t("exportComplete", [String(rows.length)]);
        userProgressText.className = "progress-text progress-text--success";
      }

      userExporting = false;
      updateUserExportBtnState();
    }

    if (msg.type === "error") {
      userProgressText.textContent = t("error", [msg.error]);
      userProgressText.className = "progress-text progress-text--error";
      setUserStatus(t("exportFailed"), "error");
      userExporting = false;
      updateUserExportBtnState();
    }
  });

  port.onDisconnect.addListener(() => {
    if (userExporting) {
      userProgressText.textContent = t("connectionLost");
      userProgressText.className = "progress-text progress-text--error";
      setUserStatus(t("exportFailed"), "error");
      userExporting = false;
      updateUserExportBtnState();
    }
  });

  port.postMessage({
    entityLogicalName,
    userGuid: selectedUser.id,
    dateFrom,
    dateTo,
  });
}

exportBtn.addEventListener("click", () => startExport());
exportViewBtn.addEventListener("click", startExportEntireView);
userExportBtn.addEventListener("click", startUserExport);

// ──────────────────────────────────────────────────────────────────────────
// Deleted Records tab — list records that were deleted in a date range
// ──────────────────────────────────────────────────────────────────────────

const delStatusEl = document.getElementById("del-status-msg");
const delEntityInput = document.getElementById("del-entity-search-input");
const delEntityDropdown = document.getElementById("del-entity-search-dropdown");
const delUserInput = document.getElementById("del-user-search-input");
const delUserDropdown = document.getElementById("del-user-search-dropdown");
const delSelectedUserEl = document.getElementById("del-selected-user");
const delSelectedUserNameEl = document.getElementById("del-selected-user-name");
const delClearUserBtn = document.getElementById("del-clear-user-btn");
const delDateFromInput = document.getElementById("del-date-from");
const delDateToInput = document.getElementById("del-date-to");
const delExportBtn = document.getElementById("del-export-btn");
const delProgressSection = document.getElementById("del-progress-section");
const delProgressFill = document.getElementById("del-progress-fill");
const delProgressText = document.getElementById("del-progress-text");

const setDelStatus = makeStatusSetter(delStatusEl);

let delSelectedUser = null;
let delExporting = false;

function updateDelExportBtnState() {
  const entityOk = delEntityInput.value.trim().length >= MIN_SEARCH_LENGTH;
  const dateOk = !!(delDateFromInput.value && delDateToInput.value);
  delExportBtn.disabled = !entityOk || !dateOk || delExporting;
}

function updateDelProgress(done, total) {
  const pct = total > 0 ? Math.min(100, (done / total) * 100) : 0;
  delProgressFill.style.width = `${pct}%`;
  delProgressText.textContent = t("processedOf", [String(done), String(total)]);
  delProgressText.className = "progress-text";
}

async function performDelEntitySearch(query) {
  if (!currentTabId) return;
  try {
    const response = await sendToTab(currentTabId, { type: "SEARCH_ENTITIES", query });
    if (!response?.ok || !response.entities?.length) {
      renderDelEntityDropdown([]);
      return;
    }
    renderDelEntityDropdown(response.entities);
  } catch {
    delEntityDropdown.hidden = true;
  }
}

function renderDelEntityDropdown(entities) {
  renderDropdown(
    delEntityDropdown,
    entities.map((ent) => ({
      name: ent.displayName || ent.logicalName,
      sub: ent.displayName !== ent.logicalName ? `(${ent.logicalName})` : "",
      raw: ent,
    })),
    t("noEntitiesFound"),
    (ent) => {
      delEntityInput.value = ent.logicalName;
      delEntityDropdown.hidden = true;
      updateDelExportBtnState();
    },
  );
}

wireSearchInput({
  input: delEntityInput,
  dropdown: delEntityDropdown,
  searchFn: performDelEntitySearch,
  renderFn: renderDelEntityDropdown,
  onChange: updateDelExportBtnState,
});

async function performDelUserSearch(query) {
  if (!currentTabId) return;
  try {
    const response = await sendToTab(currentTabId, { type: "SEARCH_USERS", query });
    if (!response?.ok || !response.users?.length) {
      renderDelUserDropdown([]);
      return;
    }
    renderDelUserDropdown(response.users);
  } catch {
    delUserDropdown.hidden = true;
  }
}

function renderDelUserDropdown(users) {
  renderDropdown(
    delUserDropdown,
    users.map((user) => ({
      name: user.fullname || t("unnamed"),
      sub: user.email ? `(${user.email})` : "",
      raw: user,
    })),
    t("noUsersFound"),
    (user) => selectDelUser(user),
  );
}

wireSearchInput({
  input: delUserInput,
  dropdown: delUserDropdown,
  searchFn: performDelUserSearch,
  renderFn: renderDelUserDropdown,
});

function selectDelUser(user) {
  delSelectedUser = user;
  delUserInput.value = "";
  delUserDropdown.hidden = true;
  delSelectedUserNameEl.textContent =
    user.fullname || user.email || (user.id ?? "").slice(0, 8);
  delSelectedUserEl.hidden = false;
}

delClearUserBtn.addEventListener("click", () => {
  delSelectedUser = null;
  delSelectedUserEl.hidden = true;
  delSelectedUserNameEl.textContent = "";
});

delDateFromInput.addEventListener("change", updateDelExportBtnState);
delDateToInput.addEventListener("change", updateDelExportBtnState);

async function startDeletedExport() {
  if (delExporting || !currentTabId) return;

  const entityLogicalName = delEntityInput.value.trim().toLowerCase();
  const dateFrom = delDateFromInput.value || null;
  const dateTo = delDateToInput.value || null;
  if (!entityLogicalName || !dateFrom || !dateTo) return;

  delExporting = true;
  delExportBtn.disabled = true;
  delProgressSection.hidden = false;
  delProgressText.textContent = t("queryingDeletions");
  delProgressText.className = "progress-text";
  updateDelProgress(0, 0);

  const port = chrome.tabs.connect(currentTabId, { name: "deleted-export" });

  port.onMessage.addListener((msg) => {
    if (msg.type === "phase") {
      delProgressText.textContent = msg.text;
      delProgressText.className = "progress-text";
    }
    if (msg.type === "progress") {
      updateDelProgress(msg.done, msg.total);
    }
    if (msg.type === "done") {
      const rows = msg.rows ?? [];
      if (rows.length === 0) {
        delProgressText.textContent = t("noDeletionsFound");
        delProgressText.className = "progress-text progress-text--empty";
      } else {
        generateExcel(rows, entityLogicalName, "deleted", {
          sheetName: "Deleted Records",
          filenamePrefix: "DeletedRecords",
        });
        delProgressText.textContent = t("deletedExportComplete", [String(rows.length)]);
        delProgressText.className = "progress-text progress-text--success";
      }
      delExporting = false;
      updateDelExportBtnState();
    }
    if (msg.type === "error") {
      delProgressText.textContent = t("error", [msg.error]);
      delProgressText.className = "progress-text progress-text--error";
      setDelStatus(t("exportFailed"), "error");
      delExporting = false;
      updateDelExportBtnState();
    }
  });

  port.onDisconnect.addListener(() => {
    if (delExporting) {
      delProgressText.textContent = t("connectionLost");
      delProgressText.className = "progress-text progress-text--error";
      setDelStatus(t("exportFailed"), "error");
      delExporting = false;
      updateDelExportBtnState();
    }
  });

  port.postMessage({
    entityLogicalName,
    dateFrom,
    dateTo,
    userGuid: delSelectedUser?.id ?? null,
  });
}

delExportBtn.addEventListener("click", startDeletedExport);

// Initial date defaults for Deleted tab: today for both From and To.
// Use *local* date strings (todayISO uses getDate/getMonth/getFullYear) so the
// values match the user's calendar; otherwise late-night users in non-UTC
// timezones see yesterday's date pre-filled.
(function seedDelDates() {
  const today = new Date();
  const toLocal = (d) =>
    `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  const todayLocal = toLocal(today);
  if (!delDateToInput.value) delDateToInput.value = todayLocal;
  if (!delDateFromInput.value) delDateFromInput.value = todayLocal;
  updateDelExportBtnState();
})();

const fillDataBtn = document.getElementById("fill-data-btn");
const fillStatusEl = document.getElementById("fill-status");

let fillStatusTimer = null;

function showFillStatus(text, type) {
  clearTimeout(fillStatusTimer);
  fillStatusEl.textContent = text;
  fillStatusEl.className = "fill-status fill-status--" + type;
  fillStatusEl.hidden = false;
  if (type === "ok" || type === "err") {
    fillStatusTimer = setTimeout(() => { fillStatusEl.hidden = true; }, 4000);
  }
}

fillDataBtn.addEventListener("click", async () => {
  if (!currentTabId) return;

  fillDataBtn.disabled = true;
  showFillStatus(t("fillingFields"), "loading");

  try {
    const response = await sendToTab(currentTabId, { type: "FILL_DATA" });

    if (response?.ok) {
      const lookupErrs = response.lookupErrors || [];
      if (response.filled > 0) {
        let msg = t("filledFields", [String(response.filled), String(response.total), String(response.skipped)]);
        if (lookupErrs.length > 0) {
          msg += " " + t("lookupIssues", [lookupErrs.slice(0, 2).join("; ")]);
        }
        showFillStatus(msg, "ok");
        fillDataBtn.classList.add("fill-data-flash");
        setTimeout(() => fillDataBtn.classList.remove("fill-data-flash"), 700);
      } else {
        let msg = t("zeroFieldsFilled", [String(response.skipped)]);
        if (lookupErrs.length > 0) {
          msg += " " + t("lookupErrors", [lookupErrs.join("; ")]);
        }
        showFillStatus(msg, "err");
      }
      if (lookupErrs.length > 0 || (response.errors || []).length > 0) {
        console.log("[Audit Lens] Fill debug — formType:", response.formType,
          "sample:", (response.sample || []).join(", "),
          "errors:", (response.errors || []).join("; "),
          "lookupErrors:", lookupErrs.join("; "));
      }
    } else {
      showFillStatus(response?.error || t("failedFill"), "err");
    }
  } catch {
    showFillStatus(t("couldNotReachContent"), "err");
  }

  fillDataBtn.disabled = false;
});

const settingsBtn      = document.getElementById("settings-btn");
const settingsMenu     = document.getElementById("settings-menu");
const themeToggleBtn   = document.getElementById("theme-toggle-btn");
const themeLabel       = document.getElementById("theme-label");
const aboutBtn         = document.getElementById("about-btn");
const aboutModal       = document.getElementById("about-modal");
const aboutCloseBtn    = document.getElementById("about-close-btn");
const langToggleBtn    = document.getElementById("lang-toggle-btn");

function applyTheme(theme) {
  if (theme === "light") {
    document.documentElement.setAttribute("data-theme", "light");
    themeLabel.textContent = t("darkMode");
  } else {
    document.documentElement.removeAttribute("data-theme");
    themeLabel.textContent = t("lightMode");
  }
}

async function loadTheme() {
  try {
    const result = await chrome.storage.local.get(THEME_STORAGE_KEY);
    applyTheme(result?.[THEME_STORAGE_KEY] ?? "light");
  } catch {
    applyTheme("light");
  }
}

async function toggleTheme() {
  const isLight = document.documentElement.getAttribute("data-theme") === "light";
  const next = isLight ? "dark" : "light";
  applyTheme(next);
  try {
    await chrome.storage.local.set({ [THEME_STORAGE_KEY]: next });
  } catch { /* storage unavailable */ }
}

settingsBtn.addEventListener("click", (e) => {
  e.stopPropagation();
  settingsMenu.hidden = !settingsMenu.hidden;
});

document.addEventListener("click", () => {
  settingsMenu.hidden = true;
});

settingsMenu.addEventListener("click", (e) => {
  e.stopPropagation();
});

themeToggleBtn.addEventListener("click", () => {
  settingsMenu.hidden = true;
  toggleTheme();
});

langToggleBtn.addEventListener("click", () => {
  settingsMenu.hidden = true;
  toggleLang();
});

aboutBtn.addEventListener("click", () => {
  settingsMenu.hidden = true;
  aboutModal.hidden = false;
});

const auditSettingsBtn = document.getElementById("audit-settings-btn");

async function resolveAppId(tab) {
  try {
    const appId = new URL(tab.url).searchParams.get("appid");
    if (appId) return appId;
  } catch (_) {}

  try {
    const resp = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: () =>
        fetch("/api/data/v9.2/appmodules?$select=appmoduleid&$top=1", {
          headers: { Accept: "application/json" },
        }).then((r) => r.json()),
    });
    const apps = resp?.[0]?.result?.value;
    if (apps?.length) return apps[0].appmoduleid;
  } catch (_) {}

  return null;
}

auditSettingsBtn.addEventListener("click", async () => {
  settingsMenu.hidden = true;
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  const origin =
    tab?.url && DYNAMICS_PATTERN.test(tab.url)
      ? new URL(tab.url).origin
      : null;

  if (!origin) {
    chrome.tabs.create({ url: "https://admin.powerplatform.microsoft.com/" });
    return;
  }

  const appId = tab ? await resolveAppId(tab) : null;
  const encodedData = encodeURIComponent('{"area":"nav_audit"}');
  const url = `${origin}/main.aspx?appid=${appId}&pagetype=control&controlName=PowerAdmin.EnvironmentSettings.NavigatorPage&data=${encodedData}`;
  chrome.tabs.create({ url });
});

aboutCloseBtn.addEventListener("click", () => {
  aboutModal.hidden = true;
});

aboutModal.addEventListener("click", (e) => {
  if (e.target === aboutModal) {
    aboutModal.hidden = true;
  }
});

document.getElementById("linkedin-btn").addEventListener("click", () => {
  chrome.tabs.create({ url: "https://www.linkedin.com/in/mahmoudzidan55" });
});

document.getElementById("github-repo-btn").addEventListener("click", () => {
  chrome.tabs.create({ url: "https://github.com/Mhmoud-Zidan/dynamics-audit-lens/releases" });
});

// ──────────────────────────────────────────────────────────────────────────
// Feedback banner — shown once per extension version
// ──────────────────────────────────────────────────────────────────────────

const FEEDBACK_DEV_EMAIL = "MahmoudZidan05@Gmail.com";
const FEEDBACK_LINKEDIN = "https://www.linkedin.com/in/mahmoudzidan55";
const FEEDBACK_DISMISS_KEY = "feedbackDismissedVersion";

const feedbackBanner = document.getElementById("feedback-banner");

async function initFeedbackBanner() {
  if (!feedbackBanner) return;
  const currentVersion = chrome.runtime.getManifest().version;
  try {
    const stored = await chrome.storage.local.get(FEEDBACK_DISMISS_KEY);
    if (stored?.[FEEDBACK_DISMISS_KEY] === currentVersion) return;
  } catch { /* storage unavailable — show anyway */ }
  feedbackBanner.hidden = false;
}

// ─── Feedback modal ─────────────────────────────────────────────────────────
// We deliberately avoid mailto: links — many users have no default mail
// client configured (especially enterprise users on browser-only mail).
// Instead we open the form, let them write the message in the popup, then
// hand the composed message off to a web mail composer or the clipboard.

const feedbackModal = document.getElementById("feedback-modal");
const feedbackKindSelect = document.getElementById("feedback-kind");
const feedbackUserEmailInput = document.getElementById("feedback-user-email");
const feedbackMessageInput = document.getElementById("feedback-message");
const feedbackMessageError = document.getElementById("feedback-message-error");
const feedbackToast = document.getElementById("feedback-toast");

function buildFeedbackPayload() {
  const version = chrome.runtime.getManifest().version;
  const kindKey = feedbackKindSelect.value;
  const kindLabel = t(
    { feature: "feedbackKindFeature", bug: "feedbackKindBug", general: "feedbackKindGeneral" }[kindKey]
      ?? "feedbackKindGeneral",
  );
  const message = feedbackMessageInput.value.trim();
  const userEmail = feedbackUserEmailInput.value.trim();

  const subject = `${t("feedbackEmailSubject")} [${kindLabel}]`;
  const bodyLines = [
    message,
    "",
    "—",
    `Type: ${kindLabel}`,
  ];
  if (userEmail) bodyLines.push(`From: ${userEmail}`);
  bodyLines.push(`Extension version: ${version}`);
  const body = bodyLines.join("\n");

  return { subject, body, message, userEmail };
}

function validateFeedbackForm() {
  const empty = feedbackMessageInput.value.trim().length === 0;
  feedbackMessageError.hidden = !empty;
  if (empty) feedbackMessageInput.focus();
  return !empty;
}

function showFeedbackToast(text, isError = false) {
  if (!feedbackToast) return;
  feedbackToast.textContent = text;
  feedbackToast.className = "feedback-form__toast" + (isError ? " feedback-form__toast--error" : "");
  feedbackToast.hidden = false;
  clearTimeout(showFeedbackToast._t);
  showFeedbackToast._t = setTimeout(() => { feedbackToast.hidden = true; }, 3500);
}

async function copyToClipboard(text) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    // Fallback for older browsers / restrictive contexts.
    try {
      const ta = document.createElement("textarea");
      ta.value = text;
      ta.style.position = "fixed";
      ta.style.opacity = "0";
      document.body.appendChild(ta);
      ta.select();
      const ok = document.execCommand("copy");
      document.body.removeChild(ta);
      return ok;
    } catch {
      return false;
    }
  }
}

function openFeedbackModal() {
  if (!feedbackModal) return;
  // Close the settings menu if it's open.
  const menu = document.getElementById("settings-menu");
  if (menu) menu.hidden = true;
  // Reset transient state.
  feedbackMessageError.hidden = true;
  feedbackToast.hidden = true;
  feedbackModal.hidden = false;
  // Focus message field for immediate typing.
  setTimeout(() => feedbackMessageInput?.focus(), 50);
}

function closeFeedbackModal() {
  if (feedbackModal) feedbackModal.hidden = true;
}

document.getElementById("feedback-email-btn")?.addEventListener("click", openFeedbackModal);
document.getElementById("send-feedback-btn")?.addEventListener("click", openFeedbackModal);
document.getElementById("feedback-modal-close-btn")?.addEventListener("click", closeFeedbackModal);
feedbackModal?.addEventListener("click", (e) => {
  // Click on the overlay (not the modal box itself) closes.
  if (e.target === feedbackModal) closeFeedbackModal();
});

// Clear the "required" error as soon as the user starts typing.
feedbackMessageInput?.addEventListener("input", () => {
  if (feedbackMessageInput.value.trim().length > 0) {
    feedbackMessageError.hidden = true;
  }
});

// ── Send via Gmail (web compose) ─────────────────────────────────────────
document.getElementById("feedback-send-gmail")?.addEventListener("click", () => {
  if (!validateFeedbackForm()) return;
  const { subject, body } = buildFeedbackPayload();
  // Gmail compose URL — works regardless of the user's default mail client.
  const url =
    "https://mail.google.com/mail/?view=cm&fs=1" +
    `&to=${encodeURIComponent(FEEDBACK_DEV_EMAIL)}` +
    `&su=${encodeURIComponent(subject)}` +
    `&body=${encodeURIComponent(body)}`;
  chrome.tabs.create({ url });
  showFeedbackToast(t("openingGmail"));
});

// ── Send via Outlook (web compose) ───────────────────────────────────────
document.getElementById("feedback-send-outlook")?.addEventListener("click", () => {
  if (!validateFeedbackForm()) return;
  const { subject, body } = buildFeedbackPayload();
  // Outlook deeplink compose — works for both personal (outlook.live.com) and
  // work (outlook.office.com) accounts; outlook.office.com handles both.
  const url =
    "https://outlook.office.com/mail/deeplink/compose" +
    `?to=${encodeURIComponent(FEEDBACK_DEV_EMAIL)}` +
    `&subject=${encodeURIComponent(subject)}` +
    `&body=${encodeURIComponent(body)}`;
  chrome.tabs.create({ url });
  showFeedbackToast(t("openingOutlook"));
});

// ── Copy formatted message to clipboard ──────────────────────────────────
document.getElementById("feedback-send-copy")?.addEventListener("click", async () => {
  if (!validateFeedbackForm()) return;
  const { subject, body } = buildFeedbackPayload();
  const composed =
    `To: ${FEEDBACK_DEV_EMAIL}\n` +
    `Subject: ${subject}\n` +
    `\n` +
    body;
  const ok = await copyToClipboard(composed);
  if (ok) showFeedbackToast(t("messageCopied"));
  else showFeedbackToast(t("copyFailed"), true);
});

// ── Copy just the email address (one-click chip) ─────────────────────────
document.getElementById("feedback-copy-address")?.addEventListener("click", async () => {
  const ok = await copyToClipboard(FEEDBACK_DEV_EMAIL);
  if (ok) showFeedbackToast(t("addressCopied"));
  else showFeedbackToast(t("copyFailed"), true);
});

document.getElementById("feedback-linkedin-btn")?.addEventListener("click", () => {
  chrome.tabs.create({ url: FEEDBACK_LINKEDIN });
});

document.getElementById("feedback-dismiss-btn")?.addEventListener("click", async () => {
  if (feedbackBanner) feedbackBanner.hidden = true;
  try {
    const v = chrome.runtime.getManifest().version;
    await chrome.storage.local.set({ [FEEDBACK_DISMISS_KEY]: v });
  } catch { /* storage unavailable */ }
});

document.addEventListener("DOMContentLoaded", () => {
  loadLang();
  loadTheme();
  fetchContext();
  initFeedbackBanner();
});

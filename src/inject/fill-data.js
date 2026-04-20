(function fillDataIIFE() {
  "use strict";

  var T_FILL_DATA_RESPONSE = "__DAL__FILL_DATA_RESPONSE";
  var TARGET_ORIGIN = window.location.origin;
  var API_VERSION = "9.2";

  var FIRST_NAMES = [
    "Alex", "Jordan", "Taylor", "Morgan", "Casey", "Riley", "Sarah",
    "Michael", "Emma", "David", "Lisa", "James", "Emily", "Daniel",
    "Sophia", "Oliver", "Ava", "William", "Isabella", "Benjamin",
  ];
  var LAST_NAMES = [
    "Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia",
    "Miller", "Davis", "Rodriguez", "Martinez", "Anderson", "Taylor",
    "Thomas", "Hernandez", "Moore", "Martin", "Jackson", "Thompson",
    "White", "Lopez",
  ];
  var COMPANIES = [
    "Contoso Ltd", "Adventure Works", "Fabrikam Inc", "Tailspin Toys",
    "Litware Inc", "Proseware Corp", "Wide World Importers",
    "Northwind Traders", "Alpine Ski House", "Consolidated Messenger",
    "Graphic Design Institute", "Humongous Insurance",
    "Woodgrove Bank", "Margie's Travel", "VanArsdel Ltd",
  ];
  var DOMAINS = [
    "contoso.com", "adventure-works.com", "fabrikam.com",
    "tailspintoys.com", "litwareinc.com", "proseware.com",
  ];
  var CITIES = [
    "Seattle", "New York", "London", "Chicago", "San Francisco",
    "Boston", "Atlanta", "Denver", "Austin", "Portland", "Toronto",
    "Sydney", "Berlin", "Tokyo", "Paris",
  ];
  var STATES_LIST = [
    "Washington", "New York", "California", "Texas", "Florida",
    "Illinois", "Colorado", "Georgia", "Massachusetts", "Oregon",
  ];
  var COUNTRIES = [
    "United States", "United Kingdom", "Canada", "Australia",
    "Germany", "France", "Japan",
  ];
  var INDUSTRIES = [
    "Technology", "Finance", "Healthcare", "Manufacturing", "Retail",
    "Energy", "Education", "Consulting", "Media", "Transportation",
  ];
  var JOB_TITLES = [
    "Manager", "Senior Manager", "Director", "Vice President",
    "Analyst", "Senior Analyst", "Consultant", "Senior Consultant",
    "Executive", "Administrator", "Coordinator", "Specialist",
    "Lead", "Architect", "Engineer",
  ];
  var STREETS = [
    "123 Main Street", "456 Oak Avenue", "789 Elm Drive",
    "321 Maple Court", "654 Pine Boulevard", "987 Cedar Lane",
    "147 Birch Road", "258 Willow Way", "369 Spruce Circle",
    "741 Aspen Terrace",
  ];
  var DESCRIPTIONS = [
    "Follow-up required to review the current status and next steps with the team.",
    "Initial contact has been made. Awaiting response from the primary decision maker.",
    "Quarterly review completed. Performance metrics are within expected range.",
    "New opportunity identified during the last client meeting. Needs further qualification.",
    "Service request submitted. Technical team has been notified and is investigating.",
    "Proposal sent. Follow-up meeting scheduled for next week to discuss terms.",
    "Onboarding process initiated. All required documents have been collected.",
    "Annual contract renewal is pending. Terms have been reviewed and approved.",
    "Issue escalated to engineering team. Workaround provided to the customer.",
    "Demo scheduled for the product showcase. All stakeholders have been invited.",
  ];
  var SUBJECTS = [
    "Quarterly Review Discussion", "Project Kickoff Meeting",
    "Contract Renewal Follow-up", "Technical Assessment Request",
    "Partnership Opportunity", "Product Demo Session",
    "Service Level Agreement Review", "Budget Planning Workshop",
    "Customer Success Check-in", "Integration Requirements Analysis",
  ];

  var PRIMARY_NAMES = {
    account: "name", contact: "fullname", lead: "fullname",
    systemuser: "fullname", team: "name", businessunit: "name",
    transactioncurrency: "currencyname", territory: "name",
    campaign: "name", opportunity: "name", quote: "name",
    salesorder: "name", invoice: "name", product: "name",
    pricelevel: "name", competitor: "name", knowledgearticle: "title",
    queue: "title", connectionrole: "name", entitlement: "name",
    contract: "title", site: "name", uomschedule: "name",
    uom: "name", discounttype: "name", list: "listname",
  };

  var ENTITY_SET_NAMES = {
    account: "accounts", contact: "contacts", lead: "leads",
    systemuser: "systemusers", team: "teams", businessunit: "businessunits",
    transactioncurrency: "transactioncurrencies", territory: "territories",
    campaign: "campaigns", opportunity: "opportunities", quote: "quotes",
    salesorder: "salesorders", invoice: "invoices", product: "products",
    pricelevel: "pricelevels", competitor: "competitors",
    knowledgearticle: "knowledgearticles", queue: "queues",
    connectionrole: "connectionroles", entitlement: "entitlements",
    contract: "contracts", site: "sites", uomschedule: "uomschedules",
    uom: "uoms", discounttype: "discounttypes", list: "lists",
    email: "emails", phonecall: "phonecalls", task: "tasks",
    appointment: "appointments", letter: "letters", fax: "faxes",
    activitypointer: "activitypointers", goal: "goals", metric: "metrics",
    rollupfield: "rollupfields", post: "posts", postfollow: "postfollows",
    position: "positions", connection: "connections",
    mailmergetemplate: "mailmergetemplates", kbarticle: "kbarticles",
    report: "reports", dashboard: "dashboards", systemform: "systemforms",
    webresource: "webresources", sdkmessageprocessingstep: "sdkmessageprocessingsteps",
    pluginassembly: "pluginassemblies", plugintype: "plugintypes",
    workflow: "workflows", asyncoperation: "asyncoperations",
    import: "imports", importfile: "importfiles", importmap: "importmaps",
    sharepointsite: "sharepointsites", sharepointdocumentlocation: "sharepointdocumentlocations",
    documenttemplate: "documenttemplates", emailserverprofile: "emailserverprofiles",
    mailbox: "mailboxes", queuesettings: "queuesettings",
    SLA: "slas", slaitem: "slaitems", actioncard: "actioncards",
    actioncarduserstate: "actioncarduserstates", channelaccessprofile: "channelaccessprofiles",
    externalparty: "externalparties", interactionforemail: "interactionforemails",
    knowledgearticleincident: "knowledgearticleincidents",
    bookmark: "bookmarks", userform: "userforms",
  };

  function toEntitySet(name) {
    return ENTITY_SET_NAMES[name] || name + "s";
  }

  var ODATA_HEADERS = {
    Accept: "application/json; odata.metadata=minimal",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    "Content-Type": "application/json; charset=utf-8",
    Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
  };

  var SKIP_PATTERNS = [
    /^_/, /id$/, /^created/, /^modified/, /^owning/,
    /^importsequencenumber/, /^overriddencreatedon/,
    /^timezoneruleversionnumber/, /^utcconversiontimezonecode/,
    /^versionnumber/, /^processid$/, /^stageid$/, /^traversedpath/,
    /^statecode$/, /^statuscode$/, /^exchangerate$/,
    /^transactioncurrencyid$/, /^sdkmessage/, /^solutionid/,
    /^solutioncomponent/, /^suppressionsolution/, /^overwritetime/,
    /^componentstate/, /^ismanaged$/, /^iscustomizable$/, /^isretired$/,
  ];

  function pick(arr) { return arr[Math.floor(Math.random() * arr.length)]; }
  function randInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }
  function randDec(min, max, d) {
    return parseFloat((Math.random() * (max - min) + min).toFixed(d || 2));
  }

  function shouldSkip(lower) {
    for (var i = 0; i < SKIP_PATTERNS.length; i++) {
      if (SKIP_PATTERNS[i].test(lower)) return true;
    }
    return false;
  }

  function genEmail() {
    return pick(FIRST_NAMES).toLowerCase() + "." +
      pick(LAST_NAMES).toLowerCase() + "@" + pick(DOMAINS);
  }
  function genPhone() {
    return "+1 (" + randInt(200, 999) + ") " +
      randInt(100, 999) + "-" + randInt(1000, 9999);
  }
  function genUrl() {
    return "https://www." + pick(DOMAINS).replace(".com", "") + ".com";
  }

  function genString(lower, format, attr) {
    if (format === "email" || /email|e-mail/.test(lower)) return genEmail();
    if (format === "phone" || /phone|telephone|fax|mobile|cell/.test(lower))
      return genPhone();
    if (format === "url" || /website|webpage|web_site|url|domain/.test(lower))
      return genUrl();
    if (/firstname|first_name|fname/.test(lower)) return pick(FIRST_NAMES);
    if (/lastname|last_name|lname|surname/.test(lower)) return pick(LAST_NAMES);
    if (/fullname|full_name/.test(lower))
      return pick(FIRST_NAMES) + " " + pick(LAST_NAMES);
    if (/company|organization|employer/.test(lower)) return pick(COMPANIES);
    if (/address.*line|address.*street|streetaddress/.test(lower))
      return pick(STREETS);
    if (/address.*city|city$|city_/.test(lower)) return pick(CITIES);
    if (/address.*state|stateprovince|stateorprovince/.test(lower))
      return pick(STATES_LIST);
    if (/postalcode|zip/.test(lower)) return String(randInt(10000, 99999));
    if (/address.*country|country/.test(lower)) return pick(COUNTRIES);
    if (/jobtitle|job_title/.test(lower) && !/sub/.test(lower))
      return (pick(["Senior ", "Lead ", "Principal ", ""]) +
        pick(JOB_TITLES)).trim();
    if (/industry/.test(lower)) return pick(INDUSTRIES);
    if (/subject|topic/.test(lower)) return pick(SUBJECTS);
    if (/salutation/.test(lower)) return pick(["Mr.", "Ms.", "Dr."]);
    if (/suffix/.test(lower)) return pick(["Jr.", "Sr.", "III", ""]);
    if (/nickname/.test(lower)) return pick(FIRST_NAMES);
    if (/department/.test(lower))
      return pick(["Sales", "Marketing", "Engineering", "Finance", "HR"]);
    if (/territory|region/.test(lower))
      return pick(["North", "South", "East", "West", "Central"]);
    if (/source/.test(lower))
      return pick(["Web", "Referral", "Cold Call", "Trade Show", "Advertisement"]);
    if (/rating/.test(lower)) return pick(["Hot", "Warm", "Cold"]);
    if (/category|classification/.test(lower))
      return pick(["Standard", "Premium", "Enterprise", "Basic"]);
    if (/code/.test(lower))
      return pick(["PRJ", "ORD", "INV", "CAS", "TKT"]) + "-" + randInt(1000, 9999);
    if (/description|notes|comments|details/.test(lower))
      return pick(DESCRIPTIONS);
    if (/account/.test(lower)) return pick(COMPANIES);
    if (/contact/.test(lower))
      return pick(FIRST_NAMES) + " " + pick(LAST_NAMES);
    if (/product/.test(lower))
      return pick(["Pro License", "Enterprise Suite", "Basic Plan",
        "Premium Package", "Starter Kit"]);
    if (/campaign/.test(lower))
      return pick(["Spring", "Summer", "Fall", "Winter"]) + " " +
        pick(["Promo", "Campaign", "Drive"]) + " " + new Date().getFullYear();
    if (/project/.test(lower))
      return pick(["Atlas", "Horizon", "Phoenix", "Pulse", "Nexus"]) +
        " - Phase " + randInt(1, 3);
    if (/task/.test(lower))
      return pick(["Review documentation", "Update records",
        "Schedule meeting", "Prepare report"]);
    if (/quote/.test(lower))
      return pick(COMPANIES) + " - " +
        ["Q1", "Q2", "Q3", "Q4"][Math.floor(new Date().getMonth() / 3)] +
        " Quote";
    if (/invoice/.test(lower)) return "INV-" + randInt(10000, 99999);
    if (/order/.test(lower)) return "ORD-" + randInt(10000, 99999);
    if (/ticket|case/.test(lower))
      return pick(["Login issue", "Performance inquiry",
        "Feature request", "Billing question"]);
    if (/name/.test(lower)) return pick(COMPANIES);
    try {
      var label = attr.getLabel && attr.getLabel();
      if (label)
        return label + " " + pick(["Sample", "Test", "Demo"]) + " " + randInt(1, 999);
    } catch (_) {}
    return "Sample " + randInt(100, 999);
  }

  function genDecimal(lower, fmt) {
    if (fmt === "percentage" || /percent|pct|percentage|rate|score/.test(lower)) {
      return randDec(1, 100, 2);
    }
    if (/discount|savings/.test(lower)) return randDec(5, 50, 2);
    if (/tax|fee/.test(lower)) return randDec(1, 500, 2);
    if (/latitude/.test(lower)) return randDec(-90, 90, 6);
    if (/longitude/.test(lower)) return randDec(-180, 180, 6);
    if (/exchange/.test(lower)) return randDec(0.5, 5, 4);
    return randDec(1, 1000, 2);
  }

  function genInteger(lower) {
    if (/employee|headcount|staff/.test(lower)) return randInt(10, 5000);
    if (/quantity|qty|count/.test(lower)) return randInt(1, 100);
    if (/age|year/.test(lower)) return randInt(1, 30);
    if (/duration|minutes|hours/.test(lower)) return randInt(1, 120);
    if (/percent|rate|score/.test(lower)) return randInt(1, 100);
    if (/priority/.test(lower)) return randInt(1, 5);
    return randInt(1, 100);
  }

  function genMoney(lower) {
    if (/revenue|turnover|income/.test(lower)) return randDec(50000, 5000000, 2);
    if (/budget|cost|expense/.test(lower)) return randDec(1000, 500000, 2);
    if (/price|amount|value|rate|fee/.test(lower)) return randDec(100, 50000, 2);
    if (/discount|savings/.test(lower)) return randDec(10, 5000, 2);
    if (/tax/.test(lower)) return randDec(50, 10000, 2);
    return randDec(100, 100000, 2);
  }

  function genDate() {
    var now = new Date();
    var d = new Date(now.getTime() + randInt(-30, 60) * 86400000);
    d.setHours(randInt(8, 17), randInt(0, 59), 0, 0);
    return d;
  }

  function genOptionValue(attr) {
    try {
      var opts = attr.getOptions && attr.getOptions();
      if (opts && opts.length) {
        var valid = [];
        for (var i = 0; i < opts.length; i++) {
          if (opts[i].value !== -1 && opts[i].value !== null &&
              opts[i].value !== undefined) valid.push(opts[i]);
        }
        if (valid.length) return pick(valid).value;
      }
    } catch (_) {}
    return null;
  }

  function genMultiSelectValue(attr) {
    try {
      var opts = attr.getOptions && attr.getOptions();
      if (opts && opts.length) {
        var valid = [];
        for (var i = 0; i < opts.length; i++) {
          if (opts[i].value !== -1 && opts[i].value !== null &&
              opts[i].value !== undefined) valid.push(opts[i]);
        }
        if (valid.length) {
          var count = randInt(1, Math.min(3, valid.length));
          var result = [];
          var pool = valid.slice();
          for (var j = 0; j < count; j++) {
            var idx = randInt(0, pool.length - 1);
            result.push(pool[idx].value);
            pool.splice(idx, 1);
          }
          return result;
        }
      }
    } catch (_) {}
    return null;
  }

  function getFormContext() {
    var fc = null;

    if (window.Xrm && window.Xrm.Page && window.Xrm.Page.data &&
        window.Xrm.Page.data.entity &&
        window.Xrm.Page.data.entity.attributes) {
      try {
        var testAttrs = window.Xrm.Page.data.entity.attributes.get();
        if (testAttrs && testAttrs.length > 0) {
          fc = window.Xrm.Page;
        }
      } catch (_) {}
    }

    if (!fc) {
      try {
        var app = window.Xrm && window.Xrm.App;
        if (app && app.formContext && app.formContext.data &&
            app.formContext.data.entity &&
            app.formContext.data.entity.attributes) {
          var testAttrs2 = app.formContext.data.entity.attributes.get();
          if (testAttrs2 && testAttrs2.length > 0) {
            fc = app.formContext;
          }
        }
      } catch (_) {}
    }

    if (!fc) {
      try {
        var frames = window.frames;
        for (var fi = 0; fi < frames.length; fi++) {
          try {
            var fXrm = frames[fi].Xrm;
            if (fXrm && fXrm.Page && fXrm.Page.data &&
                fXrm.Page.data.entity &&
                fXrm.Page.data.entity.attributes) {
              var testAttrs3 = fXrm.Page.data.entity.attributes.get();
              if (testAttrs3 && testAttrs3.length > 0) {
                fc = fXrm.Page;
                break;
              }
            }
          } catch (_) {}
        }
      } catch (_) {}
    }

    return fc;
  }

  var lookupCache = {};
  var entityMetaCache = {};

  function lookupPrimaryName(entityName) {
    return PRIMARY_NAMES[entityName] || null;
  }

  function fetchEntityMeta(entityName) {
    if (entityMetaCache[entityName]) return entityMetaCache[entityName];

    var knownName = PRIMARY_NAMES[entityName];
    var knownSet = ENTITY_SET_NAMES[entityName];
    if (knownName && knownSet) {
      entityMetaCache[entityName] = Promise.resolve({
        entitySetName: knownSet,
        primaryNameAttr: knownName,
      });
      return entityMetaCache[entityName];
    }

    var url = TARGET_ORIGIN + "/api/data/v" + API_VERSION +
      "/EntityDefinitions(LogicalName='" + encodeURIComponent(entityName) + "')" +
      "?$select=EntitySetName,PrimaryNameAttribute";

    var promise = fetch(url, {
      method: "GET",
      credentials: "include",
      headers: {
        Accept: "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
      },
    }).then(function (resp) {
      if (!resp.ok) throw new Error("HTTP " + resp.status);
      return resp.json();
    }).then(function (data) {
      return {
        entitySetName: data.EntitySetName || (entityName + "s"),
        primaryNameAttr: data.PrimaryNameAttribute || "name",
      };
    }).catch(function () {
      return {
        entitySetName: ENTITY_SET_NAMES[entityName] || (entityName + "s"),
        primaryNameAttr: knownName || "name",
      };
    });

    entityMetaCache[entityName] = promise;
    return promise;
  }

  function getLookupTargets(attr, aType) {
    try {
      var t = attr.getTargetEntityType && attr.getTargetEntityType();
      if (typeof t === "string" && t) return [t.toLowerCase()];
      if (Array.isArray(t) && t.length) {
        return t.filter(function(s) { return typeof s === "string" && s; })
                .map(function(s) { return s.toLowerCase(); });
      }
    } catch (_) {}

    try {
      var ctrls = attr.controls;
      if (ctrls && ctrls.get) {
        for (var ci = 0; ci < ctrls.getLength(); ci++) {
          try {
            var ctrl = ctrls.get(ci);
            var et = ctrl.getEntityTypes && ctrl.getEntityTypes();
            if (Array.isArray(et) && et.length) {
              return et.filter(function(s) { return typeof s === "string" && s; })
                       .map(function(s) { return s.toLowerCase(); });
            }
          } catch (_) {}
        }
      }
    } catch (_) {}

    if (aType === "owner") return ["systemuser", "team"];
    if (aType === "customer") return ["account", "contact"];

    var n = (attr.getName && attr.getName() || "").toLowerCase();
    if (/parentaccountid/.test(n)) return ["account"];
    if (/primarycontactid|preferredsystemuserid/.test(n)) return ["contact"];
    if (n === "originatingleadid") return ["lead"];
    if (n === "pricelevelid") return ["pricelevel"];
    if (n === "campaignid") return ["campaign"];
    if (n === "territoryid") return ["territory"];
    if (n === "transactioncurrencyid") return ["transactioncurrency"];
    if (n === "defaultpricelevelid") return ["pricelevel"];
    if (n === "parentbusinessunitid") return ["businessunit"];
    if (n === "slaid" || n === "slainvokedid") return ["sla"];
    if (n === "masterid") return [n.replace("id", "")];
    if (/_accountid$/.test(n)) return ["account"];
    if (/_contactid$/.test(n)) return ["contact"];
    if (/_leadid$/.test(n)) return ["lead"];
    if (/_systemuserid$/.test(n)) return ["systemuser"];
    if (/_teamid$/.test(n)) return ["team"];
    if (/_businessunitid$/.test(n)) return ["businessunit"];
    if (/_territoryid$/.test(n)) return ["territory"];
    if (/_productid$/.test(n)) return ["product"];
    if (/_opportunityid$/.test(n)) return ["opportunity"];
    if (/_campaignid$/.test(n)) return ["campaign"];
    if (/_quoteid$/.test(n)) return ["quote"];
    if (/_salesorderid$/.test(n)) return ["salesorder"];
    if (/_invoiceid$/.test(n)) return ["invoice"];
    if (/_pricelevelid$/.test(n)) return ["pricelevel"];

    return [];
  }

  function fetchLookupRecords(entityName) {
    if (lookupCache[entityName]) return lookupCache[entityName];

    var idField = entityName + "id";

    var promise = fetchEntityMeta(entityName).then(function (meta) {
      var nameField = meta.primaryNameAttr;
      var entitySet = meta.entitySetName;

      var filterParts = [];
      if (entityName === "systemuser") {
        filterParts.push("isdisabled eq false");
      }

      var queryString = "?$select=" + nameField + "&$top=5";
      if (filterParts.length) {
        queryString += "&$filter=" + filterParts.join(" and ");
      }

      var url = TARGET_ORIGIN + "/api/data/v" + API_VERSION + "/" +
                entitySet + queryString;

      return fetch(url, {
        method: "GET",
        credentials: "include",
        headers: ODATA_HEADERS,
      }).then(function (resp) {
        if (!resp.ok) throw new Error("HTTP " + resp.status + " for " + entitySet);
        return resp.json();
      }).then(function (data) {
        var entities = data.value || [];
        return entities.map(function (e) {
          return {
            id: e[idField] || "",
            name: e[nameField] || entityName,
            entityType: entityName,
          };
        }).filter(function (r) { return r.id; });
      });
    }).catch(function () {
      return tryXrmWebApiFallback(entityName, idField);
    });

    lookupCache[entityName] = promise;
    return promise;
  }

  function tryXrmWebApiFallback(entityName, idField) {
    try {
      if (!window.Xrm || !window.Xrm.WebApi ||
          !window.Xrm.WebApi.retrieveMultipleRecords) {
        return [];
      }
    } catch (_) {
      return [];
    }

    var cached = entityMetaCache[entityName];
    if (cached && cached.then) {
      return cached.then(function (meta) {
        return doXrmWebApiLookup(entityName, meta.primaryNameAttr, idField);
      }).catch(function () {
        return [];
      });
    }

    var knownName = PRIMARY_NAMES[entityName] || "name";
    return doXrmWebApiLookup(entityName, knownName, idField);
  }

  function doXrmWebApiLookup(entityName, nameField, idField) {
    var filter = "";
    if (entityName === "systemuser") {
      filter = "&$filter=isdisabled eq false";
    }

    return window.Xrm.WebApi.retrieveMultipleRecords(
      entityName,
      "?$select=" + nameField + "&$top=5" + filter
    ).then(function (results) {
      return (results.entities || []).map(function (e) {
        return {
          id: e[idField],
          name: e[nameField] || entityName,
          entityType: entityName,
        };
      }).filter(function (r) { return r.id; });
    }).catch(function () {
      return [];
    });
  }

  async function resolveLookup(attr, aType) {
    var targets = getLookupTargets(attr, aType);
    if (!targets.length) return null;

    for (var t = 0; t < targets.length; t++) {
      try {
        var records = await fetchLookupRecords(targets[t]);
        if (records.length > 0) {
          var rec = pick(records);
          return [{ id: rec.id, name: rec.name, entityType: rec.entityType }];
        }
      } catch (_) {}
    }

    return null;
  }

  async function fillFormData() {
    try {
      var formContext = getFormContext();
      if (!formContext) return { ok: false, error: "No form context available. Open a Dynamics record form first." };

      var entity = formContext.data.entity;
      var formType;
      try { formType = formContext.ui.getFormType(); } catch (_) {}

      if (formType === 3 || formType === 4) {
        return { ok: false, error: "Form is read-only. Switch to an editable form." };
      }

      var attributes;
      try { attributes = entity.attributes.get(); } catch (_) {}
      if (!attributes || !attributes.length) {
        return { ok: false, error: "No form attributes found. Form may still be loading." };
      }

      var filled = 0, skipped = 0, errors = [], sample = [];
      var lookupErrors = [];
      var i, attr, name, aType, fmt, curVal, val;

      for (i = 0; i < attributes.length; i++) {
        attr = attributes[i];
        try { name = attr.getName(); } catch (_) { skipped++; continue; }
        try { aType = attr.getAttributeType(); } catch (_) { aType = null; }
        try { fmt = attr.getFormat(); } catch (_) { fmt = null; }

        if (sample.length < 5) sample.push(name + ":" + (aType || "?"));

        if (shouldSkip(name.toLowerCase())) { skipped++; continue; }

        if (aType === "partylist") { skipped++; continue; }

        try {
          curVal = attr.getValue();
          if (curVal !== null && curVal !== undefined &&
              curVal !== "" && !(Array.isArray(curVal) && curVal.length === 0)) {
            skipped++; continue;
          }

          if (aType === "lookup" || aType === "owner" || aType === "customer") {
            var lookupVal = await resolveLookup(attr, aType);
            if (lookupVal) {
              try {
                attr.setValue(lookupVal);
                try { attr.fireOnChange(); } catch (_) {}
                filled++;
              } catch (setErr) {
                if (lookupErrors.length < 5) {
                  lookupErrors.push(name + " setValue: " + (setErr.message || "err"));
                }
                skipped++;
              }
            } else {
              var targets = getLookupTargets(attr, aType);
              if (lookupErrors.length < 5) {
                lookupErrors.push(name + "(" + aType + "): no records for [" + targets.join(",") + "]");
              }
              skipped++;
            }
            continue;
          }

          val = null;
          switch (aType) {
            case "string":    val = genString(name.toLowerCase(), fmt, attr); break;
            case "memo":      val = pick(DESCRIPTIONS); break;
            case "boolean":   val = Math.random() > 0.5; break;
            case "integer":
            case "bigint":    val = genInteger(name.toLowerCase()); break;
            case "decimal":
            case "double":    val = genDecimal(name.toLowerCase(), fmt); break;
            case "money":     val = genMoney(name.toLowerCase()); break;
            case "datetime":  val = genDate(); break;
            case "optionset": val = genOptionValue(attr); break;
            case "multiselectoptionset": val = genMultiSelectValue(attr); break;
            default: skipped++; continue;
          }

          if (val === null || val === undefined) { skipped++; continue; }

          attr.setValue(val);
          try { attr.fireOnChange(); } catch (_) {}
          filled++;
        } catch (e) {
          skipped++;
          if (errors.length < 3) errors.push(name + ": " + (e.message || "err"));
        }
      }

      return {
        ok: true,
        filled: filled,
        skipped: skipped,
        total: attributes.length,
        formType: formType,
        sample: sample,
        errors: errors,
        lookupErrors: lookupErrors,
      };
    } catch (err) {
      return { ok: false, error: err.message || String(err) };
    }
  }

  fillFormData().then(function (result) {
    window.postMessage({ type: T_FILL_DATA_RESPONSE, payload: result }, TARGET_ORIGIN);
  }).catch(function (err) {
    window.postMessage(
      { type: T_FILL_DATA_RESPONSE, payload: { ok: false, error: err.message || String(err) } },
      TARGET_ORIGIN
    );
  });
})();

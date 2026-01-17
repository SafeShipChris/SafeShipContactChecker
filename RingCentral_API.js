/**************************************************************
 * 🚢 SAFE SHIP CONTACT CHECKER — RingCentral_API.gs
 * ULTIMATE COMPLETE EDITION v9.3.2
 * 
 * 
 * v9.3.2 CHANGES:
 * ✅ FIX: Bounded retry in makeRCRequest_ (max 5 retries)
 * ✅ FIX: Prevents infinite recursion on 429/401 responses
 * 
 * v9.3 CHANGES:
 * 
 * ✅ NEW: syncToken incremental sync (ISync) for 5-minute updates
 * ✅ NEW: RC_API_quickSync() - fast incremental (~5-15 sec vs 60-90 sec)
 * ✅ NEW: Automatic FSync fallback if token expired/invalid
 * ✅ NEW: Quick Sync sidebar with minimal UI
 * ✅ KEPT: Full Smart Sync for daily/initial syncs
 * ✅ KEPT: Direction column, premium UI, yesterday caching
 * 
 * HOW IT WORKS:
 * - Smart Sync (daily): Full 4-stage sync, caches yesterday
 * - Quick Sync (5-min): Uses syncToken, only gets NEW records
 * 
 * @version 9.3.2
 **************************************************************/


/* 
 * CONFIGURATION
 * ══════════════════════════════════════════════════════════════ */

var RC_API_CONFIG = {
  CALL_LOG_SHEET: "RC CALL LOG",
  SMS_LOG_SHEET: "RC SMS LOG",
  AUTH_URL: "https://platform.ringcentral.com/restapi/oauth/authorize",
  TOKEN_URL: "https://platform.ringcentral.com/restapi/oauth/token",
  API_BASE: "https://platform.ringcentral.com/restapi/v1.0",
  PAGE_SIZE: 1000,
  PAGE_DELAY: 300,
  CALL_HEADERS: ["Direction", "Type", "Phone Number", "Name", "Date", "Time", "Action", "Result", "Reason", "Duration"],
  SMS_HEADERS: ["Direction", "Type", "Message Type", "Sender Number", "Sender Name", "Recipient Number", "Recipient Name", "Date / Time", "Segment Count", "Message Status", "Detailed Error Code", "Api Type", "Cost", "Included and Bundle", "Pay as You Go"],
  SYNC_TOKENS: {
    CALL: "RC_SYNC_TOKEN_CALL",
    CALL_TIME: "RC_SYNC_TOKEN_CALL_TIME",
    SMS: "RC_SYNC_TOKEN_SMS",
    SMS_TIME: "RC_SYNC_TOKEN_SMS_TIME"
  }
};

var RC_SMART_CONFIG = {
  SMS_TODAY: "RC SMS LOG",
  SMS_YESTERDAY: "RC SMS LOG YESTERDAY",
  CALL_TODAY: "RC CALL LOG",
  CALL_YESTERDAY: "RC CALL LOG YESTERDAY",
  TZ: "America/New_York",
  PROP_LAST_YESTERDAY_SMS_SYNC: "RC_LAST_YESTERDAY_SMS_SYNC",
  PROP_LAST_YESTERDAY_CALL_SYNC: "RC_LAST_YESTERDAY_CALL_SYNC",
  CACHE_PROGRESS: "RC_SYNC_PROGRESS_V9"
};

var SAFE_SHIP_BRAND = {
  name: "Safe Ship Moving Services",
  shortName: "Safe Ship",
  tagline: "Contact Checker Pro",
  primaryColor: "#8B1538",
  secondaryColor: "#1E3A5F",
  accentColor: "#D4AF37"
};


/* 
 * 🕐 TIMEZONE HELPER
 * ══════════════════════════════════════════════════════════════ */

function RC_getESTDayRange_(dayOffset) {
  var tz = RC_SMART_CONFIG.TZ;
  var now = new Date();

  // Base local date string in target tz
  var baseDateStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");

  // Convert base date to a Date (in tz) without Date constructor math
  var baseNoon = Utilities.parseDate(baseDateStr + " 12:00:00", tz, "yyyy-MM-dd HH:mm:ss");

  // Apply offset using string regeneration (still uses Date internally, but avoids constructor edge cases)
  var target = new Date(baseNoon.getTime() + (dayOffset * 24 * 60 * 60 * 1000));
  var targetDateStr = Utilities.formatDate(target, tz, "yyyy-MM-dd");

  var startDate = Utilities.parseDate(targetDateStr + " 00:00:00", tz, "yyyy-MM-dd HH:mm:ss");
  var endDate = Utilities.parseDate(targetDateStr + " 23:59:59", tz, "yyyy-MM-dd HH:mm:ss");

  var dateFrom = Utilities.formatDate(startDate, "UTC", "yyyy-MM-dd'T'HH:mm:ss") + ".000Z";
  var dateTo = Utilities.formatDate(endDate, "UTC", "yyyy-MM-dd'T'HH:mm:ss") + ".999Z";

  return {
    dateFrom: dateFrom,
    dateTo: dateTo,
    displayDate: targetDateStr,
    displayLabel: Utilities.formatDate(startDate, tz, "EEEE, MMMM d")
  };
}

function convertLocalToUTC_(localDateTimeStr, tz) {
  // Parse local datetime string and convert to UTC ISO
  var localDate = Utilities.parseDate(localDateTimeStr, tz, "yyyy-MM-dd'T'HH:mm:ss");
  return Utilities.formatDate(localDate, "UTC", "yyyy-MM-dd'T'HH:mm:ss") + "Z";
}


/* 
 * 🎨 TOAST HELPER
 * ══════════════════════════════════════════════════════════════ */

function SS_toast_(message, title, duration) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, "🚢 " + (title || SAFE_SHIP_BRAND.shortName), duration || 5);
}


/* 
 * 📊 PROGRESS TRACKING
 * ══════════════════════════════════════════════════════════════ */

function RC_setProgress_(data) {
  CacheService.getUserCache().put(RC_SMART_CONFIG.CACHE_PROGRESS, JSON.stringify(data), 600);
}

function RC_getProgress() {
  var raw = CacheService.getUserCache().get(RC_SMART_CONFIG.CACHE_PROGRESS);
  return raw ? JSON.parse(raw) : null;
}

function RC_clearProgress_() {
  CacheService.getUserCache().remove(RC_SMART_CONFIG.CACHE_PROGRESS);
}


/* 
 * 🔐 OAUTH2 AUTHENTICATION
 * ══════════════════════════════════════════════════════════════ */

function getRingCentralService_() {
  var props = PropertiesService.getScriptProperties();
  var clientId = props.getProperty("RC_CLIENT_ID");
  var clientSecret = props.getProperty("RC_CLIENT_SECRET");
  
  if (!clientId || !clientSecret) {
    throw new Error("Missing RC_CLIENT_ID or RC_CLIENT_SECRET in Script Properties");
  }
  
  return OAuth2.createService("ringcentral")
    .setAuthorizationBaseUrl(RC_API_CONFIG.AUTH_URL)
    .setTokenUrl(RC_API_CONFIG.TOKEN_URL)
    .setClientId(clientId)
    .setClientSecret(clientSecret)
    .setCallbackFunction("authCallback")
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope("ReadCallLog ReadAccounts ReadMessages")
    .setTokenHeaders({
      "Authorization": "Basic " + Utilities.base64Encode(clientId + ":" + clientSecret)
    });
}

function authCallback(request) {
  var service = getRingCentralService_();
  var authorized = service.handleCallback(request);
  
  if (authorized) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;text-align:center;padding:60px;background:linear-gradient(135deg,#0f172a,#1e293b);color:#fff;min-height:100vh;">' +
      '<div style="font-size:80px;margin-bottom:20px;">✅</div>' +
      '<h1 style="color:#10b981;margin:0;">Connected!</h1>' +
      '<p style="color:#94a3b8;margin-top:10px;">RingCentral is now linked to Safe Ship.<br>You can close this tab.</p>' +
      '</div>'
    );
  } else {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;text-align:center;padding:60px;background:linear-gradient(135deg,#0f172a,#1e293b);color:#fff;min-height:100vh;">' +
      '<div style="font-size:80px;margin-bottom:20px;">❌</div>' +
      '<h1 style="color:#ef4444;margin:0;">Authorization Failed</h1>' +
      '<p style="color:#94a3b8;margin-top:10px;">Please close this tab and try again.</p>' +
      '</div>'
    );
  }
}

function RC_API_authorize() {
  var service = getRingCentralService_();
  if (service.hasAccess()) {
    SpreadsheetApp.getUi().alert("✅ Already Connected!", "RingCentral is already authorized.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  var authUrl = service.getAuthorizationUrl();
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;padding:30px;text-align:center;background:linear-gradient(135deg,#0f172a,#1e293b);color:#fff;border-radius:12px;">' +
    '<div style="font-size:48px;margin-bottom:15px;">🔐</div>' +
    '<h2 style="color:' + SAFE_SHIP_BRAND.accentColor + ';margin:0 0 10px;">Connect RingCentral</h2>' +
    '<p style="color:#94a3b8;margin:0 0 20px;font-size:14px;">Authorize Safe Ship to access your call and SMS data</p>' +
    '<a href="' + authUrl + '" target="_blank" style="display:inline-block;padding:14px 32px;background:linear-gradient(135deg,' + SAFE_SHIP_BRAND.primaryColor + ',#6B1028);color:#fff;text-decoration:none;border-radius:8px;font-weight:bold;font-size:16px;">🚀 Authorize Now</a>' +
    '</div>'
  ).setWidth(400).setHeight(280);
  
  SpreadsheetApp.getUi().showModalDialog(html, "🚢 Safe Ship - RingCentral");
}

function RC_API_testConnection() {
  try {
    var service = getRingCentralService_();
    if (!service.hasAccess()) {
      SpreadsheetApp.getUi().alert("❌ Not Connected", "Please run 'Authorize RingCentral' first.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    var data = makeRCRequest_("/account/~");
    SpreadsheetApp.getUi().alert("✅ Connection Successful!", "Account: " + (data.mainNumber || data.id || "Connected"), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Connection Failed", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function RC_API_logout() {
  getRingCentralService_().reset();
  SpreadsheetApp.getUi().alert("🚪 Logged Out", "You have been disconnected from RingCentral.", SpreadsheetApp.getUi().ButtonSet.OK);
}

function RC_API_checkSetup() {
  var props = PropertiesService.getScriptProperties();
  var lines = [
    "🚢 SAFE SHIP - RingCentral Setup Check",
    "═".repeat(40),
    "",
    "CLIENT_ID: " + (props.getProperty("RC_CLIENT_ID") ? "✅ Set" : "❌ Missing"),
    "CLIENT_SECRET: " + (props.getProperty("RC_CLIENT_SECRET") ? "✅ Set" : "❌ Missing"),
    ""
  ];
  
  try {
    var service = getRingCentralService_();
    lines.push("OAuth2: ✅ Configured");
    lines.push("Access: " + (service.hasAccess() ? "✅ Authorized" : "❌ Not authorized"));
  } catch (e) {
    lines.push("OAuth2: ❌ " + e.message);
  }
  
  lines.push("");
  lines.push("─── syncToken Status ───");
  var callToken = props.getProperty(RC_API_CONFIG.SYNC_TOKENS.CALL);
  var callTime = props.getProperty(RC_API_CONFIG.SYNC_TOKENS.CALL_TIME);
  var smsToken = props.getProperty(RC_API_CONFIG.SYNC_TOKENS.SMS);
  var smsTime = props.getProperty(RC_API_CONFIG.SYNC_TOKENS.SMS_TIME);
  lines.push("Call Token: " + (callToken ? "✅ Stored" : "❌ None"));
  lines.push("Call Token Time: " + (callTime || "Never"));
  lines.push("SMS Token: " + (smsToken ? "✅ Stored" : "❌ None"));
  lines.push("SMS Token Time: " + (smsTime || "Never"));
  
  lines.push("");
  lines.push("Timezone: " + RC_SMART_CONFIG.TZ);
  var range = RC_getESTDayRange_(0);
  lines.push("Today: " + range.displayDate);
  
  SpreadsheetApp.getUi().alert(lines.join("\n"));
}


/* 
 * 🌐 API REQUEST HELPER (v9.3.2: BOUNDED RETRY)
 * ══════════════════════════════════════════════════════════════ */

function makeRCRequest_(endpoint, params, options, _retryCount) {
  // v9.3.2: Bounded retry to prevent infinite recursion
  var MAX_RETRIES = 5;
  _retryCount = _retryCount || 0;
  
  if (_retryCount >= MAX_RETRIES) {
    throw new Error("API request failed after " + MAX_RETRIES + " retries - check RingCentral connection and rate limits");
  }
  
  var service = getRingCentralService_();
  if (!service.hasAccess()) {
    throw new Error("Not authorized - please run 'Authorize RingCentral' first");
  }
  
  var url = RC_API_CONFIG.API_BASE + endpoint;
  if (params) {
    var queryString = Object.keys(params).map(function(key) {
      return encodeURIComponent(key) + "=" + encodeURIComponent(params[key]);
    }).join("&");
    url += "?" + queryString;
  }
  
  var fetchOptions = {
    method: (options && options.method) || "GET",
    headers: { "Authorization": "Bearer " + service.getAccessToken() },
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(url, fetchOptions);
  var responseCode = response.getResponseCode();
  
  if (responseCode === 429) {
    Logger.log("Rate limited - waiting 60 seconds (retry " + (_retryCount + 1) + "/" + MAX_RETRIES + ")");
    Utilities.sleep(60000);
    return makeRCRequest_(endpoint, params, options, _retryCount + 1);
  }
  
  if (responseCode === 401) {
    Logger.log("Token expired - refreshing with lock (retry " + (_retryCount + 1) + "/" + MAX_RETRIES + ")");
    var lock = LockService.getScriptLock();
    var gotLock = false;

    try {
      gotLock = lock.tryLock(30000);

      if (gotLock) {
        service.refresh();
      } else {
        // Another execution may be refreshing; back off then retry
        Utilities.sleep(1000);
      }

    } catch (e) {
      Logger.log("Refresh failed: " + e.message);
      // Back off then retry (bounded by MAX_RETRIES)
      Utilities.sleep(1000);

    } finally {
      if (gotLock) lock.releaseLock();
    }

    return makeRCRequest_(endpoint, params, options, _retryCount + 1);
  }
  
  if (responseCode !== 200) {
    throw new Error("API Error: " + responseCode + " - " + response.getContentText().substring(0, 200));
  }
  
  return JSON.parse(response.getContentText());
}


/* 
 * 📞 v9.3: CALL LOG INCREMENTAL SYNC
 * 
 * Uses /account/~/call-log-sync endpoint with syncToken
 * - FSync: Initial sync, establishes token
 * - ISync: Incremental, only new records
 * ══════════════════════════════════════════════════════════════ */

function RC_syncCallsIncremental_() {
  var props = PropertiesService.getScriptProperties();
  var existingToken = props.getProperty(RC_API_CONFIG.SYNC_TOKENS.CALL);
  
  if (existingToken) {
    try {
      var result = RC_doCallISync_(existingToken);
      if (result.success) return result;
      if (result.needsFSync) {
        addLog_("⚠️ Call token expired, doing FSync", "warn");
      }
    } catch (e) {
      addLog_("⚠️ Call ISync failed: " + e.message, "warn");
    }
  }
  
  return RC_doCallFSync_();
}

function RC_doCallFSync_() {
  var props = PropertiesService.getScriptProperties();
  var range = RC_getESTDayRange_(0);
  
  addLog_("📞 FSync: " + range.displayDate, "info");
  
  var params = {
    syncType: "FSync",
    dateFrom: range.dateFrom,
    recordCount: 250,
    view: "Simple"
  };
  
  var data = makeRCRequest_("/account/~/call-log-sync", params);
  var records = data.records || [];
  var syncInfo = data.syncInfo || {};
  
  if (syncInfo.syncToken) {
    props.setProperty(RC_API_CONFIG.SYNC_TOKENS.CALL, syncInfo.syncToken);
    props.setProperty(RC_API_CONFIG.SYNC_TOKENS.CALL_TIME, new Date().toISOString());
  }
  
  var written = RC_writeCallsToSheet_(records, true);
  
  return { success: true, count: written, syncType: "FSync" };
}

function RC_doCallISync_(syncToken) {
  var props = PropertiesService.getScriptProperties();
  
  var params = {
    syncType: "ISync",
    syncToken: syncToken
  };
  
  try {
    var data = makeRCRequest_("/account/~/call-log-sync", params);
    var records = data.records || [];
    var syncInfo = data.syncInfo || {};
    
    if (syncInfo.syncToken) {
      props.setProperty(RC_API_CONFIG.SYNC_TOKENS.CALL, syncInfo.syncToken);
      props.setProperty(RC_API_CONFIG.SYNC_TOKENS.CALL_TIME, new Date().toISOString());
    }
    
    var written = 0;
    if (records.length > 0) {
      written = RC_writeCallsToSheet_(records, false);
    }
    
    return { success: true, count: written, syncType: "ISync" };
    
  } catch (e) {
    if (e.message && (e.message.indexOf("syncToken") >= 0 || e.message.indexOf("InvalidParameter") >= 0)) {
      props.deleteProperty(RC_API_CONFIG.SYNC_TOKENS.CALL);
      return { success: false, needsFSync: true };
    }
    throw e;
  }
}


/* 
 * 💬 v9.3: SMS INCREMENTAL SYNC
 * 
 * Uses /account/~/extension/~/message-sync endpoint
 * ══════════════════════════════════════════════════════════════ */

function RC_syncSMSIncremental_() {
  var props = PropertiesService.getScriptProperties();
  var existingToken = props.getProperty(RC_API_CONFIG.SYNC_TOKENS.SMS);
  
  if (existingToken) {
    try {
      var result = RC_doSMSISync_(existingToken);
      if (result.success) return result;
      if (result.needsFSync) {
        addLog_("⚠️ SMS token expired, doing FSync", "warn");
      }
    } catch (e) {
      addLog_("⚠️ SMS ISync failed: " + e.message, "warn");
    }
  }
  
  return RC_doSMSFSync_();
}

function RC_doSMSFSync_() {
  var props = PropertiesService.getScriptProperties();
  var range = RC_getESTDayRange_(0);
  
  addLog_("💬 SMS FSync: " + range.displayDate, "info");
  
  var params = {
    syncType: "FSync",
    messageType: "SMS",
    dateFrom: range.dateFrom,
    recordCount: 250
  };
  
  var data = makeRCRequest_("/account/~/extension/~/message-sync", params);
  var records = data.records || [];
  var syncInfo = data.syncInfo || {};
  
  records = records.filter(function(r) { return r.type === "SMS"; });
  
  if (syncInfo.syncToken) {
    props.setProperty(RC_API_CONFIG.SYNC_TOKENS.SMS, syncInfo.syncToken);
    props.setProperty(RC_API_CONFIG.SYNC_TOKENS.SMS_TIME, new Date().toISOString());
  }
  
  var written = RC_writeSMSToSheet_(records, true);
  
  return { success: true, count: written, syncType: "FSync" };
}

function RC_doSMSISync_(syncToken) {
  var props = PropertiesService.getScriptProperties();
  
  var params = {
    syncType: "ISync",
    syncToken: syncToken,
    messageType: "SMS"
  };
  
  try {
    var data = makeRCRequest_("/account/~/extension/~/message-sync", params);
    var records = data.records || [];
    var syncInfo = data.syncInfo || {};
    
    records = records.filter(function(r) { return r.type === "SMS"; });
    
    if (syncInfo.syncToken) {
      props.setProperty(RC_API_CONFIG.SYNC_TOKENS.SMS, syncInfo.syncToken);
      props.setProperty(RC_API_CONFIG.SYNC_TOKENS.SMS_TIME, new Date().toISOString());
    }
    
    var written = 0;
    if (records.length > 0) {
      written = RC_writeSMSToSheet_(records, false);
    }
    
    return { success: true, count: written, syncType: "ISync" };
    
  } catch (e) {
    if (e.message && (e.message.indexOf("syncToken") >= 0 || e.message.indexOf("InvalidParameter") >= 0)) {
      props.deleteProperty(RC_API_CONFIG.SYNC_TOKENS.SMS);
      return { success: false, needsFSync: true };
    }
    throw e;
  }
}


/* 
 * 📝 v9.3: SHEET WRITERS (support append mode)
 * ══════════════════════════════════════════════════════════════ */

function RC_writeCallsToSheet_(records, clearFirst) {
  if (!records || !records.length) return 0;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RC_SMART_CONFIG.CALL_TODAY);
  if (!sheet) sheet = ss.insertSheet(RC_SMART_CONFIG.CALL_TODAY);

  var headers = RC_API_CONFIG.CALL_HEADERS;

  if (clearFirst) {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground(SAFE_SHIP_BRAND.primaryColor)
      .setFontColor("#FFFFFF")
      .setFontWeight("bold");
  }

  // Canonical dedupe key based ONLY on the formatted row (stable across runs)
  function buildCallKeyFromRow(row) {
    // row: [direction, type, phone, name, date, time, action, result, reason, duration]
    var direction = row[0] || "";
    var phone = row[2] || "";
    var date = row[4] || "";
    var time = row[5] || "";
    var duration = row[9] || "";
    return date + " " + time + "|" + direction + "|" + phone + "|" + duration;
  }

  var rows = records.map(function (call) { return formatCallRow_(call); });

  // When appending, dedupe against existing data using the same key format
  if (!clearFirst) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var existingData = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
      var existingKeys = Object.create(null);
      for (var i = 0; i < existingData.length; i++) {
        existingKeys[buildCallKeyFromRow(existingData[i])] = true;
      }

      var deduped = [];
      for (var j = 0; j < rows.length; j++) {
        var k = buildCallKeyFromRow(rows[j]);
        if (!existingKeys[k]) {
          existingKeys[k] = true; // also dedupe within this batch
          deduped.push(rows[j]);
        }
      }
      rows = deduped;
    }
  }

  var startRow = clearFirst ? 2 : Math.max(2, sheet.getLastRow() + 1);

  if (rows.length > 0) {
    sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
  }

  var timestamp = Utilities.formatDate(new Date(), RC_SMART_CONFIG.TZ, "MM/dd/yyyy hh:mm:ss a");
  sheet.getRange(1, headers.length + 2).setValue("Updated: " + timestamp);

  return rows.length;
}

function RC_writeSMSToSheet_(records, clearFirst) {
  if (!records || !records.length) return 0;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RC_SMART_CONFIG.SMS_TODAY);
  if (!sheet) sheet = ss.insertSheet(RC_SMART_CONFIG.SMS_TODAY);

  var headers = RC_API_CONFIG.SMS_HEADERS;

  if (clearFirst) {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground(SAFE_SHIP_BRAND.primaryColor)
      .setFontColor("#FFFFFF")
      .setFontWeight("bold");
  }

  // Canonical dedupe key based ONLY on the formatted row (stable across runs)
  function buildSMSKeyFromRow(row) {
    // row: [direction, type, msgType, senderNum, senderName, recipNum, recipName, dateTime, ...]
    var direction = row[0] || "";
    var senderNum = row[3] || "";
    var recipNum = row[5] || "";
    var dateTime = row[7] || "";
    return dateTime + "|" + direction + "|" + senderNum + "|" + recipNum;
  }

  var rows = records.map(function (msg) { return formatSMSRow_(msg); });

  // When appending, dedupe against existing data using the same key format
  if (!clearFirst) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var existingData = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
      var existingKeys = Object.create(null);
      for (var i = 0; i < existingData.length; i++) {
        existingKeys[buildSMSKeyFromRow(existingData[i])] = true;
      }

      var deduped = [];
      for (var j = 0; j < rows.length; j++) {
        var k = buildSMSKeyFromRow(rows[j]);
        if (!existingKeys[k]) {
          existingKeys[k] = true; // also dedupe within this batch
          deduped.push(rows[j]);
        }
      }
      rows = deduped;
    }
  }

  var startRow = clearFirst ? 2 : Math.max(2, sheet.getLastRow() + 1);

  if (rows.length > 0) {
    sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
  }

  var timestamp = Utilities.formatDate(new Date(), RC_SMART_CONFIG.TZ, "MM/dd/yyyy hh:mm:ss a");
  sheet.getRange(1, headers.length + 2).setValue("Updated: " + timestamp);

  return rows.length;
}


/* 
 * 📝 ROW FORMATTERS
 * ══════════════════════════════════════════════════════════════ */

function formatCallRow_(call) {
  var tz = RC_SMART_CONFIG.TZ;
  var startTime = call.startTime ? new Date(call.startTime) : null;
  var direction = call.direction || "Unknown";
  var isInbound = direction.toLowerCase() === "inbound";
  
  var phone = "", name = "";
  
  if (isInbound) {
    if (call.from) {
      if (Array.isArray(call.from) && call.from.length > 0) {
        phone = call.from[0].phoneNumber || call.from[0].extensionNumber || "";
        name = call.from[0].name || "";
      } else if (typeof call.from === 'object' && call.from !== null) {
        phone = call.from.phoneNumber || call.from.extensionNumber || "";
        name = call.from.name || "";
      }
    }
  } else {
    if (call.to) {
      if (Array.isArray(call.to) && call.to.length > 0) {
        phone = call.to[0].phoneNumber || call.to[0].extensionNumber || "";
        name = call.to[0].name || "";
      } else if (typeof call.to === 'object' && call.to !== null) {
        phone = call.to.phoneNumber || call.to.extensionNumber || "";
        name = call.to.name || "";
      }
    }
  }
  
  if (!phone && call.legs && call.legs.length > 0) {
    var leg = call.legs[0];
    var legField = isInbound ? leg.from : leg.to;
    if (legField) {
      if (Array.isArray(legField) && legField.length > 0) {
        phone = legField[0].phoneNumber || legField[0].extensionNumber || "";
        name = legField[0].name || "";
      } else if (typeof legField === 'object' && legField !== null) {
        phone = legField.phoneNumber || legField.extensionNumber || "";
        name = legField.name || "";
      }
    }
  }
  
  return [
    direction,
    call.type || "",
    phone,
    name,
    startTime ? Utilities.formatDate(startTime, tz, "MM/dd/yyyy") : "",
    startTime ? Utilities.formatDate(startTime, tz, "h:mm:ss a") : "",
    call.action || "",
    call.result || "",
    call.reason || "",
    call.duration ? formatDuration_(call.duration) : ""
  ];
}

function formatSMSRow_(msg) {
  var fromInfo = msg.from || {};
  var toInfo = (msg.to && msg.to.length > 0) ? msg.to[0] : {};
  
  return [
    msg.direction || "",
    msg.type || "",
    msg.messageType || "",
    fromInfo.phoneNumber || fromInfo.extensionNumber || "",
    fromInfo.name || "",
    toInfo.phoneNumber || toInfo.extensionNumber || "",
    toInfo.name || "",
    msg.creationTime || "",
    1,
    msg.messageStatus || "",
    msg.smsSendingAttemptsErrorCode || "",
    "", "", "", ""
  ];
}

function formatDuration_(seconds) {
  if (!seconds) return "";
  var h = Math.floor(seconds / 3600);
  var m = Math.floor((seconds % 3600) / 60);
  var s = seconds % 60;
  return (h > 0 ? h + ":" : "") + (m < 10 ? "0" : "") + m + ":" + (s < 10 ? "0" : "") + s;
}

function getTimeStr_() {
  return Utilities.formatDate(new Date(), RC_SMART_CONFIG.TZ, "HH:mm:ss");
}

function addLog_(msg, type) {
  var progress = RC_getProgress();
  if (progress) {
    progress.log = progress.log || [];
    progress.log.push({ time: getTimeStr_(), msg: msg, type: type || "info" });
    if (progress.log.length > 30) progress.log = progress.log.slice(-30);
    RC_setProgress_(progress);
  }
}




/* 
 * 🎯 FULL SMART SYNC (Original 4-Stage - for daily use)
 * ══════════════════════════════════════════════════════════════ */

function RC_API_smartSyncAll() {
  var service = getRingCentralService_();
  if (!service.hasAccess()) {
    SpreadsheetApp.getUi().alert("⚠️ Not Authorized", "Please run 'Authorize RingCentral' first.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  RC_clearProgress_();
  var html = HtmlService.createHtmlOutput(getSmartSyncSidebarHtml_())
    .setTitle("🚢 Safe Ship - Smart Sync")
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

function RC_startSync() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var startTime = new Date().getTime();
  var tz = RC_SMART_CONFIG.TZ;
  var todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var results = { sms1: 0, calls1: 0, sms2: 0, calls2: 0 };
  
  RC_setProgress_({
    percent: 0,
    currentStage: "Initializing...",
    status: "Starting sync engine...",
    stage1: { state: "waiting", status: "Waiting..." },
    stage2: { state: "waiting", status: "Waiting..." },
    stage3: { state: "waiting", status: "Waiting..." },
    stage4: { state: "waiting", status: "Waiting..." },
    log: [{ time: getTimeStr_(), msg: "🚀 Smart Sync started", type: "info" }],
    complete: false,
    startTime: startTime
  });
  
  Utilities.sleep(500);
  
  try {
    // STAGE 1: Yesterday's SMS
    SS_toast_("Stage 1/4: Yesterday's SMS...", "Smart Sync", 30);
    updateStageProgress_(1, "active", "Checking...", 5, "Yesterday's SMS", "Checking if sync needed...");
    addLog_("📱 Stage 1: Yesterday's SMS", "info");
    
    var smsYesterdayResult = RC_smartSyncYesterdaySMS_WithProgress_(todayStr);
    results.sms1 = smsYesterdayResult.count || 0;
    
    updateStageProgress_(1, smsYesterdayResult.skipped ? "skipped" : "complete", 
      smsYesterdayResult.message, 25, "Yesterday's SMS Complete", 
      smsYesterdayResult.message, results.sms1);
    addLog_(smsYesterdayResult.skipped ? "⭐ Cached: " + results.sms1 + " SMS" : "✅ Synced: " + results.sms1 + " SMS", 
      smsYesterdayResult.skipped ? "skip" : "success");
    
    Utilities.sleep(300);
    
    // STAGE 2: Yesterday's Calls
    SS_toast_("Stage 2/4: Yesterday's Calls...", "Smart Sync", 30);
    updateStageProgress_(2, "active", "Checking...", 30, "Yesterday's Calls", "Checking if sync needed...");
    addLog_("📞 Stage 2: Yesterday's Calls", "info");
    
    var callsYesterdayResult = RC_smartSyncYesterdayCalls_WithProgress_(todayStr);
    results.calls1 = callsYesterdayResult.count || 0;
    
    updateStageProgress_(2, callsYesterdayResult.skipped ? "skipped" : "complete",
      callsYesterdayResult.message, 50, "Yesterday's Calls Complete",
      callsYesterdayResult.message, results.calls1);
    addLog_(callsYesterdayResult.skipped ? "⭐ Cached: " + results.calls1 + " calls" : "✅ Synced: " + results.calls1 + " calls",
      callsYesterdayResult.skipped ? "skip" : "success");
    
    Utilities.sleep(300);
    
    // STAGE 3: Today's SMS
    SS_toast_("Stage 3/4: Today's SMS...", "Smart Sync", 60);
    updateStageProgress_(3, "active", "Creating export...", 55, "Today's SMS", "Creating export task...");
    addLog_("💬 Stage 3: Today's SMS", "info");
    
    results.sms2 = RC_API_syncSMS_ToSheet_WithProgress_(0, RC_SMART_CONFIG.SMS_TODAY, "TODAY", 3);
    
    updateStageProgress_(3, "complete", "Synced!", 75, "Today's SMS Complete", 
      results.sms2 + " messages synced", results.sms2);
    addLog_("✅ Synced: " + results.sms2 + " SMS", "success");
    
    Utilities.sleep(300);
    
    // STAGE 4: Today's Calls
    SS_toast_("Stage 4/4: Today's Calls...", "Smart Sync", 60);
    updateStageProgress_(4, "active", "Fetching...", 80, "Today's Calls", "Fetching all calls...");
    addLog_("☎️ Stage 4: Today's Calls", "info");
    
    results.calls2 = RC_API_syncCalls_ToSheet_WithProgress_(0, RC_SMART_CONFIG.CALL_TODAY, "TODAY", 4);
    
    updateStageProgress_(4, "complete", "Synced!", 95, "Today's Calls Complete",
      results.calls2 + " calls synced", results.calls2);
    addLog_("✅ Synced: " + results.calls2 + " calls", "success");
    
    // COMPLETE
    var elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(1);
    var totalSMS = results.sms1 + results.sms2;
    var totalCalls = results.calls1 + results.calls2;
    var total = totalSMS + totalCalls;
    
    var progress = RC_getProgress();
    progress.percent = 100;
    progress.currentStage = "Complete!";
    progress.status = "All stages completed!";
    progress.complete = true;
    progress.stats = { sms: totalSMS, calls: totalCalls, total: total, time: elapsed };
    progress.log.push({ time: getTimeStr_(), msg: "🎉 COMPLETE! " + total + " records in " + elapsed + "s", type: "success" });
    RC_setProgress_(progress);
    
    SS_toast_("✅ Synced " + total + " records in " + elapsed + "s!", "Complete", 10);
    
    return { success: true, total: total, sms: totalSMS, calls: totalCalls, time: elapsed };
    
  } catch (e) {
    Logger.log("Sync error: " + e);
    var progress = RC_getProgress() || {};
    progress.percent = 0;
    progress.currentStage = "Error";
    progress.status = e.message;
    progress.complete = true;
    progress.error = true;
    progress.log = progress.log || [];
    progress.log.push({ time: getTimeStr_(), msg: "❌ ERROR: " + e.message, type: "error" });
    RC_setProgress_(progress);
    
    return { success: false, error: e.message };
  }
}

function updateStageProgress_(stageNum, state, stageStatus, percent, currentStage, status, count) {
  var progress = RC_getProgress() || {};
  progress.percent = percent;
  progress.currentStage = currentStage;
  progress.status = status;
  progress["stage" + stageNum] = { state: state, status: stageStatus };
  if (count !== undefined) progress["stage" + stageNum].count = count;
  RC_setProgress_(progress);
}


/* 
 * YESTERDAY SYNC HELPERS (with caching) - For Smart Sync
 * ══════════════════════════════════════════════════════════════ */

function RC_smartSyncYesterdaySMS_WithProgress_(todayStr) {
  var props = PropertiesService.getScriptProperties();
  var lastSync = props.getProperty(RC_SMART_CONFIG.PROP_LAST_YESTERDAY_SMS_SYNC);
  
  if (lastSync === todayStr) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(RC_SMART_CONFIG.SMS_YESTERDAY);
    if (sheet && sheet.getLastRow() > 1) {
      return { message: "Already synced today", count: sheet.getLastRow() - 1, skipped: true };
    }
  }
  
  var count = RC_API_syncSMS_ToSheet_WithProgress_(-1, RC_SMART_CONFIG.SMS_YESTERDAY, "YESTERDAY", 1);
  props.setProperty(RC_SMART_CONFIG.PROP_LAST_YESTERDAY_SMS_SYNC, todayStr);
  return { message: "Synced!", count: count, skipped: false };
}

function RC_smartSyncYesterdayCalls_WithProgress_(todayStr) {
  var props = PropertiesService.getScriptProperties();
  var lastSync = props.getProperty(RC_SMART_CONFIG.PROP_LAST_YESTERDAY_CALL_SYNC);
  
  if (lastSync === todayStr) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(RC_SMART_CONFIG.CALL_YESTERDAY);
    if (sheet && sheet.getLastRow() > 1) {
      return { message: "Already synced today", count: sheet.getLastRow() - 1, skipped: true };
    }
  }
  
  var count = RC_API_syncCalls_ToSheet_WithProgress_(-1, RC_SMART_CONFIG.CALL_YESTERDAY, "YESTERDAY", 2);
  props.setProperty(RC_SMART_CONFIG.PROP_LAST_YESTERDAY_CALL_SYNC, todayStr);
  return { message: "Synced!", count: count, skipped: false };
}


/* 
 * 📱 SMS SYNC - Message Store Export API (For Smart Sync)
 * ══════════════════════════════════════════════════════════════ */

function RC_API_syncSMS_ToSheet_WithProgress_(dayOffset, sheetName, label, stageNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  
  var service = getRingCentralService_();
  if (!service.hasAccess()) throw new Error("Not authorized");
  
  var range = RC_getESTDayRange_(dayOffset);
  
  var progress = RC_getProgress();
  progress.status = "Creating export task for " + range.displayDate + "...";
  progress["stage" + stageNum] = { state: "active", status: "Creating task..." };
  RC_setProgress_(progress);
  addLog_("📤 Creating export: " + range.displayDate, "info");
  
  var taskId = createMessageStoreReport_(range.dateFrom, range.dateTo);
  if (!taskId) throw new Error("Failed to create task");
  
  var maxAttempts = 60;
  var attempt = 0;
  var status = "";
  
  while (attempt < maxAttempts) {
    attempt++;
    Utilities.sleep(5000);
    status = checkMessageStoreReportStatus_(taskId);
    
    progress = RC_getProgress();
    progress.status = "Export: " + status + " (" + attempt + "/" + maxAttempts + ")";
    progress["stage" + stageNum] = { state: "active", status: status + "..." };
    RC_setProgress_(progress);
    
    if (status === "Completed") break;
    if (status === "Failed" || status === "Cancelled" || status === "AttemptFailed") {
      throw new Error("Export failed: " + status);
    }
  }
  
  if (status !== "Completed") throw new Error("Export timed out");
  
  progress = RC_getProgress();
  progress.status = "Downloading archive...";
  progress["stage" + stageNum] = { state: "active", status: "Downloading..." };
  RC_setProgress_(progress);
  addLog_("📥 Downloading archive...", "info");
  
  var smsRecords = [];
  try {
    smsRecords = downloadAndParseMessageStoreArchive_(taskId);
  } catch (e) {
    Logger.log("Archive error: " + e);
  }
  
  progress = RC_getProgress();
  progress.status = "Writing " + smsRecords.length + " records...";
  progress["stage" + stageNum] = { state: "active", status: "Writing " + smsRecords.length + "..." };
  RC_setProgress_(progress);
  addLog_("📝 Writing " + smsRecords.length + " records", "info");
  
  var headers = RC_API_CONFIG.SMS_HEADERS;
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground(SAFE_SHIP_BRAND.primaryColor).setFontColor("#FFFFFF").setFontWeight("bold");
  
  if (smsRecords.length > 0) {
    var rows = smsRecords.map(function(msg) { return formatSMSRow_(msg); });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  
  var timestamp = Utilities.formatDate(new Date(), RC_SMART_CONFIG.TZ, "MM/dd/yyyy hh:mm:ss a");
  sheet.getRange(1, headers.length + 2).setValue("Updated: " + timestamp + " | " + label + " - " + range.displayDate);
  
  return smsRecords.length;
}


/* 
 * 📞 CALLS SYNC (For Smart Sync)
 * ══════════════════════════════════════════════════════════════ */

function RC_API_syncCalls_ToSheet_WithProgress_(dayOffset, sheetName, label, stageNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  
  var service = getRingCentralService_();
  if (!service.hasAccess()) throw new Error("Not authorized");
  
  var range = RC_getESTDayRange_(dayOffset);
  
  var headers = RC_API_CONFIG.CALL_HEADERS;
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground(SAFE_SHIP_BRAND.primaryColor).setFontColor("#FFFFFF").setFontWeight("bold");
  
  var totalRecords = 0;
  var currentRow = 2;
  var page = 1;
  var pageSize = RC_API_CONFIG.PAGE_SIZE;
  
  while (true) {
    var progress = RC_getProgress();
    progress.status = "Page " + page + ": " + totalRecords + " calls";
    progress["stage" + stageNum] = { state: "active", status: "Page " + page + "..." };
    RC_setProgress_(progress);
    
    var params = {
      dateFrom: range.dateFrom,
      dateTo: range.dateTo,
      perPage: pageSize,
      page: page,
      view: "Detailed"
    };
    
    var data = makeRCRequest_("/account/~/call-log", params);
    var records = data.records || [];
    
    if (records.length > 0) {
      var rows = records.map(function(call) { return formatCallRow_(call); });
      sheet.getRange(currentRow, 1, rows.length, headers.length).setValues(rows);
      currentRow += rows.length;
      totalRecords += records.length;
      
      if (page % 2 === 0) addLog_("🔄 Page " + page + ": " + totalRecords + " calls", "info");
      
      if (records.length < pageSize) break;
      page++;
      Utilities.sleep(RC_API_CONFIG.PAGE_DELAY);
    } else {
      break;
    }
  }
  
  var timestamp = Utilities.formatDate(new Date(), RC_SMART_CONFIG.TZ, "MM/dd/yyyy hh:mm:ss a");
  sheet.getRange(1, headers.length + 2).setValue("Updated: " + timestamp + " | " + label + " - " + range.displayDate);
  
  return totalRecords;
}


/* 
 * 📤 MESSAGE STORE EXPORT API HELPERS (WITH 429 HANDLING)
 * ══════════════════════════════════════════════════════════════ */

function createMessageStoreReport_(dateFrom, dateTo) {
  var service = getRingCentralService_();
  var url = RC_API_CONFIG.API_BASE + "/account/~/message-store-report";
  var maxRetries = 3;
  var retryCount = 0;

  while (retryCount < maxRetries) {
    var response = UrlFetchApp.fetch(url, {
      method: "POST",
      headers: {
        "Authorization": "Bearer " + service.getAccessToken(),
        "Content-Type": "application/json"
      },
      payload: JSON.stringify({ dateFrom: dateFrom, dateTo: dateTo }),
      muteHttpExceptions: true
    });

    var responseCode = response.getResponseCode();

    if (responseCode === 200 || responseCode === 202) {
      var data = JSON.parse(response.getContentText());
      return data.id;
    }

    if (responseCode === 429) {
      retryCount++;
      var waitTime = Math.min(60000 * retryCount, 120000);
      Logger.log(
        "Rate limited (429) on createMessageStoreReport - waiting " +
        (waitTime / 1000) + "s before retry " +
        retryCount + "/" + maxRetries
      );
      addLog_("⚠️ Rate limited - waiting " + (waitTime / 1000) + "s...", "warn");
      Utilities.sleep(waitTime);
      continue;
    }

    if (responseCode === 401) {
      retryCount++;
      var lock = LockService.getScriptLock();
      var gotLock = false;

      try {
        gotLock = lock.tryLock(30000);
        if (gotLock) {
          service.refresh();
        } else {
          // Another execution may be refreshing
          Utilities.sleep(1000);
        }
        continue;
      } catch (e) {
        throw new Error("Failed to refresh token: " + e.message);
      } finally {
        if (gotLock) lock.releaseLock();
      }
    }

    throw new Error(
      "Failed to create report: " +
      responseCode + " - " +
      response.getContentText().substring(0, 200)
    );
  }

  throw new Error(
    "Failed to create report after " +
    maxRetries + " retries (rate limited)"
  );
}

function checkMessageStoreReportStatus_(taskId) {
  var service = getRingCentralService_();
  var url = RC_API_CONFIG.API_BASE + "/account/~/message-store-report/" + taskId;
  var maxRetries = 3;
  var retryCount = 0;
  
  while (retryCount < maxRetries) {
    var response = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: { "Authorization": "Bearer " + service.getAccessToken() },
      muteHttpExceptions: true
    });
    
    var responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      var data = JSON.parse(response.getContentText());
      return data.status;
    }
    
    if (responseCode === 429) {
      retryCount++;
      var waitTime = 30000 * retryCount;
      Logger.log("Rate limited (429) on checkMessageStoreReportStatus - waiting " + (waitTime/1000) + "s");
      Utilities.sleep(waitTime);
      continue;
    }
    
    if (responseCode === 401) {
      retryCount++;
      try {
        service.refresh();
        continue;
      } catch(e) {
        return "Failed";
      }
    }
    
    return "Failed";
  }
  
  return "Failed";
}

function downloadAndParseMessageStoreArchive_(taskId) {
  var service = getRingCentralService_();
  var url = RC_API_CONFIG.API_BASE + "/account/~/message-store-report/" + taskId + "/archive";
  var maxRetries = 3;
  var retryCount = 0;
  
  var archiveData = null;
  while (retryCount < maxRetries) {
    var response = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: { "Authorization": "Bearer " + service.getAccessToken() },
      muteHttpExceptions: true
    });
    
    var responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      archiveData = JSON.parse(response.getContentText());
      break;
    }
    
    if (responseCode === 429) {
      retryCount++;
      var waitTime = 30000 * retryCount;
      Logger.log("Rate limited (429) on archive fetch - waiting " + (waitTime/1000) + "s");
      Utilities.sleep(waitTime);
      continue;
    }
    
    if (responseCode === 401) {
      retryCount++;
      try { service.refresh(); continue; } catch(e) { /* ignore */ }
    }
    
    throw new Error("Failed to get archive: " + responseCode);
  }
  
  if (!archiveData) throw new Error("Failed to get archive after retries");
  
  var records = archiveData.records || [];
  if (records.length === 0) return [];
  if (!records[0].size) return [];
  
  var archiveUri = records[0].uri;
  retryCount = 0;
  var archiveResponse = null;
  
  while (retryCount < maxRetries) {
    archiveResponse = UrlFetchApp.fetch(archiveUri, {
      method: "GET",
      headers: { "Authorization": "Bearer " + service.getAccessToken() },
      muteHttpExceptions: true
    });
    
    var code = archiveResponse.getResponseCode();
    
    if (code === 200) break;
    
    if (code === 429) {
      retryCount++;
      var waitTime = 30000 * retryCount;
      Logger.log("Rate limited (429) on archive download - waiting " + (waitTime/1000) + "s");
      Utilities.sleep(waitTime);
      continue;
    }
    
    throw new Error("Failed to download archive: " + code);
  }
  
  if (!archiveResponse || archiveResponse.getResponseCode() !== 200) {
    throw new Error("Failed to download archive after retries");
  }
  
  var bytes = archiveResponse.getBlob().getBytes();
  var zipBlob = Utilities.newBlob(bytes, "application/zip", "archive.zip");
  var zipFiles = Utilities.unzip(zipBlob);
  
  var allMessages = [];
  for (var i = 0; i < zipFiles.length; i++) {
    if (zipFiles[i].getName().endsWith(".json")) {
      try {
        var content = zipFiles[i].getDataAsString().trim();
        if (content.charCodeAt(0) === 0xFEFF) content = content.substring(1);
        var data = JSON.parse(content);
        if (data.records) {
          for (var j = 0; j < data.records.length; j++) {
            if (data.records[j].type === "SMS") allMessages.push(data.records[j]);
          }
        }
      } catch (e) { Logger.log("Parse error: " + e); }
    }
  }
  return allMessages;
}


/* 
 * 🔧 FORCE SYNC & UTILITY FUNCTIONS
 * ══════════════════════════════════════════════════════════════ */

function RC_API_forceSyncYesterdaySMS() {
  PropertiesService.getScriptProperties().deleteProperty(RC_SMART_CONFIG.PROP_LAST_YESTERDAY_SMS_SYNC);
  var count = RC_API_syncSMS_ToSheet_WithProgress_(-1, RC_SMART_CONFIG.SMS_YESTERDAY, "YESTERDAY", 1);
  var todayStr = Utilities.formatDate(new Date(), RC_SMART_CONFIG.TZ, "yyyy-MM-dd");
  PropertiesService.getScriptProperties().setProperty(RC_SMART_CONFIG.PROP_LAST_YESTERDAY_SMS_SYNC, todayStr);
  SpreadsheetApp.getUi().alert("✅ Force synced " + count + " SMS for yesterday");
}

function RC_API_forceSyncYesterdayCalls() {
  PropertiesService.getScriptProperties().deleteProperty(RC_SMART_CONFIG.PROP_LAST_YESTERDAY_CALL_SYNC);
  var count = RC_API_syncCalls_ToSheet_WithProgress_(-1, RC_SMART_CONFIG.CALL_YESTERDAY, "YESTERDAY", 2);
  var todayStr = Utilities.formatDate(new Date(), RC_SMART_CONFIG.TZ, "yyyy-MM-dd");
  PropertiesService.getScriptProperties().setProperty(RC_SMART_CONFIG.PROP_LAST_YESTERDAY_CALL_SYNC, todayStr);
  SpreadsheetApp.getUi().alert("✅ Force synced " + count + " calls for yesterday");
}

function RC_API_resetSyncCache() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(RC_SMART_CONFIG.PROP_LAST_YESTERDAY_SMS_SYNC);
  props.deleteProperty(RC_SMART_CONFIG.PROP_LAST_YESTERDAY_CALL_SYNC);
  props.deleteProperty(RC_API_CONFIG.SYNC_TOKENS.CALL);
  props.deleteProperty(RC_API_CONFIG.SYNC_TOKENS.CALL_TIME);
  props.deleteProperty(RC_API_CONFIG.SYNC_TOKENS.SMS);
  props.deleteProperty(RC_API_CONFIG.SYNC_TOKENS.SMS_TIME);
  RC_clearProgress_();
  SpreadsheetApp.getUi().alert("🔄 Sync Cache Cleared!\n\nYesterday cache + syncTokens reset.\nNext sync will be a full sync.");
}

function RC_API_showSyncStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();
  var todayStr = Utilities.formatDate(new Date(), RC_SMART_CONFIG.TZ, "yyyy-MM-dd");
  
  var lines = [
    "🚢 SYNC STATUS v9.3.2",
    "═".repeat(40),
    "",
    "Today: " + todayStr,
    ""
  ];
  
  var sheets = [
    RC_SMART_CONFIG.SMS_TODAY,
    RC_SMART_CONFIG.SMS_YESTERDAY,
    RC_SMART_CONFIG.CALL_TODAY,
    RC_SMART_CONFIG.CALL_YESTERDAY
  ];
  
  sheets.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    var count = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
    lines.push(sheetName + ": " + count + " records");
  });
  
  lines.push("");
  lines.push("─── Yesterday Cache ───");
  var lastSMS = props.getProperty(RC_SMART_CONFIG.PROP_LAST_YESTERDAY_SMS_SYNC) || "Never";
  var lastCalls = props.getProperty(RC_SMART_CONFIG.PROP_LAST_YESTERDAY_CALL_SYNC) || "Never";
  lines.push("SMS: " + lastSMS + (lastSMS === todayStr ? " ✅" : ""));
  lines.push("Calls: " + lastCalls + (lastCalls === todayStr ? " ✅" : ""));
  
  lines.push("");
  lines.push("─── syncToken (Quick Sync) ───");
  var callToken = props.getProperty(RC_API_CONFIG.SYNC_TOKENS.CALL);
  var callTime = props.getProperty(RC_API_CONFIG.SYNC_TOKENS.CALL_TIME);
  var smsToken = props.getProperty(RC_API_CONFIG.SYNC_TOKENS.SMS);
  var smsTime = props.getProperty(RC_API_CONFIG.SYNC_TOKENS.SMS_TIME);
  lines.push("Call Token: " + (callToken ? "✅ Active" : "❌ None"));
  if (callTime) lines.push("  Last: " + callTime.substring(0, 19).replace("T", " "));
  lines.push("SMS Token: " + (smsToken ? "✅ Active" : "❌ None"));
  if (smsTime) lines.push("  Last: " + smsTime.substring(0, 19).replace("T", " "));
  
  SpreadsheetApp.getUi().alert(lines.join("\n"));
}


/* 
 * 🎨 SMART SYNC SIDEBAR HTML (Full 4-Stage)
 * ══════════════════════════════════════════════════════════════ */

function getSmartSyncSidebarHtml_() {
  return '<!DOCTYPE html>\
<html>\
<head>\
  <base target="_top">\
  <style>\
    *{box-sizing:border-box;margin:0;padding:0}\
    body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;background:#0f172a;color:#e2e8f0;min-height:100vh;padding:20px}\
    .header{text-align:center;margin-bottom:16px}\
    .header h1{font-size:20px;color:#D4AF37;margin-bottom:4px}\
    .header p{font-size:11px;color:#64748b}\
    .info-box{background:rgba(59,130,246,0.15);border-left:3px solid #3b82f6;border-radius:8px;padding:10px 12px;margin-bottom:16px;font-size:11px;color:#3b82f6}\
    .progress-ring{width:120px;height:120px;margin:0 auto 16px;position:relative}\
    .progress-ring svg{transform:rotate(-90deg)}\
    .progress-ring .bg{fill:none;stroke:#334155;stroke-width:8}\
    .progress-ring .fill{fill:none;stroke:url(#grad);stroke-width:8;stroke-linecap:round;stroke-dasharray:314;stroke-dashoffset:314;transition:stroke-dashoffset 0.3s}\
    .progress-ring .center{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);text-align:center}\
    .progress-ring .pct{font-size:28px;font-weight:800;color:#fff}\
    .progress-ring .label{font-size:10px;color:#64748b}\
    .stage-text{text-align:center;margin-bottom:16px}\
    .stage-text .title{font-size:14px;font-weight:600;color:#fff}\
    .stage-text .sub{font-size:11px;color:#64748b}\
    .stages{background:#1e293b;border-radius:12px;padding:12px;margin-bottom:16px}\
    .stages-title{font-size:10px;color:#D4AF37;font-weight:600;margin-bottom:10px;text-transform:uppercase;letter-spacing:1px}\
    .stage{display:flex;align-items:center;padding:8px;border-radius:8px;margin-bottom:4px;background:rgba(255,255,255,0.02);border-left:3px solid #334155}\
    .stage:last-child{margin-bottom:0}\
    .stage.active{background:rgba(212,175,55,0.15);border-left-color:#D4AF37}\
    .stage.complete{background:rgba(16,185,129,0.15);border-left-color:#10b981}\
    .stage.skipped{background:rgba(100,116,139,0.15);border-left-color:#64748b}\
    .stage .icon{font-size:14px;margin-right:8px}\
    .stage .info{flex:1}\
    .stage .name{font-size:11px;font-weight:600;color:#fff}\
    .stage .status{font-size:9px;color:#64748b}\
    .stage .count{font-size:10px;font-weight:700;color:#10b981;display:none}\
    .log{background:#1e293b;border-radius:12px;padding:12px;max-height:120px;overflow-y:auto;margin-bottom:16px}\
    .log-title{font-size:10px;color:#D4AF37;font-weight:600;margin-bottom:8px;text-transform:uppercase;letter-spacing:1px}\
    .log-entry{font-size:9px;padding:3px 0;border-bottom:1px solid rgba(255,255,255,0.05);display:flex;gap:6px}\
    .log-entry:last-child{border-bottom:none}\
    .log-entry .time{color:#475569;font-family:monospace;min-width:50px}\
    .log-entry .msg{color:#94a3b8}\
    .log-entry.success .msg{color:#10b981}\
    .log-entry.error .msg{color:#ef4444}\
    .log-entry.warn .msg,.log-entry.skip .msg{color:#f59e0b}\
    .stats{display:grid;grid-template-columns:repeat(4,1fr);gap:6px;margin-bottom:16px;display:none}\
    .stat{background:#1e293b;border-radius:8px;padding:10px 6px;text-align:center}\
    .stat .val{font-size:16px;font-weight:800;color:#fff}\
    .stat .lbl{font-size:8px;color:#64748b;text-transform:uppercase}\
    .stat.hl{background:linear-gradient(135deg,#8B1538,#6B1028)}\
    .stat.hl .val{color:#D4AF37}\
    .btn{width:100%;padding:14px;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer}\
    .btn-start{background:linear-gradient(135deg,#D4AF37,#F4D03F);color:#000}\
    .btn-start:disabled{background:#334155;color:#64748b;cursor:not-allowed}\
    .btn-close{background:#334155;color:#fff;margin-top:8px;display:none}\
    .footer{text-align:center;font-size:9px;color:#475569;margin-top:16px}\
  </style>\
</head>\
<body>\
  <div class="header">\
    <h1>🔄 Smart Sync</h1>\
    <p>Full sync: Yesterday + Today</p>\
  </div>\
  <div class="info-box">🔄 Syncs yesterday (cached) + today\'s data. Use for daily refresh or recovery.</div>\
  <div class="progress-ring">\
    <svg viewBox="0 0 120 120" width="120" height="120">\
      <defs><linearGradient id="grad" x1="0%" y1="0%" x2="100%" y2="0%"><stop offset="0%" stop-color="#8B1538"/><stop offset="100%" stop-color="#D4AF37"/></linearGradient></defs>\
      <circle class="bg" cx="60" cy="60" r="50"/>\
      <circle class="fill" id="ringFill" cx="60" cy="60" r="50"/>\
    </svg>\
    <div class="center"><div class="pct" id="pct">0%</div><div class="label">Complete</div></div>\
  </div>\
  <div class="stage-text">\
    <div class="title" id="stageName">Ready</div>\
    <div class="sub" id="stageStatus">Click Start to begin</div>\
  </div>\
  <div class="stages">\
    <div class="stages-title">📋 Sync Stages</div>\
    <div class="stage" id="st1"><div class="icon">📱</div><div class="info"><div class="name">Yesterday\'s SMS</div><div class="status" id="st1s">Waiting...</div></div><div class="count" id="st1c">0</div></div>\
    <div class="stage" id="st2"><div class="icon">📞</div><div class="info"><div class="name">Yesterday\'s Calls</div><div class="status" id="st2s">Waiting...</div></div><div class="count" id="st2c">0</div></div>\
    <div class="stage" id="st3"><div class="icon">💬</div><div class="info"><div class="name">Today\'s SMS</div><div class="status" id="st3s">Waiting...</div></div><div class="count" id="st3c">0</div></div>\
    <div class="stage" id="st4"><div class="icon">☎️</div><div class="info"><div class="name">Today\'s Calls</div><div class="status" id="st4s">Waiting...</div></div><div class="count" id="st4c">0</div></div>\
  </div>\
  <div class="log" id="logBox">\
    <div class="log-title">📜 Activity Log</div>\
    <div id="logList"><div class="log-entry"><span class="time">--:--:--</span><span class="msg">Waiting to start...</span></div></div>\
  </div>\
  <div class="stats" id="stats">\
    <div class="stat hl"><div class="val" id="statTotal">0</div><div class="lbl">Total</div></div>\
    <div class="stat"><div class="val" id="statTime">0s</div><div class="lbl">Time</div></div>\
    <div class="stat"><div class="val" id="statSMS">0</div><div class="lbl">SMS</div></div>\
    <div class="stat"><div class="val" id="statCalls">0</div><div class="lbl">Calls</div></div>\
  </div>\
  <button class="btn btn-start" id="startBtn">🚀 Start Smart Sync</button>\
  <button class="btn btn-close" id="closeBtn">Close</button>\
  <div class="footer">Safe Ship Contact Checker • Smart Sync v9.3.2</div>\
  <script>\
    var polling=null,circ=2*Math.PI*50,ring=document.getElementById("ringFill");\
    ring.style.strokeDasharray=circ;ring.style.strokeDashoffset=circ;\
    document.getElementById("startBtn").onclick=function(){\
      this.disabled=true;this.textContent="Syncing...";\
      document.getElementById("stageName").textContent="Starting...";\
      document.getElementById("stageStatus").textContent="Connecting to RingCentral...";\
      google.script.run.withSuccessHandler(function(r){console.log("Done",r)}).withFailureHandler(function(e){alert("Error: "+e.message)}).RC_startSync();\
      polling=setInterval(checkProgress,700);\
    };\
    document.getElementById("closeBtn").onclick=function(){google.script.host.close()};\
    function checkProgress(){google.script.run.withSuccessHandler(updateUI).withFailureHandler(function(e){console.log(e)}).RC_getProgress()}\
    function updateUI(d){\
      if(!d)return;\
      var pct=d.percent||0;\
      ring.style.strokeDashoffset=circ-(pct/100)*circ;\
      document.getElementById("pct").textContent=pct+"%";\
      document.getElementById("stageName").textContent=d.currentStage||"Processing...";\
      document.getElementById("stageStatus").textContent=d.status||"";\
      for(var i=1;i<=4;i++){\
        var s=d["stage"+i];\
        if(s){\
          var el=document.getElementById("st"+i);\
          el.className="stage "+(s.state||"");\
          document.getElementById("st"+i+"s").textContent=s.status||"Waiting...";\
          if(s.count!=null){document.getElementById("st"+i+"c").textContent=s.count.toLocaleString();document.getElementById("st"+i+"c").style.display="block"}\
        }\
      }\
      if(d.log&&d.log.length){var html="";d.log.forEach(function(l){html+="<div class=\\"log-entry "+(l.type||"")+"\\"><span class=\\"time\\">"+l.time+"</span><span class=\\"msg\\">"+l.msg+"</span></div>"});document.getElementById("logList").innerHTML=html;document.getElementById("logBox").scrollTop=999999}\
      if(d.complete){\
        clearInterval(polling);\
        document.getElementById("startBtn").style.display="none";\
        document.getElementById("closeBtn").style.display="block";\
        if(d.stats){document.getElementById("stats").style.display="grid";document.getElementById("statTotal").textContent=(d.stats.total||0).toLocaleString();document.getElementById("statTime").textContent=(d.stats.time||0)+"s";document.getElementById("statSMS").textContent=(d.stats.sms||0).toLocaleString();document.getElementById("statCalls").textContent=(d.stats.calls||0).toLocaleString()}\
      }\
    }\
  </script>\
</body>\
</html>';
}
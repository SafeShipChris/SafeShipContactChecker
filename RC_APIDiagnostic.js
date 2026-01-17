/**************************************************************
 * 🔍 RC_APIDiagnostic.gs - RingCentral API Verification Tool
 * ══════════════════════════════════════════════════════════════
 * 
 * Runs diagnostic checks to verify:
 * 1. APIs are actually returning data
 * 2. Data is being written to sheets correctly
 * 3. No background processes are running/stuck
 * 
 * Menu: RingCentral API → Diagnostics → 🩺 Full API Diagnostic
 * 
 **************************************************************/


/**
 * 🩺 Run Full API Diagnostic
 */
function RC_runFullDiagnostic() {
  var ui = SpreadsheetApp.getUi();
  var results = [];
  
  results.push("═══════════════════════════════════════════════════════════");
  results.push("🩺 RINGCENTRAL API DIAGNOSTIC REPORT");
  results.push("═══════════════════════════════════════════════════════════");
  results.push("Time: " + new Date().toLocaleString("en-US", { timeZone: "America/New_York" }));
  results.push("");
  
  // 1. Check Authorization
  results.push("─────────────────────────────────────────────────────────────");
  results.push("1️⃣ AUTHORIZATION CHECK");
  results.push("─────────────────────────────────────────────────────────────");
  
  try {
    var service = getRingCentralService_();
    if (service.hasAccess()) {
      results.push("✅ Authorization: VALID");
      results.push("   Token expires: " + (service.getAccessToken() ? "Has token" : "No token"));
    } else {
      results.push("❌ Authorization: EXPIRED or MISSING");
      results.push("   → Run: RingCentral API → Authorize RingCentral");
      ui.alert("Diagnostic Aborted", "Please authorize RingCentral first.", ui.ButtonSet.OK);
      return;
    }
  } catch (e) {
    results.push("❌ Authorization Error: " + e.message);
    ui.alert("Diagnostic Aborted", "Authorization error: " + e.message, ui.ButtonSet.OK);
    return;
  }
  
  results.push("");
  
  // 2. Test Call Log API
  results.push("─────────────────────────────────────────────────────────────");
  results.push("2️⃣ CALL LOG API TEST (/account/~/call-log)");
  results.push("─────────────────────────────────────────────────────────────");
  
  try {
    var range = RC_getESTDayRange_(0);
    results.push("   Date range: " + range.displayDate);
    results.push("   DateFrom: " + range.dateFrom);
    results.push("   DateTo: " + range.dateTo);
    
    var callParams = {
      dateFrom: range.dateFrom,
      dateTo: range.dateTo,
      perPage: 100,
      view: "Simple"
    };
    
    var startTime = new Date().getTime();
    var callData = makeRCRequest_("/account/~/call-log", callParams);
    var elapsed = new Date().getTime() - startTime;
    
    var callRecords = callData.records || [];
    results.push("✅ API Response: " + callRecords.length + " calls returned");
    results.push("   Response time: " + elapsed + "ms");
    
    if (callRecords.length > 0) {
      results.push("   Sample call:");
      var sample = callRecords[0];
      results.push("     - Direction: " + (sample.direction || "?"));
      results.push("     - Type: " + (sample.type || "?"));
      results.push("     - Time: " + (sample.startTime || "?"));
      results.push("     - To: " + (sample.to ? sample.to.phoneNumber : "?"));
    } else {
      results.push("   ⚠️ No calls found for today (this may be normal if no one has called)");
    }
    
    // Count by direction
    var inbound = callRecords.filter(function(c) { return c.direction === "Inbound"; }).length;
    var outbound = callRecords.filter(function(c) { return c.direction === "Outbound"; }).length;
    results.push("   Breakdown: " + inbound + " inbound, " + outbound + " outbound");
    
  } catch (e) {
    results.push("❌ Call Log API Error: " + e.message);
  }
  
  results.push("");
  
  // 3. Test SMS API (Account-level)
  results.push("─────────────────────────────────────────────────────────────");
  results.push("3️⃣ SMS API TEST (/account/~/message-store)");
  results.push("─────────────────────────────────────────────────────────────");
  
  try {
    var smsParams = {
      dateFrom: range.dateFrom,
      dateTo: range.dateTo,
      perPage: 100,
      messageType: "SMS"
    };
    
    var startTime = new Date().getTime();
    var smsData = makeRCRequest_("/account/~/message-store", smsParams);
    var elapsed = new Date().getTime() - startTime;
    
    var smsRecords = (smsData.records || []).filter(function(r) { return r.type === "SMS"; });
    results.push("✅ API Response: " + smsRecords.length + " SMS returned");
    results.push("   Response time: " + elapsed + "ms");
    
    if (smsRecords.length > 0) {
      results.push("   Sample SMS:");
      var sample = smsRecords[0];
      results.push("     - Direction: " + (sample.direction || "?"));
      results.push("     - Type: " + (sample.type || "?"));
      results.push("     - Time: " + (sample.creationTime || "?"));
      results.push("     - From: " + (sample.from ? sample.from.phoneNumber : "?"));
      results.push("     - Status: " + (sample.messageStatus || "?"));
    } else {
      results.push("   ⚠️ No SMS found for today");
    }
    
    // Count by direction
    var inbound = smsRecords.filter(function(s) { return s.direction === "Inbound"; }).length;
    var outbound = smsRecords.filter(function(s) { return s.direction === "Outbound"; }).length;
    results.push("   Breakdown: " + inbound + " inbound, " + outbound + " outbound");
    
  } catch (e) {
    results.push("❌ SMS API Error: " + e.message);
  }
  
  results.push("");
  
  // 4. Test Extension-level API (for comparison)
  results.push("─────────────────────────────────────────────────────────────");
  results.push("4️⃣ EXTENSION-LEVEL SMS TEST (/account/~/extension/~/message-sync)");
  results.push("─────────────────────────────────────────────────────────────");
  
  try {
    var extParams = {
      syncType: "FSync",
      messageType: "SMS",
      dateFrom: range.dateFrom,
      recordCount: 100
    };
    
    var startTime = new Date().getTime();
    var extData = makeRCRequest_("/account/~/extension/~/message-sync", extParams);
    var elapsed = new Date().getTime() - startTime;
    
    var extRecords = (extData.records || []).filter(function(r) { return r.type === "SMS"; });
    results.push("⚠️ Extension-level: " + extRecords.length + " SMS (single extension only!)");
    results.push("   Response time: " + elapsed + "ms");
    results.push("   → This API only sees ONE extension's messages");
    results.push("   → Use account-level /message-store for all company SMS");
    
  } catch (e) {
    results.push("❌ Extension SMS Error: " + e.message);
  }
  
  results.push("");
  
  // 5. Check Sheet Data
  results.push("─────────────────────────────────────────────────────────────");
  results.push("5️⃣ SHEET DATA CHECK");
  results.push("─────────────────────────────────────────────────────────────");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var callSheet = ss.getSheetByName(RC_SMART_CONFIG.CALL_TODAY);
  if (callSheet) {
    var callRows = callSheet.getLastRow() - 1;
    results.push("✅ RC CALL LOG: " + callRows + " records");
    if (callRows > 0) {
      var lastCallCell = callSheet.getRange(1, RC_API_CONFIG.CALL_HEADERS.length + 2).getValue();
      results.push("   Last update: " + lastCallCell);
    }
  } else {
    results.push("⚠️ RC CALL LOG sheet not found");
  }
  
  var smsSheet = ss.getSheetByName(RC_SMART_CONFIG.SMS_TODAY);
  if (smsSheet) {
    var smsRows = smsSheet.getLastRow() - 1;
    results.push("✅ RC SMS LOG: " + smsRows + " records");
    if (smsRows > 0) {
      var lastSmsCell = smsSheet.getRange(1, RC_API_CONFIG.SMS_HEADERS.length + 2).getValue();
      results.push("   Last update: " + lastSmsCell);
    }
  } else {
    results.push("⚠️ RC SMS LOG sheet not found");
  }
  
  results.push("");
  
  // 6. Check for stuck processes
  results.push("─────────────────────────────────────────────────────────────");
  results.push("6️⃣ BACKGROUND PROCESS CHECK");
  results.push("─────────────────────────────────────────────────────────────");
  
  var cache = CacheService.getUserCache();
  var progress = cache.get(RC_SMART_CONFIG.CACHE_PROGRESS);
  
  if (progress) {
    try {
      var progressData = JSON.parse(progress);
      if (progressData.complete) {
        results.push("✅ No active sync in progress (last one completed)");
      } else {
        results.push("⚠️ ACTIVE SYNC DETECTED:");
        results.push("   Stage: " + (progressData.currentStage || progressData.stage || "?"));
        results.push("   Progress: " + (progressData.percent || 0) + "%");
        results.push("   → May be stuck. Clear cache if needed.");
      }
    } catch (e) {
      results.push("⚠️ Progress cache exists but is malformed");
    }
  } else {
    results.push("✅ No active sync process (cache empty)");
  }
  
  // Check script cache for other keys
  var syncTokenSMS = PropertiesService.getScriptProperties().getProperty(RC_API_CONFIG.SYNC_TOKENS.SMS);
  var syncTokenCall = PropertiesService.getScriptProperties().getProperty(RC_API_CONFIG.SYNC_TOKENS.CALL);
  
  results.push("");
  results.push("Sync Tokens:");
  results.push("   SMS token: " + (syncTokenSMS ? "EXISTS" : "none"));
  results.push("   Call token: " + (syncTokenCall ? "EXISTS" : "none"));
  
  results.push("");
  results.push("═══════════════════════════════════════════════════════════");
  results.push("END OF DIAGNOSTIC REPORT");
  results.push("═══════════════════════════════════════════════════════════");
  
  // Show results
  var output = results.join("\n");
  
  // Create HTML dialog for better formatting
  var htmlOutput = HtmlService.createHtmlOutput(
    '<pre style="font-family: monospace; font-size: 11px; background: #1e293b; color: #e2e8f0; padding: 16px; border-radius: 8px; overflow: auto; max-height: 500px;">' + 
    output.replace(/✅/g, '<span style="color:#10b981">✅</span>')
          .replace(/❌/g, '<span style="color:#ef4444">❌</span>')
          .replace(/⚠️/g, '<span style="color:#f59e0b">⚠️</span>') +
    '</pre>'
  )
  .setWidth(600)
  .setHeight(550);
  
  ui.showModalDialog(htmlOutput, "🩺 RingCentral API Diagnostic Report");
}


/**
 * 🔄 Clear Stuck Progress Cache
 */
function RC_clearStuckProgress() {
  var cache = CacheService.getUserCache();
  cache.remove(RC_SMART_CONFIG.CACHE_PROGRESS);
  cache.remove("RC_SYNC_PROGRESS");
  cache.remove("CC_PROGRESS_V1");
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Progress cache cleared. You can now run a new sync.",
    "Cache Cleared",
    5
  );
}


/**
 * 📊 Quick Data Count Check
 */
function RC_quickDataCheck() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];
  
  var sheets = [
    { name: RC_SMART_CONFIG.SMS_TODAY, label: "SMS Today" },
    { name: RC_SMART_CONFIG.SMS_YESTERDAY, label: "SMS Yesterday" },
    { name: RC_SMART_CONFIG.CALL_TODAY, label: "Calls Today" },
    { name: RC_SMART_CONFIG.CALL_YESTERDAY, label: "Calls Yesterday" }
  ];
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = ss.getSheetByName(sheets[i].name);
    if (sheet) {
      var rows = Math.max(0, sheet.getLastRow() - 1);
      results.push(sheets[i].label + ": " + rows + " records");
    } else {
      results.push(sheets[i].label + ": (sheet not found)");
    }
  }
  
  SpreadsheetApp.getUi().alert(
    "📊 Quick Data Check\n\n" + results.join("\n")
  );
}
/**
 * ═══════════════════════════════════════════════════════════════════════════
 * SAFE SHIP CONTACT CHECKER - DATA VALIDATOR v2.0
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * Enhanced validation with:
 * - Lead-to-RC matching analysis (calls & SMS per lead)
 * - Phone column auto-detection with more variations
 * - Coverage statistics and distribution charts
 * - Duplicate detection
 * - Data freshness checks
 * 
 * ═══════════════════════════════════════════════════════════════════════════
 */

var VALIDATOR_VERSION = "2.0";

/**
 * Main entry point - Run full validation suite
 */
function DV_runFullValidation() {
  var html = HtmlService.createHtmlOutput(DV_getValidatorSidebarHtml_())
    .setTitle('🔍 Data Validator')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Execute validation and return results (called from sidebar)
 */
function DV_executeValidation() {
  var startTime = new Date().getTime();
  var results = {
    timestamp: new Date().toISOString(),
    version: VALIDATOR_VERSION,
    duration: 0,
    summary: { total: 0, passed: 0, warnings: 0, errors: 0 },
    categories: []
  };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Try to load config safely
    var CFG = {};
    try {
      if (typeof getConfig_ === 'function') {
        CFG = getConfig_() || {};
      }
    } catch (e) { /* Config not available */ }
    
    // CATEGORY 1: Lead Data Integrity
    try {
      results.categories.push(DV_validateLeadData_(ss));
    } catch (e) {
      results.categories.push({
        name: "Lead Data Integrity", icon: "📋",
        checks: [{ name: "Scan Error", status: "error", message: e.message, details: [] }]
      });
    }
    
    // CATEGORY 2: RingCentral Data
    try {
      results.categories.push(DV_validateRingCentralData_(ss));
    } catch (e) {
      results.categories.push({
        name: "RingCentral Data", icon: "📞",
        checks: [{ name: "Scan Error", status: "error", message: e.message, details: [] }]
      });
    }
    
    // CATEGORY 3: Lead ↔ RC Matching (NEW!)
    try {
      results.categories.push(DV_validateLeadRCMatching_(ss));
    } catch (e) {
      results.categories.push({
        name: "Lead ↔ RC Matching", icon: "🔗",
        checks: [{ name: "Scan Error", status: "error", message: e.message, details: [] }]
      });
    }
    
    // CATEGORY 4: Rosters
    try {
      results.categories.push(DV_validateRosters_(ss));
    } catch (e) {
      results.categories.push({
        name: "Roster Validation", icon: "👥",
        checks: [{ name: "Scan Error", status: "error", message: e.message, details: [] }]
      });
    }
    
    // CATEGORY 5: Data Quality
    try {
      results.categories.push(DV_validateDataQuality_(ss));
    } catch (e) {
      results.categories.push({
        name: "Data Quality", icon: "✨",
        checks: [{ name: "Scan Error", status: "error", message: e.message, details: [] }]
      });
    }
    
    // CATEGORY 6: Configuration
    try {
      results.categories.push(DV_validateConfiguration_(CFG));
    } catch (e) {
      results.categories.push({
        name: "Configuration", icon: "⚙️",
        checks: [{ name: "Scan Error", status: "error", message: e.message, details: [] }]
      });
    }
    
    // Calculate summary
    for (var i = 0; i < results.categories.length; i++) {
      var cat = results.categories[i];
      if (cat.checks) {
        for (var j = 0; j < cat.checks.length; j++) {
          results.summary.total++;
          var status = cat.checks[j].status;
          if (status === "pass") results.summary.passed++;
          else if (status === "warning") results.summary.warnings++;
          else if (status === "error") results.summary.errors++;
        }
      }
    }
    
  } catch (e) {
    results.categories.push({
      name: "System Error", icon: "💥",
      checks: [{
        name: "Validation Engine", status: "error",
        message: "Validation failed: " + e.message,
        details: [String(e.stack || "")]
      }]
    });
    results.summary.errors++;
    results.summary.total++;
  }
  
  results.duration = ((new Date().getTime() - startTime) / 1000).toFixed(1);
  return results;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * CATEGORY 1: LEAD DATA INTEGRITY
 * ═══════════════════════════════════════════════════════════════════════════ */

function DV_validateLeadData_(ss) {
  var category = { name: "Lead Data Integrity", icon: "📋", checks: [] };
  
  // Auto-discover tracker sheets
  var commonNames = [
    "SMS TRACKER", 
    "CALL & VOICEMAIL TRACKER", 
    "CONTACTED LEADS",
    "PRIORITY 1 CALL & VOICEMAIL TRA",
    "SAME DAY TRANSFERS"
  ];
  
  var trackerSheets = [];
  for (var i = 0; i < commonNames.length; i++) {
    try {
      var sheet = ss.getSheetByName(commonNames[i]);
      if (sheet) {
        trackerSheets.push({ name: commonNames[i], sheet: sheet });
      }
    } catch (e) { }
  }
  
  // Report found sheets
  var sheetDetails = [];
  for (var j = 0; j < trackerSheets.length; j++) {
    sheetDetails.push("✓ " + trackerSheets[j].name);
  }
  
  category.checks.push({
    name: "Tracker Sheets Found",
    status: trackerSheets.length > 0 ? "pass" : "warning",
    message: trackerSheets.length + " tracker sheets found",
    details: sheetDetails
  });
  
  // Validate each sheet
  for (var k = 0; k < trackerSheets.length; k++) {
    try {
      var results = DV_validateTrackerSheet_(trackerSheets[k].sheet, trackerSheets[k].name);
      if (results && results.length) {
        for (var m = 0; m < results.length; m++) {
          category.checks.push(results[m]);
        }
      }
    } catch (e) {
      category.checks.push({
        name: trackerSheets[k].name + " - Error",
        status: "error", message: e.message, details: []
      });
    }
  }
  
  return category;
}

function DV_validateTrackerSheet_(sheet, sheetName) {
  var checks = [];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  if (lastRow < 2 || lastCol < 1) {
    checks.push({ name: sheetName, status: "warning", message: "Empty or no data", details: [] });
    return checks;
  }
  
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (!headers || headers.length === 0) {
    checks.push({ name: sheetName, status: "warning", message: "No headers found", details: [] });
    return checks;
  }
  
  var sampleSize = Math.min(lastRow - 1, 500);
  var data = sheet.getRange(2, 1, sampleSize, lastCol).getValues();
  
  // Use enhanced phone column finder
  var phoneCol = DV_findPhoneColumn_(headers);
  
  // Count issues
  var blankPhones = 0, invalidPhones = 0;
  for (var i = 0; i < data.length; i++) {
    if (phoneCol >= 0) {
      var phone = String(data[i][phoneCol] || "").trim();
      if (!phone) blankPhones++;
      else if (phone.replace(/\D/g, "").length < 10) invalidPhones++;
    }
  }
  
  var phoneIssues = blankPhones + invalidPhones;
  var colStatus = phoneCol >= 0 ? "Found (col " + (phoneCol + 1) + ": \"" + headers[phoneCol] + "\")" : "NOT FOUND";
  
  // Show first few headers for debugging if phone not found
  var details = ["Phone col: " + colStatus];
  if (phoneCol < 0) {
    var headerSample = [];
    for (var h = 0; h < Math.min(headers.length, 8); h++) {
      if (headers[h]) headerSample.push((h+1) + ":" + String(headers[h]).substring(0, 15));
    }
    details.push("Headers: " + headerSample.join(", "));
  } else {
    details.push("Blank phones: " + blankPhones);
    details.push("Invalid phones: " + invalidPhones);
  }
  
  checks.push({
    name: sheetName + " - Data",
    status: phoneCol < 0 ? "warning" : (phoneIssues === 0 ? "pass" : (phoneIssues < 50 ? "warning" : "error")),
    message: (lastRow - 1) + " rows" + (phoneCol >= 0 ? ", " + phoneIssues + " phone issues" : ", phone col not found"),
    details: details
  });
  
  return checks;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * CATEGORY 2: RINGCENTRAL DATA
 * ═══════════════════════════════════════════════════════════════════════════ */

function DV_validateRingCentralData_(ss) {
  var category = { name: "RingCentral Data", icon: "📞", checks: [] };
  
  var rcCallLog = ss.getSheetByName("RC CALL LOG");
  var rcSmsLog = ss.getSheetByName("RC SMS LOG");
  var rcCallYesterday = ss.getSheetByName("RC CALL LOG YESTERDAY");
  var rcSmsYesterday = ss.getSheetByName("RC SMS LOG YESTERDAY");
  
  // Call Log Today
  if (rcCallLog) {
    var stats = DV_analyzeRCSheet_(rcCallLog, "call");
    category.checks.push({
      name: "Today's Calls",
      status: "pass",
      message: stats.records.toLocaleString() + " records → " + stats.uniquePhones.toLocaleString() + " unique phones",
      details: [
        "Outbound: " + stats.outbound.toLocaleString(),
        "Inbound: " + stats.inbound.toLocaleString(),
        "Avg duration: " + stats.avgDuration + "s"
      ]
    });
  } else {
    category.checks.push({
      name: "Today's Calls", status: "warning",
      message: "RC CALL LOG sheet not found", details: []
    });
  }
  
  // SMS Log Today
  if (rcSmsLog) {
    var smsStats = DV_analyzeRCSheet_(rcSmsLog, "sms");
    category.checks.push({
      name: "Today's SMS",
      status: "pass",
      message: smsStats.records.toLocaleString() + " records → " + smsStats.uniquePhones.toLocaleString() + " unique phones",
      details: [
        "Outbound: " + smsStats.outbound.toLocaleString(),
        "Inbound: " + smsStats.inbound.toLocaleString()
      ]
    });
  } else {
    category.checks.push({
      name: "Today's SMS", status: "warning",
      message: "RC SMS LOG sheet not found", details: []
    });
  }
  
  // Yesterday data
  var yesterdayDetails = [];
  if (rcCallYesterday) yesterdayDetails.push("✓ Calls: " + (rcCallYesterday.getLastRow() - 1) + " records");
  if (rcSmsYesterday) yesterdayDetails.push("✓ SMS: " + (rcSmsYesterday.getLastRow() - 1) + " records");
  
  category.checks.push({
    name: "Yesterday's Data",
    status: (rcCallYesterday || rcSmsYesterday) ? "pass" : "warning",
    message: yesterdayDetails.length > 0 ? "Available" : "Not found",
    details: yesterdayDetails
  });
  
  // Duration parsing test
  category.checks.push(DV_testDurationParsing_());
  
  return category;
}

function DV_analyzeRCSheet_(sheet, type) {
  var result = { records: 0, uniquePhones: 0, outbound: 0, inbound: 0, avgDuration: 0 };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return result;
  
  result.records = lastRow - 1;
  
  var lastCol = Math.min(sheet.getLastColumn(), 20);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var sampleSize = Math.min(lastRow - 1, 2000);
  var data = sheet.getRange(2, 1, sampleSize, lastCol).getValues();
  
  var phoneCol = DV_findColumn_(headers, ["phone number", "to", "from", "phone", "recipient", "contact"]);
  var dirCol = DV_findColumn_(headers, ["direction", "type", "call direction"]);
  var durCol = DV_findColumn_(headers, ["duration", "length", "call duration", "talk time"]);
  
  var phones = {};
  var totalDuration = 0;
  var durationCount = 0;
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    
    if (phoneCol >= 0) {
      var phone = String(row[phoneCol] || "").replace(/\D/g, "");
      if (phone.length >= 10) phones[phone.slice(-10)] = true;
    }
    
    if (dirCol >= 0) {
      var dir = String(row[dirCol] || "").toLowerCase();
      if (dir.indexOf("outbound") >= 0 || dir.indexOf("out") >= 0) result.outbound++;
      else if (dir.indexOf("inbound") >= 0 || dir.indexOf("in") >= 0) result.inbound++;
    }
    
    if (durCol >= 0 && type === "call") {
      var dur = DV_parseDuration_(row[durCol]);
      if (dur > 0) {
        totalDuration += dur;
        durationCount++;
      }
    }
  }
  
  result.uniquePhones = Object.keys(phones).length;
  result.avgDuration = durationCount > 0 ? Math.round(totalDuration / durationCount) : 0;
  
  return result;
}

function DV_testDurationParsing_() {
  var tests = [
    { input: "0:27:00", expected: 27 },
    { input: "1:30:00", expected: 90 },
    { input: "2:48:00", expected: 168 }
  ];
  
  var passed = 0;
  var details = [];
  
  for (var i = 0; i < tests.length; i++) {
    var result = DV_parseDuration_(tests[i].input);
    if (result === tests[i].expected) {
      passed++;
      details.push(tests[i].input + " → " + result + "s ✓");
    } else {
      details.push(tests[i].input + " → " + result + "s (expected " + tests[i].expected + ") ✗");
    }
  }
  
  return {
    name: "Duration Parser",
    status: passed === tests.length ? "pass" : "error",
    message: passed + "/" + tests.length + " tests passed",
    details: details
  };
}

function DV_parseDuration_(val) {
  if (!val) return 0;
  if (val instanceof Date) return val.getHours() * 60 + val.getMinutes();
  var str = String(val).trim();
  var parts = str.split(":");
  if (parts.length >= 2) {
    return (parseInt(parts[0], 10) || 0) * 60 + (parseInt(parts[1], 10) || 0);
  }
  return 0;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * CATEGORY 3: LEAD ↔ RC MATCHING (NEW!)
 * ═══════════════════════════════════════════════════════════════════════════ */

function DV_validateLeadRCMatching_(ss) {
  var category = { name: "Lead ↔ RC Matching", icon: "🔗", checks: [] };
  
  // Build RC phone index (calls + SMS combined)
  var rcIndex = DV_buildRCPhoneIndex_(ss);
  
  category.checks.push({
    name: "RC Phone Index",
    status: "pass",
    message: rcIndex.totalPhones.toLocaleString() + " unique phones indexed",
    details: [
      "From calls: " + rcIndex.callPhones.toLocaleString() + " phones, " + rcIndex.totalCalls.toLocaleString() + " records",
      "From SMS: " + rcIndex.smsPhones.toLocaleString() + " phones, " + rcIndex.totalSms.toLocaleString() + " records"
    ]
  });
  
  // Analyze each tracker sheet against RC data
  var trackerNames = [
    "SMS TRACKER", 
    "CALL & VOICEMAIL TRACKER", 
    "CONTACTED LEADS",
    "SAME DAY TRANSFERS"
  ];
  
  for (var i = 0; i < trackerNames.length; i++) {
    var sheet = ss.getSheetByName(trackerNames[i]);
    if (sheet) {
      var matchResult = DV_analyzeLeadRCMatch_(sheet, trackerNames[i], rcIndex);
      category.checks.push(matchResult);
    }
  }
  
  // Overall coverage summary
  var allLeadsResult = DV_analyzeAllLeadsCoverage_(ss, rcIndex);
  category.checks.push(allLeadsResult);
  
  return category;
}

function DV_buildRCPhoneIndex_(ss) {
  var index = {
    phones: {},        // phone -> { calls: count, sms: count, totalDuration: seconds, lastActivity: date }
    totalPhones: 0,
    callPhones: 0,
    smsPhones: 0,
    totalCalls: 0,
    totalSms: 0
  };
  
  // Process call logs (today + yesterday)
  var callSheets = ["RC CALL LOG", "RC CALL LOG YESTERDAY"];
  var callPhonesSet = {};
  
  for (var i = 0; i < callSheets.length; i++) {
    var sheet = ss.getSheetByName(callSheets[i]);
    if (!sheet || sheet.getLastRow() < 2) continue;
    
    var lastCol = Math.min(sheet.getLastColumn(), 20);
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    
    var phoneCol = DV_findColumn_(headers, ["phone number", "to", "from", "phone", "recipient"]);
    var durCol = DV_findColumn_(headers, ["duration", "length", "call duration"]);
    var dateCol = DV_findColumn_(headers, ["date", "call date", "time", "start time"]);
    
    for (var j = 0; j < data.length; j++) {
      var row = data[j];
      if (phoneCol < 0) continue;
      
      var phone = DV_normalizePhone_(row[phoneCol]);
      if (!phone) continue;
      
      index.totalCalls++;
      callPhonesSet[phone] = true;
      
      if (!index.phones[phone]) {
        index.phones[phone] = { calls: 0, sms: 0, totalDuration: 0, lastActivity: null };
      }
      
      index.phones[phone].calls++;
      
      if (durCol >= 0) {
        index.phones[phone].totalDuration += DV_parseDuration_(row[durCol]);
      }
      
      if (dateCol >= 0 && row[dateCol]) {
        var actDate = new Date(row[dateCol]);
        if (!index.phones[phone].lastActivity || actDate > index.phones[phone].lastActivity) {
          index.phones[phone].lastActivity = actDate;
        }
      }
    }
  }
  
  index.callPhones = Object.keys(callPhonesSet).length;
  
  // Process SMS logs (today + yesterday)
  var smsSheets = ["RC SMS LOG", "RC SMS LOG YESTERDAY"];
  var smsPhonesSet = {};
  
  for (var k = 0; k < smsSheets.length; k++) {
    var sheet = ss.getSheetByName(smsSheets[k]);
    if (!sheet || sheet.getLastRow() < 2) continue;
    
    var lastCol = Math.min(sheet.getLastColumn(), 20);
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    
    var phoneCol = DV_findColumn_(headers, ["phone number", "to", "from", "phone", "recipient"]);
    var dateCol = DV_findColumn_(headers, ["date", "time", "sent time", "received time"]);
    
    for (var m = 0; m < data.length; m++) {
      var row = data[m];
      if (phoneCol < 0) continue;
      
      var phone = DV_normalizePhone_(row[phoneCol]);
      if (!phone) continue;
      
      index.totalSms++;
      smsPhonesSet[phone] = true;
      
      if (!index.phones[phone]) {
        index.phones[phone] = { calls: 0, sms: 0, totalDuration: 0, lastActivity: null };
      }
      
      index.phones[phone].sms++;
      
      if (dateCol >= 0 && row[dateCol]) {
        var actDate = new Date(row[dateCol]);
        if (!index.phones[phone].lastActivity || actDate > index.phones[phone].lastActivity) {
          index.phones[phone].lastActivity = actDate;
        }
      }
    }
  }
  
  index.smsPhones = Object.keys(smsPhonesSet).length;
  index.totalPhones = Object.keys(index.phones).length;
  
  return index;
}

function DV_analyzeLeadRCMatch_(sheet, sheetName, rcIndex) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  if (lastRow < 2 || lastCol < 1) {
    return {
      name: sheetName + " - RC Match",
      status: "warning",
      message: "No data to analyze",
      details: []
    };
  }
  
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var sampleSize = Math.min(lastRow - 1, 1000);
  var data = sheet.getRange(2, 1, sampleSize, lastCol).getValues();
  
  // Use enhanced phone column finder
  var phoneCol = DV_findPhoneColumn_(headers);
  
  if (phoneCol < 0) {
    // Show headers for debugging
    var headerSample = [];
    for (var h = 0; h < Math.min(headers.length, 10); h++) {
      if (headers[h]) headerSample.push(String(headers[h]).substring(0, 20));
    }
    return {
      name: sheetName + " - RC Match",
      status: "warning",
      message: "Phone column not found",
      details: [
        "Cannot match leads without phone column",
        "Available headers: " + headerSample.join(", ")
      ]
    };
  }
  
  // Analyze matches
  var stats = {
    total: 0,
    withCalls: 0,
    withSms: 0,
    withBoth: 0,
    withNone: 0,
    distribution: { zero: 0, low: 0, med: 0, high: 0 }  // 0, 1-3, 4-10, 10+
  };
  
  var noActivitySamples = [];
  var highActivitySamples = [];
  
  for (var i = 0; i < data.length; i++) {
    var phone = DV_normalizePhone_(data[i][phoneCol]);
    if (!phone) continue;
    
    stats.total++;
    var rc = rcIndex.phones[phone];
    
    if (rc) {
      var totalActivity = rc.calls + rc.sms;
      
      if (rc.calls > 0) stats.withCalls++;
      if (rc.sms > 0) stats.withSms++;
      if (rc.calls > 0 && rc.sms > 0) stats.withBoth++;
      
      if (totalActivity === 0) {
        stats.distribution.zero++;
      } else if (totalActivity <= 3) {
        stats.distribution.low++;
      } else if (totalActivity <= 10) {
        stats.distribution.med++;
      } else {
        stats.distribution.high++;
        if (highActivitySamples.length < 3) {
          highActivitySamples.push(phone + " (" + rc.calls + " calls, " + rc.sms + " sms)");
        }
      }
    } else {
      stats.withNone++;
      stats.distribution.zero++;
      if (noActivitySamples.length < 3) {
        noActivitySamples.push(DV_maskPhone_(phone));
      }
    }
  }
  
  var matchRate = stats.total > 0 ? Math.round(((stats.total - stats.withNone) / stats.total) * 100) : 0;
  var status = matchRate >= 70 ? "pass" : (matchRate >= 40 ? "warning" : "error");
  
  var details = [
    "📊 Coverage: " + matchRate + "% have RC activity",
    "📞 With calls: " + stats.withCalls + " (" + Math.round(stats.withCalls / stats.total * 100) + "%)",
    "💬 With SMS: " + stats.withSms + " (" + Math.round(stats.withSms / stats.total * 100) + "%)",
    "📵 No activity: " + stats.withNone,
    "─────────────────",
    "Activity distribution:",
    "  0 contacts: " + stats.distribution.zero,
    "  1-3 contacts: " + stats.distribution.low,
    "  4-10 contacts: " + stats.distribution.med,
    "  10+ contacts: " + stats.distribution.high
  ];
  
  if (noActivitySamples.length > 0) {
    details.push("─────────────────");
    details.push("Sample phones with no RC data:");
    for (var s = 0; s < noActivitySamples.length; s++) {
      details.push("  " + noActivitySamples[s]);
    }
  }
  
  return {
    name: sheetName + " - RC Match",
    status: status,
    message: stats.total + " leads analyzed, " + matchRate + "% have RC activity",
    details: details
  };
}

function DV_analyzeAllLeadsCoverage_(ss, rcIndex) {
  // Get all unique lead phones across all trackers
  var allLeadPhones = {};
  var trackerNames = [
    "SMS TRACKER", "CALL & VOICEMAIL TRACKER", "CONTACTED LEADS", "SAME DAY TRANSFERS"
  ];
  
  var debugInfo = [];
  
  for (var i = 0; i < trackerNames.length; i++) {
    var sheet = ss.getSheetByName(trackerNames[i]);
    if (!sheet || sheet.getLastRow() < 2) continue;
    
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    
    // Use enhanced phone finder
    var phoneCol = DV_findPhoneColumn_(headers);
    
    if (phoneCol >= 0) {
      var sheetPhones = 0;
      for (var j = 0; j < data.length; j++) {
        var phone = DV_normalizePhone_(data[j][phoneCol]);
        if (phone) {
          allLeadPhones[phone] = true;
          sheetPhones++;
        }
      }
      debugInfo.push(trackerNames[i] + ": " + sheetPhones + " phones");
    } else {
      debugInfo.push(trackerNames[i] + ": phone col not found");
    }
  }
  
  var totalLeadPhones = Object.keys(allLeadPhones).length;
  var matchedPhones = 0;
  var totalCalls = 0;
  var totalSms = 0;
  
  for (var phone in allLeadPhones) {
    var rc = rcIndex.phones[phone];
    if (rc && (rc.calls > 0 || rc.sms > 0)) {
      matchedPhones++;
      totalCalls += rc.calls;
      totalSms += rc.sms;
    }
  }
  
  var overallRate = totalLeadPhones > 0 ? Math.round((matchedPhones / totalLeadPhones) * 100) : 0;
  
  var details = [
    "Unique lead phones: " + totalLeadPhones.toLocaleString(),
    "Phones with RC data: " + matchedPhones.toLocaleString() + " (" + overallRate + "%)",
    "Total calls to leads: " + totalCalls.toLocaleString(),
    "Total SMS to leads: " + totalSms.toLocaleString(),
    "─────────────────",
    "RC phones not in leads: " + (rcIndex.totalPhones - matchedPhones).toLocaleString()
  ];
  
  // Add debug info about which sheets contributed
  if (debugInfo.length > 0) {
    details.push("─────────────────");
    details.push("Source breakdown:");
    for (var d = 0; d < debugInfo.length; d++) {
      details.push("  " + debugInfo[d]);
    }
  }
  
  return {
    name: "Overall Coverage Summary",
    status: totalLeadPhones > 0 ? (overallRate >= 50 ? "pass" : "warning") : "error",
    message: totalLeadPhones.toLocaleString() + " unique lead phones, " + overallRate + "% have RC activity",
    details: details
  };
}


/* ═══════════════════════════════════════════════════════════════════════════
 * CATEGORY 4: ROSTERS
 * ═══════════════════════════════════════════════════════════════════════════ */

function DV_validateRosters_(ss) {
  var category = { name: "Roster Validation", icon: "👥", checks: [] };
  
  // Sales Roster
  var salesRoster = ss.getSheetByName("Sales_Roster") || ss.getSheetByName("Sales Roster");
  if (salesRoster) {
    var lastRow = salesRoster.getLastRow();
    var repCount = Math.max(0, lastRow - 1);
    
    // Check for missing emails
    var lastCol = salesRoster.getLastColumn();
    var headers = salesRoster.getRange(1, 1, 1, lastCol).getValues()[0];
    var data = lastRow > 1 ? salesRoster.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
    
    var emailCol = DV_findColumn_(headers, ["e-mail 1 - value", "email", "email address"]);
    var missingEmails = 0;
    
    if (emailCol >= 0) {
      for (var i = 0; i < data.length; i++) {
        var email = String(data[i][emailCol] || "").trim();
        if (!email || email.indexOf("@") < 0) missingEmails++;
      }
    }
    
    category.checks.push({
      name: "Sales Roster",
      status: missingEmails === 0 ? "pass" : (missingEmails < 5 ? "warning" : "error"),
      message: repCount + " reps found" + (missingEmails > 0 ? ", " + missingEmails + " missing email" : ""),
      details: [
        "Total reps: " + repCount,
        "Missing emails: " + missingEmails
      ]
    });
  } else {
    category.checks.push({
      name: "Sales Roster",
      status: "error",
      message: "Sales_Roster sheet not found",
      details: []
    });
  }
  
  // Team Roster
  var teamRoster = ss.getSheetByName("Team_Roster") || ss.getSheetByName("Team Roster");
  if (teamRoster) {
    var headers = teamRoster.getRange(1, 1, 1, teamRoster.getLastColumn()).getValues()[0];
    var managers = 0;
    for (var j = 0; j < headers.length; j++) {
      var h = String(headers[j] || "").trim();
      if (h && h.toLowerCase() !== "sales managers") managers++;
    }
    
    // Count total rep assignments
    var teamData = teamRoster.getLastRow() > 1 ? 
      teamRoster.getRange(2, 1, teamRoster.getLastRow() - 1, teamRoster.getLastColumn()).getValues() : [];
    var totalAssignments = 0;
    for (var k = 0; k < teamData.length; k++) {
      for (var m = 0; m < teamData[k].length; m++) {
        if (String(teamData[k][m] || "").trim()) totalAssignments++;
      }
    }
    
    category.checks.push({
      name: "Team Roster",
      status: "pass",
      message: managers + " managers, " + totalAssignments + " rep assignments",
      details: []
    });
  } else {
    category.checks.push({
      name: "Team Roster",
      status: "warning",
      message: "Team_Roster not found (optional)",
      details: []
    });
  }
  
  return category;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * CATEGORY 5: DATA QUALITY (NEW!)
 * ═══════════════════════════════════════════════════════════════════════════ */

function DV_validateDataQuality_(ss) {
  var category = { name: "Data Quality", icon: "✨", checks: [] };
  
  // Check for duplicate phones in CONTACTED LEADS
  var contactedLeads = ss.getSheetByName("CONTACTED LEADS");
  if (contactedLeads && contactedLeads.getLastRow() > 1) {
    var lastCol = contactedLeads.getLastColumn();
    var headers = contactedLeads.getRange(1, 1, 1, lastCol).getValues()[0];
    var data = contactedLeads.getRange(2, 1, contactedLeads.getLastRow() - 1, lastCol).getValues();
    
    var phoneCol = DV_findColumn_(headers, ["phone", "customer phone", "phone number"]);
    
    if (phoneCol >= 0) {
      var phoneCounts = {};
      var duplicates = 0;
      
      for (var i = 0; i < data.length; i++) {
        var phone = DV_normalizePhone_(data[i][phoneCol]);
        if (phone) {
          phoneCounts[phone] = (phoneCounts[phone] || 0) + 1;
          if (phoneCounts[phone] === 2) duplicates++;
        }
      }
      
      var topDupes = [];
      for (var phone in phoneCounts) {
        if (phoneCounts[phone] > 1) {
          topDupes.push({ phone: phone, count: phoneCounts[phone] });
        }
      }
      topDupes.sort(function(a, b) { return b.count - a.count; });
      
      var dupeDetails = ["Phones appearing multiple times:"];
      for (var d = 0; d < Math.min(5, topDupes.length); d++) {
        dupeDetails.push("  " + DV_maskPhone_(topDupes[d].phone) + " × " + topDupes[d].count);
      }
      
      category.checks.push({
        name: "Duplicate Phones (Contacted)",
        status: duplicates === 0 ? "pass" : (duplicates < 50 ? "warning" : "error"),
        message: duplicates + " duplicate phone numbers",
        details: duplicates > 0 ? dupeDetails : ["No duplicates found ✓"]
      });
    }
  }
  
  // Check notification log freshness
  var notifLog = ss.getSheetByName("Notification_Log");
  if (notifLog && notifLog.getLastRow() > 1) {
    var logHeaders = notifLog.getRange(1, 1, 1, notifLog.getLastColumn()).getValues()[0];
    var dateCol = DV_findColumn_(logHeaders, ["timestamp", "date", "sent", "time"]);
    
    if (dateCol >= 0) {
      var lastLogRow = notifLog.getRange(notifLog.getLastRow(), dateCol + 1).getValue();
      var lastLogDate = new Date(lastLogRow);
      var hoursSince = Math.round((new Date() - lastLogDate) / (1000 * 60 * 60));
      
      category.checks.push({
        name: "Last Notification",
        status: hoursSince < 24 ? "pass" : (hoursSince < 72 ? "warning" : "error"),
        message: hoursSince + " hours ago",
        details: [
          "Last notification: " + lastLogDate.toLocaleString(),
          "Total notifications: " + (notifLog.getLastRow() - 1).toLocaleString()
        ]
      });
    }
  } else {
    category.checks.push({
      name: "Notification Log",
      status: "warning",
      message: "No notification history found",
      details: []
    });
  }
  
  // RC Data Freshness - check if today's data exists
  var rcCallLog = ss.getSheetByName("RC CALL LOG");
  if (rcCallLog && rcCallLog.getLastRow() > 1) {
    var rcHeaders = rcCallLog.getRange(1, 1, 1, rcCallLog.getLastColumn()).getValues()[0];
    var rcDateCol = DV_findColumn_(rcHeaders, ["date", "call date", "time", "start time"]);
    
    if (rcDateCol >= 0) {
      // Check a sample of recent rows
      var sampleRow = Math.max(2, rcCallLog.getLastRow() - 10);
      var sampleDate = rcCallLog.getRange(sampleRow, rcDateCol + 1).getValue();
      var rcDate = new Date(sampleDate);
      var today = new Date();
      var isToday = rcDate.toDateString() === today.toDateString();
      
      category.checks.push({
        name: "RC Data Freshness",
        status: isToday ? "pass" : "warning",
        message: isToday ? "Today's data present" : "Data may be stale",
        details: [
          "Sample record date: " + rcDate.toLocaleDateString(),
          "Today: " + today.toLocaleDateString()
        ]
      });
    }
  }
  
  return category;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * CATEGORY 6: CONFIGURATION
 * ═══════════════════════════════════════════════════════════════════════════ */

function DV_validateConfiguration_(CFG) {
  var category = { name: "Configuration", icon: "⚙️", checks: [] };
  
  var adminEmail = (CFG && CFG.ADMIN_EMAIL) || "";
  category.checks.push({
    name: "Admin Email",
    status: adminEmail.indexOf("@") >= 0 ? "pass" : "warning",
    message: adminEmail || "Not configured",
    details: []
  });
  
  var safeMode = CFG && CFG.SAFE_MODE;
  var forwardAll = CFG && CFG.FORWARD_ALL;
  var modeLabel = safeMode ? "🛡️ Safe Mode (Dry Run)" : (forwardAll ? "↪️ Forward All" : "🚀 Live Mode");
  
  category.checks.push({
    name: "Current Mode",
    status: "pass",
    message: modeLabel,
    details: [
      "Safe Mode: " + (safeMode ? "ON" : "OFF"),
      "Forward All: " + (forwardAll ? "ON" : "OFF")
    ]
  });
  
  return category;
}


/* ═══════════════════════════════════════════════════════════════════════════
 * UTILITIES
 * ═══════════════════════════════════════════════════════════════════════════ */

function DV_findColumn_(headers, names) {
  if (!headers || !names) return -1;
  
  // First pass: exact match
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || "").trim().toLowerCase();
    for (var j = 0; j < names.length; j++) {
      if (h === names[j].toLowerCase()) return i;
    }
  }
  
  // Second pass: contains match
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || "").trim().toLowerCase();
    for (var j = 0; j < names.length; j++) {
      if (h.indexOf(names[j].toLowerCase()) >= 0) return i;
    }
  }
  
  return -1;
}

/**
 * Special phone column finder - tries harder to find phone columns
 */
function DV_findPhoneColumn_(headers) {
  if (!headers) return -1;
  
  // Specific known names first
  var knownNames = [
    "phone", "customer phone", "phone number", "phone #", "phone#",
    "cell", "mobile", "contact phone", "primary phone", "telephone", 
    "tel", "phone1", "cust phone", "client phone", "lead phone"
  ];
  
  // Try exact match first
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || "").trim().toLowerCase();
    for (var j = 0; j < knownNames.length; j++) {
      if (h === knownNames[j]) return i;
    }
  }
  
  // Try contains "phone" anywhere
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || "").trim().toLowerCase();
    if (h.indexOf("phone") >= 0) return i;
  }
  
  // Try contains "cell" or "mobile"
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || "").trim().toLowerCase();
    if (h.indexOf("cell") >= 0 || h.indexOf("mobile") >= 0) return i;
  }
  
  // Last resort: look for a column with mostly 10-digit numbers
  // (Skip this for now as it's expensive)
  
  return -1;
}

function DV_normalizePhone_(val) {
  if (!val) return null;
  var digits = String(val).replace(/\D/g, "");
  if (digits.length === 11 && digits[0] === "1") {
    digits = digits.substring(1);
  }
  return digits.length === 10 ? digits : null;
}

function DV_maskPhone_(phone) {
  if (!phone || phone.length < 10) return phone;
  return "***-***-" + phone.slice(-4);
}


/* ═══════════════════════════════════════════════════════════════════════════
 * SIDEBAR HTML
 * ═══════════════════════════════════════════════════════════════════════════ */

function DV_getValidatorSidebarHtml_() {
  return '<!DOCTYPE html>\
<html><head><base target="_top">\
<style>\
*{box-sizing:border-box;margin:0;padding:0}\
body{font-family:Inter,-apple-system,sans-serif;background:#0f172a;color:#e2e8f0;padding:16px;min-height:100vh}\
.header{background:linear-gradient(135deg,#8B1538,#6B1028);padding:20px;border-radius:12px;text-align:center;margin-bottom:16px}\
.header h1{color:#D4AF37;font-size:18px;margin-bottom:4px}\
.header p{color:#fca5a5;font-size:12px}\
.progress-area{background:#1e293b;border-radius:12px;padding:20px;margin-bottom:16px;text-align:center;display:none}\
.spinner{width:48px;height:48px;border:4px solid #334155;border-top-color:#D4AF37;border-radius:50%;margin:0 auto 12px;animation:spin 1s linear infinite}\
@keyframes spin{to{transform:rotate(360deg)}}\
.progress-text{color:#94a3b8;font-size:13px}\
.summary{display:none;background:#1e293b;border-radius:12px;padding:16px;margin-bottom:16px}\
.summary-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;text-align:center}\
.summary-item{padding:12px 8px;border-radius:8px;background:#0f172a}\
.summary-item.passed .summary-val{color:#22c55e}\
.summary-item.warnings .summary-val{color:#eab308}\
.summary-item.errors .summary-val{color:#ef4444}\
.summary-val{font-size:24px;font-weight:700}\
.summary-lbl{font-size:10px;color:#94a3b8;margin-top:4px}\
.category{background:#1e293b;border-radius:12px;margin-bottom:12px;overflow:hidden}\
.cat-header{display:flex;align-items:center;padding:12px 16px;background:#334155;cursor:pointer;transition:background 0.2s}\
.cat-header:hover{background:#475569}\
.cat-icon{font-size:18px;margin-right:10px}\
.cat-name{flex:1;font-weight:600;font-size:14px}\
.cat-badge{padding:4px 10px;border-radius:12px;font-size:11px;font-weight:600}\
.cat-badge.pass{background:#166534;color:#86efac}\
.cat-badge.warning{background:#854d0e;color:#fde047}\
.cat-badge.error{background:#991b1b;color:#fca5a5}\
.cat-body{padding:12px;display:none}\
.check{padding:10px;border-radius:8px;background:#0f172a;margin-bottom:8px}\
.check-header{display:flex;align-items:center}\
.check-status{width:20px;height:20px;border-radius:50%;margin-right:10px;display:flex;align-items:center;justify-content:center;font-size:12px;flex-shrink:0}\
.check-status.pass{background:#166534;color:#86efac}\
.check-status.warning{background:#854d0e;color:#fde047}\
.check-status.error{background:#991b1b;color:#fca5a5}\
.check-name{font-size:13px;font-weight:500}\
.check-msg{font-size:12px;color:#94a3b8;margin-top:4px;margin-left:30px}\
.check-details{font-size:11px;color:#64748b;margin-top:6px;margin-left:30px;line-height:1.6}\
.check-details div{padding:1px 0}\
.btn{width:100%;padding:14px;border:none;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer;margin-top:8px;transition:transform 0.1s,opacity 0.2s}\
.btn:hover{opacity:0.9}\
.btn:active{transform:scale(0.98)}\
.btn-start{background:linear-gradient(135deg,#8B1538,#6B1028);color:#fff}\
.btn-close{background:#334155;color:#fff;display:none}\
.duration{text-align:center;color:#64748b;font-size:12px;margin-top:12px}\
.version{text-align:center;color:#475569;font-size:10px;margin-top:8px}\
</style></head><body>\
<div class="header"><h1>🔍 Data Validator</h1><p>Comprehensive data accuracy check</p></div>\
<div class="progress-area" id="progressArea"><div class="spinner"></div><div class="progress-text">Scanning data & building indexes...</div></div>\
<div class="summary" id="summary"><div class="summary-grid">\
<div class="summary-item"><div class="summary-val" id="sumTotal">0</div><div class="summary-lbl">TOTAL</div></div>\
<div class="summary-item passed"><div class="summary-val" id="sumPassed">0</div><div class="summary-lbl">PASSED</div></div>\
<div class="summary-item warnings"><div class="summary-val" id="sumWarnings">0</div><div class="summary-lbl">WARNINGS</div></div>\
<div class="summary-item errors"><div class="summary-val" id="sumErrors">0</div><div class="summary-lbl">ERRORS</div></div>\
</div></div>\
<div id="categories"></div>\
<div class="duration" id="duration"></div>\
<div class="version" id="version"></div>\
<button class="btn btn-start" id="startBtn">🔍 Start Validation</button>\
<button class="btn btn-close" id="closeBtn">Close</button>\
<script>\
document.getElementById("startBtn").onclick=function(){this.style.display="none";document.getElementById("progressArea").style.display="block";google.script.run.withSuccessHandler(showResults).withFailureHandler(showError).DV_executeValidation();};\
document.getElementById("closeBtn").onclick=function(){google.script.host.close();};\
function showResults(r){\
document.getElementById("progressArea").style.display="none";\
document.getElementById("summary").style.display="block";\
document.getElementById("sumTotal").textContent=r.summary.total;\
document.getElementById("sumPassed").textContent=r.summary.passed;\
document.getElementById("sumWarnings").textContent=r.summary.warnings;\
document.getElementById("sumErrors").textContent=r.summary.errors;\
document.getElementById("duration").textContent="Completed in "+r.duration+"s";\
document.getElementById("version").textContent="v"+r.version;\
var html="";\
for(var ci=0;ci<r.categories.length;ci++){\
var cat=r.categories[ci],st="pass",errs=0,warns=0;\
if(cat.checks){for(var x=0;x<cat.checks.length;x++){if(cat.checks[x].status==="error")errs++;else if(cat.checks[x].status==="warning")warns++;}}\
if(errs>0)st="error";else if(warns>0)st="warning";\
html+="<div class=\\"category\\"><div class=\\"cat-header\\" onclick=\\"toggleCat("+ci+")\\">";\
html+="<span class=\\"cat-icon\\">"+(cat.icon||"📋")+"</span><span class=\\"cat-name\\">"+cat.name+"</span>";\
html+="<span class=\\"cat-badge "+st+"\\">"+(cat.checks?cat.checks.length:0)+" checks</span></div>";\
html+="<div class=\\"cat-body\\" id=\\"cat"+ci+"\\">";\
if(cat.checks){for(var j=0;j<cat.checks.length;j++){var c=cat.checks[j];\
html+="<div class=\\"check\\"><div class=\\"check-header\\">";\
html+="<div class=\\"check-status "+c.status+"\\">"+(c.status==="pass"?"✓":c.status==="warning"?"!":"✗")+"</div>";\
html+="<div class=\\"check-name\\">"+c.name+"</div></div>";\
html+="<div class=\\"check-msg\\">"+c.message+"</div>";\
if(c.details&&c.details.length){html+="<div class=\\"check-details\\">";for(var d=0;d<c.details.length;d++)html+="<div>"+c.details[d]+"</div>";html+="</div>";}\
html+="</div>";}}\
html+="</div></div>";}\
document.getElementById("categories").innerHTML=html;\
document.getElementById("closeBtn").style.display="block";\
for(var i=0;i<r.categories.length;i++){var ct=r.categories[i],hi=false;if(ct.checks){for(var k=0;k<ct.checks.length;k++){if(ct.checks[k].status!=="pass"){hi=true;break;}}}if(hi)toggleCat(i);}}\
function toggleCat(i){var e=document.getElementById("cat"+i);e.style.display=e.style.display==="block"?"none":"block";}\
function showError(e){document.getElementById("progressArea").innerHTML="<div style=\\"color:#ef4444\\">❌ Error: "+e.message+"</div>";document.getElementById("closeBtn").style.display="block";}\
</script></body></html>';
}
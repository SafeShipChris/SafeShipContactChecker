/**************************************************************
 * [SHIP] SAFE SHIP CONTACT CHECKER -- Code.gs v4.0 (MASTER CLEAN)
 * 
 * 
 * AUTHOR: Safe Ship Moving Services
 * VERSION: 4.0.0 (Consolidated & Optimized)
 * LAST UPDATED: 2025
 * 
 * 
 * [CLIPBOARD] TABLE OF CONTENTS
 * 
 * 
 * SECTION 1:  Configuration & Constants
 * SECTION 2:  Public Entry Points
 * SECTION 3:  Pipeline Runner (Main Engine)
 * SECTION 4:  Gate Checks & Validation
 * SECTION 5:  Source Scanners
 * SECTION 6:  Data Processing (Grouping, Integrity, Aggregation)
 * SECTION 7:  Sending Logic (Rep Alerts, Manager Reports)
 * SECTION 8:  Admin Dashboard
 * SECTION 9:  Premium Email Renderers
 * SECTION 10: Contact Activity Renderer (v4.0)
 * SECTION 11: Slack Integration
 * SECTION 12: Roster Management
 * SECTION 13: Deduplication Logic
 * SECTION 14: Logging (Notification & Run Logs)
 * SECTION 15: Cancellation & Test Mode
 * SECTION 16: RC Data Freshness
 * SECTION 17: Utility Functions
 * SECTION 18: Menu Utilities
 * 
 * 
 * [NEW] v4.0 FEATURES
 * 
 * [OK] Clean consolidated codebase (no duplicates)
 * [OK] Test Mode: First 5 leads to admin only
 * [OK] Cancellation support with cache flag
 * [OK] RC Data Freshness timestamps in emails
 * [OK] "Yesterday - Today" stat bubble labels
 * [OK] Failed SMS tracking with counter
 * [OK] Process Direction with contextual tips
 * [OK] Alert banners (replied, long call, all SMS failed)
 * [OK] Optimized performance with batch operations
 * [OK] Comprehensive error handling
 * 
 * NOTE: Menu is defined in MainMenu.gs - do not add onOpen here
 **************************************************************/

"use strict";

/* 
 * SECTION 1: CONFIGURATION & CONSTANTS
 * =============================================================== */

/**
 * Master configuration object
 * Centralizes all settings, sheet names, column indices, and feature flags
 * @returns {Object} Configuration object
 */
function getConfig_() {
  var props = PropertiesService.getScriptProperties();
  
  return {
    // 
    // CORE SETTINGS
    // 
    TZ: "America/New_York",
    ADMIN_EMAIL: sanitizeSingleEmail_(props.getProperty("ADMIN_EMAIL")),
    ADMIN_REPORT_EMAILS: sanitizeEmailListRaw_(props.getProperty("ADMIN_REPORT_EMAILS")),
    
    // 
    // MODE FLAGS
    // 
    SAFE_MODE: props.getProperty("UNCONTACTED_DRY_RUN") === "true",
    FORWARD_ALL: props.getProperty("UNCONTACTED_REDIRECT_ALL") === "true",
    
    // 
    // INTEGRATIONS
    // 
    SLACK: {
      BOT_TOKEN: String(props.getProperty("SLACK_BOT_TOKEN") || "").trim()
    },
    
    // 
    // TIMING & LIMITS
    // 
    WINDOW_DAYS: toInt_(props.getProperty("WINDOW_DAYS"), 5),
    DEDUPE_HOURS: toInt_(props.getProperty("DEDUPE_HOURS"), 0),
    CACHE_TTL_SECONDS: toInt_(props.getProperty("CACHE_TTL_SECONDS"), 21600),
    TEST_MODE_LEAD_LIMIT: 5,
    
    // 
    // SHEET NAMES
    // 
    SHEETS: {
      SALES_ROSTER: "Sales_Roster",
      TEAM_ROSTER: "Team_Roster",
      NOTIFICATION_LOG: "Notification_Log",
      RUN_LOG: "Run_Log",
      SMS: "SMS TRACKER",
      CALL: "CALL & VOICEMAIL TRACKER",
      P1CALL: "PRIORITY 1 CALL & VOICEMAIL TRACKER",
      CONTACTED: "CONTACTED LEADS",
      TRANSFERS: "SAME DAY TRANSFERS",
      RC_CALL_LOG: "RC CALL LOG",
      RC_SMS_LOG: "RC SMS LOG"
    },
    
    // 
    // PRIORITY CELLS
    // 
    PRIORITY: {
      SMS_CELL: "B7",
      CALL_CELL: "A8",
      P1CALL_CELL: "A8"
    },
    
    // 
    // DATA RANGES
    // 
    TRACKER_RANGE_A1: "A3:AE",
    
    // 
    // COLUMN INDICES (0-based)
    // 
    TRACKER_COLS: {
      MOVE_DATE_IDX: 5,
      PHONE_IDX: 6,
      JOB_IDX: 8,
      USERNAME_IDX: 9,
      CUBIC_FEET_IDX: 10,
      SOURCE_IDX: 11,
      EXCLUDE_IDX: 12
    },
    
    CONTACTED_COLS: {
      MOVE_DATE_IDX: 3,
      PHONE_IDX: 5,
      JOB_IDX: 7,
      USERNAME_IDX: 8,
      CUBIC_FEET_IDX: 10,
      SOURCE_IDX: 11,
      EXCLUDE_IDX: 12
    },
    
    TRANSFERS: {
      START_ROW: 4,
      USERNAME_COL: 3,
      MOVE_DATE_COL: 0
    },
    
    // 
    // BRAND COLORS
    // 
    BRAND: {
      BLUE: "#1E3A5F",
      DARK: "#0F172A",
      MUTED: "#64748B",
      BORDER: "#E5E7EB",
      LIGHT_BG: "#F8FAFC",
      GREEN: "#10B981",
      YELLOW: "#F59E0B",
      RED: "#EF4444"
    }
  };
}


/* 
 * SECTION 2: PUBLIC ENTRY POINTS
 * =============================================================== */

/**
 * Run Uncontacted Leads Alert (SMS + Call/VM Trackers)
 */
function SSCCP_runUncontactedLeadAlerts() {
  SSCCP_runPipeline_({
    reportType: "UNCONTACTED",
    title: "Uncontacted Leads Dashboard",
    sources: ["SMS", "CALL"]
  });
}

/**
 * Run Quoted Follow-Up Alert (Contacted Leads)
 */
function SSCCP_runQuotedFollowUpAlerts() {
  SSCCP_runPipeline_({
    reportType: "QUOTED_FOLLOWUP",
    title: "Quoted Follow-Up Dashboard",
    sources: ["CONTACTED"]
  });
}

/**
 * Run Priority 1 Call/VM Alert
 */
function SSCCP_runPriority1CallVmAlerts() {
  SSCCP_runPipeline_({
    reportType: "PRIORITY1_CALLVM",
    title: "Priority 1 Call/VM Dashboard",
    sources: ["P1CALL"]
  });
}

/**
 * Run Same Day Transfers Alert
 */
function SSCCP_runSameDayTransfersAlerts() {
  SSCCP_runPipeline_({
    reportType: "SAME_DAY_TRANSFERS",
    title: "Same Day Transfers Dashboard",
    sources: ["TRANSFERS"]
  });
}

/**
 * View contact history for selected phone number (menu action)
 */
/**
 * View contact history for selected phone - WRAPPER
 * Canonical implementation is in RC_Enrichment.js
 * This wrapper is kept for backward compatibility with any old menu references
 */
function RC_viewSelectedContactHistory_Code_() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getActiveCell();
  var phone = String(cell.getValue() || "").trim();
  
  if (!phone) {
    SpreadsheetApp.getUi().alert("Please select a cell with a phone number first.");
    return;
  }
  
  var history = typeof RC_getContactHistory_ === "function" ? RC_getContactHistory_(phone) : null;
  
  if (!history) {
    SpreadsheetApp.getUi().alert("No RingCentral history found for: " + phone);
    return;
  }
  
  var lines = [
    "📱 CONTACT HISTORY FOR: " + phone,
    "=".repeat(40),
    ""
  ];
  
  if (history.lastCall) {
    lines.push("[PHONE] LAST CALL:");
    lines.push("   Time: " + history.lastCall.timeAgo + " (" + history.lastCall.dateTime + ")");
    lines.push("   Result: " + history.lastCall.result);
    lines.push("   Duration: " + history.lastCall.duration);
    lines.push("");
  } else {
    lines.push("[PHONE] No call history found", "");
  }
  
  if (history.lastSMS) {
    lines.push("[MSG] LAST SMS SENT:");
    lines.push("   Time: " + history.lastSMS.timeAgo + " (" + history.lastSMS.dateTime + ")");
    lines.push("   Status: " + history.lastSMS.status);
    lines.push("");
  } else {
    lines.push("[MSG] No outbound SMS found", "");
  }
  
  if (history.lastReply) {
    lines.push("📥 LAST REPLY RECEIVED:");
    lines.push("   Time: " + history.lastReply.timeAgo + " (" + history.lastReply.dateTime + ")");
    lines.push("");
  } else {
    lines.push("📥 No inbound reply found");
  }
  
  if (history.counts) {
    lines.push("[CHART] COUNTS (Yesterday -> Today):");
    lines.push("   SMS Sent: " + history.counts.smsYesterday + " -> " + history.counts.smsToday);
    lines.push("   Calls: " + history.counts.callsYesterday + " -> " + history.counts.callsToday);
    lines.push("   Replies: " + history.counts.repliesYesterday + " -> " + history.counts.repliesToday);
    lines.push("   Longest Call: " + history.counts.longestCall);
  }
  
  SpreadsheetApp.getUi().alert(lines.join("\n"));
}


/* 
 * SECTION 3: PIPELINE RUNNER (MAIN ENGINE)
 * =============================================================== */

/**
 * Main pipeline runner - orchestrates the entire alert process
 * @param {Object} ctx - Context object with reportType, title, sources
 */
function SSCCP_runPipeline_(ctx) {
  var CFG = getConfig_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var toast = function(msg, title, duration) {
    ss.toast(msg, title || "Safe Ship Contact Checker", duration || 5);
  };
  
  // 
  // INITIALIZATION
  // 
  SSCCP_clearCancellation_();
  
  var testMode = SSCCP_isTestMode_();
  if (testMode) {
    toast("🧪 TEST MODE: First " + CFG.TEST_MODE_LEAD_LIMIT + " leads to Admin only", "Test Mode Active", 6);
  }
  
  var run = SSCCP_startRun_(ctx.reportType);
  var now = new Date();
  var windowInfo = SSCCP_computeWindow_(CFG, now);
  windowInfo.label = SSCCP_getTrackerWindowLabel_(CFG, ctx);
  
  toast("Starting " + ctx.reportType + " -- Run " + run.runId, "Initializing", 6);
  
  // 
  // ENSURE LOG SHEETS EXIST
  // 
  SSCCP_ensureNotificationLog_(CFG);
  SSCCP_ensureRunLog_(CFG);
  
  // 
  // STEP 1: GATE CHECKS
  // 
  var gate = SSCCP_runGateChecks_(CFG, ctx, toast);
  if (!gate.ok) {
    SSCCP_logRunEvent_(CFG, run, "GATE_FAIL", gate.summary, gate.details);
    SSCCP_sendAdminDashboard_(CFG, run, ctx, windowInfo, "RED", {
      metrics: { totalLeadsScanned: 0, repsWithLeads: 0, managersSummarized: 0 },
      notes: [],
      exceptions: [{ title: "Gate Check Failure", items: [gate.summary].concat(gate.details || []) }],
      teamAgg: {},
      performance: SSCCP_buildAdminPerformancePack_(CFG, { leads: [] }, {}, {}, {}, { duplicates: [], unassigned: [] }, { exceptions: [] }, { exceptions: [] })
    });
    toast("Blocked: " + gate.summary, "Gate Checks Failed", 12);
    return;
  }
  
  // 
  // STEP 2: LOAD ROSTERS
  // 
  toast("Loading Sales_Roster...", "Step 1 of 7", 6);
  var roster = SSCCP_buildSalesRosterIndex_(CFG);
  
  toast("Loading Team_Roster...", "Step 2 of 7", 6);
  var team = SSCCP_buildTeamRosterIndex_(CFG);
  
  // 
  // STEP 3: LOAD RINGCENTRAL INDEX
  // 
  toast("Loading RingCentral logs...", "Step 3 of 7", 6);
  var rcIndex = {};
  try {
    if (typeof RC_buildLookupIndex_ === "function") {
      rcIndex = RC_buildLookupIndex_();
    }
  } catch (e) {
    console.warn("RC index load failed (continuing without RC data):", e);
  }
  
  // 
  // STEP 4: SCAN SOURCES
  // 
  toast("Scanning sources...", "Step 4 of 7", 7);
  var scan = SSCCP_scanSources_(CFG, ctx, windowInfo, roster, toast);
  
  if (!scan.ok) {
    SSCCP_logRunEvent_(CFG, run, "SCAN_FAIL", scan.summary, scan.details);
    SSCCP_sendAdminDashboard_(CFG, run, ctx, windowInfo, "RED", {
      metrics: { totalLeadsScanned: 0, repsWithLeads: 0, managersSummarized: 0 },
      notes: [],
      exceptions: [{ title: "Scan Failure", items: [scan.summary].concat(scan.details || []) }],
      teamAgg: {},
      performance: SSCCP_buildAdminPerformancePack_(CFG, { leads: [] }, {}, roster, team, { duplicates: [], unassigned: [] }, { exceptions: [] }, { exceptions: [] })
    });
    toast("Scan failed: " + scan.summary, "Blocked", 12);
    return;
  }
  
  // 
  // STEP 5: ENRICH WITH RINGCENTRAL DATA
  // 
  try {
    if (typeof RC_enrichLeads_ === "function") {
      RC_enrichLeads_(scan.leads, rcIndex);
    }
  } catch (e) {
    console.warn("RC enrichment failed (continuing):", e);
  }
  
  // 
  // STEP 6: GROUP BY REP & COMPUTE INTEGRITY
  // 
  toast("Grouping by rep...", "Step 5 of 7", 6);
  var grouped = SSCCP_groupByRep_(scan.leads);
  var integrity = SSCCP_computeIntegrity_(team, grouped);
  var mode = SSCCP_mode_(CFG);
  var stoplightBase = integrity.duplicates.length ? "RED" : (integrity.unassigned.length ? "YELLOW" : "GREEN");
  
  // 
  // STEP 7: SEND REP ALERTS
  // 
  toast("Sending Rep Alerts...", "Step 6 of 7", 7);
  var repSend = SSCCP_sendRepAlerts_(CFG, run, ctx, windowInfo, grouped, roster, integrity, toast);
  
  // 
  // STEP 8: SEND MANAGER REPORTS (skip if cancelled or test mode)
  // 
  var mgrSend = { managersSent: 0, exceptions: [] };
  if (!repSend.wasCancelled && !repSend.wasTestMode) {
    toast("Sending Manager Reports...", "Step 7 of 7", 7);
    mgrSend = SSCCP_sendManagerReports_(CFG, run, ctx, windowInfo, grouped, roster, team, integrity, toast);
  } else if (repSend.wasTestMode) {
    toast("🧪 Test Mode: Skipping manager reports", "Test Mode", 4);
  }
  
  // 
  // STEP 9: BUILD FINAL METRICS & SEND ADMIN DASHBOARD
  // 
  var finalStoplight = SSCCP_mergeStoplight_(stoplightBase, repSend, mgrSend);
  var teamAgg = SSCCP_buildTeamAgg_(grouped, roster, team, integrity);
  var notes = [].concat(scan.notes || []).concat(["Mode: " + mode.effectiveModeLabel]);
  if (testMode) notes.push("🧪 TEST MODE ACTIVE");
  
  var metrics = {
    totalLeadsScanned: scan.leads.length,
    repsWithLeads: Object.keys(grouped).length,
    managersSummarized: mgrSend.managersSent || 0
  };
  
  // Skip admin dashboard in test mode
  if (!testMode) {
    SSCCP_sendAdminDashboard_(CFG, run, ctx, windowInfo, finalStoplight, {
      metrics: metrics,
      notes: notes,
      exceptions: SSCCP_buildExceptionBlocks_(integrity, scan, repSend, mgrSend),
      teamAgg: teamAgg,
      performance: SSCCP_buildAdminPerformancePack_(CFG, scan, grouped, roster, team, integrity, repSend, mgrSend)
    });
  }
  
  // 
  // STEP 10: LOG COMPLETION
  // 
  SSCCP_logRunEvent_(CFG, run, testMode ? "TEST_SUCCESS" : "SUCCESS", ctx.reportType + " completed", [
    "Stoplight=" + finalStoplight,
    "Leads=" + metrics.totalLeadsScanned,
    "Reps=" + metrics.repsWithLeads,
    "Managers=" + metrics.managersSummarized,
    testMode ? "TEST_MODE" : ""
  ].filter(Boolean));
  
  var completeMsg = testMode
    ? "🧪 TEST: " + repSend.repsAlerted + " leads sent to admin -- Run " + run.runId
    : "Done: " + ctx.reportType + " -- " + finalStoplight + " -- Run " + run.runId;
  toast(completeMsg, "Completed", 12);
}


/* 
 * SECTION 4: GATE CHECKS & VALIDATION
 * =============================================================== */

/**
 * Run pre-flight gate checks to ensure all requirements are met
 * @param {Object} CFG - Configuration object
 * @param {Object} ctx - Context with sources array
 * @param {Function} toast - Toast notification function
 * @returns {Object} { ok: boolean, summary: string, details: array }
 */
function SSCCP_runGateChecks_(CFG, ctx, toast) {
  var details = [];
  var fail = function(msg) {
    return { ok: false, summary: msg, details: details };
  };
  
  // Check required script properties
  if (!CFG.ADMIN_EMAIL) return fail("Missing Script Property: ADMIN_EMAIL");
  if (!CFG.ADMIN_REPORT_EMAILS || !CFG.ADMIN_REPORT_EMAILS.length) {
    return fail("Missing Script Property: ADMIN_REPORT_EMAILS");
  }
  if (!CFG.SLACK.BOT_TOKEN) return fail("Missing Script Property: SLACK_BOT_TOKEN");
  
  // Check required sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mustSheets = [
    CFG.SHEETS.SALES_ROSTER,
    CFG.SHEETS.TEAM_ROSTER,
    CFG.SHEETS.NOTIFICATION_LOG,
    CFG.SHEETS.RUN_LOG
  ];
  
  // Add source-specific sheets
  (ctx.sources || []).forEach(function(k) {
    if (CFG.SHEETS[k]) mustSheets.push(CFG.SHEETS[k]);
  });
  
  // Validate all sheets exist
  for (var i = 0; i < mustSheets.length; i++) {
    if (!ss.getSheetByName(mustSheets[i])) {
      details.push("Missing sheet: " + mustSheets[i]);
    }
  }
  
  if (details.length) return fail("One or more required sheets are missing.");
  
  return { ok: true, summary: "OK", details: [] };
}

/**
 * Interactive gate checks (menu action)
 */
function SSCCP_runGateChecksInteractive() {
  var CFG = getConfig_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var toast = function(m, t, s) { ss.toast(m, t || "Gate Checks", s || 6); };
  
  var res = SSCCP_runGateChecks_(CFG, {
    sources: ["SMS", "CALL", "CONTACTED", "P1CALL", "TRANSFERS"]
  }, toast);
  
  SpreadsheetApp.getUi().alert(
    res.ok ? "[OK] Gate Checks OK" : "[X] Gate Checks Failed",
    res.ok ? "All required properties and sheets are present." : (res.summary + "\n\n" + (res.details || []).join("\n")),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Test column reading diagnostic
 */
function SSCCP_testColumnReading() {
  var CFG = getConfig_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var results = [
    "[WRENCH] COLUMN READING DIAGNOSTIC",
    "=".repeat(50),
    ""
  ];
  
  // Test SMS TRACKER
  try {
    var sh = ss.getSheetByName(CFG.SHEETS.SMS);
    if (sh) {
      var values = sh.getRange("A3:M7").getValues();
      results.push("[CLIPBOARD] SMS TRACKER - First 5 rows:");
      results.push("   Config: CF_IDX=" + CFG.TRACKER_COLS.CUBIC_FEET_IDX + ", SOURCE_IDX=" + CFG.TRACKER_COLS.SOURCE_IDX);
      
      for (var i = 0; i < Math.min(5, values.length); i++) {
        var row = values[i];
        results.push("   Row " + (i+3) + ": USER=" + (row[CFG.TRACKER_COLS.USERNAME_IDX] || "(empty)") +
                    " | CF=" + (row[CFG.TRACKER_COLS.CUBIC_FEET_IDX] || "(empty)") +
                    " | SOURCE=" + (row[CFG.TRACKER_COLS.SOURCE_IDX] || "(empty)"));
      }
    } else {
      results.push("[X] SMS TRACKER sheet not found!");
    }
  } catch (e) {
    results.push("[X] Error reading SMS TRACKER: " + e.message);
  }
  
  results.push("");
  
  // Test CALL & VOICEMAIL TRACKER
  try {
    var sh2 = ss.getSheetByName(CFG.SHEETS.CALL);
    if (sh2) {
      var values2 = sh2.getRange("A3:M7").getValues();
      results.push("[CLIPBOARD] CALL & VOICEMAIL TRACKER - First 5 rows:");
      results.push("   Config: CF_IDX=" + CFG.TRACKER_COLS.CUBIC_FEET_IDX + ", SOURCE_IDX=" + CFG.TRACKER_COLS.SOURCE_IDX);
      
      for (var j = 0; j < Math.min(5, values2.length); j++) {
        var row2 = values2[j];
        results.push("   Row " + (j+3) + ": USER=" + (row2[CFG.TRACKER_COLS.USERNAME_IDX] || "(empty)") +
                    " | CF=" + (row2[CFG.TRACKER_COLS.CUBIC_FEET_IDX] || "(empty)") +
                    " | SOURCE=" + (row2[CFG.TRACKER_COLS.SOURCE_IDX] || "(empty)"));
      }
    } else {
      results.push("[X] CALL & VOICEMAIL TRACKER sheet not found!");
    }
  } catch (e) {
    results.push("[X] Error reading CALL TRACKER: " + e.message);
  }
  
  ui.alert(results.join("\n"));
}


/* 
 * SECTION 5: SOURCE SCANNERS
 * =============================================================== */

/**
 * Scan all configured sources based on report type
 * @param {Object} CFG - Configuration
 * @param {Object} ctx - Context with reportType
 * @param {Object} windowInfo - Window information
 * @param {Object} roster - Roster index
 * @param {Function} toast - Toast function
 * @returns {Object} { ok: boolean, leads: array, notes: array }
 */
function SSCCP_scanSources_(CFG, ctx, windowInfo, roster, toast) {
  try {
    var leads = [];
    var notes = [];
    
    switch (ctx.reportType) {
      case "UNCONTACTED":
        // Verify Priority 0
        var smsPri = SSCCP_readCell_(CFG.SHEETS.SMS, CFG.PRIORITY.SMS_CELL);
        var callPri = SSCCP_readCell_(CFG.SHEETS.CALL, CFG.PRIORITY.CALL_CELL);
        
        if (!SSCCP_isPriority0_(smsPri) || !SSCCP_isPriority0_(callPri)) {
          return {
            ok: false,
            summary: "Priority cells do not indicate Priority 0 (required for Uncontacted run).",
            details: [
              "Expected Priority 0",
              CFG.SHEETS.SMS + "!" + CFG.PRIORITY.SMS_CELL + "=" + smsPri,
              CFG.SHEETS.CALL + "!" + CFG.PRIORITY.CALL_CELL + "=" + callPri
            ]
          };
        }
        
        leads = leads.concat(SSCCP_scanTracker_(CFG, CFG.SHEETS.SMS, "SMS"));
        leads = leads.concat(SSCCP_scanTracker_(CFG, CFG.SHEETS.CALL, "CALL/VM"));
        notes.push("Filtering: NONE (Tracker visibility is the source of truth)");
        break;
        
      case "QUOTED_FOLLOWUP":
        leads = leads.concat(SSCCP_scanContacted_(CFG));
        notes.push("Filtering: NONE (Tracker visibility is the source of truth)");
        break;
        
      case "PRIORITY1_CALLVM":
        var p1Pri = SSCCP_readCell_(CFG.SHEETS.P1CALL, CFG.PRIORITY.P1CALL_CELL);
        
        if (!SSCCP_isPriority1_(p1Pri)) {
          return {
            ok: false,
            summary: "Priority 1 cell does not indicate Priority 1.",
            details: [
              "Expected Priority 1",
              CFG.SHEETS.P1CALL + "!" + CFG.PRIORITY.P1CALL_CELL + "=" + p1Pri
            ]
          };
        }
        
        leads = leads.concat(SSCCP_scanTracker_(CFG, CFG.SHEETS.P1CALL, "Priority 1 Call/VM"));
        notes.push("Filtering: NONE (Tracker visibility is the source of truth)");
        break;
        
      case "SAME_DAY_TRANSFERS":
        leads = leads.concat(SSCCP_scanTransfers_(CFG));
        notes.push("Filtering: NONE (Tracker visibility is the source of truth)");
        break;
    }
    
    return { ok: true, summary: "OK", leads: leads, notes: notes };
    
  } catch (e) {
    return {
      ok: false,
      summary: "Exception during scan",
      details: [String(e && e.stack ? e.stack : e)]
    };
  }
}

/**
 * Scan a tracker sheet (SMS, CALL, P1CALL)
 * @param {Object} CFG - Configuration
 * @param {string} sheetName - Sheet name
 * @param {string} sourceLabel - Label for source type
 * @returns {Array} Array of lead objects
 */
function SSCCP_scanTracker_(CFG, sheetName, sourceLabel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  
  var values = sh.getRange(CFG.TRACKER_RANGE_A1).getValues();
  var cols = CFG.TRACKER_COLS;
  var out = [];
  
  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    
    var username = normalizeUsername_(row[cols.USERNAME_IDX]);
    var job = String(row[cols.JOB_IDX] || "").trim();
    var phone = SSCCP_normPhone_(row[cols.PHONE_IDX]);
    var exclude = row[cols.EXCLUDE_IDX];
    
    // Skip invalid rows
    if (!username || !job || !phone) continue;
    if (exclude === true || String(exclude).toUpperCase() === "TRUE") continue;
    
    var moveDate = SSCCP_fmtMoveDateDisplay_(CFG, cols.MOVE_DATE_IDX >= 0 ? row[cols.MOVE_DATE_IDX] : "");
    var cubicFeet = row[cols.CUBIC_FEET_IDX];
    cubicFeet = (cubicFeet !== null && cubicFeet !== undefined && cubicFeet !== "") ? String(cubicFeet) : "";
    
    var leadSource = row[cols.SOURCE_IDX];
    leadSource = (leadSource && String(leadSource).trim()) ? String(leadSource).trim() : "";
    
    out.push({
      username: username,
      job: job,
      trackerType: sourceLabel,
      source: leadSource || sourceLabel,
      rowDate: null,
      moveDate: moveDate,
      phone: phone,
      cubicFeet: cubicFeet
    });
  }
  
  return out;
}

/**
 * Scan CONTACTED LEADS sheet
 * @param {Object} CFG - Configuration
 * @returns {Array} Array of lead objects
 */
function SSCCP_scanContacted_(CFG) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEETS.CONTACTED);
  if (!sh) return [];
  
  var values = sh.getRange("A3:AE").getValues();
  var cols = CFG.CONTACTED_COLS;
  var out = [];
  
  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    
    var username = normalizeUsername_(row[cols.USERNAME_IDX]);
    var job = String(row[cols.JOB_IDX] || "").trim();
    var phone = SSCCP_normPhone_(row[cols.PHONE_IDX]);
    var exclude = row[cols.EXCLUDE_IDX];
    
    if (!username || !job || !phone) continue;
    if (exclude === true || String(exclude).toUpperCase() === "TRUE") continue;
    
    var moveDate = SSCCP_fmtMoveDateDisplay_(CFG, cols.MOVE_DATE_IDX >= 0 ? row[cols.MOVE_DATE_IDX] : "");
    var cubicFeet = row[cols.CUBIC_FEET_IDX] || "";
    var leadSource = row[cols.SOURCE_IDX];
    leadSource = (leadSource && String(leadSource).trim()) ? String(leadSource).trim() : "";
    
    out.push({
      username: username,
      job: job,
      trackerType: "Quoted Follow-Up",
      source: leadSource || "Quoted Follow-Up",
      rowDate: null,
      moveDate: moveDate,
      phone: phone,
      cubicFeet: cubicFeet,
      priorityLabel: "Already quoted"
    });
  }
  
  return out;
}

/**
 * Scan SAME DAY TRANSFERS sheet
 * @param {Object} CFG - Configuration
 * @returns {Array} Array of lead objects
 */
function SSCCP_scanTransfers_(CFG) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEETS.TRANSFERS);
  if (!sh) return [];
  
  var lastRow = sh.getLastRow();
  if (lastRow < CFG.TRANSFERS.START_ROW) return [];
  
  var numRows = lastRow - CFG.TRANSFERS.START_ROW + 1;
  var lastCol = sh.getLastColumn();
  var data = sh.getRange(CFG.TRANSFERS.START_ROW, 1, numRows, lastCol).getValues();
  
  var out = [];
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var username = normalizeUsername_(row[CFG.TRANSFERS.USERNAME_COL - 1]);
    if (!username) continue;
    
    var job = String(row[1] || "").trim() || ("TRANSFER_ROW_" + (CFG.TRANSFERS.START_ROW + i));
    var moveDate = "";
    if (CFG.TRANSFERS.MOVE_DATE_COL && CFG.TRANSFERS.MOVE_DATE_COL > 0) {
      moveDate = SSCCP_fmtMoveDateDisplay_(CFG, row[CFG.TRANSFERS.MOVE_DATE_COL - 1]);
    }
    
    out.push({
      username: username,
      job: job,
      trackerType: "Same Day Transfer",
      source: "Same Day Transfer",
      rowDate: null,
      moveDate: moveDate,
      phone: "",
      cubicFeet: ""
    });
  }
  
  return out;
}


/* 
 * SECTION 6: DATA PROCESSING
 * =============================================================== */

/**
 * Group leads by rep username
 * @param {Array} leads - Array of lead objects
 * @returns {Object} Map of username -> leads array
 */
function SSCCP_groupByRep_(leads) {
  var grouped = {};
  
  (leads || []).forEach(function(l) {
    var u = normalizeUsername_(l.username);
    if (!u) return;
    if (!grouped[u]) grouped[u] = [];
    grouped[u].push(l);
  });
  
  return grouped;
}

/**
 * Compute data integrity (find duplicates and unassigned reps)
 * @param {Object} team - Team roster index
 * @param {Object} grouped - Grouped leads by rep
 * @returns {Object} { duplicates: array, unassigned: array }
 */
function SSCCP_computeIntegrity_(team, grouped) {
  var reps = Object.keys(grouped || {});
  var unassigned = [];
  
  reps.forEach(function(u) {
    if (!team.repToManager[u]) unassigned.push(u);
  });
  
  return {
    duplicates: (team.duplicates || []).slice(),
    unassigned: unassigned
  };
}

/**
 * Build team aggregation for manager reports
 * @param {Object} grouped - Grouped leads
 * @param {Object} roster - Sales roster
 * @param {Object} team - Team roster
 * @param {Object} integrity - Integrity data
 * @returns {Object} Manager -> team data map
 */
function SSCCP_buildTeamAgg_(grouped, roster, team, integrity) {
  var dupSet = new Set(integrity.duplicates || []);
  var unassignedSet = new Set(integrity.unassigned || []);
  var agg = {};
  
  Object.keys(grouped || {}).forEach(function(username) {
    if (dupSet.has(username)) return;
    if (unassignedSet.has(username)) return;
    
    var leads = grouped[username] || [];
    var repProfile = roster.byUsername[username] || {};
    var managerName = team.repToManager[username];
    if (!managerName) return;
    
    if (!agg[managerName]) {
      agg[managerName] = { managerName: managerName, teamTotal: 0, reps: [] };
    }
    
    var repRow = {
      username: username,
      repName: repProfile.repName || username,
      total: leads.length,
      bySource: SSCCP_countBySource_(leads)
    };
    
    agg[managerName].teamTotal += repRow.total;
    agg[managerName].reps.push(repRow);
  });
  
  // Sort reps within each manager by lead count
  Object.keys(agg).forEach(function(m) {
    agg[m].reps.sort(function(a, b) {
      return (b.total || 0) - (a.total || 0);
    });
  });
  
  return agg;
}

/**
 * Count leads by source/tracker type
 * @param {Array} leads - Array of leads
 * @returns {Object} Source -> count map
 */
function SSCCP_countBySource_(leads) {
  var m = {};
  (leads || []).forEach(function(l) {
    var k = String(l.trackerType || l.source || "Unknown");
    m[k] = (m[k] || 0) + 1;
  });
  return m;
}

/**
 * Calculate Safe Ship Standard attempts
 * Formula: 2 dials + VM on 2nd dial (30+ sec) + 1 SMS = 1 SS Standard
 * @param {Object} rc - RingCentral data
 * @returns {number} SS Standard count
 */
function SSCCP_calcSSStandard_(rc) {
  if (!rc) return 0;
  
  var totalCalls = (rc.callsYesterday || 0) + (rc.callsToday || 0);
  var totalVM = (rc.vmYesterday || 0) + (rc.vmToday || 0);
  var totalSMS = (rc.smsYesterday || 0) + (rc.smsToday || 0);
  
  if (totalCalls < 2 || totalVM < 1 || totalSMS < 1) return 0;
  
  var potentialFromCalls = Math.floor(totalCalls / 2);
  var potentialFromVM = totalVM;
  var potentialFromSMS = totalSMS;
  
  return Math.min(potentialFromCalls, potentialFromVM, potentialFromSMS);
}


/* 
 * SECTION 7: SENDING LOGIC (REP ALERTS, MANAGER REPORTS)
 * =============================================================== */

/**
 * Send rep alerts with full features
 * - Test Mode: First N leads to admin only
 * - Cancellation support
 * - RC data freshness attachment
 * @returns {Object} { repsAlerted, exceptions, wasCancelled, wasTestMode }
 */
function SSCCP_sendRepAlerts_(CFG, run, ctx, windowInfo, grouped, roster, integrity, toast) {
  var mode = SSCCP_mode_(CFG);
  var dedupeHours = CFG.DEDUPE_HOURS;
  var exceptions = [];
  var repsAlerted = 0;
  
  // Test Mode configuration
  var testMode = SSCCP_isTestMode_();
  var testLeadLimit = CFG.TEST_MODE_LEAD_LIMIT;
  var testLeadsProcessed = 0;
  
  // Build exclusion sets
  var dupSet = new Set(integrity.duplicates || []);
  var unassignedSet = new Set(integrity.unassigned || []);
  var reps = Object.keys(grouped || {});
  
  // Get RC data freshness for display
  var rcFreshness = SSCCP_getRCDataFreshness_();
  
  for (var i = 0; i < reps.length; i++) {
    // 
    // CHECK FOR CANCELLATION
    // 
    if (SSCCP_isCancelled_()) {
      toast("[x] Cancelled by user after " + repsAlerted + " rep alerts", "Cancelled", 6);
      exceptions.push({
        title: "Run Cancelled",
        items: ["Stopped after " + repsAlerted + " of " + reps.length + " reps"]
      });
      break;
    }
    
    // 
    // TEST MODE LIMIT CHECK
    // 
    if (testMode && testLeadsProcessed >= testLeadLimit) {
      toast("🧪 Test Mode: Sent " + testLeadsProcessed + " leads to admin", "Test Complete", 6);
      break;
    }
    
    var username = reps[i];
    var leads = grouped[username] || [];
    
    // Progress toast every 6 reps
    if (i % 6 === 0) toast("Rep " + (i + 1) + "/" + reps.length, "Rep Alerts", 4);
    
    // Skip duplicates
    if (dupSet.has(username)) {
      exceptions.push({ title: "Duplicate Assignment (Skipped Rep)", items: [username] });
      continue;
    }
    
    // Get rep profile
    var profile = roster.byUsername[username];
    if (!profile || !profile.email) {
      exceptions.push({ title: "Missing Sales_Roster Email", items: [username] });
      continue;
    }
    
    // Determine routing
    var isUnassigned = unassignedSet.has(username);
    var intendedEmail = isUnassigned ? CFG.ADMIN_EMAIL : profile.email;
    
    // Apply deduplication
    var filtered = SSCCP_applyDedupe_(CFG, ctx.reportType, username, leads, dedupeHours);
    if (!filtered.length) continue;
    
    // Test Mode: limit leads and route to admin
    if (testMode) {
      filtered = filtered.slice(0, testLeadLimit - testLeadsProcessed);
      testLeadsProcessed += filtered.length;
    }
    
    // Determine final recipient
    var toEmail = testMode ? CFG.ADMIN_EMAIL : (mode.routeAllToAdmin ? CFG.ADMIN_EMAIL : intendedEmail);
    
    // Attach RC freshness to leads
    for (var l = 0; l < filtered.length; l++) {
      filtered[l]._rcFreshness = rcFreshness;
    }
    
    // Build payload
    var payload = {
      reportType: ctx.reportType,
      reportTitle: ctx.title,
      runId: run.runId,
      window: windowInfo,
      rep: { username: username, name: profile.repName || username, email: intendedEmail },
      leads: filtered,
      mode: mode,
      routedBecauseUnassigned: isUnassigned,
      showPhone: true,
      isTestMode: testMode,
      rcFreshness: rcFreshness
    };
    
    // Build subject
    var testPrefix = testMode ? "🧪 [TEST] " : "";
    var subject = testPrefix + SSCCP_stopEmoji_("GREEN") + " Safe Ship -- " + ctx.title + " -- " + payload.rep.name + " (Run " + run.runId + ")";
    
    // 
    // SEND EMAIL
    // 
    try {
      var html = SSCCP_renderRepEmailHtml_PREMIUM_(CFG, payload);
      
      MailApp.sendEmail({
        to: toEmail,
        subject: testMode ? "[TEST MODE] " + subject : (mode.routeAllToAdmin ? "[FORWARDED:" + intendedEmail + "] " + subject : subject),
        htmlBody: html,
        body: SSCCP_stripHtml_(html)
      });
      
      // Log success
      SSCCP_logNotification_(CFG, {
        timestamp: new Date(),
        runId: run.runId,
        type: "Rep Alert (" + ctx.reportType + ")" + (testMode ? " [TEST]" : ""),
        route: testMode ? "Test->Admin" : (mode.routeAllToAdmin ? "Forward-All/Admin" : (isUnassigned ? "Unassigned->Admin" : "Direct")),
        stoplight: "GREEN",
        username: username,
        repName: payload.rep.name,
        email: toEmail,
        manager: "",
        sourceSheets: payload.leads.map(function(x) { return x.source; }).filter(Boolean).join(", "),
        leadCount: payload.leads.length,
        jobNumbers: payload.leads.map(function(x) { return x.job; }).slice(0, 60).join(", "),
        emailSent: "YES",
        slackSent: testMode ? "SKIP" : "ATTEMPT",
        slackLookup: testMode ? "SKIP" : "ATTEMPT",
        error: ""
      });
      
    } catch (e) {
      exceptions.push({ title: "Email Send Error", items: [username + ": " + String(e)] });
      
      SSCCP_logNotification_(CFG, {
        timestamp: new Date(),
        runId: run.runId,
        type: "Rep Alert (" + ctx.reportType + ")",
        route: "ERROR",
        stoplight: "RED",
        username: username,
        repName: "",
        email: toEmail,
        manager: "",
        sourceSheets: "",
        leadCount: filtered.length,
        jobNumbers: "",
        emailSent: "NO",
        slackSent: "NO",
        slackLookup: "NO",
        error: String(e)
      });
      continue;
    }
    
    // 
    // SEND SLACK (skip in test mode)
    // 
    if (!testMode) {
      var slackRes = SSCCP_sendSlackDMToEmail_(CFG, toEmail, SSCCP_renderRepSlackText_(payload));
      if (!slackRes.ok) {
        exceptions.push({ title: "Slack DM Failure", items: [username + " -> " + toEmail + ": " + slackRes.error] });
      }
    }
    
    repsAlerted++;
    Utilities.sleep(140); // Rate limiting
  }
  
  return {
    repsAlerted: repsAlerted,
    exceptions: exceptions,
    wasCancelled: SSCCP_isCancelled_(),
    wasTestMode: testMode
  };
}

/**
 * Send manager reports with cancellation support
 * @returns {Object} { managersSent, exceptions, wasCancelled }
 */
function SSCCP_sendManagerReports_(CFG, run, ctx, windowInfo, grouped, roster, team, integrity, toast) {
  var mode = SSCCP_mode_(CFG);
  var managerAgg = SSCCP_buildTeamAgg_(grouped, roster, team, integrity);
  var managers = Object.keys(managerAgg || {});
  
  // Sort by team total (highest first)
  managers.sort(function(a, b) {
    return (managerAgg[b].teamTotal || 0) - (managerAgg[a].teamTotal || 0);
  });
  
  var exceptions = [];
  var managersSent = 0;
  
  // Skip in test mode
  if (SSCCP_isTestMode_()) {
    toast("🧪 Test Mode: Skipping manager reports", "Test Mode", 4);
    return { managersSent: 0, exceptions: [], wasCancelled: false, wasTestMode: true };
  }
  
  for (var i = 0; i < managers.length; i++) {
    // 
    // CHECK FOR CANCELLATION
    // 
    if (SSCCP_isCancelled_()) {
      toast("[x] Cancelled by user after " + managersSent + " manager reports", "Cancelled", 6);
      exceptions.push({
        title: "Run Cancelled",
        items: ["Stopped after " + managersSent + " of " + managers.length + " managers"]
      });
      break;
    }
    
    var managerName = managers[i];
    var mp = managerAgg[managerName];
    
    // Progress toast
    if (i % 4 === 0) toast("Manager " + (i + 1) + "/" + managers.length, "Manager Reports", 4);
    
    // Resolve manager email
    var managerEmail = SSCCP_resolveManagerEmail_(roster, managerName);
    if (!managerEmail) {
      exceptions.push({ title: "Missing Manager Email (Sales_Roster)", items: [managerName] });
      continue;
    }
    
    var toEmail = mode.routeAllToAdmin ? CFG.ADMIN_EMAIL : managerEmail;
    
    var payload = {
      reportType: ctx.reportType,
      reportTitle: ctx.title,
      runId: run.runId,
      window: windowInfo,
      manager: { name: managerName, email: managerEmail },
      team: mp,
      mode: mode
    };
    
    var subject = SSCCP_stopEmoji_("GREEN") + " Safe Ship -- Team " + ctx.title + " -- " + managerName + " (Run " + run.runId + ")";
    
    // 
    // SEND EMAIL
    // 
    try {
      var html = SSCCP_renderManagerEmailHtml_(CFG, payload);
      
      MailApp.sendEmail({
        to: toEmail,
        subject: mode.routeAllToAdmin ? "[FORWARDED:" + managerEmail + "] " + subject : subject,
        htmlBody: html,
        body: SSCCP_stripHtml_(html)
      });
      
      SSCCP_logNotification_(CFG, {
        timestamp: new Date(),
        runId: run.runId,
        type: "Manager Report (" + ctx.reportType + ")",
        route: mode.routeAllToAdmin ? "Forward-All/Admin" : "Direct",
        stoplight: "GREEN",
        username: "",
        repName: "",
        email: toEmail,
        manager: managerName,
        sourceSheets: "",
        leadCount: mp.teamTotal || 0,
        jobNumbers: "",
        emailSent: "YES",
        slackSent: "ATTEMPT",
        slackLookup: "ATTEMPT",
        error: ""
      });
      
    } catch (e) {
      exceptions.push({ title: "Manager Email Send Error", items: [managerName + ": " + String(e)] });
      continue;
    }
    
    // 
    // SEND SLACK
    // 
    var slackRes = SSCCP_sendSlackDMToEmail_(CFG, toEmail, SSCCP_renderManagerSlackText_(payload));
    if (!slackRes.ok) {
      exceptions.push({ title: "Manager Slack DM Failure", items: [managerName + " -> " + toEmail + ": " + slackRes.error] });
    }
    
    managersSent++;
    Utilities.sleep(180); // Rate limiting
  }
  
  return {
    managersSent: managersSent,
    exceptions: exceptions,
    wasCancelled: SSCCP_isCancelled_()
  };
}


/* 
 * SECTION 8: ADMIN DASHBOARD
 * =============================================================== */

/**
 * Send admin dashboard email and Slack
 */
function SSCCP_sendAdminDashboard_(CFG, run, ctx, windowInfo, finalStoplight, pack) {
  var recipients = (CFG.ADMIN_REPORT_EMAILS || []).join(",");
  var html = SSCCP_renderAdminDashboardHtml_(CFG, run, ctx, windowInfo, finalStoplight, pack);
  var subject = finalStoplight + " Safe Ship -- " + ctx.title + " (Run " + run.runId + ")";
  
  MailApp.sendEmail({
    to: recipients,
    subject: subject,
    htmlBody: html,
    body: SSCCP_stripHtml_(html)
  });
  
  // Send Slack to each admin
  (CFG.ADMIN_REPORT_EMAILS || []).forEach(function(email) {
    SSCCP_sendSlackDMToEmail_(CFG, email, SSCCP_renderAdminSlackText_(CFG, run, ctx, windowInfo, finalStoplight, pack));
    Utilities.sleep(120);
  });
}

/**
 * Build admin performance pack
 */
function SSCCP_buildAdminPerformancePack_(CFG, scan, grouped, roster, team, integrity, repSend, mgrSend) {
  var leads = (scan && scan.leads) ? scan.leads : [];
  var reps = Object.keys(grouped || {});
  
  var bySource = {};
  leads.forEach(function(l) {
    var k = String(l.source || "Unknown");
    bySource[k] = (bySource[k] || 0) + 1;
  });
  
  var exceptionsCount = ((repSend && repSend.exceptions) ? repSend.exceptions.length : 0) +
                        ((mgrSend && mgrSend.exceptions) ? mgrSend.exceptions.length : 0);
  
  return {
    totals: {
      totalLeads: leads.length,
      repsImpacted: reps.length,
      duplicates: (integrity && integrity.duplicates) ? integrity.duplicates.length : 0,
      unassigned: (integrity && integrity.unassigned) ? integrity.unassigned.length : 0,
      exceptions: exceptionsCount
    },
    bySource: bySource
  };
}

/**
 * Build exception blocks for admin dashboard
 */
function SSCCP_buildExceptionBlocks_(integrity, scan, repSend, mgrSend) {
  var repEx = (repSend && repSend.exceptions) ? repSend.exceptions : [];
  var mgrEx = (mgrSend && mgrSend.exceptions) ? mgrSend.exceptions : [];
  
  var blocks = [];
  blocks.push({ title: "Duplicate Assignments", items: (integrity && integrity.duplicates) || [] });
  blocks.push({ title: "Unassigned Reps", items: (integrity && integrity.unassigned) || [] });
  
  return blocks;
}

/**
 * Preview manifest (dry run)
 */
function SSCCP_previewManifest() {
  var CFG = getConfig_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var toast = function(m, t, s) { ss.toast(m, t || "Pre-Flight Manifest", s || 5); };
  
  var run = SSCCP_startRun_("MANIFEST");
  var windowInfo = SSCCP_computeWindow_(CFG, new Date());
  
  try {
    toast("Loading rosters...", "Step 1/3", 6);
    var roster = SSCCP_buildSalesRosterIndex_(CFG);
    var team = SSCCP_buildTeamRosterIndex_(CFG);
    
    toast("Scanning all sources (dry)...", "Step 2/3", 7);
    
    var all = [
      { reportType: "UNCONTACTED", sources: ["SMS", "CALL"] },
      { reportType: "QUOTED_FOLLOWUP", sources: ["CONTACTED"] },
      { reportType: "PRIORITY1_CALLVM", sources: ["P1CALL"] },
      { reportType: "SAME_DAY_TRANSFERS", sources: ["TRANSFERS"] }
    ];
    
    var lines = ["🚀 SAFE SHIP -- PRE-FLIGHT MANIFEST", "Run ID: " + run.runId, ""];
    
    all.forEach(function(item) {
      var ctx = { reportType: item.reportType, title: item.reportType, sources: item.sources };
      var scan;
      try {
        scan = SSCCP_scanSources_(CFG, ctx, windowInfo, roster, toast);
      } catch (e) {
        scan = { ok: false, summary: String(e), details: [] };
      }
      lines.push("* " + item.reportType + ": " + (scan.ok ? "OK (leads=" + (scan.leads || []).length + ")" : "FAIL"));
    });
    
    ui.alert(lines.join("\n"));
    
  } catch (e) {
    ui.alert("Manifest Error:\n\n" + String(e));
  }
}


/* 
 * SECTION 9: PREMIUM EMAIL RENDERERS
 * =============================================================== */

/**
 * Render Premium Rep Email HTML (v4.0)
 */
function SSCCP_renderRepEmailHtml_PREMIUM_(CFG, payload) {
  var leads = payload.leads || [];
  var nowStr = new Date().toLocaleString("en-US", { timeZone: CFG.TZ });
  var rcFreshness = payload.rcFreshness || {};
  
  // Helper to get initials
  var getInitials = function(name) {
    var parts = String(name || "?").trim().split(/\s+/);
    if (parts.length >= 2) {
      return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
    }
    return String(name || "?").charAt(0).toUpperCase();
  };
  
  // Build lead cards
  var leadCardsHtml = leads.slice(0, 50).map(function(l) {
    var rc = l.rc || {};
    
    // Calculate metrics
    var totalAttempts = (rc.callsYesterday || 0) + (rc.callsToday || 0);
    var attemptsToday = rc.callsToday || 0;
    var ssStandard = SSCCP_calcSSStandard_(rc);
    
    var phoneDisplay = l.phone ? SSCCP_maskPhoneForEmail_(CFG, l.phone) : "--";
    var cubicDisplay = (l.cubicFeet && l.cubicFeet !== "" && l.cubicFeet !== "0") ? String(l.cubicFeet) + " cu ft" : "--";
    var dateDisplay = l.moveDate || "--";
    
    // Urgency indicator
    var urgencyColor = totalAttempts === 0 ? "#DC2626" : (totalAttempts < 3 ? "#F59E0B" : "#10B981");
    
    return '<div style="background:#fff;border-radius:16px;border:1px solid #E5E7EB;margin-bottom:16px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.06);">' +
      // Card Header
      '<div style="background:linear-gradient(135deg, #1E3A5F 0%, #2D5A87 100%);padding:14px 18px;">' +
        '<div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;">' +
          '<div style="display:flex;align-items:center;gap:10px;">' +
            '<div style="width:10px;height:10px;border-radius:50%;background:' + urgencyColor + ';box-shadow:0 0 8px ' + urgencyColor + '50;"></div>' +
            '<span style="color:#fff;font-size:18px;font-weight:800;letter-spacing:0.5px;">JOB #' + escapeHtml_(String(l.job || "")) + '</span>' +
          '</div>' +
          '<div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;">' +
            '<span style="background:rgba(255,255,255,0.2);color:#fff;padding:4px 10px;border-radius:20px;font-size:11px;font-weight:700;">' + escapeHtml_(String(l.source || "")) + '</span>' +
            '<span style="background:#3b82f6;color:#fff;padding:4px 10px;border-radius:20px;font-size:11px;font-weight:700;">' + totalAttempts + ' Attempts</span>' +
            '<span style="background:#10b981;color:#fff;padding:4px 10px;border-radius:20px;font-size:11px;font-weight:700;">' + attemptsToday + ' Today</span>' +
            (ssStandard > 0 ? '<span style="background:#f59e0b;color:#fff;padding:4px 10px;border-radius:20px;font-size:11px;font-weight:700;">[x] ' + ssStandard + ' SS Standard</span>' : '') +
          '</div>' +
        '</div>' +
      '</div>' +
      // Lead Info
      '<div style="padding:14px 18px;background:#F8FAFC;border-bottom:1px solid #E5E7EB;">' +
        '<div style="display:flex;flex-wrap:wrap;gap:16px;">' +
          '<div style="flex:1;min-width:130px;">' +
            '<div style="font-size:10px;color:#64748B;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;">📱 Phone</div>' +
            '<div style="font-size:14px;color:#0F172A;font-weight:600;font-family:ui-monospace,monospace;margin-top:2px;">' + escapeHtml_(phoneDisplay) + '</div>' +
          '</div>' +
          '<div style="flex:1;min-width:100px;">' +
            '<div style="font-size:10px;color:#64748B;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;">[CALENDAR] Move Date</div>' +
            '<div style="font-size:14px;color:#0F172A;font-weight:600;margin-top:2px;">' + escapeHtml_(dateDisplay) + '</div>' +
          '</div>' +
          '<div style="flex:1;min-width:80px;">' +
            '<div style="font-size:10px;color:#64748B;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;">📦 Size</div>' +
            '<div style="font-size:14px;color:#0F172A;font-weight:600;margin-top:2px;">' + escapeHtml_(cubicDisplay) + '</div>' +
          '</div>' +
        '</div>' +
      '</div>' +
      // Contact Activity
      renderLeadContactActivity_(l) +
    '</div>';
  }).join("");
  
  var moreLeads = leads.length > 50 ? '<div style="text-align:center;padding:16px;color:#64748B;font-weight:600;">+ ' + (leads.length - 50) + ' more leads</div>' : '';
  var repInitials = getInitials(payload.rep.name);
  
  // Calculate totals
  var totalSMS = 0, totalCalls = 0, totalVM = 0;
  leads.forEach(function(l) {
    var rc = l.rc || {};
    totalSMS += (rc.smsYesterday || 0) + (rc.smsToday || 0);
    totalCalls += (rc.callsYesterday || 0) + (rc.callsToday || 0);
    totalVM += (rc.vmYesterday || 0) + (rc.vmToday || 0);
  });
  
  // Full email template
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
  '<body style="margin:0;padding:0;background:linear-gradient(180deg, #F1F5F9 0%, #E2E8F0 100%);font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,\'Helvetica Neue\',Arial,sans-serif;">' +
  '<div style="max-width:700px;margin:0 auto;padding:20px 16px;">' +
    
    // HEADER CARD
    '<div style="background:linear-gradient(135deg, #1E3A5F 0%, #2D5A87 50%, #3B82F6 100%);border-radius:20px;padding:0;margin-bottom:20px;overflow:hidden;box-shadow:0 10px 40px rgba(30,58,95,0.3);">' +
      '<div style="padding:20px 24px 16px;border-bottom:1px solid rgba(255,255,255,0.1);">' +
        '<div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;">' +
          '<div>' +
            '<div style="font-size:11px;color:rgba(255,255,255,0.7);font-weight:700;letter-spacing:1px;text-transform:uppercase;">[SHIP] SAFE SHIP CONTACT CHECKER</div>' +
            '<div style="font-size:24px;color:#fff;font-weight:800;margin-top:4px;letter-spacing:-0.5px;">' + escapeHtml_(payload.reportTitle) + '</div>' +
          '</div>' +
          '<div style="background:rgba(255,255,255,0.15);backdrop-filter:blur(10px);border-radius:12px;padding:10px 16px;text-align:center;">' +
            '<div style="font-size:10px;color:rgba(255,255,255,0.8);font-weight:700;letter-spacing:0.5px;">LEADS</div>' +
            '<div style="font-size:32px;color:#fff;font-weight:900;line-height:1;">' + leads.length + '</div>' +
          '</div>' +
        '</div>' +
      '</div>' +
      // Rep Info
      '<div style="padding:16px 24px;background:rgba(0,0,0,0.1);">' +
        '<div style="display:flex;flex-wrap:wrap;gap:16px;align-items:center;">' +
          '<div style="display:flex;align-items:center;gap:10px;">' +
            '<div style="width:42px;height:42px;line-height:42px;background:linear-gradient(135deg, #10B981 0%, #059669 100%);border-radius:50%;text-align:center;font-size:16px;color:#fff;font-weight:800;box-shadow:0 4px 12px rgba(16,185,129,0.4);">' + escapeHtml_(repInitials) + '</div>' +
            '<div>' +
              '<div style="font-size:16px;color:#fff;font-weight:700;">' + escapeHtml_(payload.rep.name) + '</div>' +
              '<div style="font-size:12px;color:rgba(255,255,255,0.7);">@' + escapeHtml_(payload.rep.username) + '</div>' +
            '</div>' +
          '</div>' +
          '<div style="flex:1;"></div>' +
          '<div style="display:flex;gap:8px;flex-wrap:wrap;">' +
            '<div style="background:rgba(59,130,246,0.3);padding:6px 12px;border-radius:20px;text-align:center;">' +
              '<div style="font-size:14px;color:#fff;font-weight:800;">' + totalSMS + '</div>' +
              '<div style="font-size:9px;color:rgba(255,255,255,0.7);">SMS</div>' +
            '</div>' +
            '<div style="background:rgba(245,158,11,0.3);padding:6px 12px;border-radius:20px;text-align:center;">' +
              '<div style="font-size:14px;color:#fff;font-weight:800;">' + totalCalls + '</div>' +
              '<div style="font-size:9px;color:rgba(255,255,255,0.7);">CALLS</div>' +
            '</div>' +
            '<div style="background:rgba(139,92,246,0.3);padding:6px 12px;border-radius:20px;text-align:center;">' +
              '<div style="font-size:14px;color:#fff;font-weight:800;">' + totalVM + '</div>' +
              '<div style="font-size:9px;color:rgba(255,255,255,0.7);">VM</div>' +
            '</div>' +
          '</div>' +
        '</div>' +
      '</div>' +
    '</div>' +
    
    // RC FRESHNESS & FORMAT INFO
    '<div style="background:#fff;border-radius:14px;padding:14px 18px;margin-bottom:16px;border:1px solid #E5E7EB;box-shadow:0 1px 3px rgba(0,0,0,0.05);">' +
      '<div style="display:flex;flex-wrap:wrap;gap:12px;align-items:center;justify-content:space-between;">' +
        '<div style="font-size:11px;color:#64748B;font-weight:700;">[CHART] Stats Format: Yesterday - Today</div>' +
        '<div style="font-size:10px;color:#94A3B8;">' +
          '[PHONE] Calls: ' + (rcFreshness.callLogLabel || "--") + ' * ' +
          '[MSG] SMS: ' + (rcFreshness.smsLogLabel || "--") +
        '</div>' +
      '</div>' +
      (payload.isTestMode ? '<div style="margin-top:8px;padding:6px 10px;background:#fef3c7;border-radius:6px;font-size:11px;color:#92400e;font-weight:600;">🧪 TEST MODE -- This is a preview email sent to admin only</div>' : '') +
    '</div>' +
    
    // LEAD CARDS
    '<div style="margin-bottom:20px;">' + leadCardsHtml + moreLeads + '</div>' +
    
    // FOOTER
    '<div style="text-align:center;padding:16px;color:#94A3B8;font-size:11px;">' +
      '<div style="margin-bottom:4px;">Safe Ship Contact Checker Pro v4.0 * Generated ' + escapeHtml_(nowStr) + '</div>' +
      '<div>Run: ' + escapeHtml_(payload.runId) + ' * ' + escapeHtml_(payload.window.label || "") + '</div>' +
    '</div>' +
    
  '</div></body></html>';
}

/**
 * Render Admin Dashboard HTML
 */
function SSCCP_renderAdminDashboardHtml_(CFG, run, ctx, windowInfo, stoplight, pack) {
  var stopColor = stoplight === "GREEN" ? "#10B981" : stoplight === "YELLOW" ? "#F59E0B" : "#EF4444";
  var nowStr = new Date().toLocaleString("en-US", { timeZone: CFG.TZ });
  var mode = SSCCP_mode_(CFG);
  
  var getInitials = function(name) {
    var parts = String(name || "?").trim().split(/\s+/);
    if (parts.length >= 2) {
      return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
    }
    return String(name || "?").charAt(0).toUpperCase();
  };
  
  var stat = function(label, value, icon) {
    return '<div style="flex:1;min-width:120px;background:#fff;border-radius:14px;padding:16px;border:1px solid #E5E7EB;box-shadow:0 1px 3px rgba(0,0,0,0.05);">' +
      '<div style="font-size:11px;color:#64748B;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;">' + icon + ' ' + escapeHtml_(label) + '</div>' +
      '<div style="font-size:28px;font-weight:800;color:#0F172A;margin-top:4px;">' + escapeHtml_(String(value)) + '</div>' +
    '</div>';
  };
  
  var colorStat = function(label, value, icon, bgColor, textColor) {
    return '<div style="flex:1;min-width:120px;background:' + bgColor + ';border-radius:14px;padding:16px;border:1px solid ' + textColor + '30;box-shadow:0 1px 3px rgba(0,0,0,0.05);">' +
      '<div style="font-size:11px;color:' + textColor + ';font-weight:700;text-transform:uppercase;letter-spacing:0.5px;">' + icon + ' ' + escapeHtml_(label) + '</div>' +
      '<div style="font-size:28px;font-weight:800;color:' + textColor + ';margin-top:4px;">' + escapeHtml_(String(value)) + '</div>' +
    '</div>';
  };
  
  var managerNames = Object.keys(pack.teamAgg || {});
  managerNames.sort(function(a, b) {
    return (pack.teamAgg[b].teamTotal || 0) - (pack.teamAgg[a].teamTotal || 0);
  });
  
  var grandTotalSMS = 0, grandTotalCall = 0;
  managerNames.forEach(function(mgr) {
    var mp = pack.teamAgg[mgr];
    (mp.reps || []).forEach(function(r) {
      var src = r.bySource || {};
      grandTotalSMS += (src["SMS"] || 0);
      grandTotalCall += (src["CALL/VM"] || 0) + (src["Priority 1 Call/VM"] || 0);
    });
  });
  
  // Build team cards
  var teamCardsHtml = managerNames.length ? managerNames.map(function(mgr) {
    var mp = pack.teamAgg[mgr];
    var reps = (mp.reps || []).slice(0, 15);
    
    var teamSMS = 0, teamCall = 0;
    reps.forEach(function(r) {
      var src = r.bySource || {};
      teamSMS += (src["SMS"] || 0);
      teamCall += (src["CALL/VM"] || 0) + (src["Priority 1 Call/VM"] || 0);
    });
    
    var rows = reps.map(function(r, idx) {
      var src = r.bySource || {};
      var smsCount = src["SMS"] || 0;
      var callCount = (src["CALL/VM"] || 0) + (src["Priority 1 Call/VM"] || 0);
      var initials = getInitials(r.repName);
      var rowBg = idx % 2 === 0 ? '#FFFFFF' : '#F8FAFC';
      
      var badges = [];
      if (smsCount > 0) badges.push('<span style="display:inline-flex;align-items:center;gap:3px;background:#DBEAFE;color:#1E40AF;padding:3px 8px;border-radius:12px;font-size:11px;font-weight:700;">[MSG] ' + smsCount + '</span>');
      if (callCount > 0) badges.push('<span style="display:inline-flex;align-items:center;gap:3px;background:#FEF3C7;color:#92400E;padding:3px 8px;border-radius:12px;font-size:11px;font-weight:700;">[PHONE] ' + callCount + '</span>');
      var badgeHtml = badges.length ? '<div style="display:flex;gap:4px;margin-top:4px;">' + badges.join('') + '</div>' : '';
      
      return '<tr style="background:' + rowBg + ';">' +
        '<td style="padding:10px 14px;border-top:1px solid #F1F5F9;">' +
          '<div style="display:flex;align-items:center;gap:10px;">' +
            '<div style="width:36px;height:36px;line-height:36px;background:linear-gradient(135deg,#10B981 0%,#059669 100%);border-radius:50%;text-align:center;color:#fff;font-weight:700;font-size:13px;flex-shrink:0;">' + escapeHtml_(initials) + '</div>' +
            '<div>' +
              '<div style="font-weight:600;color:#0F172A;">' + escapeHtml_(r.repName) + '</div>' +
              '<div style="font-size:11px;color:#64748B;">@' + escapeHtml_(r.username) + '</div>' +
              badgeHtml +
            '</div>' +
          '</div>' +
        '</td>' +
        '<td style="padding:10px 14px;border-top:1px solid #F1F5F9;text-align:right;font-weight:800;font-size:18px;color:#0F172A;vertical-align:top;">' + r.total + '</td>' +
      '</tr>';
    }).join("");
    
    var teamBadges = [];
    if (teamSMS > 0) teamBadges.push('<span style="background:rgba(219,234,254,0.4);color:#fff;padding:4px 10px;border-radius:12px;font-size:12px;font-weight:700;">[MSG] ' + teamSMS + ' SMS</span>');
    if (teamCall > 0) teamBadges.push('<span style="background:rgba(254,243,199,0.4);color:#fff;padding:4px 10px;border-radius:12px;font-size:12px;font-weight:700;">[PHONE] ' + teamCall + ' Call</span>');
    var teamBadgeHtml = teamBadges.length ? '<div style="display:flex;gap:6px;margin-top:6px;">' + teamBadges.join('') + '</div>' : '';
    
    return '<div style="background:#fff;border-radius:16px;border:1px solid #E5E7EB;margin-bottom:16px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.05);">' +
      '<div style="background:linear-gradient(135deg, #1E3A5F 0%, #2D5A87 100%);padding:14px 18px;">' +
        '<div style="display:flex;justify-content:space-between;align-items:flex-start;">' +
          '<div>' +
            '<div style="color:#fff;font-weight:700;font-size:16px;">' + escapeHtml_(mgr) + '</div>' +
            teamBadgeHtml +
          '</div>' +
          '<div style="background:rgba(255,255,255,0.2);color:#fff;padding:6px 14px;border-radius:20px;font-size:14px;font-weight:800;">' + (mp.teamTotal || 0) + ' leads</div>' +
        '</div>' +
      '</div>' +
      '<table style="width:100%;border-collapse:collapse;"><tbody>' + rows + '</tbody></table>' +
    '</div>';
  }).join("") : '<div style="background:#fff;border-radius:14px;padding:20px;text-align:center;color:#64748B;">No team data</div>';
  
  return '<!DOCTYPE html><html><head><meta charset="utf-8"></head>' +
  '<body style="margin:0;padding:0;background:#F1F5F9;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,Arial,sans-serif;">' +
  '<div style="max-width:900px;margin:0 auto;padding:20px 16px;">' +
    '<div style="background:linear-gradient(135deg, #1E3A5F 0%, #2D5A87 50%, #3B82F6 100%);border-radius:20px;padding:24px;margin-bottom:20px;box-shadow:0 10px 40px rgba(30,58,95,0.3);">' +
      '<div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:16px;">' +
        '<div>' +
          '<div style="font-size:11px;color:rgba(255,255,255,0.7);font-weight:700;letter-spacing:1px;">[SHIP] ADMIN DASHBOARD</div>' +
          '<div style="font-size:26px;color:#fff;font-weight:800;margin-top:4px;">' + escapeHtml_(ctx.title) + '</div>' +
          '<div style="font-size:12px;color:rgba(255,255,255,0.7);margin-top:8px;">Run: ' + escapeHtml_(run.runId) + ' * ' + escapeHtml_(nowStr) + '</div>' +
        '</div>' +
        '<div style="background:' + stopColor + ';color:#fff;padding:12px 24px;border-radius:30px;font-weight:800;font-size:14px;box-shadow:0 4px 12px ' + stopColor + '50;">' + stoplight + '</div>' +
      '</div>' +
    '</div>' +
    '<div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:12px;">' +
      stat("Leads", (pack.metrics && pack.metrics.totalLeadsScanned) || 0, "[CHART]") +
      stat("Reps", (pack.metrics && pack.metrics.repsWithLeads) || 0, "👤") +
      stat("Managers", (pack.metrics && pack.metrics.managersSummarized) || 0, "👥") +
    '</div>' +
    '<div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:20px;">' +
      colorStat("Total SMS", grandTotalSMS, "[MSG]", "#DBEAFE", "#1E40AF") +
      colorStat("Total Call/VM", grandTotalCall, "[PHONE]", "#FEF3C7", "#92400E") +
    '</div>' +
    '<div style="font-size:18px;font-weight:800;color:#0F172A;margin-bottom:12px;">Team Breakdown</div>' +
    teamCardsHtml +
  '</div></body></html>';
}

/**
 * Render Manager Email HTML
 */
function SSCCP_renderManagerEmailHtml_(CFG, payload) {
  var reps = (payload.team.reps || []).slice(0, 25);
  
  var teamBySource = {};
  reps.forEach(function(r) {
    var src = r.bySource || {};
    Object.keys(src).forEach(function(k) {
      teamBySource[k] = (teamBySource[k] || 0) + src[k];
    });
  });
  var teamSMS = teamBySource["SMS"] || 0;
  var teamCall = (teamBySource["CALL/VM"] || 0) + (teamBySource["Priority 1 Call/VM"] || 0);
  
  var getInitials = function(name) {
    var parts = String(name || "?").trim().split(/\s+/);
    if (parts.length >= 2) {
      return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
    }
    return String(name || "?").charAt(0).toUpperCase();
  };
  
  var rows = reps.map(function(r, idx) {
    var src = r.bySource || {};
    var smsCount = src["SMS"] || 0;
    var callCount = (src["CALL/VM"] || 0) + (src["Priority 1 Call/VM"] || 0);
    var initials = getInitials(r.repName);
    
    var badges = [];
    if (smsCount > 0) badges.push('<span style="display:inline-flex;align-items:center;gap:4px;background:linear-gradient(135deg,#DBEAFE 0%,#BFDBFE 100%);color:#1E40AF;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:700;border:1px solid #93C5FD;">[MSG] ' + smsCount + ' SMS</span>');
    if (callCount > 0) badges.push('<span style="display:inline-flex;align-items:center;gap:4px;background:linear-gradient(135deg,#FEF3C7 0%,#FDE68A 100%);color:#92400E;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:700;border:1px solid #FCD34D;">[PHONE] ' + callCount + ' Call</span>');
    var badgeHtml = badges.length ? '<div style="display:flex;gap:6px;margin-top:8px;flex-wrap:wrap;">' + badges.join('') + '</div>' : '';
    
    var rowBg = idx % 2 === 0 ? '#FFFFFF' : '#F8FAFC';
    
    return '<tr style="background:' + rowBg + ';">' +
      '<td style="padding:16px 18px;border-bottom:1px solid #E5E7EB;">' +
        '<div style="display:flex;align-items:flex-start;gap:12px;">' +
          '<div style="width:44px;height:44px;line-height:44px;background:linear-gradient(135deg,#10B981 0%,#059669 100%);border-radius:50%;text-align:center;color:#fff;font-weight:800;font-size:15px;flex-shrink:0;box-shadow:0 2px 8px rgba(16,185,129,0.3);">' + escapeHtml_(initials) + '</div>' +
          '<div style="flex:1;">' +
            '<div style="font-weight:700;color:#0F172A;font-size:15px;">' + escapeHtml_(r.repName) + '</div>' +
            '<div style="font-size:12px;color:#64748B;margin-top:2px;">@' + escapeHtml_(r.username) + '</div>' +
            badgeHtml +
          '</div>' +
        '</div>' +
      '</td>' +
      '<td style="padding:16px 18px;border-bottom:1px solid #E5E7EB;text-align:right;vertical-align:middle;">' +
        '<div style="font-weight:900;font-size:28px;color:#0F172A;line-height:1;">' + r.total + '</div>' +
        '<div style="font-size:11px;color:#64748B;margin-top:2px;">leads</div>' +
      '</td>' +
    '</tr>';
  }).join("");
  
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
  '<body style="margin:0;padding:0;background:linear-gradient(180deg,#F1F5F9 0%,#E2E8F0 100%);font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,Arial,sans-serif;">' +
  '<div style="max-width:680px;margin:0 auto;padding:24px 16px;">' +
    '<div style="background:linear-gradient(135deg, #1E3A5F 0%, #2D5A87 50%, #3B82F6 100%);border-radius:24px;padding:0;margin-bottom:20px;overflow:hidden;box-shadow:0 10px 40px rgba(30,58,95,0.3);">' +
      '<div style="padding:28px 28px 20px;">' +
        '<div style="font-size:11px;color:rgba(255,255,255,0.7);font-weight:700;letter-spacing:1.5px;text-transform:uppercase;">[SHIP] Manager Report</div>' +
        '<div style="font-size:28px;color:#fff;font-weight:800;margin-top:6px;letter-spacing:-0.5px;">Team ' + escapeHtml_(payload.reportTitle) + '</div>' +
        '<div style="font-size:15px;color:rgba(255,255,255,0.85);margin-top:8px;font-weight:500;">' + escapeHtml_(payload.manager.name) + '</div>' +
      '</div>' +
      '<div style="background:rgba(0,0,0,0.15);padding:18px 28px;display:flex;gap:16px;flex-wrap:wrap;">' +
        '<div style="background:rgba(255,255,255,0.12);backdrop-filter:blur(10px);border-radius:14px;padding:14px 20px;flex:1;min-width:100px;text-align:center;">' +
          '<div style="font-size:10px;color:rgba(255,255,255,0.7);font-weight:700;letter-spacing:0.5px;text-transform:uppercase;">Team Total</div>' +
          '<div style="font-size:36px;color:#fff;font-weight:900;line-height:1.1;margin-top:4px;">' + (payload.team.teamTotal || 0) + '</div>' +
        '</div>' +
        '<div style="background:rgba(255,255,255,0.12);backdrop-filter:blur(10px);border-radius:14px;padding:14px 20px;flex:1;min-width:80px;text-align:center;">' +
          '<div style="font-size:10px;color:rgba(255,255,255,0.7);font-weight:700;letter-spacing:0.5px;text-transform:uppercase;">Reps</div>' +
          '<div style="font-size:36px;color:#fff;font-weight:900;line-height:1.1;margin-top:4px;">' + (payload.team.reps || []).length + '</div>' +
        '</div>' +
        '<div style="background:rgba(59,130,246,0.3);border-radius:14px;padding:14px 20px;flex:1;min-width:80px;text-align:center;">' +
          '<div style="font-size:10px;color:rgba(255,255,255,0.7);font-weight:700;letter-spacing:0.5px;text-transform:uppercase;">[MSG] SMS</div>' +
          '<div style="font-size:36px;color:#fff;font-weight:900;line-height:1.1;margin-top:4px;">' + teamSMS + '</div>' +
        '</div>' +
        '<div style="background:rgba(245,158,11,0.3);border-radius:14px;padding:14px 20px;flex:1;min-width:80px;text-align:center;">' +
          '<div style="font-size:10px;color:rgba(255,255,255,0.7);font-weight:700;letter-spacing:0.5px;text-transform:uppercase;">[PHONE] Call</div>' +
          '<div style="font-size:36px;color:#fff;font-weight:900;line-height:1.1;margin-top:4px;">' + teamCall + '</div>' +
        '</div>' +
      '</div>' +
    '</div>' +
    '<div style="background:#fff;border-radius:20px;border:1px solid #E5E7EB;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.06);">' +
      '<div style="background:linear-gradient(135deg,#F8FAFC 0%,#F1F5F9 100%);padding:14px 18px;border-bottom:1px solid #E5E7EB;">' +
        '<div style="font-size:13px;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:0.5px;">[CHART] Team Breakdown</div>' +
      '</div>' +
      '<table style="width:100%;border-collapse:collapse;"><tbody>' + rows + '</tbody></table>' +
    '</div>' +
    '<div style="text-align:center;padding:20px;color:#94A3B8;font-size:11px;">' +
      '<div>Safe Ship Contact Checker Pro v4.0 * Run: ' + escapeHtml_(payload.runId || "") + '</div>' +
    '</div>' +
  '</div></body></html>';
}


/* 
 * SECTION 10: CONTACT ACTIVITY RENDERER (v4.0)
 * =============================================================== */

/**
 * Render lead contact activity section (v4.0)
 * Features:
 * - Alert banners for replied, long call, all SMS failed
 * - "Yesterday - Today" stat bubbles
 * - RC data freshness timestamps
 * - Failed SMS counter
 * - Process Direction with contextual tips
 * - Last 5 calls/SMS with badges
 */
function renderLeadContactActivity_(lead) {
  var rc = lead.rc || {};
  var freshness = lead._rcFreshness || {};
  var html = '';
  
  // 
  // ALERT BANNERS
  // 
  
  // Alert: Lead answered through text
  if (rc.hasReplied) {
    html += '<div style="background:linear-gradient(135deg,#10b981,#059669);color:#fff;padding:8px 14px;border-radius:6px;margin:10px 18px 0;display:flex;align-items:center;gap:8px;font-size:12px;">' +
      '<span style="font-size:16px;">[MSG]</span>' +
      '<span><strong>This Lead Answered Through Text!</strong> Check SMS conversation for their response</span></div>';
  }
  
  // Alert: Wrong Priority Risk (call > 4 minutes)
  if (rc.hasLongCall) {
    html += '<div style="background:linear-gradient(135deg,#f59e0b,#d97706);color:#000;padding:8px 14px;border-radius:6px;margin:10px 18px 0;display:flex;align-items:center;gap:8px;font-size:12px;">' +
      '<span style="font-size:16px;">[!]</span>' +
      '<span><strong>Wrong Priority Risk - Review File</strong> (Call duration: ' + (rc.longestCallEverDisplay || rc.longestCallDisplay || "4+ min") + ')</span></div>';
  }
  
  // Alert: All SMS attempts failed
  if (rc.allSmsFailed && (rc.smsTotal || 0) > 0) {
    html += '<div style="background:linear-gradient(135deg,#ef4444,#dc2626);color:#fff;padding:8px 14px;border-radius:6px;margin:10px 18px 0;display:flex;align-items:center;gap:8px;font-size:12px;">' +
      '<span style="font-size:16px;">🚫</span>' +
      '<span><strong>All SMS Attempts Failed!</strong> Check phone number validity or try a different method</span></div>';
  }
  
  html += '<div style="padding:14px 18px;">';
  
  // 
  // HEADER WITH RC FRESHNESS
  // 
  html += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">' +
    '<div style="font-size:10px;color:#64748b;font-weight:600;">[CHART] CONTACT ACTIVITY</div>' +
    '<div style="font-size:9px;color:#94a3b8;">[PHONE] ' + (freshness.callLogLabel || "--") + ' * [MSG] ' + (freshness.smsLogLabel || "--") + '</div>' +
  '</div>';
  
  // 
  // STAT BUBBLES ROW
  // 
  html += '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;">';
  html += renderStatBubble_("SMS SENT", rc.smsYesterday || 0, rc.smsToday || 0, "📱", "#dbeafe", "#1e40af");
  html += renderStatBubble_("CALLS", rc.callsYesterday || 0, rc.callsToday || 0, "[PHONE]", "#dcfce7", "#166534");
  html += renderStatBubble_("VOICEMAILS", rc.vmYesterday || 0, rc.vmToday || 0, "📥", "#fef3c7", "#92400e");
  
  // Longest Call bubble
  html += '<div style="background:#f1f5f9;border-radius:6px;padding:6px 10px;text-align:center;min-width:80px;">' +
    '<div style="font-size:9px;color:#64748b;font-weight:600;">[TIMER] LONGEST CALL</div>' +
    '<div style="font-size:16px;font-weight:800;color:#334155;">' + (rc.longestCallDisplay || "0:00") + '</div>' +
    '<div style="font-size:8px;color:#94a3b8;">since yesterday</div>' +
  '</div>';
  
  // Failed SMS counter
  var failedSmsCount = 0;
  if (rc.last5SMS) {
    for (var f = 0; f < rc.last5SMS.length; f++) {
      if (rc.last5SMS[f].isFailed) failedSmsCount++;
    }
  }
  
  html += '<div style="background:' + (failedSmsCount > 0 ? "#fee2e2" : "#f1f5f9") + ';border-radius:6px;padding:6px 10px;text-align:center;min-width:70px;">' +
    '<div style="font-size:9px;color:' + (failedSmsCount > 0 ? "#991b1b" : "#64748b") + ';font-weight:600;">[X] FAILED SMS</div>' +
    '<div style="font-size:16px;font-weight:800;color:' + (failedSmsCount > 0 ? "#dc2626" : "#334155") + ';">' + failedSmsCount + '</div>' +
    '<div style="font-size:8px;color:#94a3b8;">since yesterday</div>' +
  '</div>';
  
  html += '</div>'; // end stats row
  
  // 
  // BUILD CALLS HTML
  // 
  var callsHtml = '';
  var callBg = (rc.last5Calls && rc.last5Calls.length > 0) ? "#dcfce7" : "#fee2e2";
  var callColor = (rc.last5Calls && rc.last5Calls.length > 0) ? "#166534" : "#991b1b";
  
  if (rc.last5Calls && rc.last5Calls.length > 0) {
    for (var i = 0; i < rc.last5Calls.length; i++) {
      var call = rc.last5Calls[i];
      var inBadge = call.isInbound ? '<span style="display:inline-block;background:#06b6d4;color:#fff;font-size:9px;font-weight:700;padding:1px 4px;border-radius:3px;margin-right:4px;">IN</span>' : '';
      var vmBadge = (call.isVM && !call.isInbound) ? '<span style="display:inline-block;background:#f59e0b;color:#fff;font-size:9px;font-weight:700;padding:1px 4px;border-radius:3px;margin-left:4px;">VM</span>' : '';
      
      callsHtml += '<div style="font-size:10px;color:' + callColor + ';padding:1px 0;white-space:nowrap;">' +
        inBadge + call.timeAgo + ' <span style="opacity:0.7;">(' + call.timeFormatted + ')</span> ' + vmBadge + '<strong>' + call.durationDisplay + '</strong></div>';
    }
  } else {
    callsHtml = '<div style="font-size:10px;color:' + callColor + ';font-weight:600;">Never</div>';
  }
  
  // 
  // BUILD SMS HTML
  // 
  var smsHtml = '';
  var smsBg = "#fee2e2";
  var smsColor = "#991b1b";
  
  if (rc.last5SMS && rc.last5SMS.length > 0) {
    var hasSuccess = false;
    for (var k = 0; k < rc.last5SMS.length; k++) {
      if (!rc.last5SMS[k].isFailed) { hasSuccess = true; break; }
    }
    
    if (hasSuccess) {
      smsBg = "#dbeafe";
      smsColor = "#1e40af";
    } else {
      smsBg = "#fef3c7";
      smsColor = "#92400e";
    }
    
    for (var j = 0; j < rc.last5SMS.length; j++) {
      var sms = rc.last5SMS[j];
      var statusIndicator = sms.isFailed
        ? '<span style="display:inline-block;background:#ef4444;color:#fff;font-size:8px;font-weight:700;padding:1px 4px;border-radius:3px;margin-left:4px;">FAILED</span>'
        : '<span style="color:#10b981;margin-left:4px;">[x]</span>';
      
      smsHtml += '<div style="font-size:10px;color:' + smsColor + ';padding:1px 0;white-space:nowrap;">' +
        sms.timeAgo + ' <span style="opacity:0.7;">(' + sms.timeFormatted + ')</span>' + statusIndicator + '</div>';
    }
  } else {
    smsHtml = '<div style="font-size:10px;color:' + smsColor + ';font-weight:600;">Never</div>';
  }
  
  // 
  // PROCESS DIRECTION
  // 
  // v4.0.1: Use local simple version to avoid collision with Config.js enhanced version
  var processDirection = getProcessDirection_Code_(lead.trackerType);
  
  // Contextual tip for failed SMS
  var processTip = '';
  if (rc.allSmsFailed && (rc.smsTotal || 0) > 0) {
    processTip = '<div style="font-size:8px;color:#dc2626;font-weight:600;margin-top:4px;padding-top:4px;border-top:1px dashed ' + processDirection.borderColor + ';">💡 TIP: Pass to neighbor to attempt contact</div>';
  }
  
  // 
  // RENDER THREE BOXES
  // 
  html += '<div style="display:flex;gap:8px;flex-wrap:wrap;">';
  
  // Calls box
  html += '<div style="background:' + callBg + ';border-radius:6px;padding:8px 10px;flex:1;min-width:120px;">' +
    '<div style="font-size:9px;color:' + callColor + ';font-weight:700;margin-bottom:4px;">[PHONE] LAST 5 CALLS</div>' +
    callsHtml +
  '</div>';
  
  // SMS box
  html += '<div style="background:' + smsBg + ';border-radius:6px;padding:8px 10px;flex:1;min-width:120px;">' +
    '<div style="font-size:9px;color:' + smsColor + ';font-weight:700;margin-bottom:4px;">[MSG] LAST 5 SMS</div>' +
    smsHtml +
  '</div>';
  
  // Process Direction box
  html += '<div style="background:' + processDirection.bgColor + ';border-radius:6px;padding:8px 10px;flex:1;min-width:120px;border:2px solid ' + processDirection.borderColor + ';">' +
    '<div style="font-size:9px;color:' + processDirection.textColor + ';font-weight:700;margin-bottom:4px;">🎯 PROCESS DIRECTION</div>' +
    '<div style="font-size:12px;color:' + processDirection.textColor + ';font-weight:800;margin-bottom:4px;">' + processDirection.icon + ' ' + processDirection.action + '</div>' +
    '<div style="font-size:9px;color:' + processDirection.textColor + ';opacity:0.8;">Source: ' + processDirection.source + '</div>' +
    processTip +
  '</div>';
  
  html += '</div>'; // end flex container
  html += '</div>'; // end padding container
  
  return html;
}

/**
 * Render stat bubble with "Yesterday - Today" label
 */
function renderStatBubble_(label, yesterday, today, icon, bgColor, textColor) {
  return '<div style="background:' + bgColor + ';border-radius:6px;padding:6px 10px;text-align:center;min-width:85px;">' +
    '<div style="font-size:9px;color:' + textColor + ';font-weight:600;margin-bottom:2px;">' + icon + ' ' + label + '</div>' +
    '<div style="display:flex;align-items:center;justify-content:center;gap:4px;">' +
      '<span style="font-size:15px;font-weight:800;color:' + textColor + ';">' + yesterday + '</span>' +
      '<span style="font-size:11px;color:#64748b;">-</span>' +
      '<span style="font-size:15px;font-weight:800;color:' + textColor + ';">' + today + '</span>' +
    '</div>' +
    '<div style="font-size:8px;color:#94a3b8;margin-top:1px;">Yesterday - Today</div>' +
  '</div>';
}

/**
 * Get process direction based on tracker type (SIMPLE VERSION)
 * RENAMED to avoid collision with Config.js enhanced version
 * Config.js has getProcessDirection_(trackerType, rc) which analyzes actual RC activity
 */
function getProcessDirection_Code_(trackerType) {
  var type = String(trackerType || "").toUpperCase();
  
  if (type === "SMS") {
    return { action: "SEND SMS", icon: "📱", source: "SMS Tracker", bgColor: "#DBEAFE", borderColor: "#3B82F6", textColor: "#1E40AF" };
  }
  if (type === "CALL/VM" || type === "CALL" || type === "VOICEMAIL") {
    return { action: "CALL/VM", icon: "[PHONE]", source: "Call & VM Tracker", bgColor: "#DCFCE7", borderColor: "#22C55E", textColor: "#166534" };
  }
  if (type.indexOf("PRIORITY") !== -1 || type === "PRIORITY 1 CALL/VM") {
    return { action: "PRIORITY CALL", icon: "[GOAL][PHONE]", source: "Priority 1 Tracker", bgColor: "#FEE2E2", borderColor: "#EF4444", textColor: "#991B1B" };
  }
  if (type.indexOf("QUOTED") !== -1 || type.indexOf("CONTACTED") !== -1 || type.indexOf("FOLLOW") !== -1) {
    return { action: "FOLLOW UP", icon: "💰", source: "Contacted Leads", bgColor: "#FEF3C7", borderColor: "#F59E0B", textColor: "#92400E" };
  }
  if (type.indexOf("TRANSFER") !== -1 || type.indexOf("SAME DAY") !== -1) {
    return { action: "TRANSFER", icon: "⚡", source: "Same Day Transfers", bgColor: "#F3E8FF", borderColor: "#A855F7", textColor: "#7C3AED" };
  }
  
  return { action: "CONTACT LEAD", icon: "[CLIPBOARD]", source: type || "Unknown", bgColor: "#F1F5F9", borderColor: "#94A3B8", textColor: "#475569" };
}


/* 
 * SECTION 11: SLACK INTEGRATION
 * =============================================================== */

/**
 * Send Slack DM to user by email
 */
function SSCCP_sendSlackDMToEmail_(CFG, email, text) {
  try {
    if (!CFG.SLACK.BOT_TOKEN) return { ok: false, error: "Missing SLACK_BOT_TOKEN" };
    
    var userId = SSCCP_slackLookupUserIdByEmail_(CFG, email);
    if (!userId) return { ok: false, error: "Slack lookup failed" };
    
    var channelId = SSCCP_slackOpenIm_(CFG, userId);
    if (!channelId) return { ok: false, error: "Slack DM open failed" };
    
    SSCCP_slackPostMessage_(CFG, channelId, text);
    return { ok: true };
    
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

/**
 * Lookup Slack user ID by email (cached)
 */
function SSCCP_slackLookupUserIdByEmail_(CFG, email) {
  var cache = CacheService.getScriptCache();
  var key = "slack_uid_" + String(email || "").toLowerCase();
  var cached = cache.get(key);
  if (cached) return cached;
  
  var url = "https://slack.com/api/users.lookupByEmail?email=" + encodeURIComponent(email);
  var res = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true,
    headers: { Authorization: "Bearer " + CFG.SLACK.BOT_TOKEN }
  });
  
  var json = JSON.parse(res.getContentText() || "{}");
  var userId = (json && json.ok && json.user && json.user.id) ? json.user.id : "";
  
  if (userId) cache.put(key, userId, CFG.CACHE_TTL_SECONDS);
  return userId || "";
}

/**
 * Open Slack IM channel
 */
function SSCCP_slackOpenIm_(CFG, userId) {
  var res = UrlFetchApp.fetch("https://slack.com/api/conversations.open", {
    method: "post",
    muteHttpExceptions: true,
    contentType: "application/json",
    headers: { Authorization: "Bearer " + CFG.SLACK.BOT_TOKEN },
    payload: JSON.stringify({ users: userId })
  });
  
  var json = JSON.parse(res.getContentText() || "{}");
  return (json && json.ok && json.channel && json.channel.id) ? json.channel.id : "";
}

/**
 * Post message to Slack channel
 */
function SSCCP_slackPostMessage_(CFG, channelId, text) {
  UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", {
    method: "post",
    muteHttpExceptions: true,
    contentType: "application/json",
    headers: { Authorization: "Bearer " + CFG.SLACK.BOT_TOKEN },
    payload: JSON.stringify({ channel: channelId, text: text })
  });
}

/**
 * Render rep Slack text
 */
function SSCCP_renderRepSlackText_(payload) {
  var lines = [
    "[PIN] *Safe Ship -- " + payload.reportTitle + "*",
    "Rep: *" + payload.rep.name + "* (@" + payload.rep.username + ")",
    "Run: `" + payload.runId + "`",
    "",
    "*" + payload.leads.length + " leads need attention:*"
  ];
  
  payload.leads.slice(0, 20).forEach(function(l) {
    var rc = l.rc || {};
    var totalAttempts = (rc.callsYesterday || 0) + (rc.callsToday || 0);
    var cubicInfo = l.cubicFeet ? " | " + l.cubicFeet + " cu ft" : "";
    lines.push("* *" + l.job + "* (" + l.source + ") -- " + totalAttempts + " attempts" + cubicInfo);
  });
  
  if (payload.leads.length > 20) lines.push("... +" + (payload.leads.length - 20) + " more");
  lines.push("", "_[STOP] SS Standard = 2 dials + VM on 2nd + SMS._");
  
  return lines.join("\n");
}

/**
 * Render manager Slack text
 */
function SSCCP_renderManagerSlackText_(payload) {
  var lines = [
    "📣 *Safe Ship -- Team " + payload.reportTitle + "*",
    "Manager: *" + payload.manager.name + "*",
    "Team total: *" + (payload.team.teamTotal || 0) + " leads*",
    ""
  ];
  
  (payload.team.reps || []).slice(0, 15).forEach(function(r) {
    lines.push("* " + r.repName + " (@" + r.username + ") -- " + r.total + " leads");
  });
  
  if ((payload.team.reps || []).length > 15) {
    lines.push("... +" + ((payload.team.reps || []).length - 15) + " more reps");
  }
  
  return lines.join("\n");
}

/**
 * Render admin Slack text
 */
function SSCCP_renderAdminSlackText_(CFG, run, ctx, windowInfo, stoplight, pack) {
  var emoji = stoplight === "GREEN" ? "[GREEN]" : stoplight === "YELLOW" ? "[YELLOW]" : "[RED]";
  
  return [
    emoji + " *Safe Ship Admin -- " + ctx.title + "*",
    "Run: `" + run.runId + "`",
    "Mode: " + SSCCP_mode_(CFG).effectiveModeLabel,
    "",
    "[CHART] *Stats:*",
    "* Leads: " + ((pack.metrics && pack.metrics.totalLeadsScanned) || 0),
    "* Reps: " + ((pack.metrics && pack.metrics.repsWithLeads) || 0),
    "* Managers: " + ((pack.metrics && pack.metrics.managersSummarized) || 0)
  ].join("\n");
}


/* 
 * SECTION 12: ROSTER MANAGEMENT
 * =============================================================== */

/**
 * Build Sales_Roster index
 */
function SSCCP_buildSalesRosterIndex_(CFG) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEETS.SALES_ROSTER);
  if (!sh) throw new Error("Required sheet '" + CFG.SHEETS.SALES_ROSTER + "' not found.");
  var data = sh.getDataRange().getValues();
  var header = data[0].map(function(v) { return String(v || "").trim().toLowerCase(); });
  
  var unameIdx = SSCCP_findHeader_(header, ["username", "user", "rep username"]);
  var emailIdx = SSCCP_findHeader_(header, ["e-mail 1 - value", "email", "email address"]);
  var firstIdx = SSCCP_findHeader_(header, ["first name"]);
  var lastIdx = SSCCP_findHeader_(header, ["last name"]);
  var titleIdx = SSCCP_findHeader_(header, ["organization title", "title", "job title"]);
  
  if (unameIdx < 0 || emailIdx < 0) throw new Error("Sales_Roster missing required headers.");
  
  var byUsername = {};
  var byManagerName = {};
  var byManagerNameLower = {};
  var managerUsernameSet = {};
  var byFullNameLower = {};
  
  for (var i = 1; i < data.length; i++) {
    var u = normalizeUsername_(data[i][unameIdx]);
    if (!u) continue;
    
    var email = String(data[i][emailIdx] || "").trim();
    var first = String(data[i][firstIdx] || "").trim();
    var last = String(data[i][lastIdx] || "").trim();
    var fullName = (first + " " + last).trim();
    var repName = fullName || u;
    var title = String((titleIdx >= 0 ? data[i][titleIdx] : "") || "").trim().toLowerCase();
    
    byUsername[u] = { email: email, repName: repName, fullName: fullName, title: title };
    if (fullName) byFullNameLower[fullName.toLowerCase()] = u;
    
    if (title === "sales manager") {
      if (fullName) {
        byManagerName[fullName] = email;
        byManagerNameLower[fullName.toLowerCase()] = email;
      }
      managerUsernameSet[u] = true;
    }
  }
  
  return {
    byUsername: byUsername,
    byManagerName: byManagerName,
    byManagerNameLower: byManagerNameLower,
    managerUsernameSet: managerUsernameSet,
    byFullNameLower: byFullNameLower
  };
}

/**
 * Build Team_Roster index
 */
function SSCCP_buildTeamRosterIndex_(CFG) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEETS.TEAM_ROSTER);
  if (!sh) throw new Error("Required sheet '" + CFG.SHEETS.TEAM_ROSTER + "' not found.");
  var lr = sh.getLastRow();
  var lc = sh.getLastColumn();
  var values = sh.getRange(1, 1, lr, lc).getValues();
  
  var repToManager = {};
  var repToManagers = {};
  var duplicates = [];
  
  var headerRow = values[0].map(function(v) { return String(v || "").trim(); });
  
  // Find columns to ignore
  var ignoreCols = new Set();
  for (var c = 0; c < lc; c++) {
    var h = String(headerRow[c] || "").trim().toLowerCase();
    if (h === "sales managers" || h === "salesmanagers" || h === "managers") {
      ignoreCols.add(c);
    }
  }
  
  // Build rep -> manager mapping
  for (var c = 0; c < lc; c++) {
    if (ignoreCols.has(c)) continue;
    
    var managerName = String(values[0][c] || "").trim();
    if (!managerName) continue;
    
    for (var r = 1; r < lr; r++) {
      var uname = normalizeUsername_(values[r][c]);
      if (!uname) continue;
      
      if (!repToManagers[uname]) repToManagers[uname] = [];
      repToManagers[uname].push(managerName);
    }
  }
  
  // Detect duplicates and set primary manager
  Object.keys(repToManagers).forEach(function(u) {
    if (repToManagers[u].length > 1) duplicates.push(u);
    repToManager[u] = repToManagers[u][0] || "";
  });
  
  var managerHeaders = headerRow.filter(Boolean).filter(function(h) {
    var hl = h.toLowerCase();
    return !(hl === "sales managers" || hl === "salesmanagers" || hl === "managers");
  });
  
  return {
    managerHeaders: managerHeaders,
    repToManager: repToManager,
    duplicates: duplicates
  };
}

/**
 * Resolve manager email
 */
function SSCCP_resolveManagerEmail_(roster, managerNameOrUsername) {
  var raw = String(managerNameOrUsername || "").trim();
  if (!raw) return "";
  
  if (roster.byManagerName && roster.byManagerName[raw]) return roster.byManagerName[raw];
  
  var low = raw.toLowerCase();
  if (roster.byManagerNameLower && roster.byManagerNameLower[low]) return roster.byManagerNameLower[low];
  
  var u = normalizeUsername_(raw);
  if (roster.byUsername && roster.byUsername[u] && roster.byUsername[u].email) {
    return roster.byUsername[u].email;
  }
  
  return "";
}


/* 
 * SECTION 13: DEDUPLICATION LOGIC
 * =============================================================== */

/**
 * Apply deduplication to leads
 */
function SSCCP_applyDedupe_(CFG, reportType, username, leads, dedupeHours) {
  if (!dedupeHours) return leads; // No deduplication
  
  var recentKeys = SSCCP_getRecentDedupeKeys_(CFG, dedupeHours);
  
  return (leads || []).filter(function(l) {
    var k = reportType + "|" + username + "|" + String(l.job || "").trim();
    return !recentKeys[k];
  });
}

/**
 * Get recent dedupe keys from notification log
 */
function SSCCP_getRecentDedupeKeys_(CFG, dedupeHours) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEETS.NOTIFICATION_LOG);
  var lr = sh.getLastRow();
  if (lr < 2) return {};
  
  var data = sh.getRange(2, 1, lr - 1, Math.min(16, sh.getLastColumn())).getValues();
  var cutoff = Date.now() - (dedupeHours * 60 * 60 * 1000);
  var map = {};
  
  data.forEach(function(row) {
    var ts = row[0];
    var type = String(row[2] || "");
    var uname = normalizeUsername_(row[5]);
    var jobs = String(row[11] || "");
    
    var d = SSCCP_asDate_(ts);
    if (!d || d.getTime() < cutoff) return;
    
    if (type.indexOf("Rep Alert") !== -1 && uname && jobs) {
      jobs.split(",").map(function(x) { return x.trim(); }).filter(Boolean).forEach(function(job) {
        var rt = (type.match(/\(([^)]+)\)/) || [])[1] || "UNKNOWN";
        map[rt + "|" + uname + "|" + job] = true;
      });
    }
  });
  
  return map;
}


/* 
 * SECTION 14: LOGGING
 * =============================================================== */

/**
 * Ensure Notification_Log sheet exists with headers
 */
function SSCCP_ensureNotificationLog_(CFG) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEETS.NOTIFICATION_LOG);
  
  if (!sh) sh = ss.insertSheet(CFG.SHEETS.NOTIFICATION_LOG);
  
  var headers = [
    "timestamp", "runId", "type", "route", "stoplight", "username", "repName",
    "email", "manager", "sourceSheets", "leadCount", "jobNumbers",
    "emailSent", "slackSent", "slackLookup", "error"
  ];
  
  if (sh.getLastRow() === 0) sh.appendRow(headers);
}

/**
 * Log notification to sheet
 */
function SSCCP_logNotification_(CFG, row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEETS.NOTIFICATION_LOG);
  
  sh.appendRow([
    row.timestamp || new Date(),
    row.runId || "",
    row.type || "",
    row.route || "",
    row.stoplight || "",
    row.username || "",
    row.repName || "",
    row.email || "",
    row.manager || "",
    row.sourceSheets || "",
    row.leadCount || 0,
    row.jobNumbers || "",
    row.emailSent || "",
    row.slackSent || "",
    row.slackLookup || "",
    row.error || ""
  ]);
}

/**
 * Ensure Run_Log sheet exists
 */
function SSCCP_ensureRunLog_(CFG) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEETS.RUN_LOG);
  
  if (!sh) sh = ss.insertSheet(CFG.SHEETS.RUN_LOG);
  
  var headers = ["timestamp", "runId", "event", "summary", "details"];
  if (sh.getLastRow() === 0) sh.appendRow(headers);
}

/**
 * Log run event
 */
function SSCCP_logRunEvent_(CFG, run, event, summary, detailsArr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEETS.RUN_LOG);
  
  sh.appendRow([
    new Date(),
    run.runId,
    event,
    summary || "",
    (detailsArr || []).join(" | ")
  ]);
}


/* 
 * SECTION 15: CANCELLATION & TEST MODE
 * =============================================================== */

/**
 * Check if run is cancelled
 */
function SSCCP_isCancelled_() {
  var cache = CacheService.getScriptCache();
  return cache.get("SSCCP_CANCELLED") === "true";
}

/**
 * Set cancellation flag (called by sidebar)
 */
function SSCCP_setCancelled(cancelled) {
  var cache = CacheService.getScriptCache();
  if (cancelled) {
    cache.put("SSCCP_CANCELLED", "true", 300); // 5 minute expiry
  } else {
    cache.remove("SSCCP_CANCELLED");
  }
}

/**
 * Clear cancellation flag
 */
function SSCCP_clearCancellation_() {
  var cache = CacheService.getScriptCache();
  cache.remove("SSCCP_CANCELLED");
}

/**
 * Check if test mode is enabled
 */
function SSCCP_isTestMode_() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty("SSCCP_TEST_MODE") === "true";
}

/**
 * Toggle test mode
 */
function SSCCP_toggleTestMode() {
  var props = PropertiesService.getScriptProperties();
  var cur = props.getProperty("SSCCP_TEST_MODE") === "true";
  props.setProperty("SSCCP_TEST_MODE", cur ? "false" : "true");
  
  var CFG = getConfig_();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Test Mode is now " + (cur ? "OFF (Full Run)" : "ON (First " + CFG.TEST_MODE_LEAD_LIMIT + " leads to Admin only)"),
    "Safe Ship", 6
  );
}


/* 
 * SECTION 16: RC DATA FRESHNESS
 * =============================================================== */

/**
 * Get RC data freshness timestamps
 */
function SSCCP_getRCDataFreshness_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var CFG = getConfig_();
  
  var result = {
    callLogTime: null,
    smsLogTime: null,
    callLogLabel: "Unknown",
    smsLogLabel: "Unknown"
  };
  
  try {
    // Check RC CALL LOG for most recent entry
    var callSheet = ss.getSheetByName(CFG.SHEETS.RC_CALL_LOG);
    if (callSheet && callSheet.getLastRow() > 1) {
      var lastCallDate = callSheet.getRange(2, 5).getValue(); // Date column
      var lastCallTime = callSheet.getRange(2, 6).getValue(); // Time column
      
      if (lastCallDate) {
        var dt = new Date(lastCallDate);
        if (lastCallTime instanceof Date) {
          dt.setHours(lastCallTime.getHours(), lastCallTime.getMinutes());
        }
        result.callLogTime = dt;
        result.callLogLabel = Utilities.formatDate(dt, CFG.TZ, "M/d h:mm a");
      }
    }
    
    // Check RC SMS LOG for most recent entry
    var smsSheet = ss.getSheetByName(CFG.SHEETS.RC_SMS_LOG);
    if (smsSheet && smsSheet.getLastRow() > 1) {
      var lastSmsDate = smsSheet.getRange(2, 8).getValue(); // DateTime column
      
      if (lastSmsDate) {
        var dt2 = new Date(lastSmsDate);
        result.smsLogTime = dt2;
        result.smsLogLabel = Utilities.formatDate(dt2, CFG.TZ, "M/d h:mm a");
      }
    }
  } catch (e) {
    console.warn("Error getting RC freshness:", e);
  }
  
  return result;
}


/* 
 * SECTION 17: UTILITY FUNCTIONS
 * =============================================================== */

/**
 * Get current mode settings
 */
function SSCCP_mode_(CFG) {
  var routeAllToAdmin = !!(CFG.SAFE_MODE || CFG.FORWARD_ALL);
  return {
    routeAllToAdmin: routeAllToAdmin,
    effectiveModeLabel: CFG.SAFE_MODE ? "TEST (Safe Mode ON)" : (CFG.FORWARD_ALL ? "LIVE + Forward-All" : "LIVE (Direct)")
  };
}

/**
 * Merge stoplight status
 */
function SSCCP_mergeStoplight_(base, repSend, mgrSend) {
  var hasIssues = (repSend.exceptions || []).length || (mgrSend.exceptions || []).length;
  if (base === "RED") return "RED";
  if (hasIssues) return "YELLOW";
  return base;
}

/**
 * Generate run ID
 */
function SSCCP_startRun_(label) {
  var ts = Utilities.formatDate(new Date(), "America/New_York", "yyyyMMdd-HHmmss");
  var rand = Math.random().toString(36).slice(2, 6).toUpperCase();
  return { runId: label + "_" + ts + "_" + rand };
}

/**
 * Compute window info
 */
function SSCCP_computeWindow_(CFG, now) {
  var start = new Date(now);
  start.setHours(0, 0, 0, 0);
  var end = new Date(now);
  end.setHours(23, 59, 59, 999);
  return { start: start, end: end, label: "Tracker Visibility (no date filtering)" };
}

/**
 * Get tracker window label
 */
function SSCCP_getTrackerWindowLabel_(CFG, ctx) {
  if (!ctx || !ctx.reportType) return "Tracker-defined timeframe";
  switch (ctx.reportType) {
    case "SAME_DAY_TRANSFERS": return "Today (tracker-defined)";
    default: return "Last " + CFG.WINDOW_DAYS + " days (tracker-defined)";
  }
}

/**
 * Read cell value
 */
function SSCCP_readCell_(sheetName, a1) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return sh ? sh.getRange(a1).getDisplayValue() : "";
}

/**
 * Check if value indicates Priority 0
 */
function SSCCP_isPriority0_(v) {
  var s = String(v || "").trim().toLowerCase();
  return s === "0" || s === "priority 0" || s === "p0";
}

/**
 * Check if value indicates Priority 1
 */
function SSCCP_isPriority1_(v) {
  var s = String(v || "").trim().toLowerCase();
  return s === "1" || s === "priority 1" || s === "p1";
}

/**
 * Get stoplight emoji
 */
function SSCCP_stopEmoji_(stoplight) {
  return stoplight === "GREEN" ? "[GREEN]" : stoplight === "YELLOW" ? "[YELLOW]" : "[RED]";
}

/**
 * Find header index
 */
function SSCCP_findHeader_(headerLower, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var idx = headerLower.indexOf(String(candidates[i]).toLowerCase());
    if (idx > -1) return idx;
  }
  return -1;
}

/**
 * Format move date for display
 */
function SSCCP_fmtMoveDateDisplay_(CFG, v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, CFG.TZ, "MM/dd/yyyy");
  }
  return String(v).trim();
}

/**
 * Normalize phone number
 */
function SSCCP_normPhone_(v) {
  var s = String(v || "").trim();
  return s ? s.replace(/\s+/g, "") : "";
}

/**
 * Mask phone for email display
 */
function SSCCP_maskPhoneForEmail_(CFG, phone) {
  var s = String(phone || "").trim();
  if (!s) return "";
  
  var digits = s.replace(/[^\d]/g, "");
  if (digits.length < 4) return s;
  
  var showPhones = PropertiesService.getScriptProperties().getProperty("SSCCP_SHOW_REP_PHONES") === "true";
  
  if (showPhones) {
    if (digits.length === 11 && digits.charAt(0) === "1") {
      digits = digits.substring(1);
    }
    if (digits.length === 10) {
      return "(" + digits.substring(0, 3) + ") " + digits.substring(3, 6) + "-" + digits.substring(6);
    }
    return s;
  } else {
    return "***-***-" + digits.slice(-4);
  }
}

/**
 * Strip HTML tags
 */
function SSCCP_stripHtml_(html) {
  return String(html || "").replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();
}

/**
 * Escape HTML entities
 */
function escapeHtml_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/**
 * Normalize username to uppercase
 */
function normalizeUsername_(v) {
  return String(v || "").trim().toUpperCase();
}

/**
 * Sanitize single email
 */
function sanitizeSingleEmail_(v) {
  return String(v || "").trim();
}

/**
 * Sanitize email list
 */
function sanitizeEmailListRaw_(v) {
  return String(v || "").split(",").map(function(x) { return x.trim(); }).filter(Boolean);
}

/**
 * Parse integer with default
 */
function toInt_(v, def) {
  var n = parseInt(String(v || ""), 10);
  return isNaN(n) ? def : n;
}

/**
 * Parse date value
 */
function SSCCP_asDate_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) return v;
  var s = String(v).trim();
  if (!s) return null;
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}


/* 
 * SECTION 18: MENU UTILITIES
 * =============================================================== */

/**
 * Toggle Safe Mode
 */
function SSCCP_toggleSafeMode() {
  var props = PropertiesService.getScriptProperties();
  var cur = props.getProperty("UNCONTACTED_DRY_RUN") === "true";
  props.setProperty("UNCONTACTED_DRY_RUN", cur ? "false" : "true");
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Safe Mode is now " + (cur ? "OFF (Live)" : "ON (Dry Run)"),
    "Safe Ship", 6
  );
}

/**
 * Toggle Forward-All
 */
function SSCCP_toggleForwardAll() {
  var props = PropertiesService.getScriptProperties();
  var cur = props.getProperty("UNCONTACTED_REDIRECT_ALL") === "true";
  props.setProperty("UNCONTACTED_REDIRECT_ALL", cur ? "false" : "true");
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Forward-All is now " + (cur ? "OFF" : "ON"),
    "Safe Ship", 6
  );
}

/**
 * Show current status
 */
function SSCCP_showStatus() {
  var CFG = getConfig_();
  var mode = SSCCP_mode_(CFG);
  var testMode = SSCCP_isTestMode_();
  
  var lines = [
    "[SHIP] Safe Ship Contact Checker v4.0 -- Status",
    "===========================================",
    "Mode: " + mode.effectiveModeLabel,
    "Safe Mode: " + CFG.SAFE_MODE,
    "Forward-All: " + CFG.FORWARD_ALL,
    "Test Mode: " + testMode,
    "",
    "ADMIN_EMAIL: " + (CFG.ADMIN_EMAIL || "MISSING"),
    "SLACK_BOT_TOKEN: " + (CFG.SLACK.BOT_TOKEN ? "SET" : "MISSING")
  ];
  
  SpreadsheetApp.getUi().alert(lines.join("\n"));
}

/**
 * Test admin email
 */
function SSCCP_testAdminEmail() {
  var CFG = getConfig_();
  if (!CFG.ADMIN_EMAIL) {
    SpreadsheetApp.getUi().alert("ADMIN_EMAIL is missing.");
    return;
  }
  
  MailApp.sendEmail({
    to: CFG.ADMIN_EMAIL,
    subject: "[OK] Safe Ship Bot Email Test",
    body: "This is a test email from Safe Ship Contact Checker v4.0."
  });
  
  SpreadsheetApp.getActiveSpreadsheet().toast("[OK] Admin email test sent.", "Safe Ship", 6);
}

/**
 * Open Notification_Log sheet
 */
function SSCCP_openNotificationLog() {
  var CFG = getConfig_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEETS.NOTIFICATION_LOG);
  if (sh) ss.setActiveSheet(sh);
}

/**
 * Open Run_Log sheet
 */
function SSCCP_openRunLog() {
  var CFG = getConfig_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEETS.RUN_LOG);
  if (sh) ss.setActiveSheet(sh);
}
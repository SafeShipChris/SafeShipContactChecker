/**************************************************************
 * 🚢 SAFE SHIP CONTACT CHECKER PRO — TRACKER BUILDER
 * File: SSCCP_TrackerBuilder.gs
 *
 * High-performance lead-compliance pipeline that replaces
 * slow IMPORTRANGE + formula filtering with Apps Script.
 *
 * ENTRY POINT: SSCCP_buildTrackersFromGranotData()
 *
 * FEATURES:
 *   - Reads GRANOT DATA leads (current day thru last 5 days)
 *   - Filters by SCHEDULED REPS
 *   - Builds outreach indexes from RC CALL LOG & RC SMS LOG (today only)
 *   - Writes to 4 tracker sheets based on priority + outreach counts
 *   - Modular design for future sidebar integration
 *
 * SHEET NAMES (exact - Excel truncates to 31 chars):
 *   - "GRANOT DATA"
 *   - "SCHEDULED REPS"
 *   - "RC CALL LOG" / "RC SMS LOG"
 *   - "SMS TRACKER"
 *   - "CALL & VOICEMAIL TRACKER"
 *   - "PRIORITY 1 CALL & VOICEMAIL TRA" (truncated)
 *   - "PRIORITY 3,5 CALL & VOICEMAIL T" (truncated)
 *
 * TRACKER LAYOUT:
 *   - Rows 1-8: Parameters/UI (preserved)
 *   - Row 2: Headers (columns F-M visible)
 *   - Row 9+: Data rows
 *   - Core columns F-M: MOVE DATE, PHONE NUMBER, Count, JOB #, USER, CF, SOURCE, EXCLUDE
 *   - Extra columns N+: Additional GRANOT fields for future use
 *
 * PARAMETER CELLS:
 *   - SMS TRACKER: B6=requiredSmsCount, B7=priority
 *   - CALL trackers: A8=priority, B8=requiredCallCount
 *
 * v1.0 — Initial Implementation
 **************************************************************/

/* ═══════════════════════════════════════════════════════════════
 * CONFIGURATION CONSTANTS
 * ═══════════════════════════════════════════════════════════════ */

var SSCCP_TB_CONFIG = {
  // Sheet names (exact matches - full names in Google Sheets)
  SHEETS: {
    GRANOT_DATA: "GRANOT DATA",
    SCHEDULED_REPS: "SCHEDULED REPS",
    RC_CALL_LOG: "RC CALL LOG",
    RC_SMS_LOG: "RC SMS LOG",
    SMS_TRACKER: "SMS TRACKER",
    CALL_TRACKER: "CALL & VOICEMAIL TRACKER",
    PRIORITY_1_TRACKER: "PRIORITY 1 CALL & VOICEMAIL TRACKER",
    PRIORITY_35_TRACKER: "PRIORITY 3,5 CALL & VOICEMAIL TRACKER",
    RUN_LOG: "Run_Log"
  },
  
  // GRANOT DATA column indexes (0-based)
  GRANOT_COLS: {
    DEPT: 0,        // A
    JOB_NO: 1,      // B
    OPEN_DATE: 2,   // C
    OPEN_TIME: 3,   // D
    USER: 4,        // E
    FOLLOW_UP: 5,   // F
    PRIORITY: 6,    // G
    FIRST_NAME: 7,  // H
    LAST_NAME: 8,   // I
    MOVE_DATE: 9,   // J
    EMAIL: 10,      // K
    PHONE1: 11,     // L
    PHONE2: 12,     // M
    FROM_ADDRESS: 13, // N
    FROM_CITY: 14,  // O
    FROM_STATE: 15, // P
    FROM_ZIP: 16,   // Q
    TO_ADDRESS: 17, // R
    TO_CITY: 18,    // S
    TO_STATE: 19,   // T
    TO_ZIP: 20,     // U
    CF: 21,         // V
    SOURCE: 22,     // W
    REF_NO: 23      // X
  },
  
  // RC CALL LOG column indexes (0-based)
  CALL_LOG_COLS: {
    DIRECTION: 0,   // A
    TYPE: 1,        // B
    PHONE: 2,       // C
    NAME: 3,        // D
    DATE: 4,        // E
    TIME: 5,        // F
    ACTION: 6,      // G
    RESULT: 7,      // H
    REASON: 8,      // I
    DURATION: 9     // J
  },
  
  // RC SMS LOG column indexes (0-based)
  SMS_LOG_COLS: {
    DIRECTION: 0,      // A
    TYPE: 1,           // B
    MSG_TYPE: 2,       // C
    SENDER_NUM: 3,     // D
    SENDER_NAME: 4,    // E
    RECIPIENT_NUM: 5,  // F
    RECIPIENT_NAME: 6, // G
    DATETIME: 7,       // H
    SEGMENT_COUNT: 8,  // I
    MSG_STATUS: 9      // J
  },
  
  // Tracker layout configuration
  TRACKER_LAYOUT: {
    HEADER_ROW: 2,       // Row 2 has headers
    DATA_START_ROW: 3,   // Data starts row 3
    PRESERVE_ROWS: 2,    // Don't touch rows 1-2
    DATA_START_COL: 6    // Column F (1-based)
  },
  
  // Core tracker columns (F-M in row 2)
  CORE_HEADERS: ["MOVE DATE", "PHONE NUMBER", "Count", "JOB #", "USER", "CF", "SOURCE", "EXCLUDE"],
  
  // Extra columns to append from GRANOT DATA (after core columns)
  EXTRA_FIELDS: [
    "email", "first_name", "last_name", 
    "from_city", "from_state", "from_zip",
    "to_city", "to_state", "to_zip", 
    "open_date", "open_time", "ref_no", "dept", "priority"
  ]
};

/* ═══════════════════════════════════════════════════════════════
 * MAIN ENTRY POINT
 * ═══════════════════════════════════════════════════════════════ */

/**
 * 🚀 Main entry point: Build all tracker sheets from GRANOT DATA
 * 
 * Pipeline steps:
 *   1. Load scheduled reps
 *   2. Load and filter GRANOT leads
 *   3. Build outreach indexes from RC logs
 *   4. Write each tracker sheet
 *   5. Log summary
 */
function SSCCP_buildTrackersFromGranotData() {
  var startTime = new Date();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summary = {
    startTime: startTime,
    scheduledReps: 0,
    totalLeads: 0,
    filteredLeads: 0,
    callLogRows: 0,
    smsLogRows: 0,
    outboundCalls: 0,
    outboundSms: 0,
    trackers: {}
  };
  
  try {
    Logger.log("══════════════════════════════════════════════════");
    Logger.log("🚢 SSCCP Tracker Builder - Starting Pipeline");
    Logger.log("══════════════════════════════════════════════════");
    
    // ─────────────────────────────────────────────────────────────
    // STEP 1: Load scheduled reps
    // ─────────────────────────────────────────────────────────────
    Logger.log("📋 Step 1: Loading scheduled reps...");
    var scheduledReps = SSCCP_TB_readScheduledReps_(ss);
    summary.scheduledReps = scheduledReps.size;
    Logger.log("   ✓ Found " + scheduledReps.size + " scheduled reps");
    
    if (scheduledReps.size === 0) {
      throw new Error("No scheduled reps found in SCHEDULED REPS sheet");
    }
    
    // ─────────────────────────────────────────────────────────────
    // STEP 2: Load GRANOT DATA leads (filtered by scheduled reps)
    // ─────────────────────────────────────────────────────────────
    Logger.log("📊 Step 2: Loading GRANOT DATA leads...");
    var leadsResult = SSCCP_TB_loadGranotLeads_(ss, scheduledReps);
    summary.totalLeads = leadsResult.totalLeads;
    summary.filteredLeads = leadsResult.leads.length;
    Logger.log("   ✓ Total leads: " + leadsResult.totalLeads);
    Logger.log("   ✓ Filtered leads (scheduled reps only): " + leadsResult.leads.length);
    
    // ─────────────────────────────────────────────────────────────
    // STEP 3: Build outreach indexes from RC logs (TODAY only)
    // ─────────────────────────────────────────────────────────────
    Logger.log("📞 Step 3: Building outreach indexes...");
    
    var callResult = SSCCP_TB_buildCallCountsIndex_(ss);
    summary.callLogRows = callResult.totalRows;
    summary.outboundCalls = callResult.outboundCount;
    Logger.log("   ✓ Call log rows: " + callResult.totalRows);
    Logger.log("   ✓ Outbound calls indexed: " + callResult.outboundCount);
    Logger.log("   ✓ Unique phones with calls: " + callResult.index.size);
    
    var smsResult = SSCCP_TB_buildSmsCountsIndex_(ss);
    summary.smsLogRows = smsResult.totalRows;
    summary.outboundSms = smsResult.outboundCount;
    Logger.log("   ✓ SMS log rows: " + smsResult.totalRows);
    Logger.log("   ✓ Outbound SMS indexed: " + smsResult.outboundCount);
    Logger.log("   ✓ Unique phones with SMS: " + smsResult.index.size);
    
    // ─────────────────────────────────────────────────────────────
    // STEP 4: Enrich leads with outreach counts
    // ─────────────────────────────────────────────────────────────
    Logger.log("🔗 Step 4: Enriching leads with outreach counts...");
    SSCCP_TB_enrichLeadsWithOutreach_(leadsResult.leads, callResult.index, smsResult.index);
    Logger.log("   ✓ Enrichment complete");
    
    // ─────────────────────────────────────────────────────────────
    // STEP 5: Write each tracker sheet
    // ─────────────────────────────────────────────────────────────
    Logger.log("📝 Step 5: Writing tracker sheets...");
    
    // SMS TRACKER
    var smsTrackerResult = SSCCP_TB_writeTrackerSheet_(
      ss, SSCCP_TB_CONFIG.SHEETS.SMS_TRACKER, leadsResult.leads, "SMS"
    );
    summary.trackers.smsTracker = smsTrackerResult;
    Logger.log("   ✓ SMS TRACKER: " + smsTrackerResult.rowsWritten + " leads (priority=" + 
               smsTrackerResult.priority + ", req=" + smsTrackerResult.requiredCount + ")");
    
    // CALL & VOICEMAIL TRACKER
    var callTrackerResult = SSCCP_TB_writeTrackerSheet_(
      ss, SSCCP_TB_CONFIG.SHEETS.CALL_TRACKER, leadsResult.leads, "CALL"
    );
    summary.trackers.callTracker = callTrackerResult;
    Logger.log("   ✓ CALL & VOICEMAIL TRACKER: " + callTrackerResult.rowsWritten + " leads");
    
    // PRIORITY 1 CALL & VOICEMAIL TRACKER
    var p1TrackerResult = SSCCP_TB_writeTrackerSheet_(
      ss, SSCCP_TB_CONFIG.SHEETS.PRIORITY_1_TRACKER, leadsResult.leads, "CALL"
    );
    summary.trackers.priority1Tracker = p1TrackerResult;
    Logger.log("   ✓ PRIORITY 1 TRACKER: " + p1TrackerResult.rowsWritten + " leads");
    
    // PRIORITY 3,5 CALL & VOICEMAIL TRACKER (reads priorities from A8)
    var p35TrackerResult = SSCCP_TB_writeTrackerSheet_(
      ss, SSCCP_TB_CONFIG.SHEETS.PRIORITY_35_TRACKER, leadsResult.leads, "CALL_MULTI"
    );
    summary.trackers.priority35Tracker = p35TrackerResult;
    Logger.log("   ✓ PRIORITY 3,5 TRACKER: " + p35TrackerResult.rowsWritten + " leads");
    
    // ─────────────────────────────────────────────────────────────
    // STEP 6: Log summary
    // ─────────────────────────────────────────────────────────────
    var endTime = new Date();
    var duration = (endTime - startTime) / 1000;
    summary.duration = duration;
    summary.endTime = endTime;
    summary.success = true;
    
    Logger.log("══════════════════════════════════════════════════");
    Logger.log("✅ SSCCP Tracker Builder - COMPLETE");
    Logger.log("   Duration: " + duration.toFixed(2) + " seconds");
    var totalWritten = smsTrackerResult.rowsWritten + callTrackerResult.rowsWritten + 
                       p1TrackerResult.rowsWritten + p35TrackerResult.rowsWritten;
    Logger.log("   Total leads written: " + totalWritten);
    Logger.log("══════════════════════════════════════════════════");
    
    // Write to Run_Log if exists
    SSCCP_TB_writeRunLog_(ss, summary);
    
    // Toast notification
    ss.toast(
      "✅ Trackers updated: SMS=" + smsTrackerResult.rowsWritten + 
      ", Call=" + callTrackerResult.rowsWritten +
      ", P1=" + p1TrackerResult.rowsWritten +
      ", P3/5=" + p35TrackerResult.rowsWritten +
      " (" + duration.toFixed(1) + "s)",
      "SSCCP Tracker Builder", 10
    );
    
    return summary;
    
  } catch (error) {
    Logger.log("❌ ERROR: " + error.message);
    Logger.log("   Stack: " + error.stack);
    
    summary.success = false;
    summary.error = error.message;
    summary.endTime = new Date();
    summary.duration = (summary.endTime - startTime) / 1000;
    
    SSCCP_TB_writeRunLog_(ss, summary);
    
    SpreadsheetApp.getUi().alert(
      "❌ Tracker Builder Error",
      "Error: " + error.message + "\n\nCheck Logger for details.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    throw error;
  }
}

/* ═══════════════════════════════════════════════════════════════
 * MODULAR ENTRY POINTS (for sidebar/future use)
 * ═══════════════════════════════════════════════════════════════ */

/**
 * Get leads not meeting SMS outreach threshold
 * @param {number} requiredCount - Minimum SMS required
 * @param {number|Array} priorities - Priority or array of priorities
 * @returns {Array} Array of lead objects
 */
function SSCCP_TB_getLeadsNotMeetingSmsThreshold(requiredCount, priorities) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scheduledReps = SSCCP_TB_readScheduledReps_(ss);
  var leadsResult = SSCCP_TB_loadGranotLeads_(ss, scheduledReps);
  var smsResult = SSCCP_TB_buildSmsCountsIndex_(ss);
  
  SSCCP_TB_enrichLeadsWithOutreach_(leadsResult.leads, new Map(), smsResult.index);
  
  var prioritySet = new Set(Array.isArray(priorities) ? priorities : [priorities]);
  
  return leadsResult.leads.filter(function(lead) {
    return prioritySet.has(lead.priority) && lead.outboundSmsToday < requiredCount;
  });
}

/**
 * Get leads not meeting call outreach threshold
 * @param {number} requiredCount - Minimum calls required
 * @param {number|Array} priorities - Priority or array of priorities
 * @returns {Array} Array of lead objects
 */
function SSCCP_TB_getLeadsNotMeetingCallThreshold(requiredCount, priorities) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scheduledReps = SSCCP_TB_readScheduledReps_(ss);
  var leadsResult = SSCCP_TB_loadGranotLeads_(ss, scheduledReps);
  var callResult = SSCCP_TB_buildCallCountsIndex_(ss);
  
  SSCCP_TB_enrichLeadsWithOutreach_(leadsResult.leads, callResult.index, new Map());
  
  var prioritySet = new Set(Array.isArray(priorities) ? priorities : [priorities]);
  
  return leadsResult.leads.filter(function(lead) {
    return prioritySet.has(lead.priority) && lead.outboundCallsToday < requiredCount;
  });
}

/* ═══════════════════════════════════════════════════════════════
 * CORE DATA LOADING FUNCTIONS
 * ═══════════════════════════════════════════════════════════════ */

/**
 * Read scheduled reps from SCHEDULED REPS sheet
 * @param {Spreadsheet} ss - Active spreadsheet
 * @returns {Set} Set of normalized usernames (uppercase, trimmed)
 */
function SSCCP_TB_readScheduledReps_(ss) {
  var sheet = ss.getSheetByName(SSCCP_TB_CONFIG.SHEETS.SCHEDULED_REPS);
  if (!sheet) {
    throw new Error("Sheet not found: " + SSCCP_TB_CONFIG.SHEETS.SCHEDULED_REPS);
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return new Set();
  }
  
  // Read column A from row 2 onwards (skip header)
  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var reps = new Set();
  
  for (var i = 0; i < data.length; i++) {
    var rep = SSCCP_TB_normalizeUsername_(data[i][0]);
    if (rep) {
      reps.add(rep);
    }
  }
  
  return reps;
}

/**
 * Load leads from GRANOT DATA, filtered by scheduled reps
 * @param {Spreadsheet} ss - Active spreadsheet
 * @param {Set} scheduledReps - Set of normalized usernames
 * @returns {Object} { leads: array of lead objects, totalLeads: count }
 */
function SSCCP_TB_loadGranotLeads_(ss, scheduledReps) {
  var sheet = ss.getSheetByName(SSCCP_TB_CONFIG.SHEETS.GRANOT_DATA);
  if (!sheet) {
    throw new Error("Sheet not found: " + SSCCP_TB_CONFIG.SHEETS.GRANOT_DATA);
  }
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  if (lastRow < 2) {
    return { leads: [], totalLeads: 0 };
  }
  
  // Read all data at once (row 2 onwards, skip header)
  var data = sheet.getRange(2, 1, lastRow - 1, Math.max(lastCol, 24)).getValues();
  var cols = SSCCP_TB_CONFIG.GRANOT_COLS;
  
  var leads = [];
  var totalLeads = data.length;
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var user = SSCCP_TB_normalizeUsername_(row[cols.USER]);
    
    // Skip if user not in scheduled reps
    if (!user || !scheduledReps.has(user)) {
      continue;
    }
    
    // Normalize phone numbers
    var phone1 = SSCCP_TB_normalizePhone10_(row[cols.PHONE1]);
    var phone2 = SSCCP_TB_normalizePhone10_(row[cols.PHONE2]);
    
    // Skip if no valid phone
    if (!phone1 && !phone2) {
      continue;
    }
    
    // Parse priority as integer
    var priority = SSCCP_TB_parseIntSafe_(row[cols.PRIORITY], -1);
    
    // Build lead object
    var lead = {
      // Core identifiers
      jobNo: row[cols.JOB_NO],
      user: user,
      priority: priority,
      
      // Phone numbers (normalized 10-digit)
      phone1: phone1,
      phone2: phone2,
      phones: [],
      
      // For display: use phone1 if available, else phone2
      displayPhone: phone1 || phone2,
      
      // Contact info
      firstName: row[cols.FIRST_NAME] || "",
      lastName: row[cols.LAST_NAME] || "",
      email: row[cols.EMAIL] || "",
      
      // Dates
      moveDate: row[cols.MOVE_DATE],
      openDate: row[cols.OPEN_DATE],
      openTime: row[cols.OPEN_TIME],
      followUp: row[cols.FOLLOW_UP],
      
      // Location
      fromAddress: row[cols.FROM_ADDRESS] || "",
      fromCity: row[cols.FROM_CITY] || "",
      fromState: row[cols.FROM_STATE] || "",
      fromZip: row[cols.FROM_ZIP] || "",
      toAddress: row[cols.TO_ADDRESS] || "",
      toCity: row[cols.TO_CITY] || "",
      toState: row[cols.TO_STATE] || "",
      toZip: row[cols.TO_ZIP] || "",
      
      // Other
      cf: row[cols.CF] || "",
      source: row[cols.SOURCE] || "",
      refNo: row[cols.REF_NO] || "",
      dept: row[cols.DEPT] || "",
      
      // Outreach counts (to be enriched later)
      outboundCallsToday: 0,
      outboundSmsToday: 0
    };
    
    // Build phones array for matching
    if (phone1) lead.phones.push(phone1);
    if (phone2 && phone2 !== phone1) lead.phones.push(phone2);
    
    leads.push(lead);
  }
  
  return {
    leads: leads,
    totalLeads: totalLeads
  };
}

/* ═══════════════════════════════════════════════════════════════
 * OUTREACH INDEX BUILDERS
 * ═══════════════════════════════════════════════════════════════ */

/**
 * Build call counts index from RC CALL LOG (outbound only)
 * @param {Spreadsheet} ss - Active spreadsheet
 * @returns {Object} { index: Map(phone10 -> count), totalRows, outboundCount }
 */
function SSCCP_TB_buildCallCountsIndex_(ss) {
  var sheet = ss.getSheetByName(SSCCP_TB_CONFIG.SHEETS.RC_CALL_LOG);
  if (!sheet) {
    Logger.log("   ⚠️ RC CALL LOG sheet not found, returning empty index");
    return { index: new Map(), totalRows: 0, outboundCount: 0 };
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { index: new Map(), totalRows: 0, outboundCount: 0 };
  }
  
  // Read all data at once (columns A-J)
  var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var cols = SSCCP_TB_CONFIG.CALL_LOG_COLS;
  
  var index = new Map();
  var outboundCount = 0;
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var direction = String(row[cols.DIRECTION] || "").toUpperCase().trim();
    
    // Only count OUTBOUND calls
    if (!SSCCP_TB_isOutbound_(direction)) {
      continue;
    }
    
    // Normalize phone number
    var phone = SSCCP_TB_normalizePhone10_(row[cols.PHONE]);
    if (!phone) {
      continue;
    }
    
    // Increment count
    outboundCount++;
    var currentCount = index.get(phone) || 0;
    index.set(phone, currentCount + 1);
  }
  
  return {
    index: index,
    totalRows: data.length,
    outboundCount: outboundCount
  };
}

/**
 * Build SMS counts index from RC SMS LOG (outbound only)
 * @param {Spreadsheet} ss - Active spreadsheet
 * @returns {Object} { index: Map(phone10 -> count), totalRows, outboundCount, failedCount }
 */
function SSCCP_TB_buildSmsCountsIndex_(ss) {
  var sheet = ss.getSheetByName(SSCCP_TB_CONFIG.SHEETS.RC_SMS_LOG);
  if (!sheet) {
    Logger.log("   ⚠️ RC SMS LOG sheet not found, returning empty index");
    return { index: new Map(), totalRows: 0, outboundCount: 0, failedCount: 0 };
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { index: new Map(), totalRows: 0, outboundCount: 0, failedCount: 0 };
  }
  
  // Read all data at once (columns A-J)
  var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var cols = SSCCP_TB_CONFIG.SMS_LOG_COLS;
  
  var index = new Map();
  var outboundCount = 0;
  var failedCount = 0;
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var direction = String(row[cols.DIRECTION] || "").toUpperCase().trim();
    
    // Only count OUTBOUND SMS
    if (!SSCCP_TB_isOutbound_(direction)) {
      continue;
    }
    
    // For outbound, match against Recipient Number (the lead's phone)
    var phone = SSCCP_TB_normalizePhone10_(row[cols.RECIPIENT_NUM]);
    if (!phone) {
      continue;
    }
    
    // Track failed messages separately (for future diagnostics)
    var status = String(row[cols.MSG_STATUS] || "").toUpperCase();
    if (status.indexOf("FAIL") >= 0 || status.indexOf("ERROR") >= 0) {
      failedCount++;
    }
    
    // Count all outbound SMS as attempts (configurable behavior)
    outboundCount++;
    var currentCount = index.get(phone) || 0;
    index.set(phone, currentCount + 1);
  }
  
  return {
    index: index,
    totalRows: data.length,
    outboundCount: outboundCount,
    failedCount: failedCount
  };
}

/**
 * Enrich leads with outreach counts from indexes
 * @param {Array} leads - Array of lead objects
 * @param {Map} callIndex - Map of phone -> call count
 * @param {Map} smsIndex - Map of phone -> SMS count
 */
function SSCCP_TB_enrichLeadsWithOutreach_(leads, callIndex, smsIndex) {
  for (var i = 0; i < leads.length; i++) {
    var lead = leads[i];
    var callCount = 0;
    var smsCount = 0;
    
    // Check all phones for this lead
    for (var j = 0; j < lead.phones.length; j++) {
      var phone = lead.phones[j];
      callCount += callIndex.get(phone) || 0;
      smsCount += smsIndex.get(phone) || 0;
    }
    
    lead.outboundCallsToday = callCount;
    lead.outboundSmsToday = smsCount;
  }
}

/* ═══════════════════════════════════════════════════════════════
 * TRACKER SHEET WRITER
 * ═══════════════════════════════════════════════════════════════ */

/**
 * Read tracker parameters from sheet cells
 * @param {Sheet} sheet - The tracker sheet
 * @param {string} trackerType - "SMS", "CALL", or "CALL_MULTI"
 * @returns {Object} { priorities: array, requiredCount: number }
 */
function SSCCP_TB_readTrackerParams_(sheet, trackerType) {
  var result = {
    priorities: [],
    requiredCount: 1
  };
  
  try {
    if (trackerType === "CALL_MULTI") {
      // PRIORITY 3,5 (or any multi-priority) TRACKER: A8 = comma-separated priorities, B8 = required count
      var priorityVal = sheet.getRange(8, 1).getValue();
      var reqVal = sheet.getRange(8, 2).getValue();
      
      // Parse comma-separated priorities from A8 (e.g., "1, 3" or "3, 5")
      result.priorities = SSCCP_TB_parseMultiplePriorities_(priorityVal);
      
      // B8 = 0 means "show leads with 0 calls" (i.e., < 1)
      var parsed = SSCCP_TB_parseIntSafe_(reqVal, 0);
      result.requiredCount = (parsed <= 0) ? 1 : parsed + 1;  // Convert "< X" to threshold
      return result;
    }
    
    if (trackerType === "SMS") {
      // SMS TRACKER: B7 = priority, B6 = required count
      var priorityVal = sheet.getRange(7, 2).getValue();
      var reqVal = sheet.getRange(6, 2).getValue();
      
      // Parse priority (could be single or comma-separated)
      result.priorities = SSCCP_TB_parseMultiplePriorities_(priorityVal);
      if (result.priorities.length === 0) result.priorities = [0];
      
      var parsed = SSCCP_TB_parseIntSafe_(reqVal, 0);
      result.requiredCount = (parsed <= 0) ? 1 : parsed + 1;
      return result;
    }
    
    if (trackerType === "CALL") {
      // CALL TRACKER / PRIORITY 1: A8 = priority (can be comma-separated), B8 = required count
      var priorityVal = sheet.getRange(8, 1).getValue();
      var reqVal = sheet.getRange(8, 2).getValue();
      
      // Parse priority (could be single or comma-separated like "1, 3")
      result.priorities = SSCCP_TB_parseMultiplePriorities_(priorityVal);
      if (result.priorities.length === 0) result.priorities = [0];
      
      // B8 = 0 means "show leads with 0 calls" (i.e., < 1)
      var parsed = SSCCP_TB_parseIntSafe_(reqVal, 0);
      result.requiredCount = (parsed <= 0) ? 1 : parsed + 1;
      return result;
    }
  } catch (e) {
    Logger.log("   ⚠️ Error reading params: " + e.message + ", using defaults");
  }
  
  return result;
}

/**
 * Parse comma-separated priority values (e.g., "1, 3" -> [1, 3])
 * @param {*} value - Priority value(s) from cell
 * @returns {Array} Array of integer priorities
 */
function SSCCP_TB_parseMultiplePriorities_(value) {
  if (value === null || value === undefined || value === "") {
    return [];
  }
  
  var str = String(value);
  var parts = str.split(/[,\s]+/);  // Split on comma and/or whitespace
  var priorities = [];
  
  for (var i = 0; i < parts.length; i++) {
    var trimmed = parts[i].trim();
    if (trimmed) {
      var parsed = parseInt(trimmed, 10);
      if (!isNaN(parsed)) {
        priorities.push(parsed);
      }
    }
  }
  
  return priorities;
}

/**
 * Write filtered leads to a tracker sheet
 * @param {Spreadsheet} ss - Active spreadsheet
 * @param {string} sheetName - Name of the tracker sheet
 * @param {Array} leads - Array of all lead objects
 * @param {string} trackerType - "SMS", "CALL", or "CALL_35"
 * @returns {Object} { rowsWritten, priority, requiredCount }
 */
function SSCCP_TB_writeTrackerSheet_(ss, sheetName, leads, trackerType) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("   ⚠️ Sheet not found: " + sheetName + " - skipping");
    return { rowsWritten: 0, priority: null, requiredCount: null, skipped: true };
  }
  
  // Read tracker parameters from sheet cells
  var params = SSCCP_TB_readTrackerParams_(sheet, trackerType);
  
  // Filter leads based on tracker type and parameters
  var filteredLeads = SSCCP_TB_filterLeadsForTracker_(leads, params, trackerType);
  
  // Get layout configuration
  var layout = SSCCP_TB_CONFIG.TRACKER_LAYOUT;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  // Read existing headers from row 2, starting at column F
  var headerRange = sheet.getRange(layout.HEADER_ROW, layout.DATA_START_COL, 1, Math.max(lastCol - layout.DATA_START_COL + 1, 22));
  var headerValues = headerRange.getValues()[0];
  
  // Build header map: normalize header names and track positions
  var headers = [];
  var headerCount = 0;
  for (var i = 0; i < headerValues.length; i++) {
    var h = String(headerValues[i] || "").toUpperCase().trim().replace(/[\s_]+/g, "_");
    if (h) {
      headers.push(h);
      headerCount = i + 1; // Track the last non-empty header position
    } else {
      headers.push(""); // Keep empty slots to maintain positions
    }
  }
  
  // If no headers found, use default headers
  if (headerCount === 0) {
    headers = SSCCP_TB_getDefaultHeaders_();
    headerCount = headers.length;
    // Write default headers to sheet
    var defaultHeaderRange = sheet.getRange(layout.HEADER_ROW, layout.DATA_START_COL, 1, headers.length);
    defaultHeaderRange.setValues([SSCCP_TB_getDisplayHeaders_()]);
  }
  
  // Build output rows based on header positions
  var outputRows = SSCCP_TB_buildTrackerRowsByHeaders_(filteredLeads, headers, headerCount, trackerType);
  
  // Calculate columns to clear
  var totalCols = Math.max(headerCount, outputRows.length > 0 ? outputRows[0].length : 0, 22);
  
  // Clear existing data (preserve rows 1-2)
  if (lastRow >= layout.DATA_START_ROW) {
    var clearRange = sheet.getRange(
      layout.DATA_START_ROW,
      layout.DATA_START_COL,
      lastRow - layout.DATA_START_ROW + 1,
      totalCols
    );
    clearRange.clearContent();
  }
  
  // Write new data if we have rows
  if (outputRows.length > 0) {
    var outputRange = sheet.getRange(
      layout.DATA_START_ROW,
      layout.DATA_START_COL,
      outputRows.length,
      outputRows[0].length
    );
    outputRange.setValues(outputRows);
  }
  
  return {
    rowsWritten: outputRows.length,
    priority: params.priorities,
    requiredCount: params.requiredCount
  };
}

/**
 * Get default headers (normalized format for matching)
 */
function SSCCP_TB_getDefaultHeaders_() {
  return [
    "MOVE_DATE", "PHONE_NUMBER", "COUNT", "JOB_#", "USER", "CF", "SOURCE", "EXCLUDE",
    "EMAIL", "FIRST_NAME", "LAST_NAME", 
    "FROM_CITY", "FROM_STATE", "FROM_ZIP",
    "TO_CITY", "TO_STATE", "TO_ZIP", 
    "OPEN_DATE", "OPEN_TIME", "REF_NO", "DEPT", "PRIORITY"
  ];
}

/**
 * Get display headers (human-readable format for writing to sheet)
 */
function SSCCP_TB_getDisplayHeaders_() {
  return [
    "MOVE DATE", "PHONE NUMBER", "Count", "JOB #", "USER", "CF", "SOURCE", "EXCLUDE",
    "EMAIL", "FIRST_NAME", "LAST_NAME", 
    "FROM_CITY", "FROM_STATE", "FROM_ZIP",
    "TO_CITY", "TO_STATE", "TO_ZIP", 
    "OPEN_DATE", "OPEN_TIME", "REF_NO", "DEPT", "PRIORITY"
  ];
}

/**
 * Build output rows based on header positions (flexible column ordering)
 * @param {Array} leads - Filtered leads
 * @param {Array} headers - Array of normalized header names in column order
 * @param {number} headerCount - Number of columns to output
 * @param {string} trackerType - "SMS", "CALL", or "CALL_MULTI"
 * @returns {Array} 2D array of output rows
 */
function SSCCP_TB_buildTrackerRowsByHeaders_(leads, headers, headerCount, trackerType) {
  var rows = [];
  
  for (var i = 0; i < leads.length; i++) {
    var lead = leads[i];
    
    // Build a data map for this lead with all possible field names
    var leadData = SSCCP_TB_buildLeadDataMap_(lead, trackerType);
    
    // Build row based on header order
    var row = [];
    for (var j = 0; j < headerCount; j++) {
      var header = headers[j];
      if (header && leadData.hasOwnProperty(header)) {
        row.push(leadData[header]);
      } else {
        row.push(""); // Empty for unknown headers
      }
    }
    
    rows.push(row);
  }
  
  return rows;
}

/**
 * Build a data map for a lead with all possible header variations
 * @param {Object} lead - Lead object
 * @param {string} trackerType - Tracker type for count field
 * @returns {Object} Map of normalized header name -> value
 */
function SSCCP_TB_buildLeadDataMap_(lead, trackerType) {
  var count = (trackerType === "SMS") ? lead.outboundSmsToday : lead.outboundCallsToday;
  var displayPhone = SSCCP_TB_formatPhoneDisplay_(lead.displayPhone);
  
  // Map with multiple variations of header names for flexibility
  var data = {
    // Core fields - multiple variations for each
    "MOVE_DATE": lead.moveDate,
    "MOVEDATE": lead.moveDate,
    "MOVE": lead.moveDate,
    
    "PHONE_NUMBER": displayPhone,
    "PHONENUMBER": displayPhone,
    "PHONE": displayPhone,
    "PHONE_#": displayPhone,
    "PHONE#": displayPhone,
    
    "COUNT": count,
    "CNT": count,
    "CALL_COUNT": count,
    "SMS_COUNT": count,
    "TEXT_COUNT": count,
    
    "JOB_#": lead.jobNo,
    "JOB#": lead.jobNo,
    "JOB_NO": lead.jobNo,
    "JOBNO": lead.jobNo,
    "JOB": lead.jobNo,
    "JOB_NUMBER": lead.jobNo,
    
    "USER": lead.user,
    "REP": lead.user,
    "SALES_REP": lead.user,
    
    "CF": lead.cf,
    
    "SOURCE": lead.source,
    "LEAD_SOURCE": lead.source,
    
    "EXCLUDE": false,
    
    // Extra fields
    "EMAIL": lead.email,
    "E-MAIL": lead.email,
    
    "FIRST_NAME": lead.firstName,
    "FIRSTNAME": lead.firstName,
    "FIRST": lead.firstName,
    
    "LAST_NAME": lead.lastName,
    "LASTNAME": lead.lastName,
    "LAST": lead.lastName,
    
    "NAME": (lead.firstName + " " + lead.lastName).trim(),
    "FULL_NAME": (lead.firstName + " " + lead.lastName).trim(),
    
    "FROM_CITY": lead.fromCity,
    "FROMCITY": lead.fromCity,
    "ORIGIN_CITY": lead.fromCity,
    
    "FROM_STATE": lead.fromState,
    "FROMSTATE": lead.fromState,
    "ORIGIN_STATE": lead.fromState,
    
    "FROM_ZIP": lead.fromZip,
    "FROMZIP": lead.fromZip,
    "ORIGIN_ZIP": lead.fromZip,
    
    "TO_CITY": lead.toCity,
    "TOCITY": lead.toCity,
    "DEST_CITY": lead.toCity,
    "DESTINATION_CITY": lead.toCity,
    
    "TO_STATE": lead.toState,
    "TOSTATE": lead.toState,
    "DEST_STATE": lead.toState,
    "DESTINATION_STATE": lead.toState,
    
    "TO_ZIP": lead.toZip,
    "TOZIP": lead.toZip,
    "DEST_ZIP": lead.toZip,
    "DESTINATION_ZIP": lead.toZip,
    
    "OPEN_DATE": lead.openDate,
    "OPENDATE": lead.openDate,
    "CREATED_DATE": lead.openDate,
    
    "OPEN_TIME": lead.openTime,
    "OPENTIME": lead.openTime,
    "CREATED_TIME": lead.openTime,
    
    "REF_NO": lead.refNo,
    "REFNO": lead.refNo,
    "REF_#": lead.refNo,
    "REF#": lead.refNo,
    "REFERENCE": lead.refNo,
    "REFERENCE_NO": lead.refNo,
    
    "DEPT": lead.dept,
    "DEPARTMENT": lead.dept,
    
    "PRIORITY": lead.priority,
    "PRI": lead.priority,
    "P": lead.priority
  };
  
  return data;
}

/**
 * Filter leads based on tracker criteria
 * @param {Array} leads - All leads
 * @param {Object} params - { priorities: array, requiredCount: number }
 * @param {string} trackerType - "SMS", "CALL", or "CALL_MULTI"
 * @returns {Array} Filtered leads
 */
function SSCCP_TB_filterLeadsForTracker_(leads, params, trackerType) {
  var prioritySet = new Set(params.priorities);
  var requiredCount = params.requiredCount;
  
  return leads.filter(function(lead) {
    // Check priority matches
    if (!prioritySet.has(lead.priority)) {
      return false;
    }
    
    // Check outreach threshold based on tracker type
    if (trackerType === "SMS") {
      return lead.outboundSmsToday < requiredCount;
    } else {
      // CALL or CALL_MULTI
      return lead.outboundCallsToday < requiredCount;
    }
  });
}

/**
 * Build output rows for tracker sheet
 * Core columns: MOVE DATE, PHONE NUMBER, Count, JOB #, USER, CF, SOURCE, EXCLUDE
 * Extra columns: email, first_name, last_name, from_city, from_state, from_zip,
 *                to_city, to_state, to_zip, open_date, open_time, ref_no, dept, priority
 * 
 * @param {Array} leads - Filtered leads
 * @param {string} trackerType - "SMS", "CALL", or "CALL_MULTI"
 * @returns {Array} 2D array of output rows
 */
function SSCCP_TB_buildTrackerRows_(leads, trackerType) {
  var rows = [];
  
  for (var i = 0; i < leads.length; i++) {
    var lead = leads[i];
    
    // Determine count to display based on tracker type
    var count = (trackerType === "SMS") ? lead.outboundSmsToday : lead.outboundCallsToday;
    
    // Format phone for display (add dashes)
    var displayPhone = SSCCP_TB_formatPhoneDisplay_(lead.displayPhone);
    
    // Core columns (F-M): MOVE DATE, PHONE NUMBER, Count, JOB #, USER, CF, SOURCE, EXCLUDE
    var row = [
      lead.moveDate,       // F - MOVE DATE
      displayPhone,        // G - PHONE NUMBER
      count,               // H - Count
      lead.jobNo,          // I - JOB #
      lead.user,           // J - USER
      lead.cf,             // K - CF
      lead.source,         // L - SOURCE
      false                // M - EXCLUDE (default false)
    ];
    
    // Extra columns (N+): Additional GRANOT fields
    row.push(lead.email);        // N
    row.push(lead.firstName);    // O
    row.push(lead.lastName);     // P
    row.push(lead.fromCity);     // Q
    row.push(lead.fromState);    // R
    row.push(lead.fromZip);      // S
    row.push(lead.toCity);       // T
    row.push(lead.toState);      // U
    row.push(lead.toZip);        // V
    row.push(lead.openDate);     // W
    row.push(lead.openTime);     // X
    row.push(lead.refNo);        // Y
    row.push(lead.dept);         // Z
    row.push(lead.priority);     // AA
    
    rows.push(row);
  }
  
  return rows;
}

/* ═══════════════════════════════════════════════════════════════
 * UTILITY FUNCTIONS
 * ═══════════════════════════════════════════════════════════════ */

/**
 * Normalize phone number to 10 digits
 * Strips all non-digits, takes last 10 digits
 * @param {*} value - Phone number (any format)
 * @returns {string|null} 10-digit phone or null
 */
function SSCCP_TB_normalizePhone10_(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  
  // Convert to string and extract digits only
  var digits = String(value).replace(/\D/g, "");
  
  // Must have at least 10 digits
  if (digits.length < 10) {
    return null;
  }
  
  // Take last 10 digits (handles +1 prefix)
  return digits.slice(-10);
}

/**
 * Format phone number for display (XXX-XXX-XXXX)
 * @param {string} phone10 - 10-digit phone
 * @returns {string} Formatted phone or original
 */
function SSCCP_TB_formatPhoneDisplay_(phone10) {
  if (!phone10 || phone10.length !== 10) {
    return phone10 || "";
  }
  return phone10.slice(0, 3) + "-" + phone10.slice(3, 6) + "-" + phone10.slice(6);
}

/**
 * Normalize username (uppercase, trimmed)
 * @param {*} value - Username
 * @returns {string|null} Normalized username or null
 */
function SSCCP_TB_normalizeUsername_(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  var normalized = String(value).toUpperCase().trim();
  return normalized || null;
}

/**
 * Check if direction indicates outbound
 * @param {string} direction - Direction value (already uppercase)
 * @returns {boolean} True if outbound
 */
function SSCCP_TB_isOutbound_(direction) {
  if (!direction) return false;
  // Match "OUTBOUND", "OUT", etc.
  return direction.indexOf("OUT") === 0 || direction === "OUTBOUND";
}

/**
 * Safely parse integer with default
 * @param {*} value - Value to parse
 * @param {number} defaultVal - Default if invalid
 * @returns {number} Parsed integer or default
 */
function SSCCP_TB_parseIntSafe_(value, defaultVal) {
  if (value === null || value === undefined || value === "") {
    return defaultVal;
  }
  var parsed = parseInt(value, 10);
  return isNaN(parsed) ? defaultVal : parsed;
}

/* ═══════════════════════════════════════════════════════════════
 * LOGGING FUNCTIONS
 * ═══════════════════════════════════════════════════════════════ */

/**
 * Write run summary to Run_Log sheet if it exists
 * @param {Spreadsheet} ss - Active spreadsheet
 * @param {Object} summary - Run summary object
 */
function SSCCP_TB_writeRunLog_(ss, summary) {
  try {
    var sheet = ss.getSheetByName(SSCCP_TB_CONFIG.SHEETS.RUN_LOG);
    if (!sheet) {
      return; // Don't create if doesn't exist
    }
    
    var timestamp = Utilities.formatDate(summary.startTime, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    var status = summary.success ? "SUCCESS" : "ERROR";
    var details = summary.success
      ? "SMS:" + (summary.trackers.smsTracker ? summary.trackers.smsTracker.rowsWritten : 0) +
        " Call:" + (summary.trackers.callTracker ? summary.trackers.callTracker.rowsWritten : 0) +
        " P1:" + (summary.trackers.priority1Tracker ? summary.trackers.priority1Tracker.rowsWritten : 0) +
        " P3/5:" + (summary.trackers.priority35Tracker ? summary.trackers.priority35Tracker.rowsWritten : 0)
      : summary.error;
    
    // Append row to Run_Log
    sheet.appendRow([
      timestamp,
      "SSCCP_TrackerBuilder",
      status,
      summary.duration ? summary.duration.toFixed(2) + "s" : "",
      details,
      summary.filteredLeads || 0,
      summary.outboundCalls || 0,
      summary.outboundSms || 0
    ]);
    
  } catch (e) {
    Logger.log("   ⚠️ Could not write to Run_Log: " + e.message);
  }
}

/* ═══════════════════════════════════════════════════════════════
 * DIAGNOSTIC / TEST FUNCTIONS
 * ═══════════════════════════════════════════════════════════════ */

/**
 * Test function: Show outreach stats for a specific phone
 * @param {string} phone - Phone number to test
 */
function SSCCP_TB_testPhoneLookup(phone) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var phone10 = SSCCP_TB_normalizePhone10_(phone);
  
  if (!phone10) {
    SpreadsheetApp.getUi().alert("Invalid phone number: " + phone);
    return;
  }
  
  var callResult = SSCCP_TB_buildCallCountsIndex_(ss);
  var smsResult = SSCCP_TB_buildSmsCountsIndex_(ss);
  
  var callCount = callResult.index.get(phone10) || 0;
  var smsCount = smsResult.index.get(phone10) || 0;
  
  SpreadsheetApp.getUi().alert(
    "📞 Phone Lookup: " + phone10 + "\n\n" +
    "Outbound Calls Today: " + callCount + "\n" +
    "Outbound SMS Today: " + smsCount
  );
}

/**
 * Test function: Preview tracker output without writing
 */
function SSCCP_TB_previewTrackerCounts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scheduledReps = SSCCP_TB_readScheduledReps_(ss);
  var leadsResult = SSCCP_TB_loadGranotLeads_(ss, scheduledReps);
  var callResult = SSCCP_TB_buildCallCountsIndex_(ss);
  var smsResult = SSCCP_TB_buildSmsCountsIndex_(ss);
  
  SSCCP_TB_enrichLeadsWithOutreach_(leadsResult.leads, callResult.index, smsResult.index);
  
  // Count by priority
  var byPriority = {};
  var smsNeeded = { 0: 0 };
  var callNeeded = { 0: 0, 1: 0, 3: 0, 5: 0 };
  
  for (var i = 0; i < leadsResult.leads.length; i++) {
    var lead = leadsResult.leads[i];
    var p = lead.priority;
    byPriority[p] = (byPriority[p] || 0) + 1;
    
    if (p === 0 && lead.outboundSmsToday < 1) smsNeeded[0]++;
    if (p === 0 && lead.outboundCallsToday < 1) callNeeded[0]++;
    if (p === 1 && lead.outboundCallsToday < 1) callNeeded[1]++;
    if (p === 3 && lead.outboundCallsToday < 1) callNeeded[3]++;
    if (p === 5 && lead.outboundCallsToday < 1) callNeeded[5]++;
  }
  
  var msg = "📊 TRACKER PREVIEW (req=1)\n\n";
  msg += "Total Leads (scheduled reps): " + leadsResult.leads.length + "\n\n";
  msg += "By Priority:\n";
  for (var p in byPriority) {
    msg += "  P" + p + ": " + byPriority[p] + "\n";
  }
  msg += "\nNeed Outreach (count < 1):\n";
  msg += "  SMS TRACKER (P0): " + smsNeeded[0] + "\n";
  msg += "  CALL TRACKER (P0): " + callNeeded[0] + "\n";
  msg += "  PRIORITY 1 TRACKER: " + callNeeded[1] + "\n";
  msg += "  PRIORITY 3,5 TRACKER: " + (callNeeded[3] + callNeeded[5]) + "\n";
  
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * Diagnostic: Show RC log stats
 */
function SSCCP_TB_showRCLogStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var callResult = SSCCP_TB_buildCallCountsIndex_(ss);
  var smsResult = SSCCP_TB_buildSmsCountsIndex_(ss);
  
  var msg = "📊 RC LOG STATS (Today)\n\n";
  msg += "CALL LOG:\n";
  msg += "  Total rows: " + callResult.totalRows + "\n";
  msg += "  Outbound calls: " + callResult.outboundCount + "\n";
  msg += "  Unique phones: " + callResult.index.size + "\n\n";
  msg += "SMS LOG:\n";
  msg += "  Total rows: " + smsResult.totalRows + "\n";
  msg += "  Outbound SMS: " + smsResult.outboundCount + "\n";
  msg += "  Failed SMS: " + smsResult.failedCount + "\n";
  msg += "  Unique phones: " + smsResult.index.size + "\n";
  
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * Interactive test: Prompt for phone number and show lookup
 */
function SSCCP_TB_testPhoneLookupPrompt() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    "Phone Lookup Test",
    "Enter a phone number to check outreach counts:",
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    var phone = response.getResponseText();
    SSCCP_TB_testPhoneLookup(phone);
  }
}

/* ═══════════════════════════════════════════════════════════════
 * TRIGGER SETUP HELPER
 * ═══════════════════════════════════════════════════════════════ */

/**
 * Create a time-based trigger to run tracker builder every 10 minutes
 * during business hours (8 AM - 6 PM)
 * 
 * NOTE: Call this manually once to set up the trigger.
 * Do NOT auto-run this function.
 */
function SSCCP_TB_createTrigger() {
  // First, remove any existing triggers for this function
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "SSCCP_buildTrackersFromGranotData") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create new trigger - runs every 10 minutes
  ScriptApp.newTrigger("SSCCP_buildTrackersFromGranotData")
    .timeBased()
    .everyMinutes(10)
    .create();
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✅ Trigger created: SSCCP_buildTrackersFromGranotData will run every 10 minutes",
    "Trigger Setup", 10
  );
}

/**
 * Remove the time-based trigger for tracker builder
 */
function SSCCP_TB_removeTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "SSCCP_buildTrackersFromGranotData") {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✅ Removed " + removed + " trigger(s) for SSCCP_buildTrackersFromGranotData",
    "Trigger Removed", 5
  );
}

/* ═══════════════════════════════════════════════════════════════
 * MENU INTEGRATION - ADD THIS TO MainMenu.gs onOpen()
 * ═══════════════════════════════════════════════════════════════
 * 
 * Add the following submenu to an existing menu in MainMenu.gs:
 * 
 *   .addSubMenu(ui.createMenu("🔄 Tracker Builder")
 *     .addItem("▶ Build All Trackers", "SSCCP_buildTrackersFromGranotData")
 *     .addSeparator()
 *     .addItem("📊 Preview Counts", "SSCCP_TB_previewTrackerCounts")
 *     .addItem("📈 RC Log Stats", "SSCCP_TB_showRCLogStats")
 *     .addItem("🔍 Phone Lookup", "SSCCP_TB_testPhoneLookupPrompt")
 *     .addSeparator()
 *     .addItem("⏰ Create 10-min Trigger", "SSCCP_TB_createTrigger")
 *     .addItem("🛑 Remove Trigger", "SSCCP_TB_removeTrigger")
 *     .addSeparator()
 *     .addItem("🔄 Enable Auto-Run on Edit", "SSCCP_TB_createOnEditTrigger")
 *     .addItem("🛑 Disable Auto-Run on Edit", "SSCCP_TB_removeOnEditTrigger"))
 * 
 * ═══════════════════════════════════════════════════════════════ */

/* ═══════════════════════════════════════════════════════════════
 * ON-EDIT AUTO-RUN TRIGGER
 * 
 * Automatically runs the tracker when you change:
 *   - PRIORITY SELECT cell (A8 for call trackers, B7 for SMS)
 *   - CALL COUNT / TEXT COUNT cell (B8 for call trackers, B6 for SMS)
 * ═══════════════════════════════════════════════════════════════ */

/**
 * Installable onEdit trigger - detects parameter cell changes and auto-runs tracker
 * @param {Object} e - Edit event object
 */
function SSCCP_TB_onEditTrigger(e) {
  try {
    if (!e || !e.range) return;
    
    var sheet = e.range.getSheet();
    var sheetName = sheet.getName();
    var row = e.range.getRow();
    var col = e.range.getColumn();
    
    // Check if this is a tracker sheet we care about
    var trackerConfig = SSCCP_TB_getTrackerConfigBySheet_(sheetName);
    if (!trackerConfig) return;
    
    // Check if the edited cell is a parameter cell (priority or count)
    var isParamCell = false;
    
    if (trackerConfig.type === "SMS") {
      // SMS TRACKER: B7 = priority, B6 = count
      isParamCell = (row === 7 && col === 2) || (row === 6 && col === 2);
    } else {
      // CALL trackers: A8 = priority, B8 = count
      isParamCell = (row === 8 && (col === 1 || col === 2));
    }
    
    if (!isParamCell) return;
    
    // Show toast that we're updating
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast("🔄 Updating " + sheetName + "...", "Auto-Run", 3);
    
    // Run just this one tracker
    SSCCP_TB_runSingleTracker_(sheetName);
    
  } catch (err) {
    Logger.log("SSCCP_TB_onEditTrigger error: " + err.message);
  }
}

/**
 * Get tracker configuration by sheet name
 * @param {string} sheetName - Name of the sheet
 * @returns {Object|null} Tracker config or null if not a tracker sheet
 */
function SSCCP_TB_getTrackerConfigBySheet_(sheetName) {
  var configs = {
    "SMS TRACKER": { type: "SMS", key: "SMS_TRACKER" },
    "CALL & VOICEMAIL TRACKER": { type: "CALL", key: "CALL_TRACKER" },
    "PRIORITY 1 CALL & VOICEMAIL TRACKER": { type: "CALL", key: "PRIORITY_1_TRACKER" },
    "PRIORITY 3,5 CALL & VOICEMAIL TRACKER": { type: "CALL_MULTI", key: "PRIORITY_35_TRACKER" }
  };
  return configs[sheetName] || null;
}

/**
 * Run a single tracker sheet update (faster than running all 4)
 * @param {string} sheetName - Name of the tracker sheet to update
 */
function SSCCP_TB_runSingleTracker_(sheetName) {
  var startTime = new Date();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Get tracker config
    var config = SSCCP_TB_getTrackerConfigBySheet_(sheetName);
    if (!config) {
      Logger.log("Unknown tracker sheet: " + sheetName);
      return;
    }
    
    // Load data (cached if possible via script properties for performance)
    var scheduledReps = SSCCP_TB_readScheduledReps_(ss);
    var leadsResult = SSCCP_TB_loadGranotLeads_(ss, scheduledReps);
    var callResult = SSCCP_TB_buildCallCountsIndex_(ss);
    var smsResult = SSCCP_TB_buildSmsCountsIndex_(ss);
    
    // Enrich leads
    SSCCP_TB_enrichLeadsWithOutreach_(leadsResult.leads, callResult.index, smsResult.index);
    
    // Write just this tracker
    var result = SSCCP_TB_writeTrackerSheet_(ss, sheetName, leadsResult.leads, config.type);
    
    var duration = ((new Date()) - startTime) / 1000;
    
    // Show completion toast
    ss.toast(
      "✅ " + sheetName + ": " + result.rowsWritten + " leads (Priority: " + 
      result.priority.join(", ") + ", Req: <" + result.requiredCount + ") - " + 
      duration.toFixed(1) + "s",
      "Tracker Updated", 5
    );
    
    Logger.log("✅ Auto-updated " + sheetName + ": " + result.rowsWritten + " leads in " + duration.toFixed(2) + "s");
    
  } catch (err) {
    Logger.log("❌ Error updating " + sheetName + ": " + err.message);
    ss.toast("❌ Error: " + err.message, "Update Failed", 5);
  }
}

/**
 * Create installable onEdit trigger for auto-run
 * Run this once to enable the feature
 */
function SSCCP_TB_createOnEditTrigger() {
  // First remove any existing onEdit triggers for this function
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "SSCCP_TB_onEditTrigger") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create new installable onEdit trigger
  ScriptApp.newTrigger("SSCCP_TB_onEditTrigger")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✅ Auto-run enabled!\n\nChange PRIORITY SELECT or COUNT cells to auto-update that tracker.",
    "On-Edit Trigger Created", 8
  );
}

/**
 * Remove the onEdit trigger
 */
function SSCCP_TB_removeOnEditTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "SSCCP_TB_onEditTrigger") {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✅ Auto-run disabled. Removed " + removed + " trigger(s).",
    "On-Edit Trigger Removed", 5
  );
}

/* ═══════════════════════════════════════════════════════════════
 * MANUAL SINGLE-TRACKER FUNCTIONS (can also be called from menu)
 * ═══════════════════════════════════════════════════════════════ */

function SSCCP_TB_runSmsTracker() {
  SSCCP_TB_runSingleTracker_("SMS TRACKER");
}

function SSCCP_TB_runCallTracker() {
  SSCCP_TB_runSingleTracker_("CALL & VOICEMAIL TRACKER");
}

function SSCCP_TB_runPriority1Tracker() {
  SSCCP_TB_runSingleTracker_("PRIORITY 1 CALL & VOICEMAIL TRACKER");
}

function SSCCP_TB_runPriority35Tracker() {
  SSCCP_TB_runSingleTracker_("PRIORITY 3,5 CALL & VOICEMAIL TRACKER");
}
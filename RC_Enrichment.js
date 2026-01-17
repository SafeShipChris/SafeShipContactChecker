/**************************************************************
 * 🚢 SAFE SHIP CONTACT CHECKER — RC_Enrichment.gs v3.4
 * ═══════════════════════════════════════════════════════════
 * 
 * v3.4 CHANGES:
 * ─────────────────────────────────────────────────────────────
 * ✅ FIX: Show ALL SMS attempts including failed ones
 * ✅ FIX: Removed aggressive SMS filtering that hid "SendingFailed"
 * ✅ NEW: isFailed flag on SMS history entries for UI display
 * ✅ NEW: SMS history now shows red indicator for failed messages
 * 
 * THE PROBLEM WAS:
 *   - SMS with status "SendingFailed" were being filtered out
 *   - Rep needs to SEE that their SMS failed, not see "Never"
 *   - Now all attempts show, with failed ones marked
 * 
 **************************************************************/

/**
 * ═══════════════════════════════════════════════════════════════
 * CONFIGURATION
 * ═══════════════════════════════════════════════════════════════
 */
var RC_ENRICH_CONFIG = {
  SHEETS: {
    CALL_TODAY: "RC CALL LOG",
    CALL_YESTERDAY: "RC CALL LOG YESTERDAY",
    SMS_TODAY: "RC SMS LOG",
    SMS_YESTERDAY: "RC SMS LOG YESTERDAY"
  },
  
  CALL: {
    DIRECTION: 0,
    TYPE: 1,
    PHONE: 2,
    NAME: 3,
    DATE: 4,
    TIME: 5,
    ACTION: 6,
    RESULT: 7,
    REASON: 8,
    DURATION: 9
  },
  
  SMS: {
    DIRECTION: 0,
    TYPE: 1,
    MESSAGE_TYPE: 2,
    SENDER: 3,
    SENDER_NAME: 4,
    RECIPIENT: 5,
    RECIPIENT_NAME: 6,
    DATETIME: 7,
    STATUS: 9
  },
  
  VOICEMAIL_RESULTS: [
    "voicemail", "vm", "left voicemail", "voice mail", 
    "no answer", "busy", "not available"
  ],
  
  VOICEMAIL_DURATION: {
    MIN_SECONDS: 30,
    MAX_SECONDS: 90
  },
  
  LONG_CALL_THRESHOLD: 240,
  MAX_HISTORY_ITEMS: 5,
  
  // v3.4: Patterns that indicate SMS FAILURE (for marking, not filtering)
  SMS_FAILURE_PATTERNS: [
    "failed", "error", "undeliverable", "rejected", "blocked", "invalid"
  ]
};


/**
 * ═══════════════════════════════════════════════════════════════
 * PHONE NORMALIZATION
 * ═══════════════════════════════════════════════════════════════
 */
function RC_normalizePhone_(phone) {
  if (!phone) return "";
  var digits = String(phone).replace(/\D/g, "");
  
  if (digits.length === 11 && digits[0] === "1") {
    digits = digits.substring(1);
  }
  
  if (digits.length > 10) {
    digits = digits.slice(-10);
  }
  
  return digits.length === 10 ? digits : "";
}


/**
 * ═══════════════════════════════════════════════════════════════
 * DURATION PARSING & FORMATTING
 * ═══════════════════════════════════════════════════════════════
 */
function RC_parseDuration_(durationVal) {
  if (!durationVal) return 0;
  
  if (durationVal instanceof Date) {
    var h = durationVal.getHours();
    var m = durationVal.getMinutes();
    return (h * 60) + m;
  }
  
  var str = String(durationVal).trim();
  if (!str || str === "0" || str === "0:00" || str === "0:00:00") return 0;
  if (str.toLowerCase().indexOf("progress") >= 0) return 0;
  
  var parts = str.split(":");
  
  if (parts.length === 3) {
    var minutes = parseInt(parts[0], 10) || 0;
    var seconds = parseInt(parts[1], 10) || 0;
    return (minutes * 60) + seconds;
  } else if (parts.length === 2) {
    var minutes = parseInt(parts[0], 10) || 0;
    var seconds = parseInt(parts[1], 10) || 0;
    return (minutes * 60) + seconds;
  }
  
  return parseInt(str, 10) || 0;
}

function RC_formatDuration_(seconds) {
  if (!seconds || seconds <= 0) return "0:00";
  
  var h = Math.floor(seconds / 3600);
  var m = Math.floor((seconds % 3600) / 60);
  var s = Math.floor(seconds % 60);
  
  if (h > 0) {
    return h + ":" + (m < 10 ? "0" : "") + m + ":" + (s < 10 ? "0" : "") + s;
  }
  return m + ":" + (s < 10 ? "0" : "") + s;
}


/**
 * ═══════════════════════════════════════════════════════════════
 * DATE/TIME PARSING & FORMATTING
 * ═══════════════════════════════════════════════════════════════
 */
function RC_parseDateTime_(dateVal, timeVal) {
  try {
    if (!dateVal) return null;
    
    var date;
    if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
      date = new Date(dateVal);
    } else {
      var str = String(dateVal).trim();
      if (!str) return null;
      date = new Date(str);
      if (isNaN(date.getTime())) return null;
    }
    
    if (timeVal) {
      if (timeVal instanceof Date) {
        date.setHours(timeVal.getHours(), timeVal.getMinutes(), timeVal.getSeconds());
      } else if (typeof timeVal === "string" && timeVal.indexOf(":") > -1) {
        var parts = timeVal.split(":");
        date.setHours(
          parseInt(parts[0], 10) || 0,
          parseInt(parts[1], 10) || 0,
          parseInt(parts[2], 10) || 0
        );
      }
    }
    
    return date;
  } catch (e) {
    return null;
  }
}

function RC_parseISODateTime_(isoStr) {
  try {
    if (!isoStr) return null;
    var date = new Date(isoStr);
    return isNaN(date.getTime()) ? null : date;
  } catch (e) {
    return null;
  }
}

function RC_timeAgo_(date, now) {
  if (!date) return "Never";
  if (!now) now = new Date();
  
  var diffMs = now.getTime() - date.getTime();
  var diffMins = Math.floor(diffMs / (1000 * 60));
  var diffHours = Math.floor(diffMins / 60);
  var diffDays = Math.floor(diffHours / 24);
  
  if (diffMins < 1) return "Just now";
  if (diffMins < 60) return diffMins + "m ago";
  if (diffHours < 24) return diffHours + "h ago";
  if (diffDays === 1) return "Yesterday";
  if (diffDays < 7) return diffDays + "d ago";
  if (diffDays < 30) return Math.floor(diffDays / 7) + "w ago";
  return Math.floor(diffDays / 30) + "mo ago";
}

function RC_formatTime_(date, tz) {
  if (!date) return "";
  try {
    return Utilities.formatDate(date, tz || "America/New_York", "M/d h:mm a");
  } catch (e) {
    return "";
  }
}


/**
 * ═══════════════════════════════════════════════════════════════
 * VOICEMAIL DETECTION
 * ═══════════════════════════════════════════════════════════════
 */
function RC_isVoicemailByResult_(result) {
  if (!result) return false;
  var lower = String(result).toLowerCase();
  
  var patterns = RC_ENRICH_CONFIG.VOICEMAIL_RESULTS;
  for (var i = 0; i < patterns.length; i++) {
    if (lower.indexOf(patterns[i]) >= 0) {
      return true;
    }
  }
  return false;
}

function RC_isVoicemailByDuration_(durationSeconds) {
  var cfg = RC_ENRICH_CONFIG.VOICEMAIL_DURATION;
  return durationSeconds >= cfg.MIN_SECONDS && durationSeconds <= cfg.MAX_SECONDS;
}


/**
 * ═══════════════════════════════════════════════════════════════
 * v3.4: SMS FAILURE DETECTION (for marking, NOT filtering)
 * ═══════════════════════════════════════════════════════════════
 */
function RC_isSMSFailed_(status) {
  if (!status) return false;
  
  var st = String(status).toLowerCase().trim();
  if (!st) return false;
  
  var failurePatterns = RC_ENRICH_CONFIG.SMS_FAILURE_PATTERNS;
  for (var i = 0; i < failurePatterns.length; i++) {
    if (st.indexOf(failurePatterns[i]) >= 0) {
      return true;
    }
  }
  return false;
}


/**
 * ═══════════════════════════════════════════════════════════════
 * BUILD LOOKUP INDEX
 * ═══════════════════════════════════════════════════════════════
 */
function RC_buildLookupIndex_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var index = {};
  var cfg = RC_ENRICH_CONFIG;
  
  var ensurePhone = function(phone) {
    var norm = RC_normalizePhone_(phone);
    if (!norm) return null;
    
    if (!index[norm]) {
      index[norm] = {
        smsYesterday: 0,
        smsToday: 0,
        callsYesterday: 0,
        callsToday: 0,
        vmYesterday: 0,
        vmToday: 0,
        repliesYesterday: 0,
        repliesToday: 0,
        inboundCallsYesterday: 0,
        inboundCallsToday: 0,
        longestCallYesterday: 0,
        longestCallToday: 0,
        totalDurationYesterday: 0,
        totalDurationToday: 0,
        lastCall: null,
        lastCallResult: "",
        lastCallDuration: 0,
        lastCallDirection: "",
        lastSMS: null,
        lastSMSStatus: "",
        lastReply: null,
        callHistory: [],
        smsHistory: [],
        replyHistory: [],
        hasLongCall: false,
        longestCallEver: 0,
        hasInboundCall: false
      };
    }
    return norm;
  };
  
  // ═══════════════════════════════════════════════════════════════
  // PROCESS CALL LOGS
  // ═══════════════════════════════════════════════════════════════
  
  var processCallSheet = function(sheetName, isToday) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    
    var numCols = Math.max(cfg.CALL.DURATION + 1, 10);
    var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    
    var tz = "America/New_York";
    var now = new Date();
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var phone = row[cfg.CALL.PHONE];
      var norm = ensurePhone(phone);
      if (!norm) continue;
      
      var direction = String(row[cfg.CALL.DIRECTION] || "").trim().toLowerCase();
      var isInbound = direction === "inbound";
      
      var dateVal = row[cfg.CALL.DATE];
      var timeVal = row[cfg.CALL.TIME];
      var callDateTime = RC_parseDateTime_(dateVal, timeVal);
      
      var durationVal = row[cfg.CALL.DURATION];
      var durationSec = RC_parseDuration_(durationVal);
      
      var result = String(row[cfg.CALL.RESULT] || "").trim();
      
      if (isToday) {
        index[norm].callsToday++;
        index[norm].totalDurationToday += durationSec;
        if (isInbound) index[norm].inboundCallsToday++;
      } else {
        index[norm].callsYesterday++;
        index[norm].totalDurationYesterday += durationSec;
        if (isInbound) index[norm].inboundCallsYesterday++;
      }
      
      if (isInbound) {
        index[norm].hasInboundCall = true;
      }
      
      var isVmByResult = RC_isVoicemailByResult_(result);
      var isVmByDuration = !isInbound && RC_isVoicemailByDuration_(durationSec);
      var isVM = isVmByResult || isVmByDuration;
      
      if (isVM && !isInbound) {
        if (isToday) {
          index[norm].vmToday++;
        } else {
          index[norm].vmYesterday++;
        }
      }
      
      if (isToday) {
        if (durationSec > index[norm].longestCallToday) {
          index[norm].longestCallToday = durationSec;
        }
      } else {
        if (durationSec > index[norm].longestCallYesterday) {
          index[norm].longestCallYesterday = durationSec;
        }
      }
      
      if (durationSec > index[norm].longestCallEver) {
        index[norm].longestCallEver = durationSec;
      }
      if (durationSec > cfg.LONG_CALL_THRESHOLD) {
        index[norm].hasLongCall = true;
      }
      
      if (callDateTime && (!index[norm].lastCall || callDateTime > index[norm].lastCall)) {
        index[norm].lastCall = callDateTime;
        index[norm].lastCallResult = result;
        index[norm].lastCallDuration = durationSec;
        index[norm].lastCallDirection = isInbound ? "Inbound" : "Outbound";
      }
      
      index[norm].callHistory.push({
        timestamp: callDateTime,
        duration: durationSec,
        durationDisplay: RC_formatDuration_(durationSec),
        result: result,
        isVM: isVM,
        isInbound: isInbound,
        direction: isInbound ? "Inbound" : "Outbound",
        day: isToday ? "today" : "yesterday",
        timeAgo: RC_timeAgo_(callDateTime, now),
        timeFormatted: RC_formatTime_(callDateTime, tz)
      });
    }
  };
  
  processCallSheet(cfg.SHEETS.CALL_YESTERDAY, false);
  processCallSheet(cfg.SHEETS.CALL_TODAY, true);
  
  // ═══════════════════════════════════════════════════════════════
  // PROCESS SMS LOGS - v3.4: Track failed status but don't filter
  // ═══════════════════════════════════════════════════════════════
  
  var processSMSSheet = function(sheetName, isToday) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    
    var numCols = Math.max(cfg.SMS.STATUS + 1, 10);
    var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    
    var tz = "America/New_York";
    var now = new Date();
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var direction = String(row[cfg.SMS.DIRECTION] || "").trim().toLowerCase();
      var status = String(row[cfg.SMS.STATUS] || "").trim();
      var dateTimeVal = row[cfg.SMS.DATETIME];
      var msgDateTime = RC_parseISODateTime_(dateTimeVal);
      
      if (!msgDateTime && dateTimeVal) {
        msgDateTime = RC_parseDateTime_(dateTimeVal);
      }
      
      var phone, norm;
      
      if (direction === "outbound") {
        phone = row[cfg.SMS.RECIPIENT];
        norm = ensurePhone(phone);
        if (!norm) continue;
        
        if (isToday) {
          index[norm].smsToday++;
        } else {
          index[norm].smsYesterday++;
        }
        
        if (msgDateTime && (!index[norm].lastSMS || msgDateTime > index[norm].lastSMS)) {
          index[norm].lastSMS = msgDateTime;
          index[norm].lastSMSStatus = status;
        }
        
        // v3.4: Mark as failed but DON'T filter - rep needs to see this!
        var isFailed = RC_isSMSFailed_(status);
        
        index[norm].smsHistory.push({
          timestamp: msgDateTime,
          direction: "outbound",
          status: status || "Sent",
          isFailed: isFailed,  // v3.4: NEW - for UI to show red indicator
          day: isToday ? "today" : "yesterday",
          timeAgo: RC_timeAgo_(msgDateTime, now),
          timeFormatted: RC_formatTime_(msgDateTime, tz)
        });
        
      } else if (direction === "inbound") {
        phone = row[cfg.SMS.SENDER];
        norm = ensurePhone(phone);
        if (!norm) continue;
        
        if (isToday) {
          index[norm].repliesToday++;
        } else {
          index[norm].repliesYesterday++;
        }
        
        if (msgDateTime && (!index[norm].lastReply || msgDateTime > index[norm].lastReply)) {
          index[norm].lastReply = msgDateTime;
        }
        
        index[norm].replyHistory.push({
          timestamp: msgDateTime,
          direction: "inbound",
          status: status,
          day: isToday ? "today" : "yesterday",
          timeAgo: RC_timeAgo_(msgDateTime, now),
          timeFormatted: RC_formatTime_(msgDateTime, tz)
        });
      }
    }
  };
  
  processSMSSheet(cfg.SHEETS.SMS_YESTERDAY, false);
  processSMSSheet(cfg.SHEETS.SMS_TODAY, true);
  
  return index;
}


/**
 * ═══════════════════════════════════════════════════════════════
 * HELPER - Sort and trim history array
 * ═══════════════════════════════════════════════════════════════
 */
function RC_sortAndTrimHistory_(arr, maxItems) {
  if (!arr || !arr.length) return [];
  
  var copy = arr.slice();
  
  copy.sort(function(a, b) {
    var ta = a.timestamp ? a.timestamp.getTime() : 0;
    var tb = b.timestamp ? b.timestamp.getTime() : 0;
    return tb - ta;
  });
  
  return copy.slice(0, maxItems || 5);
}


/**
 * ═══════════════════════════════════════════════════════════════
 * ENRICH LEADS - v3.4: No filtering, include isFailed flag
 * ═══════════════════════════════════════════════════════════════
 */
function RC_enrichLeads_(leads, rcIndex) {
  if (!leads || !leads.length) return;
  if (!rcIndex) rcIndex = {};
  
  var tz = "America/New_York";
  var now = new Date();
  var cfg = RC_ENRICH_CONFIG;
  
  for (var i = 0; i < leads.length; i++) {
    var lead = leads[i];
    var phone = RC_normalizePhone_(lead.phone);
    var data = rcIndex[phone];
    
    if (!data) {
      data = {
        smsYesterday: 0, smsToday: 0,
        callsYesterday: 0, callsToday: 0,
        vmYesterday: 0, vmToday: 0,
        repliesYesterday: 0, repliesToday: 0,
        inboundCallsYesterday: 0, inboundCallsToday: 0,
        longestCallYesterday: 0, longestCallToday: 0,
        longestCallEver: 0,
        lastCall: null, lastSMS: null, lastReply: null,
        lastCallResult: "", lastCallDuration: 0, lastCallDirection: "",
        lastSMSStatus: "",
        callHistory: [], smsHistory: [], replyHistory: [],
        hasLongCall: false, hasInboundCall: false
      };
    }
    
    var longestYesterday = data.longestCallYesterday || 0;
    var longestToday = data.longestCallToday || 0;
    var longestTotal = Math.max(longestYesterday, longestToday);
    var longestEver = data.longestCallEver || longestTotal;
    
    var callHistory = RC_sortAndTrimHistory_(data.callHistory || [], cfg.MAX_HISTORY_ITEMS);
    
    // v3.4: NO FILTERING - show all SMS attempts, mark failures
    var smsHistory = RC_sortAndTrimHistory_(data.smsHistory || [], cfg.MAX_HISTORY_ITEMS);
    
    var replyHistory = RC_sortAndTrimHistory_(data.replyHistory || [], cfg.MAX_HISTORY_ITEMS);
    
    var hasReplied = (data.repliesYesterday || 0) + (data.repliesToday || 0) > 0;
    
    // v3.4: Check if ALL SMS attempts failed (to show warning banner)
    var allSmsFailed = false;
    var totalSms = (data.smsYesterday || 0) + (data.smsToday || 0);
    if (totalSms > 0 && smsHistory.length > 0) {
      var failedCount = smsHistory.filter(function(s) { return s.isFailed; }).length;
      allSmsFailed = (failedCount === smsHistory.length);
    }
    
    lead.rc = {
      // Counts
      smsYesterday: data.smsYesterday || 0,
      smsToday: data.smsToday || 0,
      callsYesterday: data.callsYesterday || 0,
      callsToday: data.callsToday || 0,
      vmYesterday: data.vmYesterday || 0,
      vmToday: data.vmToday || 0,
      repliesYesterday: data.repliesYesterday || 0,
      repliesToday: data.repliesToday || 0,
      
      // Inbound call counts
      inboundCallsYesterday: data.inboundCallsYesterday || 0,
      inboundCallsToday: data.inboundCallsToday || 0,
      inboundCallsTotal: (data.inboundCallsYesterday || 0) + (data.inboundCallsToday || 0),
      
      // Totals
      smsTotal: totalSms,
      callsTotal: (data.callsYesterday || 0) + (data.callsToday || 0),
      vmTotal: (data.vmYesterday || 0) + (data.vmToday || 0),
      repliesTotal: (data.repliesYesterday || 0) + (data.repliesToday || 0),
      
      // Longest call
      longestCallYesterday: longestYesterday,
      longestCallToday: longestToday,
      longestCallTotal: longestTotal,
      longestCallDisplay: RC_formatDuration_(longestTotal),
      longestCallEver: longestEver,
      longestCallEverDisplay: RC_formatDuration_(longestEver),
      
      // Last contact
      lastCall: data.lastCall || null,
      lastSMS: data.lastSMS || null,
      lastReply: data.lastReply || null,
      
      lastCallAgo: RC_timeAgo_(data.lastCall, now),
      lastSMSAgo: RC_timeAgo_(data.lastSMS, now),
      lastReplyAgo: RC_timeAgo_(data.lastReply, now),
      
      lastCallTime: RC_formatTime_(data.lastCall, tz),
      lastSMSTime: RC_formatTime_(data.lastSMS, tz),
      lastReplyTime: RC_formatTime_(data.lastReply, tz),
      
      lastCallResult: data.lastCallResult || "",
      lastCallDuration: RC_formatDuration_(data.lastCallDuration || 0),
      lastCallDirection: data.lastCallDirection || "",
      lastSMSStatus: data.lastSMSStatus || "",
      
      // Last 5 Calls
      last5Calls: callHistory.map(function(c) {
        return {
          timeAgo: c.timeAgo || RC_timeAgo_(c.timestamp, now),
          timeFormatted: c.timeFormatted || RC_formatTime_(c.timestamp, tz),
          duration: c.duration,
          durationDisplay: c.durationDisplay || RC_formatDuration_(c.duration),
          result: c.result,
          isVM: c.isVM,
          isInbound: c.isInbound || false,
          direction: c.direction || "Outbound",
          day: c.day,
          display: (c.timeAgo || RC_timeAgo_(c.timestamp, now)) + " (" + 
                   (c.timeFormatted || RC_formatTime_(c.timestamp, tz)) + ") - " + 
                   (c.durationDisplay || RC_formatDuration_(c.duration))
        };
      }),
      
      // v3.4: Last 5 SMS - now includes isFailed flag
      last5SMS: smsHistory.map(function(s) {
        return {
          timeAgo: s.timeAgo || RC_timeAgo_(s.timestamp, now),
          timeFormatted: s.timeFormatted || RC_formatTime_(s.timestamp, tz),
          status: s.status || "Sent",
          isFailed: s.isFailed || false,  // v3.4: NEW
          day: s.day,
          display: (s.timeAgo || RC_timeAgo_(s.timestamp, now)) + " (" + 
                   (s.timeFormatted || RC_formatTime_(s.timestamp, tz)) + ")"
        };
      }),
      
      // Reply history
      replyHistory: replyHistory.map(function(r) {
        return {
          timeAgo: r.timeAgo || RC_timeAgo_(r.timestamp, now),
          timeFormatted: r.timeFormatted || RC_formatTime_(r.timestamp, tz),
          day: r.day,
          display: (r.timeAgo || RC_timeAgo_(r.timestamp, now)) + " (" + 
                   (r.timeFormatted || RC_formatTime_(r.timestamp, tz)) + ")"
        };
      }),
      
      // Alert Flags
      hasReplied: hasReplied,
      hasLongCall: data.hasLongCall || false,
      hasInboundCall: data.hasInboundCall || false,
      longCallMinutes: longestEver ? Math.floor(longestEver / 60) : 0,
      
      // v3.4: NEW - SMS failure alerts
      allSmsFailed: allSmsFailed,
      hasFailedSms: smsHistory.some(function(s) { return s.isFailed; })
    };
  }
}


/**
 * ═══════════════════════════════════════════════════════════════
 * GET CONTACT HISTORY - For single phone lookup
 * ═══════════════════════════════════════════════════════════════
 */
function RC_getContactHistory_(phone) {
  var norm = RC_normalizePhone_(phone);
  if (!norm) return null;
  
  var index = RC_buildLookupIndex_();
  var data = index[norm];
  
  if (!data) return null;
  
  var tz = "America/New_York";
  var now = new Date();
  var cfg = RC_ENRICH_CONFIG;
  
  var callHistory = RC_sortAndTrimHistory_(data.callHistory || [], cfg.MAX_HISTORY_ITEMS);
  var smsHistory = RC_sortAndTrimHistory_(data.smsHistory || [], cfg.MAX_HISTORY_ITEMS);
  
  return {
    phone: norm,
    hasData: true,
    
    lastCall: data.lastCall ? {
      timeAgo: RC_timeAgo_(data.lastCall, now),
      dateTime: RC_formatTime_(data.lastCall, tz),
      result: data.lastCallResult || "Unknown",
      duration: RC_formatDuration_(data.lastCallDuration || 0),
      direction: data.lastCallDirection || "Outbound"
    } : null,
    
    lastSMS: data.lastSMS ? {
      timeAgo: RC_timeAgo_(data.lastSMS, now),
      dateTime: RC_formatTime_(data.lastSMS, tz),
      status: data.lastSMSStatus || "Sent"
    } : null,
    
    lastReply: data.lastReply ? {
      timeAgo: RC_timeAgo_(data.lastReply, now),
      dateTime: RC_formatTime_(data.lastReply, tz)
    } : null,
    
    counts: {
      smsYesterday: data.smsYesterday || 0,
      smsToday: data.smsToday || 0,
      callsYesterday: data.callsYesterday || 0,
      callsToday: data.callsToday || 0,
      vmYesterday: data.vmYesterday || 0,
      vmToday: data.vmToday || 0,
      repliesYesterday: data.repliesYesterday || 0,
      repliesToday: data.repliesToday || 0,
      inboundCallsYesterday: data.inboundCallsYesterday || 0,
      inboundCallsToday: data.inboundCallsToday || 0,
      longestCall: RC_formatDuration_(Math.max(data.longestCallYesterday || 0, data.longestCallToday || 0))
    },
    
    last5Calls: callHistory,
    last5SMS: smsHistory,
    
    hasReplied: (data.repliesYesterday || 0) + (data.repliesToday || 0) > 0,
    hasLongCall: data.hasLongCall || false,
    hasInboundCall: data.hasInboundCall || false,
    longestCallEver: data.longestCallEver || 0
  };
}


/**
 * ═══════════════════════════════════════════════════════════════
 * DIAGNOSTICS
 * ═══════════════════════════════════════════════════════════════
 */

function RC_viewSelectedContactHistory() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getActiveCell();
  var phone = String(cell.getValue() || "").trim();
  
  if (!phone) {
    SpreadsheetApp.getUi().alert("Please select a cell with a phone number first.");
    return;
  }
  
  var history = RC_getContactHistory_(phone);
  
  var lines = [];
  lines.push("📱 CONTACT HISTORY FOR: " + phone);
  lines.push("═".repeat(45));
  lines.push("");
  
  if (!history) {
    lines.push("❌ No RingCentral history found for this phone.");
    lines.push("Normalized: " + RC_normalizePhone_(phone));
    SpreadsheetApp.getUi().alert(lines.join("\n"));
    return;
  }
  
  lines.push("📊 CONTACT COUNTS (Yesterday → Today):");
  lines.push("   SMS Sent: " + history.counts.smsYesterday + " → " + history.counts.smsToday);
  lines.push("   Calls Made: " + history.counts.callsYesterday + " → " + history.counts.callsToday);
  lines.push("   Inbound Calls: " + history.counts.inboundCallsYesterday + " → " + history.counts.inboundCallsToday);
  lines.push("   Voicemails: " + history.counts.vmYesterday + " → " + history.counts.vmToday);
  lines.push("   Replies Received: " + history.counts.repliesYesterday + " → " + history.counts.repliesToday);
  lines.push("   Longest Call: " + history.counts.longestCall);
  lines.push("");
  
  lines.push("🚨 ALERT FLAGS:");
  lines.push("   Has Replied: " + (history.hasReplied ? "YES ✅" : "No"));
  lines.push("   Has Inbound Call: " + (history.hasInboundCall ? "YES 📞" : "No"));
  lines.push("   Has Long Call (>4min): " + (history.hasLongCall ? "YES ⚠️" : "No"));
  lines.push("");
  
  lines.push("📞 LAST 5 CALLS:");
  if (history.last5Calls && history.last5Calls.length > 0) {
    for (var i = 0; i < history.last5Calls.length; i++) {
      var c = history.last5Calls[i];
      var inTag = c.isInbound ? " [IN]" : "";
      var vmTag = c.isVM ? " [VM]" : "";
      lines.push("   " + (i + 1) + ". " + c.timeAgo + " (" + c.timeFormatted + ") - " + c.durationDisplay + inTag + vmTag);
    }
  } else {
    lines.push("   (none)");
  }
  
  lines.push("");
  lines.push("📱 LAST 5 SMS:");
  if (history.last5SMS && history.last5SMS.length > 0) {
    for (var j = 0; j < history.last5SMS.length; j++) {
      var s = history.last5SMS[j];
      var failTag = s.isFailed ? " ❌ FAILED" : " ✔";
      lines.push("   " + (j + 1) + ". " + s.timeAgo + " (" + s.timeFormatted + ") [" + s.status + "]" + failTag);
    }
  } else {
    lines.push("   (none)");
  }
  
  SpreadsheetApp.getUi().alert(lines.join("\n"));
}

function RC_testIndexStats() {
  var startTime = new Date().getTime();
  var index = RC_buildLookupIndex_();
  var elapsed = new Date().getTime() - startTime;
  
  var phoneCount = Object.keys(index).length;
  
  var totalSMSYesterday = 0, totalSMSToday = 0;
  var totalCallsYesterday = 0, totalCallsToday = 0;
  var totalInboundYesterday = 0, totalInboundToday = 0;
  var totalVMYesterday = 0, totalVMToday = 0;
  var totalRepliesYesterday = 0, totalRepliesToday = 0;
  var maxLongestCall = 0;
  var withCalls = 0, withSMS = 0, withVM = 0, withReplies = 0;
  var withLongCalls = 0, withInbound = 0;
  var withSMSHistory = 0, totalSMSHistoryEntries = 0;
  var failedSMSCount = 0;
  
  for (var phone in index) {
    var e = index[phone];
    
    totalSMSYesterday += e.smsYesterday;
    totalSMSToday += e.smsToday;
    totalCallsYesterday += e.callsYesterday;
    totalCallsToday += e.callsToday;
    totalInboundYesterday += e.inboundCallsYesterday || 0;
    totalInboundToday += e.inboundCallsToday || 0;
    totalVMYesterday += e.vmYesterday;
    totalVMToday += e.vmToday;
    totalRepliesYesterday += e.repliesYesterday;
    totalRepliesToday += e.repliesToday;
    
    var longest = Math.max(e.longestCallYesterday || 0, e.longestCallToday || 0);
    if (longest > maxLongestCall) maxLongestCall = longest;
    
    if (e.callsYesterday + e.callsToday > 0) withCalls++;
    if (e.smsYesterday + e.smsToday > 0) withSMS++;
    if (e.vmYesterday + e.vmToday > 0) withVM++;
    if (e.repliesYesterday + e.repliesToday > 0) withReplies++;
    if (e.hasLongCall) withLongCalls++;
    if (e.hasInboundCall) withInbound++;
    
    if (e.smsHistory && e.smsHistory.length > 0) {
      withSMSHistory++;
      totalSMSHistoryEntries += e.smsHistory.length;
      
      // v3.4: Count failed SMS
      for (var k = 0; k < e.smsHistory.length; k++) {
        if (e.smsHistory[k].isFailed) failedSMSCount++;
      }
    }
  }
  
  var msg = [
    "📊 RingCentral Contact Index Stats (v3.4)",
    "═════════════════════════════════════════",
    "Built in " + elapsed + "ms",
    "",
    "Total phones indexed: " + phoneCount,
    "",
    "📱 SMS Messages:",
    "   Yesterday: " + totalSMSYesterday,
    "   Today: " + totalSMSToday,
    "   Total: " + (totalSMSYesterday + totalSMSToday),
    "   Phones with SMS: " + withSMS,
    "   SMS history entries: " + totalSMSHistoryEntries,
    "   ⚠️ Failed SMS: " + failedSMSCount,
    "",
    "📞 Calls:",
    "   Yesterday: " + totalCallsYesterday,
    "   Today: " + totalCallsToday,
    "   Total: " + (totalCallsYesterday + totalCallsToday),
    "   Phones with calls: " + withCalls,
    "",
    "📥 Inbound Calls:",
    "   Yesterday: " + totalInboundYesterday,
    "   Today: " + totalInboundToday,
    "   Phones with inbound: " + withInbound,
    "",
    "📭 Voicemails:",
    "   Yesterday: " + totalVMYesterday,
    "   Today: " + totalVMToday,
    "   Phones with VM: " + withVM,
    "",
    "📥 Replies:",
    "   Yesterday: " + totalRepliesYesterday,
    "   Today: " + totalRepliesToday,
    "   Phones with replies: " + withReplies,
    "",
    "⏱️ Longest call: " + RC_formatDuration_(maxLongestCall),
    "⚠️ Phones with long calls (>4min): " + withLongCalls
  ].join("\n");
  
  SpreadsheetApp.getUi().alert(msg);
}

function RC_testPhoneMatching() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var activeCell = ss.getActiveCell();
  var testPhone = String(activeCell.getValue() || "").trim();
  
  if (!testPhone || !testPhone.match(/\d/)) {
    ui.alert("Please select a cell with a phone number first.");
    return;
  }
  
  var normalized = RC_normalizePhone_(testPhone);
  var lines = [];
  lines.push("🔧 RC PHONE MATCHING (v3.4)");
  lines.push("═".repeat(50));
  lines.push("");
  lines.push("Testing: " + testPhone);
  lines.push("Normalized: " + normalized);
  lines.push("");
  
  var index = RC_buildLookupIndex_();
  var data = index[normalized];
  
  if (data) {
    lines.push("✅ FOUND IN INDEX");
    lines.push("");
    lines.push("📱 SMS: " + data.smsYesterday + " → " + data.smsToday);
    lines.push("   History entries: " + (data.smsHistory ? data.smsHistory.length : 0));
    
    if (data.smsHistory && data.smsHistory.length > 0) {
      lines.push("   Sample entries:");
      for (var i = 0; i < Math.min(3, data.smsHistory.length); i++) {
        var s = data.smsHistory[i];
        var failMark = s.isFailed ? " ❌" : " ✔";
        lines.push("      " + (i+1) + ". " + s.timeFormatted + " [" + s.status + "]" + failMark);
      }
    }
    
    lines.push("");
    lines.push("📞 Calls: " + data.callsYesterday + " → " + data.callsToday);
    lines.push("   History entries: " + (data.callHistory ? data.callHistory.length : 0));
  } else {
    lines.push("❌ NOT FOUND IN INDEX");
  }
  
  ui.alert(lines.join("\n"));
}

function RC_testEnrichment() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getActiveCell();
  var phone = String(cell.getValue() || "").trim();
  
  if (!phone) {
    ui.alert("Please select a cell with a phone number first.");
    return;
  }
  
  var mockLead = { phone: phone, job: "TEST-001", username: "TEST_USER" };
  var index = RC_buildLookupIndex_();
  RC_enrichLeads_([mockLead], index);
  
  var rc = mockLead.rc || {};
  
  var lines = [];
  lines.push("🔍 RC ENRICHMENT TEST (v3.4)");
  lines.push("═".repeat(50));
  lines.push("Phone: " + phone);
  lines.push("");
  
  lines.push("📊 COUNTS:");
  lines.push("   SMS: " + rc.smsYesterday + " → " + rc.smsToday + " (Total: " + rc.smsTotal + ")");
  lines.push("   Calls: " + rc.callsYesterday + " → " + rc.callsToday);
  lines.push("");
  
  lines.push("📱 LAST 5 SMS (length=" + (rc.last5SMS ? rc.last5SMS.length : 0) + "):");
  if (rc.last5SMS && rc.last5SMS.length > 0) {
    for (var j = 0; j < rc.last5SMS.length; j++) {
      var s = rc.last5SMS[j];
      var failMark = s.isFailed ? " ❌ FAILED" : " ✔";
      lines.push("   " + (j+1) + ". " + s.timeAgo + " [" + s.status + "]" + failMark);
    }
  } else {
    lines.push("   (none)");
  }
  
  lines.push("");
  lines.push("🚨 FLAGS:");
  lines.push("   allSmsFailed: " + (rc.allSmsFailed ? "YES ⚠️" : "No"));
  lines.push("   hasFailedSms: " + (rc.hasFailedSms ? "YES" : "No"));
  
  ui.alert(lines.join("\n"));
}

function RC_debugSMSMatching() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var activeCell = ss.getActiveCell();
  var testPhone = String(activeCell.getValue() || "").trim();
  
  if (!testPhone || !testPhone.match(/\d/)) {
    ui.alert("Please select a cell with a phone number first.");
    return;
  }
  
  var normalized = RC_normalizePhone_(testPhone);
  var lines = [];
  lines.push("🔍 SMS DEBUG (v3.4): " + testPhone);
  lines.push("Normalized: " + normalized);
  lines.push("═".repeat(50));
  lines.push("");
  
  var index = RC_buildLookupIndex_();
  var data = index[normalized];
  
  if (!data) {
    lines.push("❌ Phone not found in index");
    ui.alert(lines.join("\n"));
    return;
  }
  
  lines.push("📊 SMS COUNTS:");
  lines.push("   Yesterday: " + data.smsYesterday);
  lines.push("   Today: " + data.smsToday);
  lines.push("   Total: " + (data.smsYesterday + data.smsToday));
  lines.push("");
  
  lines.push("📜 SMS HISTORY:");
  lines.push("   Length: " + (data.smsHistory ? data.smsHistory.length : 0));
  lines.push("");
  
  if (data.smsHistory && data.smsHistory.length > 0) {
    lines.push("🔍 ENTRIES (v3.4 - NO FILTERING):");
    for (var i = 0; i < data.smsHistory.length; i++) {
      var s = data.smsHistory[i];
      var marker = s.isFailed ? "❌ FAILED" : "✅ OK";
      lines.push("   " + (i+1) + ". Status: \"" + s.status + "\"");
      lines.push("      Time: " + s.timeFormatted);
      lines.push("      " + marker);
      lines.push("");
    }
  } else {
    lines.push("   (no entries)");
  }
  
  ui.alert(lines.join("\n"));
}

function RC_diagnoseSMSIssue() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var cfg = RC_ENRICH_CONFIG;
  
  var lines = [];
  lines.push("🔬 SMS DIAGNOSTIC (v3.4)");
  lines.push("═".repeat(55));
  lines.push("");
  
  var sheets = [cfg.SHEETS.SMS_TODAY, cfg.SHEETS.SMS_YESTERDAY];
  
  for (var s = 0; s < sheets.length; s++) {
    var sheetName = sheets[s];
    var sheet = ss.getSheetByName(sheetName);
    
    lines.push("📋 " + sheetName + ":");
    
    if (!sheet) {
      lines.push("   ❌ NOT FOUND");
      continue;
    }
    
    var lastRow = sheet.getLastRow();
    lines.push("   Rows: " + lastRow);
    
    if (lastRow >= 2) {
      var sample = sheet.getRange(2, 1, 1, 10).getValues()[0];
      lines.push("   Direction (col 0): " + sample[cfg.SMS.DIRECTION]);
      lines.push("   Recipient (col 5): " + sample[cfg.SMS.RECIPIENT]);
      lines.push("   DateTime (col 7): " + sample[cfg.SMS.DATETIME]);
      lines.push("   Status (col 9): " + sample[cfg.SMS.STATUS]);
    }
    lines.push("");
  }
  
  var index = RC_buildLookupIndex_();
  var withSms = 0, withHistory = 0;
  
  for (var phone in index) {
    if (index[phone].smsYesterday + index[phone].smsToday > 0) withSms++;
    if (index[phone].smsHistory.length > 0) withHistory++;
  }
  
  lines.push("📊 INDEX RESULTS:");
  lines.push("   Phones with SMS count: " + withSms);
  lines.push("   Phones with SMS history: " + withHistory);
  
  ui.alert(lines.join("\n"));
}

function RC_debugDurationParsing() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var cfg = RC_ENRICH_CONFIG;
  
  var sheet = ss.getSheetByName(cfg.SHEETS.CALL_TODAY);
  if (!sheet) {
    ui.alert("RC CALL LOG not found");
    return;
  }
  
  var lastRow = Math.min(sheet.getLastRow(), 6);
  if (lastRow < 2) {
    ui.alert("No data");
    return;
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var lines = ["⏱️ DURATION PARSING", "═".repeat(40), ""];
  
  for (var i = 0; i < data.length; i++) {
    var dur = data[i][cfg.CALL.DURATION];
    lines.push((i+1) + ". Raw: " + dur + " → " + RC_formatDuration_(RC_parseDuration_(dur)));
  }
  
  ui.alert(lines.join("\n"));
}
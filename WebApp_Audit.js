/**************************************************************
 * 🚢 SAFE SHIP CONTACT CHECKER PRO — REP WEB APP
 * File: WebApp_Audit.gs
 *
 * Handles audit counters by indexing RC CALL LOG and RC SMS LOG.
 * Counts OUTBOUND ONLY - inbound is ignored for compliance tracking.
 *
 * Performance:
 *   - Builds phone→count indexes in memory
 *   - Caches indexes for 1-3 minutes
 *   - Uses normalized last-10-digit phone matching
 *
 * v1.0 — Production Build
 **************************************************************/

/* ════════════════════════════════════════════════════════════════
 * CONFIGURATION
 * ════════════════════════════════════════════════════════════════ */

var WEBAPP_AUDIT_CONFIG = {
  // Sheet names
  RC_CALL_LOG: 'RC CALL LOG',
  RC_SMS_LOG: 'RC SMS LOG',
  
  // RC CALL LOG columns (0-based)
  CALL_COLS: {
    DIRECTION: 0,    // A
    TYPE: 1,         // B
    PHONE: 2,        // C
    NAME: 3,         // D
    DATE: 4,         // E
    TIME: 5,         // F
    ACTION: 6,       // G
    RESULT: 7,       // H
    REASON: 8,       // I
    DURATION: 9      // J
  },
  
  // RC SMS LOG columns (0-based)
  SMS_COLS: {
    DIRECTION: 0,       // A
    TYPE: 1,            // B
    MESSAGE_TYPE: 2,    // C
    SENDER_NUMBER: 3,   // D
    SENDER_NAME: 4,     // E
    RECIPIENT_NUMBER: 5,// F
    RECIPIENT_NAME: 6,  // G
    DATE_TIME: 7,       // H
    SEGMENT_COUNT: 8,   // I
    MESSAGE_STATUS: 9,  // J
    DETAILED_ERROR: 10  // K
  },
  
  // Cache settings (OPTIMIZED - longer cache)
  CACHE_KEY_CALLS: 'WEBAPP_RC_CALL_INDEX',
  CACHE_KEY_SMS: 'WEBAPP_RC_SMS_INDEX',
  CACHE_DURATION: 180,  // 3 minutes (was 2)
  
  // Direction values
  OUTBOUND: 'outbound'
};

/* ════════════════════════════════════════════════════════════════
 * MAIN AUDIT FUNCTIONS
 * ════════════════════════════════════════════════════════════════ */

/**
 * Get audit counters for a list of phone numbers
 * @param {Array} phones - Array of phone numbers to check
 * @returns {Object} Map of normalized phone -> { calls, sms, lastCall, lastSms, ... }
 */
function WebApp_Audit_getCounters(phones) {
  if (!phones || phones.length === 0) return {};
  
  // Normalize all input phones
  var phoneSet = new Set();
  phones.forEach(function(p) {
    var norm = WebApp_Audit_normalizePhone_(p);
    if (norm) phoneSet.add(norm);
  });
  
  if (phoneSet.size === 0) return {};
  
  // Build indexes from RC logs
  var callIndex = WebApp_Audit_getCallIndex_();
  var smsIndex = WebApp_Audit_getSmsIndex_();
  
  // Build result for each requested phone
  var result = {};
  
  phoneSet.forEach(function(phone) {
    var callData = callIndex[phone] || { count: 0, lastTime: null, lastResult: '' };
    var smsData = smsIndex[phone] || { count: 0, lastTime: null, lastStatus: '' };
    
    result[phone] = {
      calls: callData.count,
      sms: smsData.count,
      lastCall: callData.lastTime,
      lastCallResult: callData.lastResult,
      lastSms: smsData.lastTime,
      lastSmsStatus: smsData.lastStatus
    };
  });
  
  return result;
}

/**
 * Get rep activity summary (total calls/SMS made today)
 * @param {string} username - Rep's username (not currently used for filtering, but reserved)
 * @returns {Object} Summary stats
 */
function WebApp_Audit_getRepSummary(username) {
  var callIndex = WebApp_Audit_getCallIndex_();
  var smsIndex = WebApp_Audit_getSmsIndex_();
  
  var totalCalls = 0;
  var totalSms = 0;
  var uniquePhones = new Set();
  
  for (var phone in callIndex) {
    totalCalls += callIndex[phone].count;
    uniquePhones.add(phone);
  }
  
  for (var phone in smsIndex) {
    totalSms += smsIndex[phone].count;
    uniquePhones.add(phone);
  }
  
  return {
    totalCalls: totalCalls,
    totalSms: totalSms,
    uniqueLeadsContacted: uniquePhones.size
  };
}

/* ════════════════════════════════════════════════════════════════
 * INDEX BUILDING (with caching)
 * ════════════════════════════════════════════════════════════════ */

/**
 * Get call index (phone -> count/lastTime/result)
 * Uses caching for performance
 * @returns {Object} Call index
 */
function WebApp_Audit_getCallIndex_() {
  var cache = CacheService.getScriptCache();
  var cacheKey = WEBAPP_AUDIT_CONFIG.CACHE_KEY_CALLS;
  
  // Try cache first
  var cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      Logger.log('Cache parse error for calls: ' + e.message);
    }
  }
  
  // Build from sheet
  var index = WebApp_Audit_buildCallIndex_();
  
  // Cache the result
  try {
    cache.put(cacheKey, JSON.stringify(index), WEBAPP_AUDIT_CONFIG.CACHE_DURATION);
  } catch (e) {
    Logger.log('Cache put error: ' + e.message);
  }
  
  return index;
}

/**
 * Get SMS index (phone -> count/lastTime/status)
 * Uses caching for performance
 * @returns {Object} SMS index
 */
function WebApp_Audit_getSmsIndex_() {
  var cache = CacheService.getScriptCache();
  var cacheKey = WEBAPP_AUDIT_CONFIG.CACHE_KEY_SMS;
  
  // Try cache first
  var cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      Logger.log('Cache parse error for SMS: ' + e.message);
    }
  }
  
  // Build from sheet
  var index = WebApp_Audit_buildSmsIndex_();
  
  // Cache the result
  try {
    cache.put(cacheKey, JSON.stringify(index), WEBAPP_AUDIT_CONFIG.CACHE_DURATION);
  } catch (e) {
    Logger.log('Cache put error: ' + e.message);
  }
  
  return index;
}

/**
 * Build call index from RC CALL LOG
 * @returns {Object} Index of phone -> { count, lastTime, lastResult }
 */
function WebApp_Audit_buildCallIndex_() {
  var index = {};
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(WEBAPP_AUDIT_CONFIG.RC_CALL_LOG);
    
    if (!sheet) {
      Logger.log('RC CALL LOG sheet not found');
      return index;
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return index;
    
    var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();  // A:J
    var CC = WEBAPP_AUDIT_CONFIG.CALL_COLS;
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      
      // Only count OUTBOUND calls
      var direction = String(row[CC.DIRECTION] || '').toLowerCase().trim();
      if (direction !== WEBAPP_AUDIT_CONFIG.OUTBOUND) continue;
      
      // Get phone number
      var phone = WebApp_Audit_normalizePhone_(row[CC.PHONE]);
      if (!phone) continue;
      
      // Parse date/time
      var callDate = row[CC.DATE];
      var callTime = row[CC.TIME];
      var callDateTime = WebApp_Audit_combineDateAndTime_(callDate, callTime);
      
      // Get result
      var result = String(row[CC.RESULT] || '');
      
      // Update index
      if (!index[phone]) {
        index[phone] = {
          count: 0,
          lastTime: null,
          lastResult: ''
        };
      }
      
      index[phone].count++;
      
      // Track most recent
      if (callDateTime && (!index[phone].lastTime || callDateTime > new Date(index[phone].lastTime))) {
        index[phone].lastTime = callDateTime.toISOString();
        index[phone].lastResult = result;
      }
    }
    
  } catch (error) {
    Logger.log('Error building call index: ' + error.message);
  }
  
  return index;
}

/**
 * Build SMS index from RC SMS LOG
 * @returns {Object} Index of phone -> { count, lastTime, lastStatus }
 */
function WebApp_Audit_buildSmsIndex_() {
  var index = {};
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(WEBAPP_AUDIT_CONFIG.RC_SMS_LOG);
    
    if (!sheet) {
      Logger.log('RC SMS LOG sheet not found');
      return index;
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return index;
    
    var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();  // A:K
    var SC = WEBAPP_AUDIT_CONFIG.SMS_COLS;
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      
      // Only count OUTBOUND SMS
      var direction = String(row[SC.DIRECTION] || '').toLowerCase().trim();
      if (direction !== WEBAPP_AUDIT_CONFIG.OUTBOUND) continue;
      
      // For outbound SMS, the lead's phone is the RECIPIENT
      var phone = WebApp_Audit_normalizePhone_(row[SC.RECIPIENT_NUMBER]);
      if (!phone) continue;
      
      // Parse date/time
      var dateTime = row[SC.DATE_TIME];
      var smsDateTime = null;
      if (dateTime instanceof Date) {
        smsDateTime = dateTime;
      } else if (dateTime) {
        smsDateTime = new Date(dateTime);
      }
      
      // Get status
      var status = String(row[SC.MESSAGE_STATUS] || '');
      
      // Update index
      if (!index[phone]) {
        index[phone] = {
          count: 0,
          lastTime: null,
          lastStatus: ''
        };
      }
      
      index[phone].count++;
      
      // Track most recent
      if (smsDateTime && !isNaN(smsDateTime.getTime())) {
        if (!index[phone].lastTime || smsDateTime > new Date(index[phone].lastTime)) {
          index[phone].lastTime = smsDateTime.toISOString();
          index[phone].lastStatus = status;
        }
      }
    }
    
  } catch (error) {
    Logger.log('Error building SMS index: ' + error.message);
  }
  
  return index;
}

/* ════════════════════════════════════════════════════════════════
 * UTILITY FUNCTIONS
 * ════════════════════════════════════════════════════════════════ */

/**
 * Normalize phone number to last 10 digits
 * @param {string|number} phone - Phone number
 * @returns {string} Normalized 10-digit phone or empty string
 */
function WebApp_Audit_normalizePhone_(phone) {
  if (!phone) return '';
  var digits = String(phone).replace(/\D/g, '');
  return digits.length >= 10 ? digits.slice(-10) : '';
}

/**
 * Combine separate date and time values into a single Date
 * @param {Date|string} dateVal - Date value
 * @param {Date|string} timeVal - Time value
 * @returns {Date|null} Combined datetime or null
 */
function WebApp_Audit_combineDateAndTime_(dateVal, timeVal) {
  try {
    if (!dateVal) return null;
    
    var d = dateVal instanceof Date ? dateVal : new Date(dateVal);
    if (isNaN(d.getTime())) return null;
    
    if (timeVal) {
      var t = timeVal instanceof Date ? timeVal : new Date(timeVal);
      if (!isNaN(t.getTime())) {
        d.setHours(t.getHours(), t.getMinutes(), t.getSeconds());
      }
    }
    
    return d;
    
  } catch (e) {
    return null;
  }
}

/**
 * Format time ago string
 * @param {string} isoString - ISO timestamp
 * @returns {string} Human readable time ago
 */
function WebApp_Audit_timeAgo(isoString) {
  if (!isoString) return 'Never';
  
  var date = new Date(isoString);
  if (isNaN(date.getTime())) return 'Never';
  
  var seconds = Math.floor((new Date() - date) / 1000);
  
  if (seconds < 60) return seconds + 's ago';
  if (seconds < 3600) return Math.floor(seconds / 60) + 'm ago';
  if (seconds < 86400) return Math.floor(seconds / 3600) + 'h ago';
  return Math.floor(seconds / 86400) + 'd ago';
}

/**
 * Clear audit caches (call after RC sync)
 */
function WebApp_Audit_clearCache() {
  var cache = CacheService.getScriptCache();
  cache.removeAll([
    WEBAPP_AUDIT_CONFIG.CACHE_KEY_CALLS,
    WEBAPP_AUDIT_CONFIG.CACHE_KEY_SMS
  ]);
}
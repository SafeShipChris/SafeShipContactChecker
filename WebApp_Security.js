/**************************************************************
 * 🚢 SAFE SHIP CONTACT CHECKER PRO — REP WEB APP
 * File: WebApp_Security.gs
 *
 * Handles authentication, rep identity mapping, and access control.
 * Uses Google login + Sales_Roster mapping.
 *
 * Security Model:
 *   - User identified by Session.getActiveUser().getEmail()
 *   - Email mapped to username via Sales_Roster sheet
 *   - Only scheduled reps can access their leads
 *   - Admins have full access
 *
 * v1.0 — Production Build
 **************************************************************/

/* ════════════════════════════════════════════════════════════════
 * CONFIGURATION
 * ════════════════════════════════════════════════════════════════ */

var WEBAPP_SECURITY_CONFIG = {
  // Admin emails (full access)
  ADMIN_EMAILS: [
    'christopher@safeshipmoving.com'  // Update with actual admin emails
  ],
  
  // Sheet names
  SALES_ROSTER_SHEET: 'Sales_Roster',
  SCHEDULED_REPS_SHEET: 'SCHEDULED REPS',
  
  // Cache duration (seconds)
  CACHE_DURATION: 300,  // 5 minutes
  
  // Sales_Roster column headers we need
  ROSTER_HEADERS: {
    USERNAME: 'Username',
    EMAIL: 'Email',
    FIRST_NAME: 'First Name',
    LAST_NAME: 'Last Name',
    ROLE: 'Role',
    TEAM: 'Team',
    MANAGER: 'Manager',
    PHONE: 'Phone',
    RC_EXTENSION: 'RC Extension'
  }
};

/* ════════════════════════════════════════════════════════════════
 * MAIN SECURITY FUNCTIONS
 * ════════════════════════════════════════════════════════════════ */

/**
 * Get the current user's context including authorization status
 * @returns {Object} Context object with user info and permissions
 */
function WebApp_Security_getRepContext() {
  var result = {
    authorized: false,
    reason: '',
    email: '',
    username: '',
    displayName: '',
    firstName: '',
    lastName: '',
    role: '',
    team: '',
    manager: '',
    phone: '',
    rcExtension: '',
    isAdmin: false,
    isScheduled: false,
    lastSync: null
  };
  
  try {
    // Get current user email
    var email = Session.getActiveUser().getEmail();
    
    if (!email) {
      result.reason = 'Could not determine user email. Please ensure you are logged in.';
      return result;
    }
    
    result.email = email.toLowerCase().trim();
    
    // Check if admin
    result.isAdmin = WEBAPP_SECURITY_CONFIG.ADMIN_EMAILS.some(function(adminEmail) {
      return adminEmail.toLowerCase().trim() === result.email;
    });
    
    // Look up user in Sales_Roster
    var rosterData = WebApp_Security_getRosterData_();
    var userRow = rosterData.find(function(row) {
      return row.email && row.email.toLowerCase().trim() === result.email;
    });
    
    if (!userRow && !result.isAdmin) {
      result.reason = 'Your email (' + result.email + ') is not found in the Sales Roster. Please contact your administrator.';
      return result;
    }
    
    if (userRow) {
      result.username = (userRow.username || '').toUpperCase().trim();
      result.firstName = userRow.firstName || '';
      result.lastName = userRow.lastName || '';
      result.displayName = (result.firstName + ' ' + result.lastName).trim() || result.username;
      result.role = userRow.role || '';
      result.team = userRow.team || '';
      result.manager = userRow.manager || '';
      result.phone = userRow.phone || '';
      result.rcExtension = userRow.rcExtension || '';
    } else if (result.isAdmin) {
      // Admin without roster entry
      result.username = 'ADMIN';
      result.displayName = 'Administrator';
      result.firstName = 'Admin';
    }
    
    // Check if scheduled (unless admin)
    if (!result.isAdmin) {
      var scheduledReps = WebApp_Security_getScheduledReps_();
      result.isScheduled = scheduledReps.has(result.username);
      
      if (!result.isScheduled) {
        result.reason = 'You (' + result.username + ') are not scheduled today. Please check with your manager.';
        return result;
      }
    } else {
      result.isScheduled = true;  // Admins are always "scheduled"
    }
    
    // Get last sync time
    result.lastSync = WebApp_Security_getLastSyncTime_();
    
    // All checks passed
    result.authorized = true;
    
  } catch (error) {
    Logger.log('WebApp_Security_getRepContext error: ' + error.message);
    result.reason = 'Authentication error: ' + error.message;
  }
  
  return result;
}

/**
 * Check if a username is authorized to view a specific lead
 * @param {string} username - The rep's username
 * @param {string} leadUser - The lead's assigned user
 * @param {boolean} isAdmin - Whether the user is an admin
 * @returns {boolean} Whether access is allowed
 */
function WebApp_Security_canAccessLead(username, leadUser, isAdmin) {
  if (isAdmin) return true;
  
  var normalizedUsername = (username || '').toUpperCase().trim();
  var normalizedLeadUser = (leadUser || '').toUpperCase().trim();
  
  return normalizedUsername === normalizedLeadUser;
}

/* ════════════════════════════════════════════════════════════════
 * DATA LOADING (with caching)
 * ════════════════════════════════════════════════════════════════ */

/**
 * Get Sales_Roster data with caching
 * @returns {Array} Array of roster row objects
 */
function WebApp_Security_getRosterData_() {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'WEBAPP_SALES_ROSTER';
  
  // Try cache first
  var cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }
  
  // Load from sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(WEBAPP_SECURITY_CONFIG.SALES_ROSTER_SHEET);
  
  if (!sheet) {
    Logger.log('Sales_Roster sheet not found');
    return [];
  }
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  var headers = data[0];
  var headerMap = {};
  
  // Map header names to column indexes
  var H = WEBAPP_SECURITY_CONFIG.ROSTER_HEADERS;
  headers.forEach(function(h, i) {
    var normalized = String(h || '').trim();
    if (normalized === H.USERNAME || normalized.toLowerCase() === 'username') headerMap.username = i;
    if (normalized === H.EMAIL || normalized.toLowerCase() === 'email') headerMap.email = i;
    if (normalized === H.FIRST_NAME || normalized.toLowerCase().indexOf('first') >= 0) headerMap.firstName = i;
    if (normalized === H.LAST_NAME || normalized.toLowerCase().indexOf('last') >= 0) headerMap.lastName = i;
    if (normalized === H.ROLE || normalized.toLowerCase() === 'role') headerMap.role = i;
    if (normalized === H.TEAM || normalized.toLowerCase() === 'team') headerMap.team = i;
    if (normalized === H.MANAGER || normalized.toLowerCase() === 'manager') headerMap.manager = i;
    if (normalized === H.PHONE || normalized.toLowerCase() === 'phone') headerMap.phone = i;
    if (normalized === H.RC_EXTENSION || normalized.toLowerCase().indexOf('extension') >= 0) headerMap.rcExtension = i;
  });
  
  // Build roster array
  var roster = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var entry = {
      username: headerMap.username !== undefined ? String(row[headerMap.username] || '') : '',
      email: headerMap.email !== undefined ? String(row[headerMap.email] || '') : '',
      firstName: headerMap.firstName !== undefined ? String(row[headerMap.firstName] || '') : '',
      lastName: headerMap.lastName !== undefined ? String(row[headerMap.lastName] || '') : '',
      role: headerMap.role !== undefined ? String(row[headerMap.role] || '') : '',
      team: headerMap.team !== undefined ? String(row[headerMap.team] || '') : '',
      manager: headerMap.manager !== undefined ? String(row[headerMap.manager] || '') : '',
      phone: headerMap.phone !== undefined ? String(row[headerMap.phone] || '') : '',
      rcExtension: headerMap.rcExtension !== undefined ? String(row[headerMap.rcExtension] || '') : ''
    };
    
    if (entry.username || entry.email) {
      roster.push(entry);
    }
  }
  
  // Cache for 5 minutes
  try {
    cache.put(cacheKey, JSON.stringify(roster), WEBAPP_SECURITY_CONFIG.CACHE_DURATION);
  } catch (e) {
    // Cache might be too large, skip
  }
  
  return roster;
}

/**
 * Get scheduled reps as a Set (with caching)
 * @returns {Set} Set of uppercase usernames
 */
function WebApp_Security_getScheduledReps_() {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'WEBAPP_SCHEDULED_REPS';
  
  // Try cache first
  var cached = cache.get(cacheKey);
  if (cached) {
    try {
      return new Set(JSON.parse(cached));
    } catch (e) {}
  }
  
  // Load from sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(WEBAPP_SECURITY_CONFIG.SCHEDULED_REPS_SHEET);
  
  var reps = new Set();
  
  if (sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      data.forEach(function(row) {
        var username = String(row[0] || '').toUpperCase().trim();
        if (username) {
          reps.add(username);
        }
      });
    }
  }
  
  // Cache
  try {
    cache.put(cacheKey, JSON.stringify(Array.from(reps)), WEBAPP_SECURITY_CONFIG.CACHE_DURATION);
  } catch (e) {}
  
  return reps;
}

/**
 * Get the last RingCentral sync time
 * @returns {string} ISO timestamp or null
 */
function WebApp_Security_getLastSyncTime_() {
  try {
    // First try script property
    var props = PropertiesService.getScriptProperties();
    var lastSync = props.getProperty('WEBAPP_LAST_RC_SYNC');
    
    if (lastSync) {
      return lastSync;
    }
    
    // Fall back to checking RC SMS LOG timestamp
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var smsSheet = ss.getSheetByName('RC SMS LOG');
    
    if (smsSheet && smsSheet.getLastRow() >= 2) {
      // Assuming Date/Time is in column H
      var lastDateTime = smsSheet.getRange(2, 8).getValue();
      if (lastDateTime instanceof Date) {
        return lastDateTime.toISOString();
      }
    }
    
    return null;
    
  } catch (e) {
    Logger.log('Error getting last sync time: ' + e.message);
    return null;
  }
}

/* ════════════════════════════════════════════════════════════════
 * UTILITY FUNCTIONS
 * ════════════════════════════════════════════════════════════════ */

/**
 * Normalize a phone number to last 10 digits
 * @param {string} phone - Phone number
 * @returns {string} Normalized 10-digit phone
 */
function WebApp_Security_normalizePhone(phone) {
  if (!phone) return '';
  var digits = String(phone).replace(/\D/g, '');
  return digits.length >= 10 ? digits.slice(-10) : digits;
}

/**
 * Check if current user is an admin
 * @returns {boolean}
 */
function WebApp_Security_isAdmin() {
  var email = Session.getActiveUser().getEmail();
  if (!email) return false;
  
  return WEBAPP_SECURITY_CONFIG.ADMIN_EMAILS.some(function(adminEmail) {
    return adminEmail.toLowerCase().trim() === email.toLowerCase().trim();
  });
}
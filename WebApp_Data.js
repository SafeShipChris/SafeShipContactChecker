/**************************************************************
 * 🚢 SAFE SHIP CONTACT CHECKER PRO — REP WEB APP
 * File: WebApp_Data.gs
 *
 * OPTIMIZED VERSION - Performance focused
 *
 * Optimizations:
 *   1. Paginate BEFORE audit enrichment (huge win)
 *   2. Cache lead data for 60 seconds
 *   3. Cache dashboard stats for 30 seconds
 *   4. Batch initial load into single call
 *   5. Reduced page size to 20
 *   6. Minimal column reads
 *
 * v2.0 — Performance Optimized
 **************************************************************/

/* ════════════════════════════════════════════════════════════════
 * CONFIGURATION
 * ════════════════════════════════════════════════════════════════ */

var WEBAPP_DATA_CONFIG = {
  // Sheet names
  GRANOT_DATA: 'GRANOT DATA',
  SMS_TRACKER: 'SMS TRACKER',
  CALL_TRACKER: 'CALL & VOICEMAIL TRACKER',
  PRIORITY_1_TRACKER: 'PRIORITY 1 CALL & VOICEMAIL TRACKER',
  PRIORITY_35_TRACKER: 'PRIORITY 3,5 CALL & VOICEMAIL TRACKER',
  
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
  
  // Tracker sheet column indexes (0-based, starting from F which is index 5)
  TRACKER_COLS: {
    MOVE_DATE: 0,   // F
    PHONE: 1,       // G
    COUNT: 2,       // H
    JOB_NO: 3,      // I
    USER: 4,        // J
    CF: 5,          // K
    SOURCE: 6,      // L
    EXCLUDE: 7      // M
  },
  
  // Tracker data start row (1-based)
  TRACKER_DATA_START: 3,
  TRACKER_DATA_COL_START: 6,  // Column F
  
  // Cache settings (OPTIMIZED)
  CACHE_DURATION_LEADS: 60,    // 1 minute for lead lists
  CACHE_DURATION_STATS: 30,    // 30 seconds for stats
  CACHE_DURATION_TRACKER: 45,  // 45 seconds for tracker counts
  
  // Days of leads to show
  LEAD_DAYS_BACK: 5,
  
  // DEFAULT PAGE SIZE (REDUCED for performance)
  DEFAULT_PAGE_SIZE: 20,
  
  // Tab identifiers
  TABS: {
    P0_SMS: 'p0_sms',
    P0_CALL: 'p0_call',
    P1_CALL: 'p1_call',
    P35_CALL: 'p35_call',
    ALL_LEADS: 'all_leads',
    ACTIVITY: 'activity'
  }
};

/* ════════════════════════════════════════════════════════════════
 * BATCHED INITIAL LOAD (Single API call for everything)
 * ════════════════════════════════════════════════════════════════ */

/**
 * Get everything needed for initial app load in ONE call
 * This eliminates 4 round trips on startup
 * @param {string} username - Rep's username
 * @param {boolean} isAdmin - Admin flag
 * @returns {Object} { stats, leads, templates, lastSync }
 */
function WebApp_Data_getInitialLoad(username, isAdmin) {
  var startTime = new Date().getTime();
  
  var result = {
    stats: WebApp_Data_getDashboardStats(username, isAdmin),
    leads: WebApp_Data_getRepLeads(username, { 
      tab: WEBAPP_DATA_CONFIG.TABS.P0_SMS, 
      page: 1, 
      pageSize: WEBAPP_DATA_CONFIG.DEFAULT_PAGE_SIZE 
    }, isAdmin),
    templates: WebApp_Templates_getAll(),
    lastSync: PropertiesService.getScriptProperties().getProperty('WEBAPP_LAST_RC_SYNC'),
    loadTimeMs: 0
  };
  
  result.loadTimeMs = new Date().getTime() - startTime;
  return result;
}

/* ════════════════════════════════════════════════════════════════
 * MAIN DATA FUNCTIONS (OPTIMIZED)
 * ════════════════════════════════════════════════════════════════ */

/**
 * Get leads for a rep based on tab selection and filters
 * OPTIMIZED: Paginates FIRST, then enriches only visible leads
 * 
 * @param {string} username - Rep's username
 * @param {Object} options - { tab, filters, sort, page, pageSize }
 * @param {boolean} isAdmin - If true, show ALL leads (not filtered by user)
 * @returns {Object} { leads: [], total: number, cached: boolean, timing: {} }
 */
function WebApp_Data_getRepLeads(username, options, isAdmin) {
  var timing = { start: new Date().getTime() };
  
  options = options || {};
  var tab = options.tab || WEBAPP_DATA_CONFIG.TABS.ALL_LEADS;
  var filters = options.filters || {};
  var sort = options.sort || { field: 'moveDate', dir: 'asc' };
  var page = options.page || 1;
  var pageSize = options.pageSize || WEBAPP_DATA_CONFIG.DEFAULT_PAGE_SIZE;
  
  var result = {
    leads: [],
    total: 0,
    cached: false,
    tab: tab,
    page: page,
    pageSize: pageSize,
    timing: {}
  };
  
  try {
    // ═══════════════════════════════════════════════════════════
    // STEP 1: Get raw leads (with caching)
    // ═══════════════════════════════════════════════════════════
    timing.fetchStart = new Date().getTime();
    
    var cacheKey = 'WEBAPP_LEADS_' + tab + '_' + (isAdmin ? 'ADMIN' : username);
    var cache = CacheService.getScriptCache();
    var cachedData = cache.get(cacheKey);
    var leads = [];
    
    if (cachedData && !filters.search) {
      // Use cached data (skip cache if searching - need fresh for accuracy)
      try {
        leads = JSON.parse(cachedData);
        result.cached = true;
      } catch (e) {
        leads = WebApp_Data_fetchLeadsForTab_(tab, username, isAdmin);
      }
    } else {
      leads = WebApp_Data_fetchLeadsForTab_(tab, username, isAdmin);
      
      // Cache the results (only if small enough - max 100KB)
      // Skip caching for admin with large datasets
      if (leads.length < 500 && !isAdmin) {
        try {
          var jsonStr = JSON.stringify(leads);
          if (jsonStr.length < 100000) {  // 100KB limit
            cache.put(cacheKey, jsonStr, WEBAPP_DATA_CONFIG.CACHE_DURATION_LEADS);
          }
        } catch (e) {
          // Cache too large or other error, skip silently
        }
      }
    }
    
    timing.fetchEnd = new Date().getTime();
    
    // ═══════════════════════════════════════════════════════════
    // STEP 2: Apply filters (fast in-memory)
    // ═══════════════════════════════════════════════════════════
    timing.filterStart = new Date().getTime();
    
    if (filters.search) {
      var searchLower = filters.search.toLowerCase();
      leads = leads.filter(function(lead) {
        return (lead.firstName + ' ' + lead.lastName).toLowerCase().indexOf(searchLower) >= 0 ||
               String(lead.jobNo).indexOf(searchLower) >= 0 ||
               String(lead.phone1).indexOf(searchLower) >= 0 ||
               String(lead.phone2).indexOf(searchLower) >= 0;
      });
    }
    
    if (filters.priority !== undefined && filters.priority !== null && filters.priority !== '') {
      var priFilter = parseInt(filters.priority, 10);
      leads = leads.filter(function(lead) { return lead.priority === priFilter; });
    }
    
    if (filters.source) {
      var srcLower = filters.source.toLowerCase();
      leads = leads.filter(function(lead) { 
        return String(lead.source || '').toLowerCase().indexOf(srcLower) >= 0; 
      });
    }
    
    if (filters.moveWithinDays) {
      var today = new Date();
      today.setHours(0, 0, 0, 0);
      var futureDate = new Date(today.getTime() + (filters.moveWithinDays * 24 * 60 * 60 * 1000));
      leads = leads.filter(function(lead) {
        if (!lead.moveDate) return false;
        var md = new Date(lead.moveDate);
        return md >= today && md <= futureDate;
      });
    }
    
    timing.filterEnd = new Date().getTime();
    
    // ═══════════════════════════════════════════════════════════
    // STEP 3: Sort (fast in-memory)
    // ═══════════════════════════════════════════════════════════
    timing.sortStart = new Date().getTime();
    
    leads.sort(function(a, b) {
      var aVal, bVal;
      
      switch (sort.field) {
        case 'moveDate':
          aVal = a.moveDateTs || 0;
          bVal = b.moveDateTs || 0;
          break;
        case 'priority':
          aVal = a.priority || 0;
          bVal = b.priority || 0;
          break;
        case 'name':
          aVal = (a.firstName + ' ' + a.lastName).toLowerCase();
          bVal = (b.firstName + ' ' + b.lastName).toLowerCase();
          break;
        default:
          aVal = a.jobNo || '';
          bVal = b.jobNo || '';
      }
      
      if (sort.dir === 'desc') {
        return aVal < bVal ? 1 : aVal > bVal ? -1 : 0;
      }
      return aVal > bVal ? 1 : aVal < bVal ? -1 : 0;
    });
    
    timing.sortEnd = new Date().getTime();
    
    // ═══════════════════════════════════════════════════════════
    // STEP 4: PAGINATE FIRST (before audit enrichment!)
    // ═══════════════════════════════════════════════════════════
    result.total = leads.length;
    
    var startIdx = (page - 1) * pageSize;
    var pageLeads = leads.slice(startIdx, startIdx + pageSize);
    
    // ═══════════════════════════════════════════════════════════
    // STEP 5: Enrich ONLY the page leads with audit data
    // ═══════════════════════════════════════════════════════════
    timing.auditStart = new Date().getTime();
    
    // Collect phones for ONLY this page (max 40 phones for 20 leads)
    var phones = [];
    pageLeads.forEach(function(lead) {
      if (lead.phone1) phones.push(lead.phone1);
      if (lead.phone2) phones.push(lead.phone2);
    });
    
    var auditData = WebApp_Audit_getCounters(phones);
    
    // Enrich only page leads
    pageLeads.forEach(function(lead) {
      var phone1Norm = WebApp_Audit_normalizePhone_(lead.phone1);
      var phone2Norm = WebApp_Audit_normalizePhone_(lead.phone2);
      
      var audit1 = auditData[phone1Norm] || { calls: 0, sms: 0 };
      var audit2 = auditData[phone2Norm] || { calls: 0, sms: 0 };
      
      lead.callsToday = Math.max(audit1.calls, audit2.calls);
      lead.smsToday = Math.max(audit1.sms, audit2.sms);
      lead.lastCall = audit1.lastCall || audit2.lastCall || null;
      lead.lastSms = audit1.lastSms || audit2.lastSms || null;
      lead.lastCallResult = audit1.lastCallResult || audit2.lastCallResult || '';
      lead.lastSmsStatus = audit1.lastSmsStatus || audit2.lastSmsStatus || '';
    });
    
    timing.auditEnd = new Date().getTime();
    
    result.leads = pageLeads;
    
    // Calculate timing stats
    result.timing = {
      fetch: timing.fetchEnd - timing.fetchStart,
      filter: timing.filterEnd - timing.filterStart,
      sort: timing.sortEnd - timing.sortStart,
      audit: timing.auditEnd - timing.auditStart,
      total: new Date().getTime() - timing.start
    };
    
  } catch (error) {
    Logger.log('WebApp_Data_getRepLeads error: ' + error.message);
    throw error;
  }
  
  return result;
}

/**
 * Fetch leads for a specific tab (internal helper)
 */
function WebApp_Data_fetchLeadsForTab_(tab, username, isAdmin) {
  switch (tab) {
    case WEBAPP_DATA_CONFIG.TABS.P0_SMS:
      return WebApp_Data_getTrackerLeads_(WEBAPP_DATA_CONFIG.SMS_TRACKER, username, isAdmin);
      
    case WEBAPP_DATA_CONFIG.TABS.P0_CALL:
      return WebApp_Data_getTrackerLeads_(WEBAPP_DATA_CONFIG.CALL_TRACKER, username, isAdmin);
      
    case WEBAPP_DATA_CONFIG.TABS.P1_CALL:
      return WebApp_Data_getTrackerLeads_(WEBAPP_DATA_CONFIG.PRIORITY_1_TRACKER, username, isAdmin);
      
    case WEBAPP_DATA_CONFIG.TABS.P35_CALL:
      return WebApp_Data_getTrackerLeads_(WEBAPP_DATA_CONFIG.PRIORITY_35_TRACKER, username, isAdmin);
      
    case WEBAPP_DATA_CONFIG.TABS.ALL_LEADS:
    default:
      return WebApp_Data_getAllLeads_(username, isAdmin);
  }
}

/**
 * Get dashboard statistics for a rep (with caching)
 * @param {string} username - Rep's username
 * @param {boolean} isAdmin - If true, count ALL leads
 * @returns {Object} Dashboard stats
 */
function WebApp_Data_getDashboardStats(username, isAdmin) {
  // Try cache first
  var cache = CacheService.getScriptCache();
  var cacheKey = 'WEBAPP_STATS_' + (isAdmin ? 'ADMIN' : username);
  var cached = cache.get(cacheKey);
  
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }
  
  var stats = {
    p0_sms: 0,
    p0_call: 0,
    p1_call: 0,
    p35_call: 0,
    total_leads: 0,
    worked_today: 0,
    sms_sent_today: 0,
    calls_today: 0
  };
  
  try {
    // Count leads from each tracker (optimized batch read)
    stats.p0_sms = WebApp_Data_countTrackerLeads_(WEBAPP_DATA_CONFIG.SMS_TRACKER, username, isAdmin);
    stats.p0_call = WebApp_Data_countTrackerLeads_(WEBAPP_DATA_CONFIG.CALL_TRACKER, username, isAdmin);
    stats.p1_call = WebApp_Data_countTrackerLeads_(WEBAPP_DATA_CONFIG.PRIORITY_1_TRACKER, username, isAdmin);
    stats.p35_call = WebApp_Data_countTrackerLeads_(WEBAPP_DATA_CONFIG.PRIORITY_35_TRACKER, username, isAdmin);
    
    // Total is sum of tracker counts (avoid reading GRANOT again)
    stats.total_leads = stats.p0_sms + stats.p0_call + stats.p1_call + stats.p35_call;
    
    // Count SMS sent today by this rep from log
    if (typeof WebApp_SMSLog_countTodayByRep === 'function') {
      stats.sms_sent_today = WebApp_SMSLog_countTodayByRep(username);
    }
    
    // Get today's RC activity (uses cached index)
    var auditSummary = WebApp_Audit_getRepSummary(username);
    stats.calls_today = auditSummary.totalCalls || 0;
    stats.worked_today = auditSummary.uniqueLeadsContacted || 0;
    
  } catch (error) {
    Logger.log('WebApp_Data_getDashboardStats error: ' + error.message);
  }
  
  // Cache stats
  try {
    cache.put(cacheKey, JSON.stringify(stats), WEBAPP_DATA_CONFIG.CACHE_DURATION_STATS);
  } catch (e) {}
  
  return stats;
}

/* ════════════════════════════════════════════════════════════════
 * INTERNAL DATA LOADING FUNCTIONS (OPTIMIZED)
 * ════════════════════════════════════════════════════════════════ */

/**
 * Get leads from a tracker sheet for a specific user
 * OPTIMIZED: Reads minimal columns, pre-computes timestamps
 */
function WebApp_Data_getTrackerLeads_(sheetName, username, isAdmin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log('Sheet not found: ' + sheetName);
    return [];
  }
  
  var lastRow = sheet.getLastRow();
  var startRow = WEBAPP_DATA_CONFIG.TRACKER_DATA_START;
  
  if (lastRow < startRow) return [];
  
  // Read tracker data (F:AA = 22 columns)
  var numRows = lastRow - startRow + 1;
  var data = sheet.getRange(startRow, WEBAPP_DATA_CONFIG.TRACKER_DATA_COL_START, numRows, 22).getValues();
  
  var normalizedUsername = (username || '').toUpperCase().trim();
  var leads = [];
  
  var TC = WEBAPP_DATA_CONFIG.TRACKER_COLS;
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var leadUser = String(row[TC.USER] || '').toUpperCase().trim();
    
    // Skip if not this rep's lead (unless admin viewing all)
    if (!isAdmin && leadUser !== normalizedUsername) continue;
    
    // Skip excluded leads
    if (row[TC.EXCLUDE] === true || row[TC.EXCLUDE] === 'TRUE') continue;
    
    // Skip if no phone
    if (!row[TC.PHONE]) continue;
    
    // Pre-compute timestamp for faster sorting
    var moveDate = row[TC.MOVE_DATE];
    var moveDateTs = 0;
    if (moveDate instanceof Date) {
      moveDateTs = moveDate.getTime();
    } else if (moveDate) {
      var d = new Date(moveDate);
      if (!isNaN(d.getTime())) moveDateTs = d.getTime();
    }
    
    var lead = {
      jobNo: String(row[TC.JOB_NO] || ''),
      moveDate: moveDate,
      moveDateTs: moveDateTs,  // Pre-computed for fast sorting
      phone1: String(row[TC.PHONE] || ''),
      phone2: '',
      user: leadUser,
      cf: row[TC.CF],
      source: String(row[TC.SOURCE] || ''),
      firstName: '',
      lastName: '',
      priority: 0,
      fromCity: '',
      fromState: '',
      toCity: '',
      toState: '',
      email: '',
      trackerSource: sheetName
    };
    
    // Get extended data if available (columns N onwards = index 8+)
    if (row.length > 8) {
      lead.email = String(row[8] || '');
      lead.firstName = String(row[9] || '');
      lead.lastName = String(row[10] || '');
      lead.fromCity = String(row[11] || '');
      lead.fromState = String(row[12] || '');
      lead.fromZip = String(row[13] || '');
      lead.toCity = String(row[14] || '');
      lead.toState = String(row[15] || '');
      lead.toZip = String(row[16] || '');
      lead.phone2 = String(row[17] || '');
      lead.priority = parseInt(row[21], 10) || 0;  // AA column
    }
    
    leads.push(lead);
  }
  
  return leads;
}

/**
 * Count leads in a tracker for a user (fast - reads single column)
 */
function WebApp_Data_countTrackerLeads_(sheetName, username, isAdmin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return 0;
  
  var lastRow = sheet.getLastRow();
  var startRow = WEBAPP_DATA_CONFIG.TRACKER_DATA_START;
  
  if (lastRow < startRow) return 0;
  
  var numRows = lastRow - startRow + 1;
  
  // If admin, just count all rows
  if (isAdmin) {
    return numRows;
  }
  
  // Only read USER column (J = column 10)
  var data = sheet.getRange(startRow, 10, numRows, 1).getValues();
  
  var normalizedUsername = (username || '').toUpperCase().trim();
  var count = 0;
  
  for (var i = 0; i < data.length; i++) {
    var leadUser = String(data[i][0] || '').toUpperCase().trim();
    if (leadUser === normalizedUsername) {
      count++;
    }
  }
  
  return count;
}

/**
 * Get all leads from GRANOT DATA for a user
 * OPTIMIZED: Pre-computes timestamps, minimal processing
 */
function WebApp_Data_getAllLeads_(username, isAdmin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(WEBAPP_DATA_CONFIG.GRANOT_DATA);
  
  if (!sheet) {
    Logger.log('GRANOT DATA sheet not found');
    return [];
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var data = sheet.getRange(2, 1, lastRow - 1, 24).getValues();  // A:X
  
  var normalizedUsername = (username || '').toUpperCase().trim();
  var leads = [];
  
  var GC = WEBAPP_DATA_CONFIG.GRANOT_COLS;
  
  // Calculate date threshold (today - N days)
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var threshold = new Date(today.getTime() - (WEBAPP_DATA_CONFIG.LEAD_DAYS_BACK * 24 * 60 * 60 * 1000));
  var thresholdTs = threshold.getTime();
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var leadUser = String(row[GC.USER] || '').toUpperCase().trim();
    
    // Skip if not this rep's lead (unless admin viewing all)
    if (!isAdmin && leadUser !== normalizedUsername) continue;
    
    // Skip if no phone
    if (!row[GC.PHONE1] && !row[GC.PHONE2]) continue;
    
    // Skip if open date is too old
    var openDate = row[GC.OPEN_DATE];
    if (openDate instanceof Date && openDate.getTime() < thresholdTs) continue;
    
    // Pre-compute move date timestamp for sorting
    var moveDate = row[GC.MOVE_DATE];
    var moveDateTs = 0;
    if (moveDate instanceof Date) {
      moveDateTs = moveDate.getTime();
    } else if (moveDate) {
      var d = new Date(moveDate);
      if (!isNaN(d.getTime())) moveDateTs = d.getTime();
    }
    
    var lead = {
      jobNo: String(row[GC.JOB_NO] || ''),
      dept: String(row[GC.DEPT] || ''),
      openDate: openDate,
      user: leadUser,
      priority: parseInt(row[GC.PRIORITY], 10) || 0,
      firstName: String(row[GC.FIRST_NAME] || ''),
      lastName: String(row[GC.LAST_NAME] || ''),
      moveDate: moveDate,
      moveDateTs: moveDateTs,  // Pre-computed for fast sorting
      email: String(row[GC.EMAIL] || ''),
      phone1: String(row[GC.PHONE1] || ''),
      phone2: String(row[GC.PHONE2] || ''),
      fromCity: String(row[GC.FROM_CITY] || ''),
      fromState: String(row[GC.FROM_STATE] || ''),
      fromZip: String(row[GC.FROM_ZIP] || ''),
      toCity: String(row[GC.TO_CITY] || ''),
      toState: String(row[GC.TO_STATE] || ''),
      toZip: String(row[GC.TO_ZIP] || ''),
      cf: row[GC.CF],
      source: String(row[GC.SOURCE] || ''),
      trackerSource: 'GRANOT_DATA'
    };
    
    leads.push(lead);
  }
  
  return leads;
}

/* ════════════════════════════════════════════════════════════════
 * UTILITY FUNCTIONS
 * ════════════════════════════════════════════════════════════════ */

/**
 * Clear all data caches (call after major data changes)
 */
function WebApp_Data_clearCache() {
  var cache = CacheService.getScriptCache();
  // Can't enumerate cache keys, so clear known patterns
  var keysToRemove = [
    'WEBAPP_STATS_ADMIN',
    'WEBAPP_LEADS_p0_sms_ADMIN',
    'WEBAPP_LEADS_p0_call_ADMIN',
    'WEBAPP_LEADS_p1_call_ADMIN',
    'WEBAPP_LEADS_p35_call_ADMIN',
    'WEBAPP_LEADS_all_leads_ADMIN'
  ];
  
  try {
    cache.removeAll(keysToRemove);
  } catch (e) {}
}
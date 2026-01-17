/**************************************************************
 * 🚢 SAFE SHIP CONTACT CHECKER PRO — REP WEB APP
 * File: WebApp_Main.gs
 *
 * OPTIMIZED VERSION
 * - Batched initial load (single API call)
 * - Reduced round trips
 *
 * v2.0 — Performance Optimized
 **************************************************************/

/* ════════════════════════════════════════════════════════════════
 * WEB APP ENTRY POINT
 * ════════════════════════════════════════════════════════════════ */

/**
 * Main entry point for the web app
 */
function doGet(e) {
  try {
    var template = HtmlService.createTemplateFromFile('ui/Index');
    template.params = e ? e.parameter : {};
    
    var html = template.evaluate()
      .setTitle('🚢 Safe Ship Rep Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
    
    return html;
    
  } catch (error) {
    Logger.log('doGet error: ' + error.message);
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:sans-serif;padding:40px;text-align:center;">' +
      '<h1 style="color:#ef4444;">⚠️ Error Loading App</h1>' +
      '<p>' + error.message + '</p></body></html>'
    ).setTitle('Error');
  }
}

/**
 * Include HTML partials
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ════════════════════════════════════════════════════════════════
 * OPTIMIZED: BATCHED INITIAL LOAD
 * Single API call returns everything needed for startup
 * ════════════════════════════════════════════════════════════════ */

/**
 * Get everything for initial app load in ONE call
 * Eliminates 4 round trips: context + stats + leads + templates
 * @returns {Object} Complete initial state
 */
function WEBAPP_getInitialLoad() {
  var startTime = new Date().getTime();
  
  try {
    // Get context first (needed for everything else)
    var context = WebApp_Security_getRepContext();
    
    if (!context.authorized) {
      return {
        authorized: false,
        reason: context.reason,
        context: context
      };
    }
    
    // Batch fetch everything with individual error handling
    var result = {
      authorized: true,
      context: context,
      stats: null,
      leads: null,
      templates: [],
      lastSync: context.lastSync,
      loadTimeMs: 0,
      errors: []
    };
    
    // Get stats
    try {
      result.stats = WebApp_Data_getDashboardStats(context.username, context.isAdmin);
    } catch (e) {
      result.errors.push('Stats: ' + e.message);
      result.stats = { p0_sms: 0, p0_call: 0, p1_call: 0, p35_call: 0, total_leads: 0, worked_today: 0 };
    }
    
    // Get leads
    try {
      result.leads = WebApp_Data_getRepLeads(context.username, { 
        tab: 'p0_sms', 
        page: 1, 
        pageSize: 20 
      }, context.isAdmin);
    } catch (e) {
      result.errors.push('Leads: ' + e.message);
      result.leads = { leads: [], total: 0, timing: {} };
    }
    
    // Get templates
    try {
      result.templates = WebApp_Templates_getAll();
    } catch (e) {
      result.errors.push('Templates: ' + e.message);
      result.templates = [];
    }
    
    result.loadTimeMs = new Date().getTime() - startTime;
    return result;
    
  } catch (error) {
    Logger.log('WEBAPP_getInitialLoad critical error: ' + error.message);
    return {
      authorized: false,
      reason: 'Error loading app: ' + error.message,
      context: null
    };
  }
}

/* ════════════════════════════════════════════════════════════════
 * PUBLIC API FUNCTIONS
 * ════════════════════════════════════════════════════════════════ */

/**
 * Get the current user's context
 */
function WEBAPP_getRepContext() {
  return WebApp_Security_getRepContext();
}

/**
 * Get leads for the current rep
 * @param {Object} options - { tab, filters, sort, page, pageSize }
 */
function WEBAPP_getRepLeads(options) {
  var context = WebApp_Security_getRepContext();
  if (!context.authorized) {
    throw new Error('Unauthorized: ' + context.reason);
  }
  
  // Default page size is now 20 for performance
  options = options || {};
  options.pageSize = options.pageSize || 20;
  
  return WebApp_Data_getRepLeads(context.username, options, context.isAdmin);
}

/**
 * Get audit counters for phones (for lazy loading)
 */
function WEBAPP_getAuditCounters(phones) {
  var context = WebApp_Security_getRepContext();
  if (!context.authorized) {
    throw new Error('Unauthorized');
  }
  
  return WebApp_Audit_getCounters(phones);
}

/**
 * Generate an SMS draft
 */
function WEBAPP_generateSmsDraft(lead, templateId, options) {
  var context = WebApp_Security_getRepContext();
  if (!context.authorized) {
    throw new Error('Unauthorized');
  }
  
  return WebApp_Templates_generate(lead, templateId, context, options);
}

/**
 * Send an SMS via RingCentral
 */
function WEBAPP_sendSms(jobNo, phone, message, leadInfo) {
  var context = WebApp_Security_getRepContext();
  if (!context.authorized) {
    throw new Error('Unauthorized');
  }
  
  return WebApp_RingCentral_sendSms(context, jobNo, phone, message, leadInfo);
}

/**
 * Send bulk SMS
 */
function WEBAPP_sendBulkSms(items) {
  var context = WebApp_Security_getRepContext();
  if (!context.authorized) {
    throw new Error('Unauthorized');
  }
  
  if (!context.isAdmin && items.length > 10) {
    throw new Error('Bulk send limited to 10 for non-admins');
  }
  
  return WebApp_RingCentral_sendBulk(context, items);
}

/**
 * Refresh RingCentral sync
 */
function WEBAPP_refreshRcSync() {
  var context = WebApp_Security_getRepContext();
  if (!context.authorized) {
    throw new Error('Unauthorized');
  }
  
  // Clear data caches after sync (safe calls)
  try {
    if (typeof WebApp_Data_clearCache === 'function') WebApp_Data_clearCache();
  } catch (e) {}
  try {
    if (typeof WebApp_Audit_clearCache === 'function') WebApp_Audit_clearCache();
  } catch (e) {}
  
  return WebApp_RingCentral_refreshSync();
}

/**
 * Get activity log
 */
function WEBAPP_getActivityLog(limit) {
  var context = WebApp_Security_getRepContext();
  if (!context.authorized) {
    throw new Error('Unauthorized');
  }
  
  return WebApp_SMSLog_getRepActivity(context.username, limit || 50);
}

/**
 * Get templates (rarely changes, could be cached client-side)
 */
function WEBAPP_getTemplates() {
  return WebApp_Templates_getAll();
}

/**
 * Admin: Send test SMS
 */
function WEBAPP_adminTestSms(phone, message) {
  var context = WebApp_Security_getRepContext();
  if (!context.isAdmin) {
    throw new Error('Admin access required');
  }
  
  return WebApp_RingCentral_sendSms(context, 'TEST', phone, message, { firstName: 'Test', lastName: 'Admin' });
}

/**
 * Get dashboard stats
 */
function WEBAPP_getDashboardStats() {
  var context = WebApp_Security_getRepContext();
  if (!context.authorized) {
    throw new Error('Unauthorized');
  }
  
  return WebApp_Data_getDashboardStats(context.username, context.isAdmin);
}

/**
 * Clear all caches (admin only)
 */
function WEBAPP_clearCaches() {
  var context = WebApp_Security_getRepContext();
  if (!context.isAdmin) {
    throw new Error('Admin access required');
  }
  
  try {
    if (typeof WebApp_Data_clearCache === 'function') WebApp_Data_clearCache();
  } catch (e) {}
  try {
    if (typeof WebApp_Audit_clearCache === 'function') WebApp_Audit_clearCache();
  } catch (e) {}
  
  return { success: true, message: 'All caches cleared' };
}
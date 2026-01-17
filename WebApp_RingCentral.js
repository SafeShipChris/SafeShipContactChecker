/**************************************************************
 * 🚢 SAFE SHIP CONTACT CHECKER PRO — REP WEB APP
 * File: WebApp_RingCentral.gs
 *
 * Handles SMS sending via RingCentral API.
 * Uses existing project OAuth2 integration if available.
 *
 * Features:
 *   - Send individual SMS
 *   - Send bulk SMS (with rate limiting)
 *   - Log all attempts to WebApp_SMS_Log sheet
 *   - Handle error states gracefully
 *
 * v1.0 — Production Build
 **************************************************************/

/* ════════════════════════════════════════════════════════════════
 * CONFIGURATION
 * ════════════════════════════════════════════════════════════════ */

var WEBAPP_RC_CONFIG = {
  // RingCentral API endpoints
  API_BASE_PRODUCTION: 'https://platform.ringcentral.com',
  API_BASE_SANDBOX: 'https://platform.devtest.ringcentral.com',
  
  // SMS endpoint
  SMS_ENDPOINT: '/restapi/v1.0/account/~/extension/~/sms',
  
  // Default from number (can be overridden by rep's extension)
  DEFAULT_FROM_NUMBER: '', // Will be set from script properties
  
  // Rate limiting
  SMS_DELAY_MS: 500,      // Delay between bulk SMS
  MAX_BULK_SMS: 50,       // Max SMS per bulk operation
  
  // Script property keys
  PROP_ACCESS_TOKEN: 'RC_ACCESS_TOKEN',
  PROP_REFRESH_TOKEN: 'RC_REFRESH_TOKEN',
  PROP_TOKEN_EXPIRES: 'RC_TOKEN_EXPIRES',
  PROP_FROM_NUMBER: 'RC_FROM_NUMBER',
  PROP_USE_SANDBOX: 'RC_USE_SANDBOX',
  
  // OAuth settings (if not using existing OAuth2 library)
  CLIENT_ID_PROP: 'RC_CLIENT_ID',
  CLIENT_SECRET_PROP: 'RC_CLIENT_SECRET'
};

/* ════════════════════════════════════════════════════════════════
 * MAIN SMS FUNCTIONS
 * ════════════════════════════════════════════════════════════════ */

/**
 * Send an SMS via RingCentral
 * @param {Object} context - Rep context (from security module)
 * @param {string} jobNo - Job number for logging
 * @param {string} phone - Phone number to send to
 * @param {string} message - Message text
 * @param {Object} leadInfo - Lead info for logging (firstName, lastName, etc.)
 * @returns {Object} { ok: boolean, status: string, messageId: string, error: string, details: object }
 */
function WebApp_RingCentral_sendSms(context, jobNo, phone, message, leadInfo) {
  var result = {
    ok: false,
    status: 'PENDING',
    messageId: null,
    error: null,
    details: null,
    timestamp: new Date().toISOString()
  };
  
  try {
    // Validate inputs
    if (!phone) {
      throw new Error('Phone number is required');
    }
    
    if (!message || message.trim().length === 0) {
      throw new Error('Message is required');
    }
    
    // Normalize phone number
    var normalizedPhone = WebApp_RingCentral_normalizePhone_(phone);
    if (!normalizedPhone || normalizedPhone.length < 10) {
      throw new Error('Invalid phone number format');
    }
    
    // Format phone for RingCentral (with +1 prefix)
    var formattedPhone = '+1' + normalizedPhone;
    
    // Get access token
    var accessToken = WebApp_RingCentral_getAccessToken_();
    if (!accessToken) {
      throw new Error('RingCentral not authenticated. Please contact administrator.');
    }
    
    // Get from number
    var fromNumber = WebApp_RingCentral_getFromNumber_(context);
    if (!fromNumber) {
      throw new Error('No RingCentral sending number configured');
    }
    
    // Build API URL
    var apiBase = WebApp_RingCentral_getApiBase_();
    var url = apiBase + WEBAPP_RC_CONFIG.SMS_ENDPOINT;
    
    // Build request payload
    var payload = {
      from: { phoneNumber: fromNumber },
      to: [{ phoneNumber: formattedPhone }],
      text: message
    };
    
    // Make API request
    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + accessToken
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();
    
    // Parse response
    var responseData = {};
    try {
      responseData = JSON.parse(responseBody);
    } catch (e) {}
    
    if (responseCode >= 200 && responseCode < 300) {
      // Success
      result.ok = true;
      result.status = 'SENT';
      result.messageId = responseData.id || null;
      result.details = {
        creationTime: responseData.creationTime,
        messageStatus: responseData.messageStatus,
        direction: responseData.direction
      };
      
    } else if (responseCode === 401) {
      // Token expired - try to refresh
      var refreshed = WebApp_RingCentral_refreshAccessToken_();
      if (refreshed) {
        // Retry once
        return WebApp_RingCentral_sendSms(context, jobNo, phone, message, leadInfo);
      } else {
        result.status = 'AUTH_ERROR';
        result.error = 'Authentication failed. Please contact administrator.';
      }
      
    } else {
      // Other error
      result.status = 'FAILED';
      result.error = responseData.message || responseData.error_description || 'Unknown error';
      result.details = {
        httpCode: responseCode,
        errorCode: responseData.errorCode,
        errors: responseData.errors
      };
    }
    
  } catch (error) {
    result.status = 'ERROR';
    result.error = error.message;
  }
  
  // Log the attempt
  WebApp_SMSLog_logAttempt({
    timestamp: result.timestamp,
    rep: context.username,
    repEmail: context.email,
    jobNo: jobNo,
    phone: phone,
    leadName: (leadInfo.firstName || '') + ' ' + (leadInfo.lastName || ''),
    message: message,
    status: result.status,
    messageId: result.messageId,
    error: result.error
  });
  
  return result;
}

/**
 * Send bulk SMS to multiple leads
 * @param {Object} context - Rep context
 * @param {Array} items - Array of { jobNo, phone, message, leadInfo }
 * @returns {Object} { sent: number, failed: number, results: [] }
 */
function WebApp_RingCentral_sendBulk(context, items) {
  var results = {
    sent: 0,
    failed: 0,
    results: []
  };
  
  if (!items || items.length === 0) {
    return results;
  }
  
  // Limit bulk size
  var toProcess = items.slice(0, WEBAPP_RC_CONFIG.MAX_BULK_SMS);
  
  for (var i = 0; i < toProcess.length; i++) {
    var item = toProcess[i];
    
    try {
      var result = WebApp_RingCentral_sendSms(
        context,
        item.jobNo,
        item.phone,
        item.message,
        item.leadInfo || {}
      );
      
      results.results.push({
        jobNo: item.jobNo,
        phone: item.phone,
        ok: result.ok,
        status: result.status,
        error: result.error
      });
      
      if (result.ok) {
        results.sent++;
      } else {
        results.failed++;
      }
      
      // Rate limiting delay
      if (i < toProcess.length - 1) {
        Utilities.sleep(WEBAPP_RC_CONFIG.SMS_DELAY_MS);
      }
      
    } catch (e) {
      results.failed++;
      results.results.push({
        jobNo: item.jobNo,
        phone: item.phone,
        ok: false,
        status: 'ERROR',
        error: e.message
      });
    }
  }
  
  return results;
}

/**
 * Refresh RingCentral sync (calls existing sync function)
 * @returns {Object} { success: boolean, message: string, lastSync: string }
 */
function WebApp_RingCentral_refreshSync() {
  var result = {
    success: false,
    message: '',
    lastSync: null
  };
  
  try {
    // Try to call existing sync functions
    if (typeof RC_syncTodayCallsAndSMS === 'function') {
      RC_syncTodayCallsAndSMS();
      result.success = true;
      result.message = 'Sync completed (calls + SMS)';
    } else if (typeof SSCCP_RC_syncToday === 'function') {
      SSCCP_RC_syncToday();
      result.success = true;
      result.message = 'Sync completed';
    } else {
      result.message = 'No sync function available. RC logs may be stale.';
    }
    
    // Clear audit cache after sync
    WebApp_Audit_clearCache();
    
    // Update last sync time
    var now = new Date().toISOString();
    PropertiesService.getScriptProperties().setProperty('WEBAPP_LAST_RC_SYNC', now);
    result.lastSync = now;
    
  } catch (error) {
    result.message = 'Sync error: ' + error.message;
    Logger.log('WebApp_RingCentral_refreshSync error: ' + error.message);
  }
  
  return result;
}

/* ════════════════════════════════════════════════════════════════
 * AUTHENTICATION HELPERS
 * ════════════════════════════════════════════════════════════════ */

/**
 * Get RingCentral access token
 * Uses existing OAuth2 library if available, otherwise script properties
 * @returns {string|null} Access token or null
 */
function WebApp_RingCentral_getAccessToken_() {
  try {
    // Try existing OAuth2 service (from project)
    if (typeof getRingCentralService === 'function') {
      var service = getRingCentralService();
      if (service && service.hasAccess()) {
        return service.getAccessToken();
      }
    }
    
    // Fall back to script properties
    var props = PropertiesService.getScriptProperties();
    var token = props.getProperty(WEBAPP_RC_CONFIG.PROP_ACCESS_TOKEN);
    var expires = props.getProperty(WEBAPP_RC_CONFIG.PROP_TOKEN_EXPIRES);
    
    if (token && expires) {
      var expiresDate = new Date(parseInt(expires, 10));
      if (expiresDate > new Date()) {
        return token;
      } else {
        // Token expired, try to refresh
        return WebApp_RingCentral_refreshAccessToken_();
      }
    }
    
    return null;
    
  } catch (e) {
    Logger.log('Error getting access token: ' + e.message);
    return null;
  }
}

/**
 * Refresh the RingCentral access token
 * @returns {string|null} New access token or null
 */
function WebApp_RingCentral_refreshAccessToken_() {
  try {
    // Try existing OAuth2 service
    if (typeof getRingCentralService === 'function') {
      var service = getRingCentralService();
      if (service) {
        service.refresh();
        if (service.hasAccess()) {
          return service.getAccessToken();
        }
      }
    }
    
    // Manual refresh using script properties
    var props = PropertiesService.getScriptProperties();
    var refreshToken = props.getProperty(WEBAPP_RC_CONFIG.PROP_REFRESH_TOKEN);
    var clientId = props.getProperty(WEBAPP_RC_CONFIG.CLIENT_ID_PROP);
    var clientSecret = props.getProperty(WEBAPP_RC_CONFIG.CLIENT_SECRET_PROP);
    
    if (!refreshToken || !clientId || !clientSecret) {
      Logger.log('Missing refresh token or credentials');
      return null;
    }
    
    var apiBase = WebApp_RingCentral_getApiBase_();
    var tokenUrl = apiBase + '/restapi/oauth/token';
    
    var response = UrlFetchApp.fetch(tokenUrl, {
      method: 'post',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSecret),
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      payload: 'grant_type=refresh_token&refresh_token=' + encodeURIComponent(refreshToken),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      var data = JSON.parse(response.getContentText());
      
      // Store new tokens
      props.setProperty(WEBAPP_RC_CONFIG.PROP_ACCESS_TOKEN, data.access_token);
      props.setProperty(WEBAPP_RC_CONFIG.PROP_REFRESH_TOKEN, data.refresh_token);
      props.setProperty(WEBAPP_RC_CONFIG.PROP_TOKEN_EXPIRES, 
        String(Date.now() + (data.expires_in * 1000)));
      
      return data.access_token;
    }
    
    return null;
    
  } catch (e) {
    Logger.log('Error refreshing token: ' + e.message);
    return null;
  }
}

/**
 * Get the RingCentral API base URL
 * @returns {string} API base URL
 */
function WebApp_RingCentral_getApiBase_() {
  var props = PropertiesService.getScriptProperties();
  var useSandbox = props.getProperty(WEBAPP_RC_CONFIG.PROP_USE_SANDBOX) === 'true';
  
  return useSandbox 
    ? WEBAPP_RC_CONFIG.API_BASE_SANDBOX 
    : WEBAPP_RC_CONFIG.API_BASE_PRODUCTION;
}

/**
 * Get the "from" phone number for sending SMS
 * @param {Object} context - Rep context
 * @returns {string} Formatted phone number with +1
 */
function WebApp_RingCentral_getFromNumber_(context) {
  // First try rep's RC extension number
  if (context.rcExtension) {
    return context.rcExtension;
  }
  
  // Fall back to configured default
  var props = PropertiesService.getScriptProperties();
  var fromNumber = props.getProperty(WEBAPP_RC_CONFIG.PROP_FROM_NUMBER);
  
  if (fromNumber) {
    // Ensure proper format
    var digits = fromNumber.replace(/\D/g, '');
    if (digits.length >= 10) {
      return '+1' + digits.slice(-10);
    }
  }
  
  return null;
}

/**
 * Normalize phone number
 * @param {string} phone - Phone number
 * @returns {string} Normalized 10-digit phone
 */
function WebApp_RingCentral_normalizePhone_(phone) {
  if (!phone) return '';
  var digits = String(phone).replace(/\D/g, '');
  return digits.length >= 10 ? digits.slice(-10) : digits;
}

/* ════════════════════════════════════════════════════════════════
 * ADMIN FUNCTIONS
 * ════════════════════════════════════════════════════════════════ */

/**
 * Check RingCentral connection status
 * @returns {Object} Connection status
 */
function WebApp_RingCentral_checkStatus() {
  var status = {
    authenticated: false,
    tokenExpires: null,
    fromNumber: null,
    apiBase: WebApp_RingCentral_getApiBase_()
  };
  
  try {
    var token = WebApp_RingCentral_getAccessToken_();
    status.authenticated = !!token;
    
    var props = PropertiesService.getScriptProperties();
    var expires = props.getProperty(WEBAPP_RC_CONFIG.PROP_TOKEN_EXPIRES);
    if (expires) {
      status.tokenExpires = new Date(parseInt(expires, 10)).toISOString();
    }
    
    status.fromNumber = props.getProperty(WEBAPP_RC_CONFIG.PROP_FROM_NUMBER);
    
  } catch (e) {
    status.error = e.message;
  }
  
  return status;
}

/**
 * Set the default from number (admin function)
 * @param {string} phoneNumber - Phone number to use
 */
function WebApp_RingCentral_setFromNumber(phoneNumber) {
  var normalized = WebApp_RingCentral_normalizePhone_(phoneNumber);
  if (normalized.length >= 10) {
    PropertiesService.getScriptProperties()
      .setProperty(WEBAPP_RC_CONFIG.PROP_FROM_NUMBER, '+1' + normalized);
    return { success: true, number: '+1' + normalized };
  }
  return { success: false, error: 'Invalid phone number' };
}
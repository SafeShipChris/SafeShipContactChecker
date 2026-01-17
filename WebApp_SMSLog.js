/**************************************************************
 * 🚢 SAFE SHIP CONTACT CHECKER PRO — REP WEB APP
 * File: WebApp_SMSLog.gs
 *
 * Handles logging of all SMS attempts from the web app.
 * Provides activity log queries for reps and admins.
 *
 * Log Sheet: "WebApp_SMS_Log"
 * Columns: Timestamp, Rep, RepEmail, JobNo, Phone, LeadName, 
 *          Message, Status, MessageId, Error
 *
 * v1.0 — Production Build
 **************************************************************/

/* ════════════════════════════════════════════════════════════════
 * CONFIGURATION
 * ════════════════════════════════════════════════════════════════ */

var WEBAPP_SMSLOG_CONFIG = {
  SHEET_NAME: 'WebApp_SMS_Log',
  
  // Column headers (will auto-create if sheet doesn't exist)
  HEADERS: [
    'Timestamp',
    'Rep',
    'Rep Email',
    'Job No',
    'Phone',
    'Lead Name',
    'Message',
    'Char Count',
    'Segments',
    'Status',
    'Message ID',
    'Error',
    'Template Used'
  ],
  
  // Max rows to keep (for cleanup)
  MAX_ROWS: 10000,
  
  // Statuses
  STATUS: {
    SENT: 'SENT',
    FAILED: 'FAILED',
    ERROR: 'ERROR',
    PENDING: 'PENDING'
  }
};

/* ════════════════════════════════════════════════════════════════
 * LOGGING FUNCTIONS
 * ════════════════════════════════════════════════════════════════ */

/**
 * Log an SMS attempt
 * @param {Object} data - Log entry data
 */
function WebApp_SMSLog_logAttempt(data) {
  try {
    var sheet = WebApp_SMSLog_getOrCreateSheet_();
    
    var row = [
      data.timestamp || new Date().toISOString(),
      data.rep || '',
      data.repEmail || '',
      data.jobNo || '',
      data.phone || '',
      data.leadName || '',
      data.message || '',
      (data.message || '').length,
      Math.ceil((data.message || '').length / 160),
      data.status || WEBAPP_SMSLOG_CONFIG.STATUS.PENDING,
      data.messageId || '',
      data.error || '',
      data.templateId || ''
    ];
    
    // Insert at row 2 (after headers) so newest is at top
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, row.length).setValues([row]);
    
    // Cleanup old rows if needed
    WebApp_SMSLog_cleanup_(sheet);
    
  } catch (error) {
    Logger.log('Error logging SMS attempt: ' + error.message);
  }
}

/**
 * Get activity log for a rep
 * @param {string} username - Rep's username
 * @param {number} limit - Max rows to return
 * @returns {Array} Log entries
 */
function WebApp_SMSLog_getRepActivity(username, limit) {
  limit = limit || 50;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(WEBAPP_SMSLOG_CONFIG.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    var data = sheet.getRange(2, 1, Math.min(sheet.getLastRow() - 1, 500), 13).getValues();
    var normalizedUsername = (username || '').toUpperCase().trim();
    
    var entries = [];
    
    for (var i = 0; i < data.length && entries.length < limit; i++) {
      var row = data[i];
      var rowRep = String(row[1] || '').toUpperCase().trim();
      
      if (rowRep === normalizedUsername) {
        entries.push({
          timestamp: row[0],
          rep: row[1],
          jobNo: row[3],
          phone: row[4],
          leadName: row[5],
          messagePreview: String(row[6] || '').substring(0, 50) + '...',
          charCount: row[7],
          status: row[9],
          messageId: row[10],
          error: row[11]
        });
      }
    }
    
    return entries;
    
  } catch (error) {
    Logger.log('Error getting rep activity: ' + error.message);
    return [];
  }
}

/**
 * Count SMS sent today by a rep
 * @param {string} username - Rep's username
 * @returns {number} Count of SMS sent today
 */
function WebApp_SMSLog_countTodayByRep(username) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(WEBAPP_SMSLOG_CONFIG.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return 0;
    }
    
    // Read timestamp and rep columns only
    var data = sheet.getRange(2, 1, Math.min(sheet.getLastRow() - 1, 500), 2).getValues();
    var normalizedUsername = (username || '').toUpperCase().trim();
    
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    
    var count = 0;
    
    for (var i = 0; i < data.length; i++) {
      var timestamp = data[i][0];
      var rep = String(data[i][1] || '').toUpperCase().trim();
      
      // Check if today
      var rowDate = timestamp instanceof Date ? timestamp : new Date(timestamp);
      if (isNaN(rowDate.getTime())) continue;
      
      rowDate.setHours(0, 0, 0, 0);
      if (rowDate.getTime() !== today.getTime()) continue;
      
      // Check if this rep
      if (rep === normalizedUsername) {
        count++;
      }
    }
    
    return count;
    
  } catch (error) {
    Logger.log('Error counting today SMS: ' + error.message);
    return 0;
  }
}

/**
 * Get all activity (admin view)
 * @param {number} limit - Max rows
 * @returns {Array} Log entries
 */
function WebApp_SMSLog_getAllActivity(limit) {
  limit = limit || 100;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(WEBAPP_SMSLOG_CONFIG.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    var numRows = Math.min(sheet.getLastRow() - 1, limit);
    var data = sheet.getRange(2, 1, numRows, 13).getValues();
    
    return data.map(function(row) {
      return {
        timestamp: row[0],
        rep: row[1],
        repEmail: row[2],
        jobNo: row[3],
        phone: row[4],
        leadName: row[5],
        messagePreview: String(row[6] || '').substring(0, 50) + '...',
        charCount: row[7],
        segments: row[8],
        status: row[9],
        messageId: row[10],
        error: row[11],
        template: row[12]
      };
    });
    
  } catch (error) {
    Logger.log('Error getting all activity: ' + error.message);
    return [];
  }
}

/* ════════════════════════════════════════════════════════════════
 * INTERNAL FUNCTIONS
 * ════════════════════════════════════════════════════════════════ */

/**
 * Get or create the SMS log sheet
 * @returns {Sheet} The log sheet
 */
function WebApp_SMSLog_getOrCreateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(WEBAPP_SMSLOG_CONFIG.SHEET_NAME);
  
  if (!sheet) {
    // Create the sheet
    sheet = ss.insertSheet(WEBAPP_SMSLOG_CONFIG.SHEET_NAME);
    
    // Add headers
    sheet.getRange(1, 1, 1, WEBAPP_SMSLOG_CONFIG.HEADERS.length)
      .setValues([WEBAPP_SMSLOG_CONFIG.HEADERS])
      .setFontWeight('bold')
      .setBackground('#1e3a5f')
      .setFontColor('#ffffff');
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Set column widths
    sheet.setColumnWidth(1, 150);  // Timestamp
    sheet.setColumnWidth(2, 100);  // Rep
    sheet.setColumnWidth(3, 180);  // Email
    sheet.setColumnWidth(4, 80);   // Job No
    sheet.setColumnWidth(5, 120);  // Phone
    sheet.setColumnWidth(6, 150);  // Lead Name
    sheet.setColumnWidth(7, 300);  // Message
    sheet.setColumnWidth(8, 60);   // Char Count
    sheet.setColumnWidth(9, 60);   // Segments
    sheet.setColumnWidth(10, 80);  // Status
    sheet.setColumnWidth(11, 120); // Message ID
    sheet.setColumnWidth(12, 200); // Error
    sheet.setColumnWidth(13, 120); // Template
    
    Logger.log('Created WebApp_SMS_Log sheet');
  }
  
  return sheet;
}

/**
 * Clean up old log rows to prevent sheet bloat
 * @param {Sheet} sheet - The log sheet
 */
function WebApp_SMSLog_cleanup_(sheet) {
  try {
    var lastRow = sheet.getLastRow();
    var maxRows = WEBAPP_SMSLOG_CONFIG.MAX_ROWS;
    
    if (lastRow > maxRows + 100) {  // Add buffer before cleanup
      var rowsToDelete = lastRow - maxRows;
      sheet.deleteRows(maxRows + 1, rowsToDelete);
      Logger.log('Cleaned up ' + rowsToDelete + ' old log rows');
    }
  } catch (e) {
    Logger.log('Cleanup error: ' + e.message);
  }
}

/**
 * Get summary statistics
 * @returns {Object} Summary stats
 */
function WebApp_SMSLog_getSummary() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(WEBAPP_SMSLOG_CONFIG.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { totalSent: 0, todaySent: 0, failed: 0 };
    }
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    
    var stats = {
      totalSent: 0,
      todaySent: 0,
      failed: 0,
      byRep: {}
    };
    
    data.forEach(function(row) {
      var status = row[9];
      var timestamp = row[0];
      var rep = row[1];
      
      if (status === 'SENT') {
        stats.totalSent++;
        
        var rowDate = timestamp instanceof Date ? timestamp : new Date(timestamp);
        if (!isNaN(rowDate.getTime())) {
          rowDate.setHours(0, 0, 0, 0);
          if (rowDate.getTime() === today.getTime()) {
            stats.todaySent++;
          }
        }
        
        if (rep) {
          stats.byRep[rep] = (stats.byRep[rep] || 0) + 1;
        }
      } else if (status === 'FAILED' || status === 'ERROR') {
        stats.failed++;
      }
    });
    
    return stats;
    
  } catch (error) {
    Logger.log('Error getting summary: ' + error.message);
    return { totalSent: 0, todaySent: 0, failed: 0 };
  }
}
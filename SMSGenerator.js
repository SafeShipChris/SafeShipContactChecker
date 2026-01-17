/**************************************************************
 * 🚢 SAFE SHIP — SMS GENERATOR
 * File: SMSGenerator.gs (Add this to your Google Apps Script project)
 *
 * DO NOT use the .jsx file - that's a React demo only!
 * This file + SMSGeneratorUI.html = complete solution
 **************************************************************/

/**
 * ═══════════════════════════════════════════════════════════════
 * WEB APP ENTRY POINT
 * Deploy as: Web App (Execute as: Me, Access: Anyone)
 * ═══════════════════════════════════════════════════════════════
 */

function doGet_SMSGenerator(e) {
  var leadId = (e && e.parameter && e.parameter.leadId) ? e.parameter.leadId : '';
  var repUsername = (e && e.parameter && e.parameter.rep) ? e.parameter.rep : '';
  
  var template = HtmlService.createTemplateFromFile('SMSGeneratorUI');
  template.leadId = leadId;
  template.repUsername = repUsername;
  
  return template.evaluate()
    .setTitle('Safe Ship SMS Generator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * ═══════════════════════════════════════════════════════════════
 * SERVER-SIDE FUNCTIONS (Called from HTML via google.script.run)
 * ═══════════════════════════════════════════════════════════════
 */

/**
 * Get lead data by job number
 */
function getLeadData(jobNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var granotSheet = ss.getSheetByName('GRANOT DATA');
  
  if (!granotSheet) return null;
  
  var data = granotSheet.getDataRange().getValues();
  var headers = data[0];
  
  // Find column indices
  var cols = {
    jobNo: findColumnIndex_(headers, ['job_no', 'job']),
    firstName: findColumnIndex_(headers, ['first_name', 'firstname']),
    lastName: findColumnIndex_(headers, ['last_name', 'lastname']),
    phone1: findColumnIndex_(headers, ['phone1', 'phone']),
    phone2: findColumnIndex_(headers, ['phone2']),
    fromState: findColumnIndex_(headers, ['from_state']),
    toState: findColumnIndex_(headers, ['to_state']),
    moveDate: findColumnIndex_(headers, ['move_date', 'movedate']),
    cf: findColumnIndex_(headers, ['cf', 'cubic_feet']),
    estTotal: findColumnIndex_(headers, ['est_total', 'esttotal']),
    source: findColumnIndex_(headers, ['source']),
    user: findColumnIndex_(headers, ['user', 'username'])
  };
  
  // Search for lead
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowJobNo = cols.jobNo >= 0 ? String(row[cols.jobNo]) : '';
    
    if (rowJobNo === String(jobNo)) {
      return {
        id: rowJobNo,
        firstName: cols.firstName >= 0 ? String(row[cols.firstName] || '') : '',
        lastName: cols.lastName >= 0 ? String(row[cols.lastName] || '') : '',
        phone: cols.phone1 >= 0 ? formatPhone_(row[cols.phone1]) : '',
        phone2: cols.phone2 >= 0 ? formatPhone_(row[cols.phone2]) : '',
        fromState: cols.fromState >= 0 ? String(row[cols.fromState] || '') : '',
        toState: cols.toState >= 0 ? String(row[cols.toState] || '') : '',
        moveDate: cols.moveDate >= 0 && row[cols.moveDate] ? Utilities.formatDate(new Date(row[cols.moveDate]), 'America/New_York', 'yyyy-MM-dd') : '',
        cubicFeet: cols.cf >= 0 ? (row[cols.cf] || 0) : 0,
        estTotal: cols.estTotal >= 0 ? (row[cols.estTotal] || 0) : 0,
        source: cols.source >= 0 ? String(row[cols.source] || '') : '',
        repUsername: cols.user >= 0 ? String(row[cols.user] || '') : ''
      };
    }
  }
  
  return null;
}

/**
 * Get multiple leads for a rep from trackers
 */
function getLeadsForRep(repUsername, limit) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leads = [];
  var seenJobs = {};
  var maxLeads = limit || 20;
  
  // Normalize username
  var repUpper = String(repUsername || '').trim().toUpperCase();
  if (!repUpper) return leads;
  
  // Check SMS TRACKER
  var smsSheet = ss.getSheetByName('SMS TRACKER');
  if (smsSheet) {
    var smsData = smsSheet.getRange('A3:M' + Math.min(smsSheet.getLastRow(), 500)).getValues();
    for (var i = 0; i < smsData.length && leads.length < maxLeads; i++) {
      var row = smsData[i];
      var username = String(row[9] || '').trim().toUpperCase(); // USER column (index 9)
      var jobNo = String(row[8] || '').trim(); // JOB # column (index 8)
      
      if (username === repUpper && jobNo && !seenJobs[jobNo]) {
        var leadData = getLeadData(jobNo);
        if (leadData) {
          leadData.trackerType = 'SMS';
          leads.push(leadData);
          seenJobs[jobNo] = true;
        }
      }
    }
  }
  
  // Check CALL TRACKER
  var callSheet = ss.getSheetByName('CALL & VOICEMAIL TRACKER');
  if (callSheet) {
    var callData = callSheet.getRange('A3:M' + Math.min(callSheet.getLastRow(), 500)).getValues();
    for (var j = 0; j < callData.length && leads.length < maxLeads; j++) {
      var row2 = callData[j];
      var username2 = String(row2[9] || '').trim().toUpperCase();
      var jobNo2 = String(row2[8] || '').trim();
      
      if (username2 === repUpper && jobNo2 && !seenJobs[jobNo2]) {
        var leadData2 = getLeadData(jobNo2);
        if (leadData2) {
          leadData2.trackerType = 'CALL';
          leads.push(leadData2);
          seenJobs[jobNo2] = true;
        }
      }
    }
  }
  
  return leads;
}

/**
 * Get rep's first name from roster
 */
function getRepFirstName(username) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rosterSheet = ss.getSheetByName('Sales_Roster');
  
  if (!rosterSheet) return username;
  
  var data = rosterSheet.getDataRange().getValues();
  var headers = data[0];
  var headersLower = headers.map(function(h) { return String(h).toLowerCase(); });
  
  var usernameIdx = findColumnIndex_(headersLower, ['username', 'user']);
  var firstNameIdx = findColumnIndex_(headersLower, ['first name', 'firstname']);
  
  if (usernameIdx < 0 || firstNameIdx < 0) return username;
  
  var targetUser = String(username).toUpperCase();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][usernameIdx]).toUpperCase() === targetUser) {
      return String(data[i][firstNameIdx] || username);
    }
  }
  
  return username;
}

/**
 * Get SMS templates
 */
function getSMSTemplates() {
  return {
    initial: {
      name: '👋 Initial Outreach',
      icon: '👋',
      template: 'Hi {firstName}! This is {repFirstName} from Safe Ship Moving. I see you\'re planning a move from {fromState} to {toState} around {moveDateFormatted}. I\'d love to help make your move stress-free! When\'s a good time to chat for 5 minutes?'
    },
    followUp: {
      name: '🔄 Follow Up',
      icon: '🔄',
      template: 'Hi {firstName}, just following up on your upcoming move to {toState}. I have some great options that could save you money on your {moveDateFormatted} move. Do you have a few minutes to discuss? - {repFirstName}, Safe Ship'
    },
    quote: {
      name: '💰 Quote Ready',
      icon: '💰',
      template: 'Great news {firstName}! I\'ve put together a competitive quote for your {fromState} to {toState} move. Your estimated total is around ${estTotal} for {cubicFeet} cubic feet. Want me to walk you through the details? - {repFirstName}'
    },
    urgent: {
      name: '🔥 Urgent/Hot Lead',
      icon: '🔥',
      template: 'Hi {firstName}! I noticed your move date ({moveDateFormatted}) is coming up fast! I want to make sure we get you locked in with the best carrier. Can we connect today? I have a 5-minute window at 2pm or 4pm. - {repFirstName}, Safe Ship'
    },
    checkIn: {
      name: '📞 Check In',
      icon: '📞',
      template: 'Hey {firstName}, hope you\'re doing well! Just checking in on your move plans. Have you had a chance to think about the quote I sent? Happy to answer any questions. - {repFirstName}'
    },
    lastChance: {
      name: '⏰ Last Chance',
      icon: '⏰',
      template: '{firstName}, I want to make sure you\'re taken care of for your {moveDateFormatted} move. Carrier space fills up fast this time of year. If you\'re still interested, let\'s chat today - I can hold your spot. - {repFirstName}, Safe Ship'
    },
    reEngage: {
      name: '🔔 Re-engagement',
      icon: '🔔',
      template: 'Hi {firstName}! It\'s been a little while since we chatted about your move. Are you still planning to relocate to {toState}? If your plans have changed, no worries - just let me know! - {repFirstName}'
    },
    valueAdd: {
      name: '💎 Value Add',
      icon: '💎',
      template: '{firstName}, quick tip for your {fromState} to {toState} move: booking 2+ weeks ahead typically saves 10-15%. Your move date is {moveDateFormatted} - want me to lock in current rates? - {repFirstName}, Safe Ship'
    }
  };
}

/**
 * Log SMS generation for analytics
 */
function logSMSGeneration(leadId, templateUsed, repUsername) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName('SMS_Generation_Log');
    
    if (!logSheet) {
      logSheet = ss.insertSheet('SMS_Generation_Log');
      logSheet.appendRow(['Timestamp', 'Lead ID', 'Template', 'Rep', 'Action']);
      logSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    }
    
    logSheet.appendRow([
      new Date(),
      leadId || '',
      templateUsed || '',
      repUsername || '',
      'Generated'
    ]);
  } catch (e) {
    console.error('Error logging SMS generation:', e);
  }
}

/**
 * ═══════════════════════════════════════════════════════════════
 * HELPER FUNCTIONS
 * ═══════════════════════════════════════════════════════════════
 */

function findColumnIndex_(headers, possibleNames) {
  for (var i = 0; i < possibleNames.length; i++) {
    var name = possibleNames[i].toLowerCase();
    for (var j = 0; j < headers.length; j++) {
      if (String(headers[j]).toLowerCase() === name) {
        return j;
      }
    }
  }
  return -1;
}

function formatPhone_(phone) {
  if (!phone) return '';
  var digits = String(phone).replace(/\D/g, '');
  if (digits.length === 10) {
    return '(' + digits.substr(0, 3) + ') ' + digits.substr(3, 3) + '-' + digits.substr(6);
  }
  if (digits.length === 11 && digits[0] === '1') {
    return '(' + digits.substr(1, 3) + ') ' + digits.substr(4, 3) + '-' + digits.substr(7);
  }
  return String(phone);
}

/**
 * ═══════════════════════════════════════════════════════════════
 * MENU INTEGRATION (Add to your MainMenu.gs)
 * ═══════════════════════════════════════════════════════════════
 */

function openSMSGeneratorSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('SMSGeneratorUI')
    .setTitle('SMS Generator')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function openSMSGeneratorPopup() {
  var html = HtmlService.createHtmlOutputFromFile('SMSGeneratorUI')
    .setWidth(500)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '💬 SMS Generator');
}

/**
 * Generate URL for SMS Generator (use in email links)
 */
function getSMSGeneratorURL(jobNo, repUsername) {
  var scriptUrl = ScriptApp.getService().getUrl();
  if (!scriptUrl) {
    // Fallback - you'll need to deploy and get the URL
    scriptUrl = 'YOUR_DEPLOYED_WEB_APP_URL';
  }
  return scriptUrl + '?leadId=' + encodeURIComponent(jobNo) + '&rep=' + encodeURIComponent(repUsername);
}

/**
 * Test function - run this to verify setup
 */
function testSMSGenerator() {
  var ui = SpreadsheetApp.getUi();
  
  // Test getting templates
  var templates = getSMSTemplates();
  var templateCount = Object.keys(templates).length;
  
  // Test getting a lead (use a real job number from your data)
  var testLead = getLeadData('7142309'); // Change to a real job number
  
  var message = '🚢 SMS Generator Test Results:\n\n';
  message += '✅ Templates loaded: ' + templateCount + '\n';
  message += '✅ Lead lookup: ' + (testLead ? 'Working (' + testLead.firstName + ')' : 'No lead found (check job #)') + '\n';
  message += '\nNext step: Create SMSGeneratorUI.html file';
  
  ui.alert('SMS Generator Test', message, ui.ButtonSet.OK);
}
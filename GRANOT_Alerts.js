/**************************************************************
 * 🚢 SAFE SHIP CONTACT CHECKER — GRANOT_Alerts.gs v2.0
 * ═══════════════════════════════════════════════════════════
 * 
 * GRANOT Lead Alert System - Full Featured
 * 
 * FEATURES:
 * ✅ Multi-select Priority filter (0-8)
 * ✅ Date range filter (open_date)
 * ✅ Move date proximity filter
 * ✅ Lead age filter (stale leads)
 * ✅ Call/SMS count thresholds
 * ✅ Est. Total (deal value) filter
 * ✅ Source filter
 * ✅ Save/Load filter presets
 * ✅ Rep mapping via Sales_Roster/Team_Roster
 * ✅ CC Manager option
 * ✅ Team-based reports for managers
 * ✅ Admin health report with issues
 * ✅ Compliance scoring per rep
 * ✅ Highlight flags (Hot, High Value, Stale)
 * ✅ Same visual Lead Card style as Contact Activity
 * 
 **************************************************************/

/**
 * ═══════════════════════════════════════════════════════════════
 * CONFIGURATION
 * ═══════════════════════════════════════════════════════════════
 */
var GRANOT_CONFIG = {
  // Sheet names
  SHEET_NAME: "GRANOT DATA",
  SALES_ROSTER: "Sales_Roster",
  TEAM_ROSTER: "Team_Roster",
  PRESETS_KEY: "GRANOT_ALERT_PRESETS",
  
  // Column indices (0-based) for GRANOT DATA
  COLS: {
    DEPT: 0, JOB_NO: 1, OPEN_DATE: 2, OPEN_TIME: 3, STATUS: 4,
    USER: 5, JOB_REP: 6, FOLLOW_UP: 7, PRIORITY: 8,
    FIRST_NAME: 9, LAST_NAME: 10, EMAIL: 11, PHONE1: 12, PHONE2: 13,
    FROM_ADDRESS: 14, FROM_CITY: 15, FROM_STATE: 16, FROM_ZIP: 17,
    TO_ADDRESS: 18, TO_CITY: 19, TO_STATE: 20, TO_ZIP: 21,
    BOOK_DATE: 22, MOVE_DATE: 23, CF: 24, EST_TOTAL: 25,
    CARRIER: 26, SOURCE: 27, REF_NO: 28, FAD1: 29, FAD2: 30
  },
  
  // Contact standards
  STANDARDS: {
    MIN_TOTAL_ATTEMPTS: 5,
    LONG_CALL_THRESHOLD: 240,
    STALE_DAYS: 3,
    HOT_MOVE_DAYS: 7,
    HIGH_VALUE_THRESHOLD: 5000
  },
  
  // Priority colors
  PRIORITY_COLORS: {
    "0": "#ef4444", "1": "#f59e0b", "2": "#f97316", "3": "#eab308",
    "4": "#10b981", "5": "#f59e0b", "6": "#6b7280", "7": "#22c55e", "8": "#64748b"
  },
  
  // Admin emails for health reports
  ADMIN_EMAILS: [
    // Add admin emails here
  ]
};


/**
 * ═══════════════════════════════════════════════════════════════
 * MENU ENTRY
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_showAlertBuilder() {
  var html = HtmlService.createHtmlOutput(GRANOT_getAlertBuilderHTML_())
    .setWidth(600)
    .setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, "🚢 GRANOT Lead Alert Builder v2.0");
}


/**
 * ═══════════════════════════════════════════════════════════════
 * LOAD GRANOT DATA
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_loadLeads_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(GRANOT_CONFIG.SHEET_NAME);
  if (!sheet) throw new Error("GRANOT DATA sheet not found!");
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var data = sheet.getRange(2, 1, lastRow - 1, 31).getValues();
  var leads = [];
  var cols = GRANOT_CONFIG.COLS;
  var tz = "America/New_York";
  var now = new Date();
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[cols.JOB_NO] && !row[cols.PHONE1]) continue;
    
    var openDate = GRANOT_parseDate_(row[cols.OPEN_DATE]);
    var moveDate = GRANOT_parseDate_(row[cols.MOVE_DATE]);
    var daysUntilMove = moveDate ? Math.ceil((moveDate - now) / 86400000) : 999;
    var leadAgeDays = openDate ? Math.floor((now - openDate) / 86400000) : 0;
    
    leads.push({
      rowIndex: i + 2,
      dept: String(row[cols.DEPT] || "").trim(),
      jobNo: String(row[cols.JOB_NO] || "").trim(),
      openDate: openDate,
      openDateStr: openDate ? Utilities.formatDate(openDate, tz, "M/d/yyyy") : "",
      status: String(row[cols.STATUS] || "").trim(),
      user: String(row[cols.USER] || "").trim().toUpperCase(),
      jobRep: String(row[cols.JOB_REP] || "").trim().toUpperCase(),
      priority: String(row[cols.PRIORITY] || "0").trim(),
      firstName: String(row[cols.FIRST_NAME] || "").trim(),
      lastName: String(row[cols.LAST_NAME] || "").trim(),
      email: String(row[cols.EMAIL] || "").trim(),
      phone1: String(row[cols.PHONE1] || "").trim(),
      phone2: String(row[cols.PHONE2] || "").trim(),
      fromCity: String(row[cols.FROM_CITY] || "").trim(),
      fromState: String(row[cols.FROM_STATE] || "").trim(),
      toCity: String(row[cols.TO_CITY] || "").trim(),
      toState: String(row[cols.TO_STATE] || "").trim(),
      moveDate: moveDate,
      moveDateStr: moveDate ? Utilities.formatDate(moveDate, tz, "M/d/yyyy") : "",
      daysUntilMove: daysUntilMove,
      leadAgeDays: leadAgeDays,
      estTotal: parseFloat(row[cols.EST_TOTAL]) || 0,
      source: String(row[cols.SOURCE] || "").trim(),
      refNo: String(row[cols.REF_NO] || "").trim()
    });
  }
  return leads;
}


/**
 * ═══════════════════════════════════════════════════════════════
 * PARSE DATE
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_parseDate_(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  var str = String(val).trim();
  if (!str || str === "/  /" || str === "/ /") return null;
  var parts = str.split("/");
  if (parts.length === 3) {
    var m = parseInt(parts[0], 10) - 1;
    var d = parseInt(parts[1], 10);
    var y = parseInt(parts[2], 10);
    if (y < 100) y += 2000;
    var date = new Date(y, m, d);
    if (!isNaN(date.getTime())) return date;
  }
  var parsed = new Date(str);
  return isNaN(parsed.getTime()) ? null : parsed;
}


/**
 * ═══════════════════════════════════════════════════════════════
 * LOAD REP ROSTER
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_loadRepRoster_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var roster = {};
  
  var salesSheet = ss.getSheetByName(GRANOT_CONFIG.SALES_ROSTER);
  if (salesSheet && salesSheet.getLastRow() >= 2) {
    var data = salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, 10).getValues();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var username = String(row[0] || "").trim().toUpperCase();
      var email = String(row[1] || "").trim();
      var name = String(row[2] || "").trim();
      var team = String(row[3] || "").trim();
      var manager = String(row[4] || "").trim();
      var managerEmail = String(row[5] || "").trim();
      
      if (username && email) {
        roster[username] = {
          username: username,
          email: email,
          name: name || username,
          team: team,
          manager: manager,
          managerEmail: managerEmail
        };
      }
    }
  }
  
  var teamSheet = ss.getSheetByName(GRANOT_CONFIG.TEAM_ROSTER);
  if (teamSheet && teamSheet.getLastRow() >= 2) {
    var data2 = teamSheet.getRange(2, 1, teamSheet.getLastRow() - 1, 10).getValues();
    for (var j = 0; j < data2.length; j++) {
      var row2 = data2[j];
      var username2 = String(row2[0] || "").trim().toUpperCase();
      var email2 = String(row2[1] || "").trim();
      if (username2 && email2 && !roster[username2]) {
        roster[username2] = { username: username2, email: email2, name: row2[2] || username2 };
      }
    }
  }
  return roster;
}


/**
 * ═══════════════════════════════════════════════════════════════
 * NORMALIZE PHONE
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_normalizePhone_(phone) {
  if (!phone) return "";
  var digits = String(phone).replace(/\D/g, "");
  if (digits.length === 11 && digits[0] === "1") digits = digits.substring(1);
  return digits.length === 10 ? digits : "";
}


/**
 * ═══════════════════════════════════════════════════════════════
 * ENRICH LEADS WITH RC DATA
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_enrichLeadsWithRC_(leads) {
  var rcIndex = {};
  try {
    if (typeof RC_buildLookupIndex_ === "function") rcIndex = RC_buildLookupIndex_();
  } catch (e) { Logger.log("RC index not available"); }
  
  var now = new Date();
  var tz = "America/New_York";
  
  for (var i = 0; i < leads.length; i++) {
    var lead = leads[i];
    var phone = GRANOT_normalizePhone_(lead.phone1) || GRANOT_normalizePhone_(lead.phone2);
    var rc = rcIndex[phone] || {};
    
    var callsTotal = (rc.callsYesterday || 0) + (rc.callsToday || 0);
    var smsTotal = (rc.smsYesterday || 0) + (rc.smsToday || 0);
    var vmTotal = (rc.vmYesterday || 0) + (rc.vmToday || 0);
    var repliesTotal = (rc.repliesYesterday || 0) + (rc.repliesToday || 0);
    var inboundTotal = (rc.inboundCallsYesterday || 0) + (rc.inboundCallsToday || 0);
    var longestCall = Math.max(rc.longestCallYesterday || 0, rc.longestCallToday || 0);
    var totalAttempts = callsTotal + smsTotal;
    
    var meetsStandard = totalAttempts >= GRANOT_CONFIG.STANDARDS.MIN_TOTAL_ATTEMPTS;
    var hasLongCall = longestCall >= GRANOT_CONFIG.STANDARDS.LONG_CALL_THRESHOLD;
    var hasReplied = repliesTotal > 0;
    var hasInbound = inboundTotal > 0;
    
    var workStatus = "NOT_WORKED";
    if (hasLongCall || hasReplied || hasInbound) workStatus = "ENGAGED";
    else if (meetsStandard) workStatus = "WORKED";
    else if (totalAttempts > 0) workStatus = "PARTIAL";
    
    var isHot = lead.daysUntilMove <= GRANOT_CONFIG.STANDARDS.HOT_MOVE_DAYS && lead.daysUntilMove > 0;
    var isHighValue = lead.estTotal >= GRANOT_CONFIG.STANDARDS.HIGH_VALUE_THRESHOLD;
    var isStale = lead.leadAgeDays >= GRANOT_CONFIG.STANDARDS.STALE_DAYS && totalAttempts === 0;
    
    lead.rc = {
      phone: phone,
      callsTotal: callsTotal, smsTotal: smsTotal, vmTotal: vmTotal,
      repliesTotal: repliesTotal, inboundTotal: inboundTotal,
      longestCall: longestCall,
      longestCallDisplay: GRANOT_formatDuration_(longestCall),
      totalAttempts: totalAttempts,
      meetsStandard: meetsStandard, hasLongCall: hasLongCall,
      hasReplied: hasReplied, hasInbound: hasInbound,
      workStatus: workStatus,
      workStatusLabel: GRANOT_getWorkStatusLabel_(workStatus),
      workStatusColor: GRANOT_getWorkStatusColor_(workStatus),
      isHot: isHot, isHighValue: isHighValue, isStale: isStale,
      last5Calls: GRANOT_formatCallHistory_(rc.callHistory || [], now, tz),
      last5SMS: GRANOT_formatSMSHistory_(rc.smsHistory || [], now, tz),
      lastCallAgo: GRANOT_timeAgo_(rc.lastCall, now),
      lastSMSAgo: GRANOT_timeAgo_(rc.lastSMS, now)
    };
  }
  return leads;
}

function GRANOT_formatDuration_(s) {
  if (!s || s <= 0) return "0:00";
  return Math.floor(s/60) + ":" + (s%60 < 10 ? "0" : "") + (s%60);
}

function GRANOT_timeAgo_(date, now) {
  if (!date) return "Never";
  var mins = Math.floor((now - date) / 60000);
  if (mins < 60) return mins + "m ago";
  var hours = Math.floor(mins / 60);
  if (hours < 24) return hours + "h ago";
  return Math.floor(hours / 24) + "d ago";
}

function GRANOT_formatCallHistory_(history, now, tz) {
  if (!history || !history.length) return [];
  history.sort(function(a,b) { return (b.timestamp||0) - (a.timestamp||0); });
  return history.slice(0,5).map(function(c) {
    return {
      timeAgo: GRANOT_timeAgo_(c.timestamp, now),
      timeFormatted: c.timestamp ? Utilities.formatDate(c.timestamp, tz, "M/d h:mm a") : "",
      durationDisplay: GRANOT_formatDuration_(c.duration),
      isInbound: c.isInbound || false, isVM: c.isVM || false
    };
  });
}

function GRANOT_formatSMSHistory_(history, now, tz) {
  if (!history || !history.length) return [];
  history.sort(function(a,b) { return (b.timestamp||0) - (a.timestamp||0); });
  return history.slice(0,5).map(function(s) {
    return {
      timeAgo: GRANOT_timeAgo_(s.timestamp, now),
      timeFormatted: s.timestamp ? Utilities.formatDate(s.timestamp, tz, "M/d h:mm a") : ""
    };
  });
}

function GRANOT_getWorkStatusLabel_(s) {
  return {ENGAGED:"✅ Engaged",WORKED:"✓ Worked",PARTIAL:"⚠️ Partial",NOT_WORKED:"❌ Not Worked"}[s]||s;
}

function GRANOT_getWorkStatusColor_(s) {
  return {ENGAGED:"#10b981",WORKED:"#3b82f6",PARTIAL:"#f59e0b",NOT_WORKED:"#ef4444"}[s]||"#64748b";
}


/**
 * ═══════════════════════════════════════════════════════════════
 * FILTER LEADS
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_filterLeads_(leads, p) {
  return leads.filter(function(l) {
    if (p.priorities && p.priorities.length && p.priorities.indexOf(l.priority) === -1) return false;
    
    if (p.dateFrom) {
      var from = GRANOT_parseDate_(p.dateFrom);
      if (from && l.openDate && l.openDate < from) return false;
    }
    if (p.dateTo) {
      var to = GRANOT_parseDate_(p.dateTo);
      if (to && l.openDate) { to.setHours(23,59,59,999); if (l.openDate > to) return false; }
    }
    
    if (p.moveDaysMax !== null && p.moveDaysMax !== "" && l.daysUntilMove > parseInt(p.moveDaysMax)) return false;
    if (p.leadAgeDaysMin !== null && p.leadAgeDaysMin !== "" && l.leadAgeDays < parseInt(p.leadAgeDaysMin)) return false;
    if (p.minEstTotal !== null && p.minEstTotal !== "" && l.estTotal < parseFloat(p.minEstTotal)) return false;
    
    if (p.sources && p.sources.length && p.sources.indexOf(l.source) === -1) return false;
    
    if (p.minCalls !== null && p.minCalls !== "" && l.rc.callsTotal < parseInt(p.minCalls)) return false;
    if (p.maxCalls !== null && p.maxCalls !== "" && l.rc.callsTotal > parseInt(p.maxCalls)) return false;
    if (p.minSMS !== null && p.minSMS !== "" && l.rc.smsTotal < parseInt(p.minSMS)) return false;
    if (p.maxSMS !== null && p.maxSMS !== "" && l.rc.smsTotal > parseInt(p.maxSMS)) return false;
    
    if (p.hotOnly && !l.rc.isHot) return false;
    if (p.highValueOnly && !l.rc.isHighValue) return false;
    if (p.staleOnly && !l.rc.isStale) return false;
    
    return true;
  });
}


/**
 * ═══════════════════════════════════════════════════════════════
 * GET FILTER OPTIONS FOR UI
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_getFilterOptions() {
  var leads = GRANOT_loadLeads_();
  var sources = {};
  leads.forEach(function(l) { if (l.source) sources[l.source] = true; });
  return {
    sources: Object.keys(sources).sort(),
    totalLeads: leads.length,
    presets: GRANOT_loadPresets_()
  };
}


/**
 * ═══════════════════════════════════════════════════════════════
 * PRESET MANAGEMENT
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_loadPresets_() {
  var props = PropertiesService.getDocumentProperties();
  var data = props.getProperty(GRANOT_CONFIG.PRESETS_KEY);
  return data ? JSON.parse(data) : {};
}

function GRANOT_savePreset(name, params) {
  var presets = GRANOT_loadPresets_();
  presets[name] = params;
  PropertiesService.getDocumentProperties().setProperty(GRANOT_CONFIG.PRESETS_KEY, JSON.stringify(presets));
  return { success: true, message: "Preset '" + name + "' saved!" };
}

function GRANOT_deletePreset(name) {
  var presets = GRANOT_loadPresets_();
  delete presets[name];
  PropertiesService.getDocumentProperties().setProperty(GRANOT_CONFIG.PRESETS_KEY, JSON.stringify(presets));
  return { success: true, message: "Preset deleted" };
}

function GRANOT_getPresets() {
  return GRANOT_loadPresets_();
}


/**
 * ═══════════════════════════════════════════════════════════════
 * GET PREVIEW
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_getPreview(params) {
  try {
    var leads = GRANOT_enrichLeadsWithRC_(GRANOT_loadLeads_());
    var filtered = GRANOT_filterLeads_(leads, params);
    var roster = GRANOT_loadRepRoster_();
    
    var byRep = {}, byTeam = {}, bySource = {};
    var byStatus = { ENGAGED:0, WORKED:0, PARTIAL:0, NOT_WORKED:0 };
    var hotCount = 0, highValueCount = 0, staleCount = 0;
    
    filtered.forEach(function(l) {
      var rep = l.user || l.jobRep || "UNASSIGNED";
      var repInfo = roster[rep] || {};
      var team = repInfo.team || "No Team";
      
      if (!byRep[rep]) byRep[rep] = { total:0, engaged:0, worked:0, partial:0, notWorked:0, team:team };
      byRep[rep].total++;
      byRep[rep][l.rc.workStatus === "ENGAGED" ? "engaged" : l.rc.workStatus === "WORKED" ? "worked" : l.rc.workStatus === "PARTIAL" ? "partial" : "notWorked"]++;
      byStatus[l.rc.workStatus]++;
      
      if (!byTeam[team]) byTeam[team] = { total:0, engaged:0, worked:0, partial:0, notWorked:0 };
      byTeam[team].total++;
      byTeam[team][l.rc.workStatus === "ENGAGED" ? "engaged" : l.rc.workStatus === "WORKED" ? "worked" : l.rc.workStatus === "PARTIAL" ? "partial" : "notWorked"]++;
      
      if (l.source) {
        if (!bySource[l.source]) bySource[l.source] = { total:0, notWorked:0 };
        bySource[l.source].total++;
        if (l.rc.workStatus === "NOT_WORKED") bySource[l.source].notWorked++;
      }
      
      if (l.rc.isHot) hotCount++;
      if (l.rc.isHighValue) highValueCount++;
      if (l.rc.isStale) staleCount++;
    });
    
    var repSummary = Object.keys(byRep).map(function(r) {
      var d = byRep[r];
      d.rep = r;
      d.compliance = d.total > 0 ? Math.round(((d.engaged + d.worked) / d.total) * 100) : 0;
      return d;
    }).sort(function(a,b) { return a.compliance - b.compliance; });
    
    return {
      success: true,
      totalFiltered: filtered.length,
      totalAll: leads.length,
      byStatus: byStatus,
      repSummary: repSummary.slice(0, 20),
      teamSummary: byTeam,
      sourceSummary: bySource,
      hotCount: hotCount,
      highValueCount: highValueCount,
      staleCount: staleCount
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


/**
 * ═══════════════════════════════════════════════════════════════
 * RUN ALERT - Main entry point
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_runAlert(params) {
  var startTime = new Date().getTime();
  var senderEmail = GRANOT_getCurrentUserEmail_();
  
  Logger.log("🚀 GRANOT Alert by: " + senderEmail);
  
  try {
    var leads = GRANOT_enrichLeadsWithRC_(GRANOT_loadLeads_());
    var filtered = GRANOT_filterLeads_(leads, params);
    var roster = GRANOT_loadRepRoster_();
    
    if (filtered.length === 0) {
      return { success: true, message: "No leads matched your filters.", count: 0 };
    }
    
    var results = { success: true, emailsSent: 0, leadsProcessed: filtered.length, errors: [], skippedReps: [] };
    
    // Group leads by rep
    var byRep = {};
    filtered.forEach(function(l) {
      var rep = l.user || l.jobRep || "UNASSIGNED";
      if (!byRep[rep]) byRep[rep] = [];
      byRep[rep].push(l);
    });
    
    if (params.sendMode === "BY_REP") {
      for (var rep in byRep) {
        var repLeads = byRep[rep];
        var repInfo = roster[rep] || {};
        var repEmail = repInfo.email;
        var emailTo = params.testMode ? senderEmail : repEmail;
        
        if (emailTo) {
          try {
            var ccList = [];
            if (params.ccManager && repInfo.managerEmail && !params.testMode) {
              ccList.push(repInfo.managerEmail);
            }
            GRANOT_sendRepReport_(rep, repLeads, emailTo, ccList, params, senderEmail, roster);
            results.emailsSent++;
          } catch (e) { results.errors.push(rep + ": " + e.message); }
        } else {
          results.skippedReps.push(rep);
        }
      }
    } else if (params.sendMode === "BY_TEAM") {
      var byTeam = {};
      for (var rep2 in byRep) {
        var repInfo2 = roster[rep2] || {};
        var team = repInfo2.team || "No Team";
        var mgr = repInfo2.manager || "";
        var mgrEmail = repInfo2.managerEmail || "";
        if (!byTeam[team]) byTeam[team] = { leads: [], manager: mgr, managerEmail: mgrEmail };
        byTeam[team].leads = byTeam[team].leads.concat(byRep[rep2]);
      }
      
      for (var team2 in byTeam) {
        var teamData = byTeam[team2];
        var emailTo2 = params.testMode ? senderEmail : teamData.managerEmail;
        if (emailTo2) {
          try {
            GRANOT_sendTeamReport_(team2, teamData.leads, emailTo2, params, senderEmail, roster);
            results.emailsSent++;
          } catch (e) { results.errors.push(team2 + ": " + e.message); }
        }
      }
    } else if (params.sendMode === "ADMIN") {
      var emailTo3 = params.testMode ? senderEmail : (params.emailTo || senderEmail);
      GRANOT_sendAdminReport_(filtered, byRep, emailTo3, params, senderEmail, roster);
      results.emailsSent = 1;
    } else {
      var emailTo4 = params.testMode ? senderEmail : (params.emailTo || senderEmail);
      GRANOT_sendSummaryReport_(filtered, byRep, emailTo4, params, senderEmail, roster);
      results.emailsSent = 1;
    }
    
    var elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(1);
    results.message = "Sent " + results.emailsSent + " email(s) for " + filtered.length + " leads in " + elapsed + "s";
    if (results.skippedReps.length) results.message += ". Skipped: " + results.skippedReps.join(", ");
    
    return results;
  } catch (e) {
    return { success: false, message: "Error: " + e.message };
  }
}

function GRANOT_getCurrentUserEmail_() {
  try { return Session.getActiveUser().getEmail() || "system@safeshipmoving.com"; }
  catch (e) { return "system@safeshipmoving.com"; }
}


/**
 * ═══════════════════════════════════════════════════════════════
 * SEND REP REPORT
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_sendRepReport_(rep, leads, emailTo, ccList, params, senderEmail, roster) {
  var tz = "America/New_York";
  var now = new Date();
  var repInfo = roster[rep] || { name: rep };
  
  leads.sort(function(a,b) {
    var pa = parseInt(a.priority)||99, pb = parseInt(b.priority)||99;
    if (pa !== pb) return pa - pb;
    var order = {NOT_WORKED:0,PARTIAL:1,WORKED:2,ENGAGED:3};
    return (order[a.rc.workStatus]||0) - (order[b.rc.workStatus]||0);
  });
  
  var subject = "🚢 Lead Alert - " + (repInfo.name || rep) + " - " + leads.length + " Lead(s)";
  var html = GRANOT_buildRepReportHTML_(rep, repInfo.name || rep, leads, {
    dateStr: Utilities.formatDate(now, tz, "EEEE, MMMM d, yyyy"),
    timeStr: Utilities.formatDate(now, tz, "h:mm a"),
    senderEmail: senderEmail
  }, roster);
  
  var mailOptions = { to: emailTo, subject: subject, htmlBody: html };
  if (ccList && ccList.length) mailOptions.cc = ccList.join(",");
  
  MailApp.sendEmail(mailOptions);
}


/**
 * ═══════════════════════════════════════════════════════════════
 * SEND TEAM REPORT
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_sendTeamReport_(team, leads, emailTo, params, senderEmail, roster) {
  var tz = "America/New_York";
  var now = new Date();
  
  // Group by rep within team
  var byRep = {};
  leads.forEach(function(l) {
    var rep = l.user || l.jobRep || "UNASSIGNED";
    if (!byRep[rep]) byRep[rep] = [];
    byRep[rep].push(l);
  });
  
  var subject = "🚢 Team Report: " + team + " - " + leads.length + " Lead(s)";
  var html = GRANOT_buildTeamReportHTML_(team, leads, byRep, {
    dateStr: Utilities.formatDate(now, tz, "EEEE, MMMM d, yyyy"),
    timeStr: Utilities.formatDate(now, tz, "h:mm a"),
    senderEmail: senderEmail
  }, roster);
  
  MailApp.sendEmail({ to: emailTo, subject: subject, htmlBody: html });
}


/**
 * ═══════════════════════════════════════════════════════════════
 * SEND ADMIN HEALTH REPORT
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_sendAdminReport_(leads, byRep, emailTo, params, senderEmail, roster) {
  var tz = "America/New_York";
  var now = new Date();
  
  // Calculate health metrics
  var issues = [];
  var hotNotWorked = leads.filter(function(l) { return l.rc.isHot && l.rc.workStatus === "NOT_WORKED"; });
  var highValueNotWorked = leads.filter(function(l) { return l.rc.isHighValue && l.rc.workStatus === "NOT_WORKED"; });
  var staleLeads = leads.filter(function(l) { return l.rc.isStale; });
  
  if (hotNotWorked.length > 0) issues.push({ severity: "CRITICAL", msg: hotNotWorked.length + " HOT leads (move ≤7 days) not worked!" });
  if (highValueNotWorked.length > 0) issues.push({ severity: "HIGH", msg: highValueNotWorked.length + " HIGH VALUE leads ($5K+) not worked!" });
  if (staleLeads.length > 0) issues.push({ severity: "MEDIUM", msg: staleLeads.length + " leads are STALE (3+ days, 0 contact)" });
  
  // Rep compliance issues
  for (var rep in byRep) {
    var repLeads = byRep[rep];
    var notWorked = repLeads.filter(function(l) { return l.rc.workStatus === "NOT_WORKED"; }).length;
    var pct = Math.round((notWorked / repLeads.length) * 100);
    if (pct > 50) issues.push({ severity: "HIGH", msg: rep + " has " + pct + "% not worked (" + notWorked + "/" + repLeads.length + ")" });
  }
  
  var subject = "🚢 Admin Health Report - " + leads.length + " Leads - " + issues.length + " Issues";
  var html = GRANOT_buildAdminReportHTML_(leads, byRep, issues, {
    dateStr: Utilities.formatDate(now, tz, "EEEE, MMMM d, yyyy"),
    timeStr: Utilities.formatDate(now, tz, "h:mm a"),
    senderEmail: senderEmail,
    hotNotWorked: hotNotWorked,
    highValueNotWorked: highValueNotWorked,
    staleLeads: staleLeads
  }, roster);
  
  MailApp.sendEmail({ to: emailTo, subject: subject, htmlBody: html });
}


/**
 * ═══════════════════════════════════════════════════════════════
 * SEND SUMMARY REPORT
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_sendSummaryReport_(leads, byRep, emailTo, params, senderEmail, roster) {
  var tz = "America/New_York";
  var now = new Date();
  
  var subject = "🚢 GRANOT Summary - " + leads.length + " Leads";
  var html = GRANOT_buildSummaryReportHTML_(leads, byRep, {
    dateStr: Utilities.formatDate(now, tz, "EEEE, MMMM d, yyyy"),
    timeStr: Utilities.formatDate(now, tz, "h:mm a"),
    senderEmail: senderEmail
  }, roster);
  
  MailApp.sendEmail({ to: emailTo, subject: subject, htmlBody: html });
}


/**
 * ═══════════════════════════════════════════════════════════════
 * BUILD REP REPORT HTML - Lead Cards Style
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_buildRepReportHTML_(rep, repName, leads, info, roster) {
  var engaged = leads.filter(function(l){return l.rc.workStatus==="ENGAGED";}).length;
  var worked = leads.filter(function(l){return l.rc.workStatus==="WORKED";}).length;
  var partial = leads.filter(function(l){return l.rc.workStatus==="PARTIAL";}).length;
  var notWorked = leads.filter(function(l){return l.rc.workStatus==="NOT_WORKED";}).length;
  var compliance = leads.length > 0 ? Math.round(((engaged+worked)/leads.length)*100) : 0;
  
  var html = '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="margin:0;padding:0;font-family:Arial,sans-serif;background:#f1f5f9;">';
  
  // Header
  html += '<div style="background:linear-gradient(135deg,#8B1538,#6B1028);padding:24px;text-align:center;">';
  html += '<div style="font-size:10px;color:#D4AF37;letter-spacing:2px;">🚢 SAFE SHIP MOVING SERVICES</div>';
  html += '<div style="font-size:24px;font-weight:bold;color:#fff;margin-top:4px;">Lead Alert</div>';
  html += '<div style="font-size:16px;color:rgba(255,255,255,0.9);margin-top:4px;">' + repName + '</div>';
  html += '<div style="font-size:12px;color:rgba(255,255,255,0.7);">' + info.dateStr + '</div>';
  html += '</div>';
  
  // Compliance Score Bar
  var compColor = compliance >= 80 ? "#10b981" : compliance >= 60 ? "#f59e0b" : "#ef4444";
  html += '<div style="background:#1e293b;padding:16px;text-align:center;">';
  html += '<div style="font-size:10px;color:#94a3b8;margin-bottom:4px;">COMPLIANCE SCORE</div>';
  html += '<div style="display:inline-block;background:#0f172a;border-radius:20px;padding:8px 24px;">';
  html += '<span style="font-size:32px;font-weight:bold;color:' + compColor + ';">' + compliance + '%</span>';
  html += '</div>';
  html += '<div style="display:flex;justify-content:center;gap:16px;margin-top:12px;flex-wrap:wrap;">';
  html += '<div><span style="font-size:20px;font-weight:bold;color:#ef4444;">' + notWorked + '</span><div style="font-size:9px;color:#94a3b8;">NOT WORKED</div></div>';
  html += '<div><span style="font-size:20px;font-weight:bold;color:#f59e0b;">' + partial + '</span><div style="font-size:9px;color:#94a3b8;">PARTIAL</div></div>';
  html += '<div><span style="font-size:20px;font-weight:bold;color:#3b82f6;">' + worked + '</span><div style="font-size:9px;color:#94a3b8;">WORKED</div></div>';
  html += '<div><span style="font-size:20px;font-weight:bold;color:#10b981;">' + engaged + '</span><div style="font-size:9px;color:#94a3b8;">ENGAGED</div></div>';
  html += '</div></div>';
  
  // Lead Cards
  html += '<div style="padding:16px;max-width:700px;margin:0 auto;">';
  
  for (var i = 0; i < leads.length; i++) {
    var l = leads[i];
    var priBg = GRANOT_CONFIG.PRIORITY_COLORS[l.priority] || "#64748b";
    
    html += '<div style="background:#fff;border-radius:12px;margin-bottom:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">';
    
    // Card Header
    html += '<div style="background:linear-gradient(135deg,#1e293b,#334155);padding:12px 16px;display:flex;justify-content:space-between;align-items:center;">';
    html += '<div><div style="font-size:16px;font-weight:bold;color:#fff;">' + (l.firstName + ' ' + l.lastName).trim() + '</div>';
    html += '<div style="font-size:11px;color:#94a3b8;">Job #' + l.jobNo + ' • ' + l.openDateStr + '</div></div>';
    html += '<div style="display:flex;gap:6px;align-items:center;">';
    
    // Flags
    if (l.rc.isHot) html += '<span style="background:#ef4444;color:#fff;font-size:9px;padding:2px 6px;border-radius:4px;">🔥 HOT</span>';
    if (l.rc.isHighValue) html += '<span style="background:#8b5cf6;color:#fff;font-size:9px;padding:2px 6px;border-radius:4px;">💰 $' + Math.round(l.estTotal/1000) + 'K</span>';
    if (l.rc.isStale) html += '<span style="background:#64748b;color:#fff;font-size:9px;padding:2px 6px;border-radius:4px;">⏰ STALE</span>';
    
    html += '<span style="background:' + priBg + ';color:#fff;font-size:14px;font-weight:bold;padding:4px 10px;border-radius:6px;">P' + l.priority + '</span>';
    html += '</div></div>';
    
    // Card Body
    html += '<div style="padding:12px 16px;">';
    
    // Status badge
    html += '<div style="margin-bottom:10px;"><span style="display:inline-block;padding:3px 10px;border-radius:4px;font-size:11px;font-weight:bold;background:' + l.rc.workStatusColor + ';color:#fff;">' + l.rc.workStatusLabel + '</span></div>';
    
    // Contact & Move info
    html += '<div style="display:flex;gap:16px;margin-bottom:10px;flex-wrap:wrap;">';
    html += '<div style="flex:1;min-width:130px;"><div style="font-size:9px;color:#64748b;text-transform:uppercase;">Phone</div><div style="font-size:13px;font-weight:600;">' + (l.phone1||"N/A") + '</div></div>';
    html += '<div style="flex:1;min-width:130px;"><div style="font-size:9px;color:#64748b;text-transform:uppercase;">Email</div><div style="font-size:11px;word-break:break-all;">' + (l.email||"N/A") + '</div></div>';
    html += '</div>';
    
    html += '<div style="display:flex;gap:12px;margin-bottom:10px;flex-wrap:wrap;">';
    html += '<div><div style="font-size:9px;color:#64748b;">FROM</div><div style="font-size:11px;">' + [l.fromCity,l.fromState].filter(Boolean).join(", ") + '</div></div>';
    html += '<div><div style="font-size:9px;color:#64748b;">TO</div><div style="font-size:11px;">' + [l.toCity,l.toState].filter(Boolean).join(", ") + '</div></div>';
    html += '<div><div style="font-size:9px;color:#64748b;">MOVE</div><div style="font-size:11px;font-weight:600;color:' + (l.rc.isHot?"#ef4444":"#1e293b") + ';">' + (l.moveDateStr||"TBD") + '</div></div>';
    html += '<div><div style="font-size:9px;color:#64748b;">EST</div><div style="font-size:11px;font-weight:600;color:#10b981;">$' + (l.estTotal?l.estTotal.toLocaleString():"0") + '</div></div>';
    html += '</div>';
    
    // Activity Section
    html += '<div style="background:#f8fafc;border-radius:8px;padding:10px;margin-top:8px;">';
    html += '<div style="font-size:9px;font-weight:700;color:#64748b;margin-bottom:6px;">📊 CONTACT ACTIVITY</div>';
    
    // Bubbles
    html += '<div style="display:flex;gap:10px;flex-wrap:wrap;">';
    var callBg = l.rc.callsTotal > 0 ? "#dcfce7" : "#fee2e2";
    var callCol = l.rc.callsTotal > 0 ? "#166534" : "#991b1b";
    html += '<div style="background:' + callBg + ';padding:6px 12px;border-radius:6px;text-align:center;"><div style="font-size:18px;font-weight:bold;color:' + callCol + ';">' + l.rc.callsTotal + '</div><div style="font-size:8px;color:' + callCol + ';">📞 Calls</div></div>';
    
    var smsBg = l.rc.smsTotal > 0 ? "#dbeafe" : "#fee2e2";
    var smsCol = l.rc.smsTotal > 0 ? "#1e40af" : "#991b1b";
    html += '<div style="background:' + smsBg + ';padding:6px 12px;border-radius:6px;text-align:center;"><div style="font-size:18px;font-weight:bold;color:' + smsCol + ';">' + l.rc.smsTotal + '</div><div style="font-size:8px;color:' + smsCol + ';">💬 SMS</div></div>';
    
    if (l.rc.vmTotal > 0) html += '<div style="background:#fef3c7;padding:6px 12px;border-radius:6px;text-align:center;"><div style="font-size:18px;font-weight:bold;color:#92400e;">' + l.rc.vmTotal + '</div><div style="font-size:8px;color:#92400e;">📭 VM</div></div>';
    if (l.rc.inboundTotal > 0) html += '<div style="background:#cffafe;padding:6px 12px;border-radius:6px;text-align:center;"><div style="font-size:18px;font-weight:bold;color:#0e7490;">' + l.rc.inboundTotal + '</div><div style="font-size:8px;color:#0e7490;">📥 IN</div></div>';
    html += '</div>';
    
    // Call history
    if (l.rc.last5Calls && l.rc.last5Calls.length) {
      html += '<div style="margin-top:8px;"><div style="font-size:9px;color:#64748b;">Last Calls:</div><div style="background:#fff;border-radius:4px;padding:4px 8px;margin-top:2px;">';
      l.rc.last5Calls.forEach(function(c) {
        var inB = c.isInbound ? '<span style="background:#06b6d4;color:#fff;font-size:7px;padding:1px 3px;border-radius:2px;margin-right:3px;">IN</span>' : '';
        var vmB = c.isVM ? '<span style="background:#f59e0b;color:#fff;font-size:7px;padding:1px 3px;border-radius:2px;margin-left:3px;">VM</span>' : '';
        html += '<div style="font-size:10px;color:#334155;">' + inB + c.timeAgo + ' (' + c.timeFormatted + ') ' + c.durationDisplay + vmB + '</div>';
      });
      html += '</div></div>';
    }
    
    // Alerts
    if (l.rc.hasReplied) html += '<div style="margin-top:6px;padding:6px;background:#d1fae5;border-radius:4px;border-left:3px solid #10b981;font-size:10px;font-weight:700;color:#166534;">📥 LEAD REPLIED!</div>';
    if (l.rc.hasLongCall) html += '<div style="margin-top:6px;padding:6px;background:#dbeafe;border-radius:4px;border-left:3px solid #3b82f6;font-size:10px;font-weight:700;color:#1e40af;">⏱️ Long Call: ' + l.rc.longestCallDisplay + '</div>';
    
    html += '</div></div></div>';
  }
  
  html += '</div>';
  
  // Footer
  html += '<div style="background:#1e293b;padding:16px;text-align:center;color:#64748b;font-size:10px;">';
  html += 'Generated ' + info.timeStr + ' by ' + info.senderEmail + '<br>GRANOT Alert System v2.0</div>';
  
  html += '</body></html>';
  return html;
}


/**
 * ═══════════════════════════════════════════════════════════════
 * BUILD TEAM REPORT HTML
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_buildTeamReportHTML_(team, leads, byRep, info, roster) {
  var html = '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="margin:0;padding:0;font-family:Arial,sans-serif;background:#f1f5f9;">';
  
  html += '<div style="background:linear-gradient(135deg,#8B1538,#6B1028);padding:24px;text-align:center;">';
  html += '<div style="font-size:10px;color:#D4AF37;letter-spacing:2px;">🚢 SAFE SHIP MOVING SERVICES</div>';
  html += '<div style="font-size:24px;font-weight:bold;color:#fff;margin-top:4px;">Team Report: ' + team + '</div>';
  html += '<div style="font-size:12px;color:rgba(255,255,255,0.7);">' + info.dateStr + '</div>';
  html += '</div>';
  
  // Team stats
  var totalEngaged = leads.filter(function(l){return l.rc.workStatus==="ENGAGED";}).length;
  var totalWorked = leads.filter(function(l){return l.rc.workStatus==="WORKED";}).length;
  var totalPartial = leads.filter(function(l){return l.rc.workStatus==="PARTIAL";}).length;
  var totalNotWorked = leads.filter(function(l){return l.rc.workStatus==="NOT_WORKED";}).length;
  var teamCompliance = leads.length > 0 ? Math.round(((totalEngaged+totalWorked)/leads.length)*100) : 0;
  
  var compColor = teamCompliance >= 80 ? "#10b981" : teamCompliance >= 60 ? "#f59e0b" : "#ef4444";
  html += '<div style="background:#1e293b;padding:16px;text-align:center;">';
  html += '<div style="font-size:10px;color:#94a3b8;">TEAM COMPLIANCE</div>';
  html += '<div style="font-size:36px;font-weight:bold;color:' + compColor + ';">' + teamCompliance + '%</div>';
  html += '<div style="display:flex;justify-content:center;gap:20px;margin-top:8px;">';
  html += '<span style="color:#ef4444;">❌ ' + totalNotWorked + '</span>';
  html += '<span style="color:#f59e0b;">⚠️ ' + totalPartial + '</span>';
  html += '<span style="color:#3b82f6;">✓ ' + totalWorked + '</span>';
  html += '<span style="color:#10b981;">✅ ' + totalEngaged + '</span>';
  html += '</div></div>';
  
  // Rep breakdown table
  html += '<div style="padding:16px;"><table style="width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;">';
  html += '<tr style="background:#334155;color:#fff;font-size:11px;"><th style="padding:10px;text-align:left;">Rep</th><th style="padding:10px;text-align:center;">Total</th><th style="padding:10px;text-align:center;">❌</th><th style="padding:10px;text-align:center;">⚠️</th><th style="padding:10px;text-align:center;">✓</th><th style="padding:10px;text-align:center;">✅</th><th style="padding:10px;text-align:center;">Score</th></tr>';
  
  var repList = [];
  for (var rep in byRep) {
    var rl = byRep[rep];
    var e = rl.filter(function(l){return l.rc.workStatus==="ENGAGED";}).length;
    var w = rl.filter(function(l){return l.rc.workStatus==="WORKED";}).length;
    var p = rl.filter(function(l){return l.rc.workStatus==="PARTIAL";}).length;
    var n = rl.filter(function(l){return l.rc.workStatus==="NOT_WORKED";}).length;
    var c = rl.length > 0 ? Math.round(((e+w)/rl.length)*100) : 0;
    var ri = roster[rep] || {};
    repList.push({ rep: rep, name: ri.name||rep, total: rl.length, engaged: e, worked: w, partial: p, notWorked: n, compliance: c });
  }
  repList.sort(function(a,b){ return a.compliance - b.compliance; });
  
  repList.forEach(function(r, idx) {
    var bg = idx % 2 === 0 ? "#fff" : "#f8fafc";
    var cc = r.compliance >= 80 ? "#10b981" : r.compliance >= 60 ? "#f59e0b" : "#ef4444";
    html += '<tr style="background:' + bg + ';font-size:12px;">';
    html += '<td style="padding:8px;font-weight:bold;">' + r.name + '</td>';
    html += '<td style="padding:8px;text-align:center;">' + r.total + '</td>';
    html += '<td style="padding:8px;text-align:center;color:#ef4444;font-weight:bold;">' + r.notWorked + '</td>';
    html += '<td style="padding:8px;text-align:center;color:#f59e0b;font-weight:bold;">' + r.partial + '</td>';
    html += '<td style="padding:8px;text-align:center;color:#3b82f6;font-weight:bold;">' + r.worked + '</td>';
    html += '<td style="padding:8px;text-align:center;color:#10b981;font-weight:bold;">' + r.engaged + '</td>';
    html += '<td style="padding:8px;text-align:center;"><span style="background:' + cc + ';color:#fff;padding:2px 8px;border-radius:4px;font-weight:bold;">' + r.compliance + '%</span></td>';
    html += '</tr>';
  });
  
  html += '</table></div>';
  
  // Footer
  html += '<div style="background:#1e293b;padding:16px;text-align:center;color:#64748b;font-size:10px;">Generated ' + info.timeStr + '<br>GRANOT Alert System v2.0</div>';
  html += '</body></html>';
  return html;
}


/**
 * ═══════════════════════════════════════════════════════════════
 * BUILD ADMIN HEALTH REPORT HTML
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_buildAdminReportHTML_(leads, byRep, issues, info, roster) {
  var html = '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="margin:0;padding:0;font-family:Arial,sans-serif;background:#f1f5f9;">';
  
  html += '<div style="background:linear-gradient(135deg,#8B1538,#6B1028);padding:24px;text-align:center;">';
  html += '<div style="font-size:10px;color:#D4AF37;letter-spacing:2px;">🚢 SAFE SHIP MOVING SERVICES</div>';
  html += '<div style="font-size:24px;font-weight:bold;color:#fff;margin-top:4px;">Admin Health Report</div>';
  html += '<div style="font-size:12px;color:rgba(255,255,255,0.7);">' + info.dateStr + '</div>';
  html += '</div>';
  
  // Health Score
  var totalEngaged = leads.filter(function(l){return l.rc.workStatus==="ENGAGED";}).length;
  var totalWorked = leads.filter(function(l){return l.rc.workStatus==="WORKED";}).length;
  var healthScore = leads.length > 0 ? Math.round(((totalEngaged+totalWorked)/leads.length)*100) : 0;
  var healthColor = healthScore >= 80 ? "#10b981" : healthScore >= 60 ? "#f59e0b" : "#ef4444";
  var healthStatus = healthScore >= 80 ? "HEALTHY" : healthScore >= 60 ? "NEEDS ATTENTION" : "CRITICAL";
  
  html += '<div style="background:#1e293b;padding:20px;text-align:center;">';
  html += '<div style="font-size:10px;color:#94a3b8;">OVERALL HEALTH</div>';
  html += '<div style="font-size:42px;font-weight:bold;color:' + healthColor + ';">' + healthScore + '%</div>';
  html += '<div style="font-size:12px;color:' + healthColor + ';font-weight:bold;">' + healthStatus + '</div>';
  html += '<div style="display:flex;justify-content:center;gap:24px;margin-top:12px;">';
  html += '<div><div style="font-size:24px;font-weight:bold;color:#D4AF37;">' + leads.length + '</div><div style="font-size:9px;color:#94a3b8;">TOTAL</div></div>';
  html += '<div><div style="font-size:24px;font-weight:bold;color:#ef4444;">' + info.hotNotWorked.length + '</div><div style="font-size:9px;color:#94a3b8;">🔥 HOT MISSED</div></div>';
  html += '<div><div style="font-size:24px;font-weight:bold;color:#8b5cf6;">' + info.highValueNotWorked.length + '</div><div style="font-size:9px;color:#94a3b8;">💰 HIGH $ MISSED</div></div>';
  html += '<div><div style="font-size:24px;font-weight:bold;color:#64748b;">' + info.staleLeads.length + '</div><div style="font-size:9px;color:#94a3b8;">⏰ STALE</div></div>';
  html += '</div></div>';
  
  // Issues Section
  if (issues.length > 0) {
    html += '<div style="padding:16px;"><div style="font-size:14px;font-weight:bold;color:#ef4444;margin-bottom:8px;">⚠️ ISSUES DETECTED (' + issues.length + ')</div>';
    html += '<div style="background:#fff;border-radius:8px;overflow:hidden;">';
    issues.forEach(function(issue) {
      var issueBg = issue.severity === "CRITICAL" ? "#fef2f2" : issue.severity === "HIGH" ? "#fffbeb" : "#f8fafc";
      var issueCol = issue.severity === "CRITICAL" ? "#991b1b" : issue.severity === "HIGH" ? "#92400e" : "#475569";
      var issueBadge = issue.severity === "CRITICAL" ? "🔴" : issue.severity === "HIGH" ? "🟠" : "🟡";
      html += '<div style="padding:10px 12px;border-bottom:1px solid #e5e7eb;background:' + issueBg + ';">';
      html += '<span style="font-size:12px;">' + issueBadge + ' <strong style="color:' + issueCol + ';">' + issue.severity + ':</strong> ' + issue.msg + '</span>';
      html += '</div>';
    });
    html += '</div></div>';
  }
  
  // Rep Performance Table
  html += '<div style="padding:0 16px 16px;"><div style="font-size:14px;font-weight:bold;color:#1e293b;margin-bottom:8px;">📊 Rep Performance</div>';
  html += '<table style="width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;">';
  html += '<tr style="background:#334155;color:#fff;font-size:10px;"><th style="padding:8px;text-align:left;">Rep</th><th style="padding:8px;">Total</th><th style="padding:8px;">❌</th><th style="padding:8px;">⚠️</th><th style="padding:8px;">✓</th><th style="padding:8px;">✅</th><th style="padding:8px;">Score</th></tr>';
  
  var repList = [];
  for (var rep in byRep) {
    var rl = byRep[rep];
    var e = rl.filter(function(l){return l.rc.workStatus==="ENGAGED";}).length;
    var w = rl.filter(function(l){return l.rc.workStatus==="WORKED";}).length;
    var p = rl.filter(function(l){return l.rc.workStatus==="PARTIAL";}).length;
    var n = rl.filter(function(l){return l.rc.workStatus==="NOT_WORKED";}).length;
    var c = rl.length > 0 ? Math.round(((e+w)/rl.length)*100) : 0;
    var ri = roster[rep] || {};
    repList.push({ rep: rep, name: ri.name||rep, team: ri.team||"", total: rl.length, engaged: e, worked: w, partial: p, notWorked: n, compliance: c });
  }
  repList.sort(function(a,b){ return a.compliance - b.compliance; });
  
  repList.forEach(function(r, idx) {
    var bg = idx % 2 === 0 ? "#fff" : "#f8fafc";
    if (r.compliance < 50) bg = "#fef2f2";
    var cc = r.compliance >= 80 ? "#10b981" : r.compliance >= 60 ? "#f59e0b" : "#ef4444";
    html += '<tr style="background:' + bg + ';font-size:11px;">';
    html += '<td style="padding:6px 8px;"><strong>' + r.name + '</strong><br><span style="font-size:9px;color:#94a3b8;">' + r.team + '</span></td>';
    html += '<td style="padding:6px;text-align:center;">' + r.total + '</td>';
    html += '<td style="padding:6px;text-align:center;color:#ef4444;font-weight:bold;">' + r.notWorked + '</td>';
    html += '<td style="padding:6px;text-align:center;color:#f59e0b;">' + r.partial + '</td>';
    html += '<td style="padding:6px;text-align:center;color:#3b82f6;">' + r.worked + '</td>';
    html += '<td style="padding:6px;text-align:center;color:#10b981;">' + r.engaged + '</td>';
    html += '<td style="padding:6px;text-align:center;"><span style="background:' + cc + ';color:#fff;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:bold;">' + r.compliance + '%</span></td>';
    html += '</tr>';
  });
  html += '</table></div>';
  
  // Critical Leads List (Hot not worked)
  if (info.hotNotWorked.length > 0) {
    html += '<div style="padding:0 16px 16px;"><div style="font-size:14px;font-weight:bold;color:#ef4444;margin-bottom:8px;">🔥 HOT Leads Not Worked (' + info.hotNotWorked.length + ')</div>';
    html += '<table style="width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;font-size:10px;">';
    html += '<tr style="background:#991b1b;color:#fff;"><th style="padding:6px;">Rep</th><th style="padding:6px;">Lead</th><th style="padding:6px;">Move Date</th><th style="padding:6px;">Est $</th></tr>';
    info.hotNotWorked.slice(0,15).forEach(function(l, idx) {
      var bg = idx % 2 === 0 ? "#fff" : "#fef2f2";
      html += '<tr style="background:' + bg + ';"><td style="padding:4px 6px;">' + (l.user||l.jobRep) + '</td><td style="padding:4px 6px;">' + l.firstName + ' ' + l.lastName + '</td><td style="padding:4px 6px;font-weight:bold;color:#ef4444;">' + l.moveDateStr + '</td><td style="padding:4px 6px;">$' + l.estTotal.toLocaleString() + '</td></tr>';
    });
    html += '</table></div>';
  }
  
  // Footer
  html += '<div style="background:#1e293b;padding:16px;text-align:center;color:#64748b;font-size:10px;">Generated ' + info.timeStr + ' by ' + info.senderEmail + '<br>GRANOT Alert System v2.0</div>';
  html += '</body></html>';
  return html;
}


/**
 * ═══════════════════════════════════════════════════════════════
 * BUILD SUMMARY REPORT HTML
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_buildSummaryReportHTML_(leads, byRep, info, roster) {
  var engaged = leads.filter(function(l){return l.rc.workStatus==="ENGAGED";}).length;
  var worked = leads.filter(function(l){return l.rc.workStatus==="WORKED";}).length;
  var partial = leads.filter(function(l){return l.rc.workStatus==="PARTIAL";}).length;
  var notWorked = leads.filter(function(l){return l.rc.workStatus==="NOT_WORKED";}).length;
  
  var html = '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="margin:0;padding:0;font-family:Arial,sans-serif;background:#f1f5f9;">';
  
  html += '<div style="background:linear-gradient(135deg,#8B1538,#6B1028);padding:24px;text-align:center;">';
  html += '<div style="font-size:10px;color:#D4AF37;letter-spacing:2px;">🚢 SAFE SHIP MOVING SERVICES</div>';
  html += '<div style="font-size:24px;font-weight:bold;color:#fff;">GRANOT Lead Summary</div>';
  html += '<div style="font-size:12px;color:rgba(255,255,255,0.7);">' + info.dateStr + '</div>';
  html += '</div>';
  
  html += '<div style="background:#1e293b;padding:16px;display:flex;justify-content:center;gap:24px;flex-wrap:wrap;">';
  html += '<div style="text-align:center;"><div style="font-size:28px;font-weight:bold;color:#D4AF37;">' + leads.length + '</div><div style="font-size:9px;color:#94a3b8;">TOTAL</div></div>';
  html += '<div style="text-align:center;"><div style="font-size:28px;font-weight:bold;color:#ef4444;">' + notWorked + '</div><div style="font-size:9px;color:#94a3b8;">NOT WORKED</div></div>';
  html += '<div style="text-align:center;"><div style="font-size:28px;font-weight:bold;color:#f59e0b;">' + partial + '</div><div style="font-size:9px;color:#94a3b8;">PARTIAL</div></div>';
  html += '<div style="text-align:center;"><div style="font-size:28px;font-weight:bold;color:#3b82f6;">' + worked + '</div><div style="font-size:9px;color:#94a3b8;">WORKED</div></div>';
  html += '<div style="text-align:center;"><div style="font-size:28px;font-weight:bold;color:#10b981;">' + engaged + '</div><div style="font-size:9px;color:#94a3b8;">ENGAGED</div></div>';
  html += '</div>';
  
  // Rep table
  html += '<div style="padding:16px;"><table style="width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;">';
  html += '<tr style="background:#334155;color:#fff;font-size:11px;"><th style="padding:10px;text-align:left;">Rep</th><th>Total</th><th>❌</th><th>⚠️</th><th>✓</th><th>✅</th><th>Score</th></tr>';
  
  var repList = [];
  for (var rep in byRep) {
    var rl = byRep[rep];
    var e = rl.filter(function(l){return l.rc.workStatus==="ENGAGED";}).length;
    var w = rl.filter(function(l){return l.rc.workStatus==="WORKED";}).length;
    var p = rl.filter(function(l){return l.rc.workStatus==="PARTIAL";}).length;
    var n = rl.filter(function(l){return l.rc.workStatus==="NOT_WORKED";}).length;
    var c = rl.length > 0 ? Math.round(((e+w)/rl.length)*100) : 0;
    repList.push({ rep: rep, total: rl.length, engaged: e, worked: w, partial: p, notWorked: n, compliance: c });
  }
  repList.sort(function(a,b){ return a.compliance - b.compliance; });
  
  repList.forEach(function(r, idx) {
    var bg = idx % 2 === 0 ? "#fff" : "#f8fafc";
    var cc = r.compliance >= 80 ? "#10b981" : r.compliance >= 60 ? "#f59e0b" : "#ef4444";
    html += '<tr style="background:' + bg + ';font-size:12px;"><td style="padding:8px;font-weight:bold;">' + r.rep + '</td><td style="text-align:center;">' + r.total + '</td><td style="text-align:center;color:#ef4444;font-weight:bold;">' + r.notWorked + '</td><td style="text-align:center;color:#f59e0b;">' + r.partial + '</td><td style="text-align:center;color:#3b82f6;">' + r.worked + '</td><td style="text-align:center;color:#10b981;">' + r.engaged + '</td><td style="text-align:center;"><span style="background:' + cc + ';color:#fff;padding:2px 8px;border-radius:4px;">' + r.compliance + '%</span></td></tr>';
  });
  html += '</table></div>';
  
  html += '<div style="background:#1e293b;padding:16px;text-align:center;color:#64748b;font-size:10px;">Generated ' + info.timeStr + '<br>GRANOT Alert System v2.0</div>';
  html += '</body></html>';
  return html;
}


/**
 * ═══════════════════════════════════════════════════════════════
 * ALERT BUILDER UI HTML
 * ═══════════════════════════════════════════════════════════════
 */
function GRANOT_getAlertBuilderHTML_() {
  return '<!DOCTYPE html>\
<html>\
<head>\
  <base target="_top">\
  <style>\
    :root{--primary:#8B1538;--gold:#D4AF37;--success:#10b981;--warning:#f59e0b;--error:#ef4444;--bg:#0f172a;--card:#1e293b;--text:#e2e8f0;--muted:#94a3b8}\
    *{box-sizing:border-box;margin:0;padding:0}\
    body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;background:linear-gradient(180deg,var(--bg),#0c1322);color:var(--text);min-height:100vh}\
    .header{background:linear-gradient(135deg,var(--primary),#6B1028);padding:16px;text-align:center}\
    .brand{font-size:9px;color:var(--gold);letter-spacing:2px}\
    .title{font-size:18px;font-weight:800;color:#fff;margin-top:2px}\
    .content{padding:12px;overflow-y:auto;max-height:640px}\
    .section{background:var(--card);border-radius:10px;padding:12px;margin-bottom:10px}\
    .section-title{font-size:10px;font-weight:700;color:var(--gold);letter-spacing:1px;text-transform:uppercase;margin-bottom:8px}\
    .priority-grid{display:flex;flex-wrap:wrap;gap:6px}\
    .pri-btn{padding:10px 16px;border-radius:6px;font-size:16px;font-weight:700;cursor:pointer;border:2px solid #475569;background:#334155;color:#fff;transition:all 0.15s}\
    .pri-btn:hover{border-color:var(--gold)}\
    .pri-btn.selected{background:var(--gold);color:#000;border-color:var(--gold)}\
    .input-row{display:flex;gap:8px;margin-bottom:8px}\
    .input-group{flex:1}\
    .input-label{font-size:9px;color:var(--muted);margin-bottom:3px;text-transform:uppercase}\
    .input-field{width:100%;padding:8px;background:#0f172a;border:1px solid #334155;border-radius:6px;color:#fff;font-size:12px}\
    .input-field:focus{outline:none;border-color:var(--gold)}\
    select.input-field{cursor:pointer}\
    .checkbox-row{display:flex;align-items:center;gap:8px;padding:6px 0;font-size:11px}\
    .checkbox-row input{accent-color:var(--gold);width:16px;height:16px}\
    .flag-grid{display:flex;gap:8px;flex-wrap:wrap}\
    .flag-btn{padding:6px 12px;border-radius:6px;font-size:11px;cursor:pointer;border:1px solid #475569;background:#334155;color:#fff}\
    .flag-btn:hover{border-color:var(--gold)}\
    .flag-btn.selected{background:var(--gold);color:#000;border-color:var(--gold)}\
    .preview-box{background:#0f172a;border-radius:8px;padding:10px;max-height:160px;overflow-y:auto}\
    .preview-stat{display:flex;justify-content:space-between;padding:3px 0;font-size:11px;border-bottom:1px solid #334155}\
    .preview-stat:last-child{border-bottom:none}\
    .preview-stat .value.red{color:var(--error)}\
    .preview-stat .value.yellow{color:var(--warning)}\
    .preview-stat .value.blue{color:#3b82f6}\
    .preview-stat .value.green{color:var(--success)}\
    .btn-row{display:flex;gap:8px;padding:12px;background:var(--bg);border-top:1px solid #334155}\
    .btn{flex:1;padding:12px;border:none;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer}\
    .btn-preview{background:#334155;color:#fff}\
    .btn-send{background:linear-gradient(135deg,var(--gold),#F4D03F);color:#000}\
    .btn-send:disabled{background:#334155;color:#64748b;cursor:not-allowed}\
    .status-msg{text-align:center;padding:8px;font-size:11px;border-radius:6px;margin:8px 12px}\
    .status-msg.success{background:rgba(16,185,129,0.2);color:var(--success)}\
    .status-msg.error{background:rgba(239,68,68,0.2);color:var(--error)}\
    .status-msg.info{background:rgba(59,130,246,0.2);color:#3b82f6}\
    .preset-row{display:flex;gap:6px;margin-bottom:8px}\
    .preset-row select{flex:1}\
    .preset-row button{padding:6px 12px;border:none;border-radius:4px;cursor:pointer;font-size:11px}\
    .toggle{position:relative;display:inline-block;width:36px;height:18px;margin-left:auto}\
    .toggle input{opacity:0;width:0;height:0}\
    .toggle .slider{position:absolute;cursor:pointer;top:0;left:0;right:0;bottom:0;background:#334155;border-radius:18px;transition:0.2s}\
    .toggle .slider:before{position:absolute;content:"";height:12px;width:12px;left:3px;bottom:3px;background:#fff;border-radius:50%;transition:0.2s}\
    .toggle input:checked+.slider{background:var(--gold)}\
    .toggle input:checked+.slider:before{transform:translateX(18px)}\
  </style>\
</head>\
<body>\
  <div class="header">\
    <div class="brand">🚢 SAFE SHIP MOVING SERVICES</div>\
    <div class="title">GRANOT Alert Builder v2.0</div>\
  </div>\
  <div class="content">\
    <!-- Presets -->\
    <div class="section">\
      <div class="section-title">💾 Presets</div>\
      <div class="preset-row">\
        <select class="input-field" id="presetSelect"><option value="">-- Select Preset --</option></select>\
        <button style="background:#334155;color:#fff;" onclick="loadPreset()">Load</button>\
        <button style="background:var(--gold);color:#000;" onclick="savePreset()">Save</button>\
        <button style="background:#ef4444;color:#fff;" onclick="deletePreset()">Del</button>\
      </div>\
    </div>\
    <!-- Priority -->\
    <div class="section">\
      <div class="section-title">🎯 Priority (click to select)</div>\
      <div class="priority-grid" id="priorityGrid">\
        <button type="button" class="pri-btn" data-pri="0">0</button>\
        <button type="button" class="pri-btn" data-pri="1">1</button>\
        <button type="button" class="pri-btn" data-pri="2">2</button>\
        <button type="button" class="pri-btn" data-pri="3">3</button>\
        <button type="button" class="pri-btn" data-pri="4">4</button>\
        <button type="button" class="pri-btn" data-pri="5">5</button>\
        <button type="button" class="pri-btn" data-pri="6">6</button>\
        <button type="button" class="pri-btn" data-pri="7">7</button>\
        <button type="button" class="pri-btn" data-pri="8">8</button>\
      </div>\
    </div>\
    <!-- Date Range -->\
    <div class="section">\
      <div class="section-title">📅 Open Date Range</div>\
      <div class="input-row">\
        <div class="input-group"><div class="input-label">From</div><input type="date" class="input-field" id="dateFrom"></div>\
        <div class="input-group"><div class="input-label">To</div><input type="date" class="input-field" id="dateTo"></div>\
      </div>\
    </div>\
    <!-- Contact Thresholds -->\
    <div class="section">\
      <div class="section-title">📞 Contact Thresholds</div>\
      <div class="input-row">\
        <div class="input-group"><div class="input-label">Min Calls</div><input type="number" class="input-field" id="minCalls" min="0"></div>\
        <div class="input-group"><div class="input-label">Max Calls</div><input type="number" class="input-field" id="maxCalls" min="0"></div>\
        <div class="input-group"><div class="input-label">Min SMS</div><input type="number" class="input-field" id="minSMS" min="0"></div>\
        <div class="input-group"><div class="input-label">Max SMS</div><input type="number" class="input-field" id="maxSMS" min="0"></div>\
      </div>\
    </div>\
    <!-- Additional Filters -->\
    <div class="section">\
      <div class="section-title">🔍 Additional Filters</div>\
      <div class="input-row">\
        <div class="input-group"><div class="input-label">Move Within (days)</div><input type="number" class="input-field" id="moveDaysMax" min="0"></div>\
        <div class="input-group"><div class="input-label">Lead Age (min days)</div><input type="number" class="input-field" id="leadAgeDaysMin" min="0"></div>\
        <div class="input-group"><div class="input-label">Min Est Total ($)</div><input type="number" class="input-field" id="minEstTotal" min="0"></div>\
      </div>\
      <div class="section-title" style="margin-top:8px;">Highlight Flags</div>\
      <div class="flag-grid">\
        <button type="button" class="flag-btn" id="flagHot" onclick="toggleFlag(this)">🔥 Hot (≤7 days)</button>\
        <button type="button" class="flag-btn" id="flagHighValue" onclick="toggleFlag(this)">💰 High Value ($5K+)</button>\
        <button type="button" class="flag-btn" id="flagStale" onclick="toggleFlag(this)">⏰ Stale (3+ days)</button>\
      </div>\
    </div>\
    <!-- Send Options -->\
    <div class="section">\
      <div class="section-title">📧 Send Options</div>\
      <div class="input-row">\
        <div class="input-group"><div class="input-label">Send Mode</div>\
          <select class="input-field" id="sendMode">\
            <option value="BY_REP">Individual Rep Reports</option>\
            <option value="BY_TEAM">Team Reports (to Managers)</option>\
            <option value="SUMMARY">Summary Report</option>\
            <option value="ADMIN">Admin Health Report</option>\
          </select>\
        </div>\
      </div>\
      <div class="input-row" id="emailToRow" style="display:none;">\
        <div class="input-group"><div class="input-label">Email To</div><input type="email" class="input-field" id="emailTo" placeholder="Leave blank for your email"></div>\
      </div>\
      <div class="checkbox-row"><label for="ccManager">📋 CC Manager on rep reports</label><label class="toggle"><input type="checkbox" id="ccManager"><span class="slider"></span></label></div>\
      <div class="checkbox-row"><label for="testMode">🧪 Test Mode (send only to me)</label><label class="toggle"><input type="checkbox" id="testMode" checked><span class="slider"></span></label></div>\
    </div>\
    <!-- Preview -->\
    <div class="section">\
      <div class="section-title">👁️ Preview</div>\
      <div id="previewContent" class="preview-box"><div style="color:#94a3b8;text-align:center;padding:16px;">Select filters and click Preview</div></div>\
    </div>\
  </div>\
  <div class="btn-row">\
    <button class="btn btn-preview" onclick="loadPreview()">👁️ Preview</button>\
    <button class="btn btn-send" id="sendBtn" onclick="sendAlert()">📧 Send Alert</button>\
  </div>\
  <div id="statusMsg"></div>\
  <script>\
    var selectedPriorities = [];\
    var presets = {};\
    \
    document.querySelectorAll(".pri-btn").forEach(function(btn){\
      btn.addEventListener("click",function(){\
        var p = this.dataset.pri;\
        var i = selectedPriorities.indexOf(p);\
        if(i>=0){selectedPriorities.splice(i,1);this.classList.remove("selected");}\
        else{selectedPriorities.push(p);this.classList.add("selected");}\
      });\
    });\
    \
    document.getElementById("sendMode").addEventListener("change",function(){\
      document.getElementById("emailToRow").style.display=(this.value==="SUMMARY"||this.value==="ADMIN")?"flex":"none";\
      document.getElementById("ccManager").parentElement.parentElement.style.display=this.value==="BY_REP"?"flex":"none";\
    });\
    \
    function toggleFlag(btn){btn.classList.toggle("selected");}\
    \
    function getParams(){\
      return {\
        priorities:selectedPriorities.length?selectedPriorities:null,\
        dateFrom:document.getElementById("dateFrom").value||null,\
        dateTo:document.getElementById("dateTo").value||null,\
        minCalls:document.getElementById("minCalls").value||null,\
        maxCalls:document.getElementById("maxCalls").value||null,\
        minSMS:document.getElementById("minSMS").value||null,\
        maxSMS:document.getElementById("maxSMS").value||null,\
        moveDaysMax:document.getElementById("moveDaysMax").value||null,\
        leadAgeDaysMin:document.getElementById("leadAgeDaysMin").value||null,\
        minEstTotal:document.getElementById("minEstTotal").value||null,\
        hotOnly:document.getElementById("flagHot").classList.contains("selected"),\
        highValueOnly:document.getElementById("flagHighValue").classList.contains("selected"),\
        staleOnly:document.getElementById("flagStale").classList.contains("selected"),\
        sendMode:document.getElementById("sendMode").value,\
        emailTo:document.getElementById("emailTo").value||null,\
        ccManager:document.getElementById("ccManager").checked,\
        testMode:document.getElementById("testMode").checked\
      };\
    }\
    \
    function setParams(p){\
      selectedPriorities=p.priorities||[];\
      document.querySelectorAll(".pri-btn").forEach(function(btn){\
        btn.classList.toggle("selected",selectedPriorities.indexOf(btn.dataset.pri)>=0);\
      });\
      document.getElementById("dateFrom").value=p.dateFrom||"";\
      document.getElementById("dateTo").value=p.dateTo||"";\
      document.getElementById("minCalls").value=p.minCalls||"";\
      document.getElementById("maxCalls").value=p.maxCalls||"";\
      document.getElementById("minSMS").value=p.minSMS||"";\
      document.getElementById("maxSMS").value=p.maxSMS||"";\
      document.getElementById("moveDaysMax").value=p.moveDaysMax||"";\
      document.getElementById("leadAgeDaysMin").value=p.leadAgeDaysMin||"";\
      document.getElementById("minEstTotal").value=p.minEstTotal||"";\
      document.getElementById("flagHot").classList.toggle("selected",!!p.hotOnly);\
      document.getElementById("flagHighValue").classList.toggle("selected",!!p.highValueOnly);\
      document.getElementById("flagStale").classList.toggle("selected",!!p.staleOnly);\
    }\
    \
    function loadPreview(){\
      document.getElementById("previewContent").innerHTML=\'<div style="color:#94a3b8;text-align:center;">Loading...</div>\';\
      google.script.run.withSuccessHandler(function(d){\
        if(!d.success){document.getElementById("previewContent").innerHTML=\'<div style="color:#ef4444;">Error: \'+d.error+\'</div>\';return;}\
        var h=\'<div class="preview-stat"><span>Matching</span><span class="value">\'+d.totalFiltered+\' / \'+d.totalAll+\'</span></div>\';\
        h+=\'<div class="preview-stat"><span>❌ Not Worked</span><span class="value red">\'+d.byStatus.NOT_WORKED+\'</span></div>\';\
        h+=\'<div class="preview-stat"><span>⚠️ Partial</span><span class="value yellow">\'+d.byStatus.PARTIAL+\'</span></div>\';\
        h+=\'<div class="preview-stat"><span>✓ Worked</span><span class="value blue">\'+d.byStatus.WORKED+\'</span></div>\';\
        h+=\'<div class="preview-stat"><span>✅ Engaged</span><span class="value green">\'+d.byStatus.ENGAGED+\'</span></div>\';\
        h+=\'<div class="preview-stat"><span>🔥 Hot</span><span class="value">\'+d.hotCount+\'</span></div>\';\
        h+=\'<div class="preview-stat"><span>💰 High $</span><span class="value">\'+d.highValueCount+\'</span></div>\';\
        h+=\'<div class="preview-stat"><span>⏰ Stale</span><span class="value">\'+d.staleCount+\'</span></div>\';\
        document.getElementById("previewContent").innerHTML=h;\
      }).withFailureHandler(function(e){document.getElementById("previewContent").innerHTML=\'<div style="color:#ef4444;">\'+e.message+\'</div>\';}).GRANOT_getPreview(getParams());\
    }\
    \
    function sendAlert(){\
      var btn=document.getElementById("sendBtn");\
      btn.disabled=true;btn.textContent="Sending...";\
      showStatus("Sending...","info");\
      google.script.run.withSuccessHandler(function(r){\
        btn.disabled=false;btn.textContent="📧 Send Alert";\
        showStatus(r.message,r.success?"success":"error");\
      }).withFailureHandler(function(e){\
        btn.disabled=false;btn.textContent="📧 Send Alert";\
        showStatus(e.message,"error");\
      }).GRANOT_runAlert(getParams());\
    }\
    \
    function showStatus(m,t){\
      var e=document.getElementById("statusMsg");\
      e.className="status-msg "+t;e.textContent=m;e.style.display="block";\
      setTimeout(function(){e.style.display="none";},8000);\
    }\
    \
    function refreshPresets(){\
      google.script.run.withSuccessHandler(function(p){\
        presets=p;\
        var sel=document.getElementById("presetSelect");\
        sel.innerHTML=\'<option value="">-- Select Preset --</option>\';\
        for(var n in p)sel.innerHTML+=\'<option value="\'+n+\'">\'+n+\'</option>\';\
      }).GRANOT_getPresets();\
    }\
    \
    function loadPreset(){\
      var n=document.getElementById("presetSelect").value;\
      if(n&&presets[n])setParams(presets[n]);\
    }\
    \
    function savePreset(){\
      var n=prompt("Preset name:");\
      if(n)google.script.run.withSuccessHandler(function(r){showStatus(r.message,"success");refreshPresets();}).GRANOT_savePreset(n,getParams());\
    }\
    \
    function deletePreset(){\
      var n=document.getElementById("presetSelect").value;\
      if(n&&confirm("Delete preset \'"+n+"\'?"))google.script.run.withSuccessHandler(function(){refreshPresets();}).GRANOT_deletePreset(n);\
    }\
    \
    refreshPresets();\
  </script>\
</body>\
</html>';}
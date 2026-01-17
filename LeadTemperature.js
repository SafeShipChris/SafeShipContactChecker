/**************************************************************
 * 🚢 SAFE SHIP CONTACT CHECKER — UPGRADE IMPLEMENTATION GUIDE
 * Technical Specifications & Code Snippets
 * 
 * This file contains ready-to-implement code for the top
 * priority upgrades identified in the research document.
 **************************************************************/

/**
 * ═══════════════════════════════════════════════════════════════
 * 1. LEAD TEMPERATURE SCORING
 * Add this to Code.gs or create LeadTemperature.gs
 * ═══════════════════════════════════════════════════════════════
 */

/**
 * Calculate lead temperature score (0-100)
 * @param {Object} lead - Lead object with RC enrichment
 * @param {Object} granot - GRANOT data for the lead
 * @returns {Object} Temperature score and components
 */
function calculateLeadTemperature(lead, granot) {
  var rc = lead.rc || {};
  var now = new Date();
  
  // ═══════════════════════════════════════════════════════════════
  // RECENCY SCORE (0-30 points)
  // ═══════════════════════════════════════════════════════════════
  var recencyScore = 0;
  var lastContactDate = getLastContactDate_(rc);
  
  if (lastContactDate) {
    var daysSinceContact = Math.floor((now - lastContactDate) / (1000 * 60 * 60 * 24));
    if (daysSinceContact === 0) recencyScore = 30;
    else if (daysSinceContact === 1) recencyScore = 25;
    else if (daysSinceContact <= 3) recencyScore = 15;
    else if (daysSinceContact <= 7) recencyScore = 5;
    else recencyScore = 0;
  }
  
  // ═══════════════════════════════════════════════════════════════
  // ENGAGEMENT SCORE (0-30 points)
  // ═══════════════════════════════════════════════════════════════
  var engagementScore = 0;
  
  if (rc.hasReplied) engagementScore += 15;           // Lead replied to SMS
  if (rc.hasInbound) engagementScore += 10;           // Lead called back
  if (rc.hasLongCall) engagementScore += 5;           // Had conversation (4+ min)
  if (rc.hasConnectedCall) engagementScore += 5;      // Answered a call
  
  engagementScore = Math.min(30, engagementScore);    // Cap at 30
  
  // ═══════════════════════════════════════════════════════════════
  // VALUE SCORE (0-20 points)
  // ═══════════════════════════════════════════════════════════════
  var valueScore = 0;
  var estTotal = parseFloat(granot.estTotal) || 0;
  
  if (estTotal >= 8000) valueScore = 20;
  else if (estTotal >= 5000) valueScore = 15;
  else if (estTotal >= 3000) valueScore = 10;
  else if (estTotal >= 1500) valueScore = 5;
  else valueScore = 2;
  
  // ═══════════════════════════════════════════════════════════════
  // URGENCY SCORE (0-20 points)
  // ═══════════════════════════════════════════════════════════════
  var urgencyScore = 0;
  var moveDate = granot.moveDate ? new Date(granot.moveDate) : null;
  
  if (moveDate) {
    var daysToMove = Math.floor((moveDate - now) / (1000 * 60 * 60 * 24));
    if (daysToMove <= 3) urgencyScore = 20;
    else if (daysToMove <= 7) urgencyScore = 15;
    else if (daysToMove <= 14) urgencyScore = 10;
    else if (daysToMove <= 30) urgencyScore = 5;
    else urgencyScore = 2;
  }
  
  // ═══════════════════════════════════════════════════════════════
  // TOTAL TEMPERATURE
  // ═══════════════════════════════════════════════════════════════
  var totalTemp = recencyScore + engagementScore + valueScore + urgencyScore;
  
  return {
    temperature: totalTemp,
    recency: recencyScore,
    engagement: engagementScore,
    value: valueScore,
    urgency: urgencyScore,
    label: getTemperatureLabel_(totalTemp),
    emoji: getTemperatureEmoji_(totalTemp),
    color: getTemperatureColor_(totalTemp)
  };
}

function getTemperatureLabel_(temp) {
  if (temp >= 90) return "CRITICAL";
  if (temp >= 70) return "HOT";
  if (temp >= 50) return "WARM";
  if (temp >= 30) return "COOL";
  if (temp >= 10) return "COLD";
  return "ICE";
}

function getTemperatureEmoji_(temp) {
  if (temp >= 90) return "🔥🔥🔥🔥🔥";
  if (temp >= 70) return "🔥🔥🔥🔥⚪";
  if (temp >= 50) return "🔥🔥🔥⚪⚪";
  if (temp >= 30) return "🔥🔥⚪⚪⚪";
  if (temp >= 10) return "🔥⚪⚪⚪⚪";
  return "⚪⚪⚪⚪⚪";
}

function getTemperatureColor_(temp) {
  if (temp >= 90) return "#DC2626";  // Red
  if (temp >= 70) return "#EA580C";  // Orange
  if (temp >= 50) return "#F59E0B";  // Amber
  if (temp >= 30) return "#3B82F6";  // Blue
  if (temp >= 10) return "#6B7280";  // Gray
  return "#9CA3AF";                   // Light gray
}

function getLastContactDate_(rc) {
  // Find most recent contact from RC data
  var dates = [];
  if (rc.last5Calls && rc.last5Calls.length > 0) {
    dates.push(new Date(rc.last5Calls[0].timestamp));
  }
  if (rc.last5SMS && rc.last5SMS.length > 0) {
    dates.push(new Date(rc.last5SMS[0].timestamp));
  }
  return dates.length > 0 ? new Date(Math.max.apply(null, dates)) : null;
}


/**
 * ═══════════════════════════════════════════════════════════════
 * 2. SAFE SHIP SCORE (Rep Performance)
 * Add this to Code.gs or create RepScoring.gs
 * ═══════════════════════════════════════════════════════════════
 */

/**
 * Calculate Safe Ship Score for a rep
 * @param {string} username - Rep username
 * @param {Object} activityData - Aggregated activity data
 * @returns {Object} Score breakdown
 */
function calculateSafeShipScore(username, activityData) {
  var data = activityData || {};
  
  // ═══════════════════════════════════════════════════════════════
  // SPEED SCORE (0-25) - First contact velocity
  // ═══════════════════════════════════════════════════════════════
  var speedScore = 0;
  var avgFirstContactMinutes = data.avgFirstContactMinutes || 999;
  
  if (avgFirstContactMinutes <= 5) speedScore = 25;
  else if (avgFirstContactMinutes <= 15) speedScore = 22;
  else if (avgFirstContactMinutes <= 60) speedScore = 18;
  else if (avgFirstContactMinutes <= 240) speedScore = 12;
  else if (avgFirstContactMinutes <= 1440) speedScore = 6;
  else speedScore = 0;
  
  // ═══════════════════════════════════════════════════════════════
  // QUALITY SCORE (0-25) - Contact effectiveness
  // ═══════════════════════════════════════════════════════════════
  var qualityScore = 0;
  
  // Connect rate component (0-10)
  var connectRate = data.totalCalls > 0 ? 
    (data.connectedCalls / data.totalCalls) : 0;
  qualityScore += Math.min(10, connectRate * 15);
  
  // Talk time component (0-8)
  var avgDuration = data.avgCallDuration || 0; // in seconds
  qualityScore += Math.min(8, avgDuration / 15);
  
  // Reply rate component (0-7)
  var replyRate = data.totalSMS > 0 ? 
    (data.smsReplies / data.totalSMS) : 0;
  qualityScore += Math.min(7, replyRate * 10);
  
  // ═══════════════════════════════════════════════════════════════
  // CONSISTENCY SCORE (0-25) - Daily reliability
  // ═══════════════════════════════════════════════════════════════
  var consistencyScore = 25;
  var stdDev = data.ssStandardStdDev || 0;
  consistencyScore = Math.max(0, 25 - (stdDev * 3));
  
  // ═══════════════════════════════════════════════════════════════
  // ENGAGEMENT SCORE (0-25) - Lead response rate
  // ═══════════════════════════════════════════════════════════════
  var engagementScore = 0;
  
  // Reply rate (0-15)
  var leadReplyRate = data.totalLeads > 0 ?
    (data.leadsReplied / data.totalLeads) : 0;
  engagementScore += Math.min(15, leadReplyRate * 25);
  
  // Answer rate (0-10)
  var leadAnswerRate = data.totalLeads > 0 ?
    (data.leadsAnswered / data.totalLeads) : 0;
  engagementScore += Math.min(10, leadAnswerRate * 20);
  
  // ═══════════════════════════════════════════════════════════════
  // TOTAL SAFE SHIP SCORE
  // ═══════════════════════════════════════════════════════════════
  var totalScore = speedScore + qualityScore + consistencyScore + engagementScore;
  
  return {
    username: username,
    totalScore: Math.round(totalScore * 10) / 10,
    speedScore: Math.round(speedScore * 10) / 10,
    qualityScore: Math.round(qualityScore * 10) / 10,
    consistencyScore: Math.round(consistencyScore * 10) / 10,
    engagementScore: Math.round(engagementScore * 10) / 10,
    grade: getScoreGrade_(totalScore),
    trend: data.scoreTrend || "→"
  };
}

function getScoreGrade_(score) {
  if (score >= 90) return "A+";
  if (score >= 85) return "A";
  if (score >= 80) return "A-";
  if (score >= 75) return "B+";
  if (score >= 70) return "B";
  if (score >= 65) return "B-";
  if (score >= 60) return "C+";
  if (score >= 55) return "C";
  if (score >= 50) return "C-";
  if (score >= 45) return "D";
  return "F";
}


/**
 * ═══════════════════════════════════════════════════════════════
 * 3. DAILY ACTIVITY ARCHIVE
 * Add this as ActivityArchive.gs
 * ═══════════════════════════════════════════════════════════════
 */

/**
 * Archive daily activity for all reps
 * Run via daily trigger at 11:59 PM ET
 */
function archiveDailyActivity() {
  var CFG = getConfig_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var today = new Date();
  var todayStr = Utilities.formatDate(today, CFG.TZ, "yyyy-MM-dd");
  
  // Ensure archive sheet exists
  var archiveSheet = ensureArchiveSheet_(ss);
  
  // Get roster
  var roster = SSCCP_buildSalesRosterIndex_(CFG);
  var team = SSCCP_buildTeamRosterIndex_(CFG);
  
  // Build RC index for today
  var rcIndex = {};
  try {
    if (typeof RC_buildLookupIndex_ === "function") {
      rcIndex = RC_buildLookupIndex_();
    }
  } catch (e) {
    console.error("RC index failed:", e);
  }
  
  // Aggregate data by rep
  var repData = aggregateRepActivity_(CFG, rcIndex, roster, team);
  
  // Write archive rows
  var rows = [];
  Object.keys(repData).forEach(function(username) {
    var d = repData[username];
    rows.push([
      todayStr,
      username,
      d.manager || "",
      d.totalCalls || 0,
      d.connectedCalls || 0,
      d.totalSMS || 0,
      d.smsReplies || 0,
      d.vmLeft || 0,
      d.inboundCalls || 0,
      d.avgCallDuration || 0,
      d.ssStandards || 0,
      d.leadsWorked || 0,
      d.leadsEngaged || 0,
      d.newLeadsAssigned || 0,
      d.leadsBooked || 0,
      d.bookedValue || 0
    ]);
  });
  
  if (rows.length > 0) {
    archiveSheet.getRange(
      archiveSheet.getLastRow() + 1, 
      1, 
      rows.length, 
      rows[0].length
    ).setValues(rows);
  }
  
  // Prune old records (keep 365 days)
  pruneArchive_(archiveSheet, 365);
  
  console.log("Archived " + rows.length + " rep records for " + todayStr);
}

function ensureArchiveSheet_(ss) {
  var sheetName = "Activity_Archive";
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var headers = [
      "Date", "Username", "Manager", "TotalCalls", "ConnectedCalls",
      "TotalSMS", "SMSReplies", "VMLeft", "InboundCalls", "AvgCallDuration",
      "SSStandards", "LeadsWorked", "LeadsEngaged", "NewLeadsAssigned",
      "LeadsBooked", "BookedValue"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground("#1E3A5F")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

function aggregateRepActivity_(CFG, rcIndex, roster, team) {
  var repData = {};
  
  // Initialize all reps from roster
  Object.keys(roster.byUsername).forEach(function(username) {
    repData[username] = {
      manager: team.repToManager[username] || "",
      totalCalls: 0,
      connectedCalls: 0,
      totalSMS: 0,
      smsReplies: 0,
      vmLeft: 0,
      inboundCalls: 0,
      avgCallDuration: 0,
      ssStandards: 0,
      leadsWorked: 0,
      leadsEngaged: 0,
      newLeadsAssigned: 0,
      leadsBooked: 0,
      bookedValue: 0
    };
  });
  
  // Aggregate RC data by phone -> rep mapping
  // (This would need to be connected to actual RC log processing)
  
  return repData;
}

function pruneArchive_(sheet, daysToKeep) {
  var cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);
  
  var data = sheet.getDataRange().getValues();
  var rowsToDelete = [];
  
  for (var i = data.length - 1; i >= 1; i--) {
    var rowDate = new Date(data[i][0]);
    if (rowDate < cutoffDate) {
      rowsToDelete.push(i + 1); // 1-indexed
    }
  }
  
  // Delete in reverse order to maintain row indices
  rowsToDelete.forEach(function(row) {
    sheet.deleteRow(row);
  });
  
  if (rowsToDelete.length > 0) {
    console.log("Pruned " + rowsToDelete.length + " archive rows older than " + daysToKeep + " days");
  }
}


/**
 * ═══════════════════════════════════════════════════════════════
 * 4. ESCALATION ENGINE
 * Add this as EscalationEngine.gs
 * ═══════════════════════════════════════════════════════════════
 */

/**
 * Check for escalation triggers and send alerts
 * Run via trigger every 2 hours during business hours
 */
function checkEscalationTriggers() {
  var CFG = getConfig_();
  var now = new Date();
  var hour = now.getHours();
  
  // Only run during business hours (8 AM - 8 PM ET)
  if (hour < 8 || hour > 20) return;
  
  var escalations = [];
  
  // Check Level 1: Hot leads going stale
  var staleHotLeads = findStaleHotLeads_(CFG);
  staleHotLeads.forEach(function(lead) {
    escalations.push({
      level: 1,
      type: "STALE_HOT_LEAD",
      lead: lead,
      message: "Hot lead " + lead.jobNo + " has no contact in " + lead.hoursSinceContact + " hours"
    });
  });
  
  // Check Level 2: SLA breaches (24hr+ on hot leads)
  var slaBreaches = findSLABreaches_(CFG);
  slaBreaches.forEach(function(breach) {
    escalations.push({
      level: 2,
      type: "SLA_BREACH",
      lead: breach.lead,
      rep: breach.rep,
      manager: breach.manager,
      message: "SLA breach: " + breach.lead.jobNo + " (24+ hrs no contact)"
    });
  });
  
  // Check Level 3: High-value leads at risk
  var highValueAtRisk = findHighValueAtRisk_(CFG);
  highValueAtRisk.forEach(function(item) {
    escalations.push({
      level: 3,
      type: "HIGH_VALUE_RISK",
      lead: item.lead,
      rep: item.rep,
      message: "$" + item.lead.estTotal + " lead at risk: " + item.lead.jobNo
    });
  });
  
  // Process escalations
  processEscalations_(CFG, escalations);
}

function findStaleHotLeads_(CFG) {
  // Find P1-P3 leads with no contact in 4+ hours
  var staleLeads = [];
  // Implementation: Query trackers + RC data
  return staleLeads;
}

function findSLABreaches_(CFG) {
  // Find hot leads with 24+ hours no contact
  var breaches = [];
  // Implementation: Query trackers + RC data
  return breaches;
}

function findHighValueAtRisk_(CFG) {
  // Find $8K+ leads with 48+ hours no contact
  var atRisk = [];
  // Implementation: Query GRANOT + RC data
  return atRisk;
}

function processEscalations_(CFG, escalations) {
  escalations.forEach(function(esc) {
    switch (esc.level) {
      case 1:
        // Level 1: Add flag to next rep alert
        flagLeadForNextAlert_(esc.lead, esc.message);
        break;
      case 2:
        // Level 2: Send manager alert
        sendManagerEscalation_(CFG, esc);
        break;
      case 3:
        // Level 3: Send admin alert
        sendAdminEscalation_(CFG, esc);
        break;
    }
  });
}

function sendManagerEscalation_(CFG, escalation) {
  var subject = "🚨 [Level 2] Escalation Alert — " + escalation.type;
  var html = buildEscalationEmail_(escalation);
  
  // Get manager email
  var managerEmail = escalation.manager ? 
    resolveManagerEmail_(escalation.manager) : CFG.ADMIN_EMAIL;
  
  MailApp.sendEmail({
    to: managerEmail,
    subject: subject,
    htmlBody: html
  });
  
  // Also send Slack DM
  if (CFG.SLACK.BOT_TOKEN) {
    SSCCP_sendSlackDMToEmail_(CFG, managerEmail, 
      "🚨 *Escalation Alert*\n" + escalation.message);
  }
}

function sendAdminEscalation_(CFG, escalation) {
  var subject = "🔴 [Level 3] CRITICAL Escalation — " + escalation.type;
  var html = buildEscalationEmail_(escalation);
  
  (CFG.ADMIN_REPORT_EMAILS || [CFG.ADMIN_EMAIL]).forEach(function(email) {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: html
    });
  });
}

function buildEscalationEmail_(escalation) {
  var levelColors = {1: "#F59E0B", 2: "#EA580C", 3: "#DC2626"};
  var levelLabels = {1: "Rep Alert", 2: "Manager Alert", 3: "Admin Alert"};
  
  return '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;padding:20px;">' +
    '<div style="background:' + levelColors[escalation.level] + ';color:#fff;padding:15px;border-radius:8px;">' +
    '<h2 style="margin:0;">🚨 ' + levelLabels[escalation.level] + '</h2>' +
    '<p style="margin:10px 0 0;">' + escalation.type + '</p>' +
    '</div>' +
    '<div style="padding:20px;background:#f5f5f5;border-radius:8px;margin-top:15px;">' +
    '<p><strong>Message:</strong> ' + escalation.message + '</p>' +
    (escalation.lead ? '<p><strong>Job #:</strong> ' + escalation.lead.jobNo + '</p>' : '') +
    (escalation.rep ? '<p><strong>Rep:</strong> ' + escalation.rep + '</p>' : '') +
    '</div>' +
    '</body></html>';
}


/**
 * ═══════════════════════════════════════════════════════════════
 * 5. LEADERBOARD GENERATION
 * Add this as Leaderboard.gs
 * ═══════════════════════════════════════════════════════════════
 */

/**
 * Generate daily leaderboard
 * @returns {Array} Sorted array of rep performance objects
 */
function generateDailyLeaderboard() {
  var CFG = getConfig_();
  var roster = SSCCP_buildSalesRosterIndex_(CFG);
  var team = SSCCP_buildTeamRosterIndex_(CFG);
  
  var leaderboard = [];
  
  Object.keys(roster.byUsername).forEach(function(username) {
    var profile = roster.byUsername[username];
    
    // Get today's activity (would come from RC enrichment)
    var activity = getTodayActivity_(username);
    
    // Calculate Safe Ship Score
    var score = calculateSafeShipScore(username, activity);
    
    leaderboard.push({
      rank: 0, // Will be set after sorting
      username: username,
      name: profile.repName || username,
      manager: team.repToManager[username] || "",
      ssStandards: activity.ssStandards || 0,
      calls: activity.totalCalls || 0,
      sms: activity.totalSMS || 0,
      connects: activity.connectedCalls || 0,
      score: score.totalScore,
      grade: score.grade,
      streak: activity.streak || 0
    });
  });
  
  // Sort by score descending
  leaderboard.sort(function(a, b) {
    return b.score - a.score;
  });
  
  // Assign ranks
  leaderboard.forEach(function(rep, idx) {
    rep.rank = idx + 1;
  });
  
  return leaderboard;
}

function getTodayActivity_(username) {
  // Placeholder - would aggregate from RC logs
  return {
    ssStandards: 0,
    totalCalls: 0,
    totalSMS: 0,
    connectedCalls: 0,
    streak: 0
  };
}

/**
 * Generate leaderboard HTML for email footer
 */
function generateLeaderboardHTML(topN) {
  var leaderboard = generateDailyLeaderboard();
  var top = leaderboard.slice(0, topN || 5);
  
  var html = '<div style="background:#1E3A5F;padding:15px;border-radius:8px;margin-top:20px;">';
  html += '<div style="color:#D4AF37;font-size:12px;font-weight:bold;margin-bottom:10px;">🏆 TODAY\'S LEADERBOARD</div>';
  html += '<table style="width:100%;border-collapse:collapse;color:#fff;font-size:11px;">';
  html += '<tr style="border-bottom:1px solid rgba(255,255,255,0.2);">';
  html += '<th style="text-align:left;padding:5px;">Rank</th>';
  html += '<th style="text-align:left;padding:5px;">Rep</th>';
  html += '<th style="text-align:center;padding:5px;">SS Std</th>';
  html += '<th style="text-align:center;padding:5px;">Score</th>';
  html += '<th style="text-align:center;padding:5px;">Streak</th>';
  html += '</tr>';
  
  top.forEach(function(rep) {
    var rankEmoji = rep.rank === 1 ? "🥇" : (rep.rank === 2 ? "🥈" : (rep.rank === 3 ? "🥉" : rep.rank));
    var streakEmoji = rep.streak >= 5 ? "🔥 " + rep.streak : rep.streak;
    
    html += '<tr>';
    html += '<td style="padding:5px;">' + rankEmoji + '</td>';
    html += '<td style="padding:5px;">' + rep.name + '</td>';
    html += '<td style="text-align:center;padding:5px;">' + rep.ssStandards + '</td>';
    html += '<td style="text-align:center;padding:5px;">' + rep.score.toFixed(1) + '</td>';
    html += '<td style="text-align:center;padding:5px;">' + streakEmoji + '</td>';
    html += '</tr>';
  });
  
  html += '</table></div>';
  return html;
}


/**
 * ═══════════════════════════════════════════════════════════════
 * 6. CONVERSION FUNNEL TRACKING
 * Add this as ConversionFunnel.gs
 * ═══════════════════════════════════════════════════════════════
 */

/**
 * Calculate funnel metrics for a date range
 */
function calculateFunnelMetrics(startDate, endDate) {
  var CFG = getConfig_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get GRANOT data
  var granotSheet = ss.getSheetByName("GRANOT DATA");
  var granotData = granotSheet.getDataRange().getValues();
  
  var funnel = {
    newLeads: 0,
    contacted: 0,
    engaged: 0,
    quoted: 0,
    booked: 0,
    bookedValue: 0
  };
  
  // Process each lead
  for (var i = 1; i < granotData.length; i++) {
    var row = granotData[i];
    var openDate = row[2]; // open_date column
    var status = row[4];   // status column
    var estTotal = row[25]; // est_total column
    
    // Check if lead is in date range
    if (openDate >= startDate && openDate <= endDate) {
      funnel.newLeads++;
      
      // Check contact status (would need RC lookup)
      // funnel.contacted++;
      
      // Check engagement (would need RC lookup)
      // funnel.engaged++;
      
      // Check quoted (status = "Follow" implies contacted)
      if (status === "Follow" || status === "Booked") {
        funnel.quoted++;
      }
      
      // Check booked
      if (status === "Booked") {
        funnel.booked++;
        funnel.bookedValue += parseFloat(estTotal) || 0;
      }
    }
  }
  
  // Calculate rates
  funnel.contactRate = funnel.newLeads > 0 ? 
    (funnel.contacted / funnel.newLeads * 100).toFixed(1) + "%" : "N/A";
  funnel.engageRate = funnel.contacted > 0 ? 
    (funnel.engaged / funnel.contacted * 100).toFixed(1) + "%" : "N/A";
  funnel.quoteRate = funnel.engaged > 0 ? 
    (funnel.quoted / funnel.engaged * 100).toFixed(1) + "%" : "N/A";
  funnel.closeRate = funnel.quoted > 0 ? 
    (funnel.booked / funnel.quoted * 100).toFixed(1) + "%" : "N/A";
  funnel.overallConversion = funnel.newLeads > 0 ? 
    (funnel.booked / funnel.newLeads * 100).toFixed(1) + "%" : "N/A";
  
  return funnel;
}


/**
 * ═══════════════════════════════════════════════════════════════
 * 7. TRIGGER SETUP
 * Run once to set up all scheduled triggers
 * ═══════════════════════════════════════════════════════════════
 */

function setupUpgradeTriggers() {
  // Remove existing triggers for these functions
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    var funcName = trigger.getHandlerFunction();
    if (funcName === "archiveDailyActivity" || 
        funcName === "checkEscalationTriggers") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Daily archive at 11:59 PM ET
  ScriptApp.newTrigger("archiveDailyActivity")
    .timeBased()
    .atHour(23)
    .nearMinute(59)
    .everyDays(1)
    .inTimezone("America/New_York")
    .create();
  
  // Escalation check every 2 hours during business hours
  ScriptApp.newTrigger("checkEscalationTriggers")
    .timeBased()
    .everyHours(2)
    .create();
  
  SpreadsheetApp.getUi().alert(
    "✅ Upgrade Triggers Created\n\n" +
    "• Daily Archive: 11:59 PM ET\n" +
    "• Escalation Check: Every 2 hours"
  );
}
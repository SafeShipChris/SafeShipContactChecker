/**************************************************************
 * 🚢 SAFE SHIP CONTACT CHECKER — IndividualRepAlert.gs v1.0
 * ══════════════════════════════════════════════════════════
 * 
 * Send alerts to INDIVIDUAL sales reps instead of the whole floor
 * 
 * FEATURES:
 * ✅ Select alert type (Uncontacted, Quoted, Priority 1, Transfers)
 * ✅ Dropdown of all reps with lead counts
 * ✅ Preview leads before sending
 * ✅ Send to single rep only
 * ✅ Option to CC manager
 * ✅ Test mode support
 * 
 **************************************************************/

/**
 * ═══════════════════════════════════════════════════════════════
 * MENU ENTRY - Show the Individual Rep Alert Sidebar
 * ═══════════════════════════════════════════════════════════════
 */
function IRA_showIndividualRepAlert() {
  var html = HtmlService.createHtmlOutput(IRA_getSidebarHTML_())
    .setTitle("🚢 Send to Individual Rep")
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * ═══════════════════════════════════════════════════════════════
 * GET REPS WITH LEAD COUNTS FOR A GIVEN ALERT TYPE
 * Called from the sidebar to populate the rep dropdown
 * ═══════════════════════════════════════════════════════════════
 */
function IRA_getRepsForAlertType(alertType) {
  try {
    var CFG = getConfig_();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var roster = SSCCP_buildSalesRosterIndex_(CFG);
    var team = SSCCP_buildTeamRosterIndex_(CFG);
    
    // Define context based on alert type
    var ctx = IRA_getContextForAlertType_(alertType);
    var windowInfo = SSCCP_computeWindow_(CFG, new Date());
    
    // Scan sources to get leads
    var scan = SSCCP_scanSources_(CFG, ctx, windowInfo, roster, function(){});
    
    if (!scan.ok) {
      return { success: false, error: scan.summary };
    }
    
    // Enrich with RC data if available
    var rcIndex = {};
    try {
      if (typeof RC_buildLookupIndex_ === "function") {
        rcIndex = RC_buildLookupIndex_();
      }
    } catch (e) { /* RC not available */ }
    
    try {
      if (typeof RC_enrichLeads_ === "function") {
        RC_enrichLeads_(scan.leads, rcIndex);
      }
    } catch (e) { /* RC enrichment not available */ }
    
    // Group by rep
    var grouped = SSCCP_groupByRep_(scan.leads);
    
    // Build rep list with counts and details
    var repList = [];
    var reps = Object.keys(grouped);
    
    for (var i = 0; i < reps.length; i++) {
      var username = reps[i];
      var leads = grouped[username];
      var profile = roster.byUsername[username] || {};
      var managerName = team.repToManager[username] || "Unassigned";
      
      // Count by work status if RC data available
      var notWorked = 0, partial = 0, worked = 0, engaged = 0;
      leads.forEach(function(l) {
        if (l.rc) {
          var totalAttempts = (l.rc.callsYesterday || 0) + (l.rc.callsToday || 0) + 
                              (l.rc.smsYesterday || 0) + (l.rc.smsToday || 0);
          if (totalAttempts === 0) notWorked++;
          else if (totalAttempts < 5) partial++;
          else worked++;
          if (l.rc.hasReplied || l.rc.hasLongCall) engaged++;
        }
      });
      
      repList.push({
        username: username,
        name: profile.repName || username,
        email: profile.email || "",
        manager: managerName,
        leadCount: leads.length,
        notWorked: notWorked,
        partial: partial,
        worked: worked,
        engaged: engaged,
        hasEmail: !!(profile.email)
      });
    }
    
    // Sort by lead count (highest first)
    repList.sort(function(a, b) {
      return b.leadCount - a.leadCount;
    });
    
    return {
      success: true,
      alertType: alertType,
      alertTitle: ctx.title,
      totalLeads: scan.leads.length,
      totalReps: repList.length,
      reps: repList
    };
    
  } catch (e) {
    return { success: false, error: e.message };
  }
}


/**
 * ═══════════════════════════════════════════════════════════════
 * GET LEAD DETAILS FOR A SPECIFIC REP
 * Shows preview of what will be sent
 * ═══════════════════════════════════════════════════════════════
 */
function IRA_getLeadPreviewForRep(alertType, username) {
  try {
    var CFG = getConfig_();
    var roster = SSCCP_buildSalesRosterIndex_(CFG);
    var ctx = IRA_getContextForAlertType_(alertType);
    var windowInfo = SSCCP_computeWindow_(CFG, new Date());
    
    var scan = SSCCP_scanSources_(CFG, ctx, windowInfo, roster, function(){});
    if (!scan.ok) return { success: false, error: scan.summary };
    
    // Enrich with RC
    var rcIndex = {};
    try {
      if (typeof RC_buildLookupIndex_ === "function") rcIndex = RC_buildLookupIndex_();
      if (typeof RC_enrichLeads_ === "function") RC_enrichLeads_(scan.leads, rcIndex);
    } catch (e) { }
    
    // Filter to just this rep
    var repLeads = scan.leads.filter(function(l) {
      return normalizeUsername_(l.username) === normalizeUsername_(username);
    });
    
    // Build preview data
    var preview = repLeads.slice(0, 10).map(function(l) {
      var rc = l.rc || {};
      var totalAttempts = (rc.callsYesterday || 0) + (rc.callsToday || 0);
      var totalSMS = (rc.smsYesterday || 0) + (rc.smsToday || 0);
      return {
        job: l.job,
        source: l.source || l.trackerType,
        moveDate: l.moveDate || "",
        cubicFeet: l.cubicFeet || "",
        calls: totalAttempts,
        sms: totalSMS,
        hasReplied: rc.hasReplied || false,
        hasLongCall: rc.hasLongCall || false
      };
    });
    
    return {
      success: true,
      username: username,
      totalLeads: repLeads.length,
      preview: preview,
      moreCount: Math.max(0, repLeads.length - 10)
    };
    
  } catch (e) {
    return { success: false, error: e.message };
  }
}


/**
 * ═══════════════════════════════════════════════════════════════
 * SEND ALERT TO INDIVIDUAL REP
 * Main execution function
 * ═══════════════════════════════════════════════════════════════
 */
function IRA_sendToIndividualRep(params) {
  var startTime = new Date().getTime();
  
  try {
    var alertType = params.alertType;
    var username = params.username;
    var ccManager = params.ccManager || false;
    var testMode = params.testMode || false;
    
    var CFG = getConfig_();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var roster = SSCCP_buildSalesRosterIndex_(CFG);
    var team = SSCCP_buildTeamRosterIndex_(CFG);
    
    // Get rep profile
    var profile = roster.byUsername[normalizeUsername_(username)];
    if (!profile || !profile.email) {
      return { success: false, error: "Rep not found or missing email: " + username };
    }
    
    // Build context
    var ctx = IRA_getContextForAlertType_(alertType);
    var run = SSCCP_startRun_(ctx.reportType + "_INDIVIDUAL");
    var windowInfo = SSCCP_computeWindow_(CFG, new Date());
    windowInfo.label = SSCCP_getTrackerWindowLabel_(CFG, ctx);
    
    // Scan sources
    var scan = SSCCP_scanSources_(CFG, ctx, windowInfo, roster, function(){});
    if (!scan.ok) {
      return { success: false, error: "Scan failed: " + scan.summary };
    }
    
    // Enrich with RC data
    var rcIndex = {};
    try {
      if (typeof RC_buildLookupIndex_ === "function") rcIndex = RC_buildLookupIndex_();
      if (typeof RC_enrichLeads_ === "function") RC_enrichLeads_(scan.leads, rcIndex);
    } catch (e) { }
    
    // Filter to just this rep's leads
    var repLeads = scan.leads.filter(function(l) {
      return normalizeUsername_(l.username) === normalizeUsername_(username);
    });
    
    if (repLeads.length === 0) {
      return { success: false, error: "No leads found for rep: " + username };
    }
    
    // Apply dedupe
    var dedupeHours = CFG.DEDUPE_HOURS;
    var filtered = SSCCP_applyDedupe_(CFG, ctx.reportType, username, repLeads, dedupeHours);
    
    if (filtered.length === 0) {
      return { success: false, error: "All leads were filtered by dedupe (already notified recently)" };
    }
    
    // Determine email recipient
    var mode = SSCCP_mode_(CFG);
    var intendedEmail = profile.email;
    var toEmail = (testMode || mode.routeAllToAdmin) ? CFG.ADMIN_EMAIL : intendedEmail;
    
    // Build payload
    var payload = {
      reportType: ctx.reportType,
      reportTitle: ctx.title,
      runId: run.runId,
      window: windowInfo,
      rep: { 
        username: username, 
        name: profile.repName || username, 
        email: intendedEmail 
      },
      leads: filtered,
      mode: mode,
      routedBecauseUnassigned: false,
      showPhone: true
    };
    
    // Build subject
    var modePrefix = testMode ? "[TEST] " : (mode.routeAllToAdmin ? "[FORWARDED:" + intendedEmail + "] " : "");
    var subject = modePrefix + SSCCP_stopEmoji_("GREEN") + " Safe Ship — " + ctx.title + " — " + payload.rep.name + " (Run " + run.runId + ")";
    
    // Send email
    var html = SSCCP_renderRepEmailHtml_PREMIUM_(CFG, payload);
    var mailOptions = {
      to: toEmail,
      subject: subject,
      htmlBody: html,
      body: SSCCP_stripHtml_(html)
    };
    
    // Add CC for manager if requested
    if (ccManager && !testMode) {
      var managerName = team.repToManager[normalizeUsername_(username)];
      if (managerName) {
        var managerEmail = SSCCP_resolveManagerEmail_(roster, managerName);
        if (managerEmail) {
          mailOptions.cc = managerEmail;
        }
      }
    }
    
    MailApp.sendEmail(mailOptions);
    
    // Log notification
    SSCCP_logNotification_(CFG, {
      timestamp: new Date(),
      runId: run.runId,
      type: "Individual Rep Alert (" + ctx.reportType + ")",
      route: testMode ? "TEST" : (mode.routeAllToAdmin ? "Forward-All/Admin" : "Direct"),
      stoplight: "GREEN",
      username: username,
      repName: payload.rep.name,
      email: toEmail,
      manager: mailOptions.cc || "",
      sourceSheets: filtered.map(function(x) { return x.source; }).filter(Boolean).join(", "),
      leadCount: filtered.length,
      jobNumbers: filtered.map(function(x) { return x.job; }).slice(0, 60).join(", "),
      emailSent: "YES",
      slackSent: "NO",
      slackLookup: "NO",
      error: ""
    });
    
    // Send Slack DM if not test mode
    var slackSent = false;
    if (!testMode && !mode.routeAllToAdmin) {
      var slackRes = SSCCP_sendSlackDMToEmail_(CFG, toEmail, SSCCP_renderRepSlackText_(payload));
      slackSent = slackRes.ok;
    }
    
    var elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(1);
    
    return {
      success: true,
      message: "✅ Sent " + filtered.length + " leads to " + (testMode ? "ADMIN (test mode)" : payload.rep.name),
      details: {
        rep: payload.rep.name,
        email: toEmail,
        leadCount: filtered.length,
        ccManager: mailOptions.cc || null,
        slackSent: slackSent,
        runId: run.runId,
        elapsed: elapsed
      }
    };
    
  } catch (e) {
    return { success: false, error: e.message };
  }
}


/**
 * ═══════════════════════════════════════════════════════════════
 * HELPER - Get context object for alert type
 * ═══════════════════════════════════════════════════════════════
 */
function IRA_getContextForAlertType_(alertType) {
  switch (alertType) {
    case "UNCONTACTED":
      return { reportType: "UNCONTACTED", title: "Uncontacted Leads Dashboard", sources: ["SMS", "CALL"] };
    case "QUOTED_FOLLOWUP":
      return { reportType: "QUOTED_FOLLOWUP", title: "Quoted Follow-Up Dashboard", sources: ["CONTACTED"] };
    case "PRIORITY1_CALLVM":
      return { reportType: "PRIORITY1_CALLVM", title: "Priority 1 Call/VM Dashboard", sources: ["P1CALL"] };
    case "SAME_DAY_TRANSFERS":
      return { reportType: "SAME_DAY_TRANSFERS", title: "Same Day Transfers Dashboard", sources: ["TRANSFERS"] };
    default:
      return { reportType: "UNCONTACTED", title: "Uncontacted Leads Dashboard", sources: ["SMS", "CALL"] };
  }
}


/**
 * ═══════════════════════════════════════════════════════════════
 * SIDEBAR HTML
 * ═══════════════════════════════════════════════════════════════
 */
function IRA_getSidebarHTML_() {
  return '<!DOCTYPE html>\
<html>\
<head>\
  <base target="_top">\
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">\
  <style>\
    :root{--primary:#8B1538;--gold:#D4AF37;--success:#10b981;--warning:#f59e0b;--error:#ef4444;--bg:#0f172a;--card:#1e293b;--text:#e2e8f0;--muted:#94a3b8}\
    *{box-sizing:border-box;margin:0;padding:0}\
    body{font-family:Inter,-apple-system,sans-serif;background:linear-gradient(180deg,var(--bg),#0c1322);color:var(--text);min-height:100vh}\
    .header{background:linear-gradient(135deg,var(--primary),#6B1028);padding:20px;text-align:center}\
    .brand{font-size:9px;color:var(--gold);letter-spacing:2px;text-transform:uppercase}\
    .title{font-size:18px;font-weight:800;color:#fff;margin-top:4px}\
    .subtitle{font-size:11px;color:rgba(255,255,255,0.7);margin-top:2px}\
    .content{padding:16px}\
    .section{background:var(--card);border-radius:12px;padding:14px;margin-bottom:12px}\
    .section-title{font-size:10px;font-weight:700;color:var(--gold);letter-spacing:1px;text-transform:uppercase;margin-bottom:10px;display:flex;align-items:center;gap:6px}\
    .select-wrap{position:relative}\
    select{width:100%;padding:12px;background:#0f172a;border:1px solid #334155;border-radius:8px;color:#fff;font-size:13px;cursor:pointer;appearance:none}\
    select:focus{outline:none;border-color:var(--gold)}\
    .select-arrow{position:absolute;right:12px;top:50%;transform:translateY(-50%);color:var(--muted);pointer-events:none}\
    .rep-card{background:#0f172a;border-radius:8px;padding:12px;margin-top:10px;display:none}\
    .rep-card.visible{display:block}\
    .rep-name{font-size:15px;font-weight:700;color:#fff}\
    .rep-meta{font-size:11px;color:var(--muted);margin-top:2px}\
    .rep-stats{display:flex;gap:8px;margin-top:10px;flex-wrap:wrap}\
    .rep-stat{background:#334155;padding:6px 10px;border-radius:6px;text-align:center;min-width:50px}\
    .rep-stat-val{font-size:16px;font-weight:800;color:#fff}\
    .rep-stat-lbl{font-size:8px;color:var(--muted);text-transform:uppercase}\
    .rep-stat.red .rep-stat-val{color:var(--error)}\
    .rep-stat.yellow .rep-stat-val{color:var(--warning)}\
    .rep-stat.green .rep-stat-val{color:var(--success)}\
    .lead-preview{margin-top:10px;max-height:150px;overflow-y:auto}\
    .lead-item{display:flex;justify-content:space-between;align-items:center;padding:6px 8px;background:#1e293b;border-radius:4px;margin-bottom:4px;font-size:11px}\
    .lead-job{font-weight:600;color:#fff}\
    .lead-source{color:var(--muted);font-size:10px}\
    .lead-badges{display:flex;gap:4px}\
    .lead-badge{padding:2px 6px;border-radius:3px;font-size:9px;font-weight:600}\
    .lead-badge.calls{background:#dcfce7;color:#166534}\
    .lead-badge.sms{background:#dbeafe;color:#1e40af}\
    .lead-badge.reply{background:#d1fae5;color:#047857}\
    .checkbox-row{display:flex;align-items:center;gap:10px;padding:8px 0;font-size:12px}\
    .checkbox-row input{accent-color:var(--gold);width:18px;height:18px;cursor:pointer}\
    .toggle-wrap{margin-left:auto}\
    .btn{width:100%;padding:14px;border:none;border-radius:10px;font-size:14px;font-weight:700;cursor:pointer;transition:all 0.2s}\
    .btn-send{background:linear-gradient(135deg,var(--gold),#F4D03F);color:#000}\
    .btn-send:hover{transform:scale(1.02)}\
    .btn-send:disabled{background:#334155;color:#64748b;cursor:not-allowed;transform:none}\
    .btn-close{background:#334155;color:#fff;margin-top:8px}\
    .status{padding:10px;border-radius:8px;margin-bottom:12px;font-size:12px;text-align:center;display:none}\
    .status.success{display:block;background:rgba(16,185,129,0.2);color:var(--success)}\
    .status.error{display:block;background:rgba(239,68,68,0.2);color:var(--error)}\
    .status.loading{display:block;background:rgba(59,130,246,0.2);color:#3b82f6}\
    .loading-spinner{display:inline-block;width:14px;height:14px;border:2px solid #3b82f6;border-top-color:transparent;border-radius:50%;animation:spin 0.8s linear infinite;margin-right:6px;vertical-align:middle}\
    @keyframes spin{to{transform:rotate(360deg)}}\
    .empty-state{text-align:center;padding:20px;color:var(--muted);font-size:12px}\
    .summary-bar{background:var(--primary);padding:8px 12px;border-radius:6px;margin-bottom:10px;display:flex;justify-content:space-between;font-size:11px;color:#fff}\
    .footer{text-align:center;padding:12px;font-size:9px;color:#475569}\
  </style>\
</head>\
<body>\
  <div class="header">\
    <div class="brand">🚢 Safe Ship Contact Checker</div>\
    <div class="title">Send to Individual Rep</div>\
    <div class="subtitle">Target specific reps instead of the whole floor</div>\
  </div>\
  \
  <div class="content">\
    <div id="statusBox" class="status"></div>\
    \
    <!-- Step 1: Select Alert Type -->\
    <div class="section">\
      <div class="section-title">📋 Step 1: Select Alert Type</div>\
      <div class="select-wrap">\
        <select id="alertType" onchange="loadReps()">\
          <option value="">-- Choose Alert Type --</option>\
          <option value="UNCONTACTED">📉 Uncontacted Leads (SMS + Call/VM)</option>\
          <option value="QUOTED_FOLLOWUP">💰 Quoted Follow-Up (Contacted Leads)</option>\
          <option value="PRIORITY1_CALLVM">🔥 Priority 1 Call/VM</option>\
          <option value="SAME_DAY_TRANSFERS">⚡ Same Day Transfers</option>\
        </select>\
        <span class="select-arrow">▼</span>\
      </div>\
      <div id="alertSummary" class="summary-bar" style="display:none;">\
        <span id="summaryTitle">-</span>\
        <span><strong id="summaryTotal">0</strong> leads • <strong id="summaryReps">0</strong> reps</span>\
      </div>\
    </div>\
    \
    <!-- Step 2: Select Rep -->\
    <div class="section" id="repSection" style="opacity:0.5;pointer-events:none;">\
      <div class="section-title">👤 Step 2: Select Sales Rep</div>\
      <div class="select-wrap">\
        <select id="repSelect" onchange="loadRepDetails()">\
          <option value="">-- Select a rep --</option>\
        </select>\
        <span class="select-arrow">▼</span>\
      </div>\
      \
      <div class="rep-card" id="repCard">\
        <div class="rep-name" id="repName">-</div>\
        <div class="rep-meta" id="repMeta">-</div>\
        <div class="rep-stats">\
          <div class="rep-stat"><div class="rep-stat-val" id="statTotal">0</div><div class="rep-stat-lbl">Total</div></div>\
          <div class="rep-stat red"><div class="rep-stat-val" id="statNotWorked">0</div><div class="rep-stat-lbl">Not Worked</div></div>\
          <div class="rep-stat yellow"><div class="rep-stat-val" id="statPartial">0</div><div class="rep-stat-lbl">Partial</div></div>\
          <div class="rep-stat green"><div class="rep-stat-val" id="statWorked">0</div><div class="rep-stat-lbl">Worked</div></div>\
        </div>\
        <div class="lead-preview" id="leadPreview"></div>\
      </div>\
    </div>\
    \
    <!-- Step 3: Options -->\
    <div class="section" id="optionsSection" style="opacity:0.5;pointer-events:none;">\
      <div class="section-title">⚙️ Step 3: Options</div>\
      <div class="checkbox-row">\
        <input type="checkbox" id="ccManager">\
        <label for="ccManager">📋 CC Manager on email</label>\
      </div>\
      <div class="checkbox-row">\
        <input type="checkbox" id="testMode" checked>\
        <label for="testMode">🧪 Test Mode (send to admin only)</label>\
      </div>\
    </div>\
    \
    <!-- Send Button -->\
    <button class="btn btn-send" id="sendBtn" disabled onclick="sendAlert()">📧 Send Alert</button>\
    <button class="btn btn-close" onclick="google.script.host.close()">Close</button>\
  </div>\
  \
  <div class="footer">Individual Rep Alert System v1.0</div>\
  \
  <script>\
    var currentReps = [];\
    var selectedRep = null;\
    \
    function showStatus(msg, type) {\
      var box = document.getElementById("statusBox");\
      box.className = "status " + type;\
      box.innerHTML = (type === "loading" ? \'<span class="loading-spinner"></span>\' : "") + msg;\
    }\
    \
    function hideStatus() {\
      document.getElementById("statusBox").className = "status";\
    }\
    \
    function loadReps() {\
      var alertType = document.getElementById("alertType").value;\
      if (!alertType) {\
        document.getElementById("repSection").style.opacity = "0.5";\
        document.getElementById("repSection").style.pointerEvents = "none";\
        document.getElementById("alertSummary").style.display = "none";\
        return;\
      }\
      \
      showStatus("Loading reps...", "loading");\
      \
      google.script.run\
        .withSuccessHandler(function(result) {\
          hideStatus();\
          if (!result.success) {\
            showStatus("Error: " + result.error, "error");\
            return;\
          }\
          \
          currentReps = result.reps;\
          \
          // Show summary\
          document.getElementById("alertSummary").style.display = "flex";\
          document.getElementById("summaryTitle").textContent = result.alertTitle;\
          document.getElementById("summaryTotal").textContent = result.totalLeads;\
          document.getElementById("summaryReps").textContent = result.totalReps;\
          \
          // Populate rep dropdown\
          var select = document.getElementById("repSelect");\
          select.innerHTML = \'<option value="">-- Select a rep (\' + result.totalReps + \' with leads) --</option>\';\
          \
          result.reps.forEach(function(rep) {\
            var opt = document.createElement("option");\
            opt.value = rep.username;\
            var emailStatus = rep.hasEmail ? "" : " ⚠️ No Email";\
            opt.textContent = rep.name + " (" + rep.leadCount + " leads)" + emailStatus;\
            if (!rep.hasEmail) opt.style.color = "#f59e0b";\
            select.appendChild(opt);\
          });\
          \
          // Enable section\
          document.getElementById("repSection").style.opacity = "1";\
          document.getElementById("repSection").style.pointerEvents = "auto";\
          \
          // Reset rep card\
          document.getElementById("repCard").classList.remove("visible");\
          document.getElementById("optionsSection").style.opacity = "0.5";\
          document.getElementById("optionsSection").style.pointerEvents = "none";\
          document.getElementById("sendBtn").disabled = true;\
        })\
        .withFailureHandler(function(err) {\
          showStatus("Error: " + err.message, "error");\
        })\
        .IRA_getRepsForAlertType(alertType);\
    }\
    \
    function loadRepDetails() {\
      var username = document.getElementById("repSelect").value;\
      if (!username) {\
        document.getElementById("repCard").classList.remove("visible");\
        document.getElementById("optionsSection").style.opacity = "0.5";\
        document.getElementById("optionsSection").style.pointerEvents = "none";\
        document.getElementById("sendBtn").disabled = true;\
        return;\
      }\
      \
      // Find rep in current list\
      selectedRep = currentReps.find(function(r) { return r.username === username; });\
      if (!selectedRep) return;\
      \
      // Update rep card\
      document.getElementById("repName").textContent = selectedRep.name;\
      document.getElementById("repMeta").textContent = "@" + selectedRep.username + " • " + (selectedRep.email || "No email") + " • Manager: " + selectedRep.manager;\
      document.getElementById("statTotal").textContent = selectedRep.leadCount;\
      document.getElementById("statNotWorked").textContent = selectedRep.notWorked;\
      document.getElementById("statPartial").textContent = selectedRep.partial;\
      document.getElementById("statWorked").textContent = selectedRep.worked + selectedRep.engaged;\
      \
      document.getElementById("repCard").classList.add("visible");\
      \
      // Enable options\
      document.getElementById("optionsSection").style.opacity = "1";\
      document.getElementById("optionsSection").style.pointerEvents = "auto";\
      document.getElementById("sendBtn").disabled = !selectedRep.hasEmail;\
      \
      if (!selectedRep.hasEmail) {\
        showStatus("⚠️ This rep has no email address in Sales_Roster", "error");\
      } else {\
        hideStatus();\
      }\
      \
      // Load lead preview\
      loadLeadPreview();\
    }\
    \
    function loadLeadPreview() {\
      var alertType = document.getElementById("alertType").value;\
      var username = document.getElementById("repSelect").value;\
      if (!alertType || !username) return;\
      \
      document.getElementById("leadPreview").innerHTML = \'<div class="empty-state"><span class="loading-spinner"></span> Loading leads...</div>\';\
      \
      google.script.run\
        .withSuccessHandler(function(result) {\
          if (!result.success) {\
            document.getElementById("leadPreview").innerHTML = \'<div class="empty-state">Error loading preview</div>\';\
            return;\
          }\
          \
          var html = "";\
          result.preview.forEach(function(lead) {\
            html += \'<div class="lead-item">\';\
            html += \'<div><span class="lead-job">\' + lead.job + \'</span><br><span class="lead-source">\' + lead.source + \'</span></div>\';\
            html += \'<div class="lead-badges">\';\
            if (lead.calls > 0) html += \'<span class="lead-badge calls">\' + lead.calls + \' calls</span>\';\
            if (lead.sms > 0) html += \'<span class="lead-badge sms">\' + lead.sms + \' sms</span>\';\
            if (lead.hasReplied) html += \'<span class="lead-badge reply">💬</span>\';\
            html += \'</div></div>\';\
          });\
          \
          if (result.moreCount > 0) {\
            html += \'<div class="empty-state">+ \' + result.moreCount + \' more leads</div>\';\
          }\
          \
          document.getElementById("leadPreview").innerHTML = html || \'<div class="empty-state">No leads to preview</div>\';\
        })\
        .withFailureHandler(function(err) {\
          document.getElementById("leadPreview").innerHTML = \'<div class="empty-state">Error: \' + err.message + \'</div>\';\
        })\
        .IRA_getLeadPreviewForRep(alertType, username);\
    }\
    \
    function sendAlert() {\
      var alertType = document.getElementById("alertType").value;\
      var username = document.getElementById("repSelect").value;\
      var ccManager = document.getElementById("ccManager").checked;\
      var testMode = document.getElementById("testMode").checked;\
      \
      if (!alertType || !username) {\
        showStatus("Please select alert type and rep", "error");\
        return;\
      }\
      \
      var btn = document.getElementById("sendBtn");\
      btn.disabled = true;\
      btn.textContent = "Sending...";\
      showStatus("Sending alert to " + selectedRep.name + "...", "loading");\
      \
      google.script.run\
        .withSuccessHandler(function(result) {\
          btn.textContent = "📧 Send Alert";\
          if (result.success) {\
            showStatus(result.message, "success");\
            btn.disabled = false;\
          } else {\
            showStatus("Error: " + result.error, "error");\
            btn.disabled = false;\
          }\
        })\
        .withFailureHandler(function(err) {\
          btn.disabled = false;\
          btn.textContent = "📧 Send Alert";\
          showStatus("Error: " + err.message, "error");\
        })\
        .IRA_sendToIndividualRep({\
          alertType: alertType,\
          username: username,\
          ccManager: ccManager,\
          testMode: testMode\
        });\
    }\
  </script>\
</body>\
</html>';
}
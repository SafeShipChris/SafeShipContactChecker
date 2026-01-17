/**************************************************************
 * 🚢 SAFE SHIP CONTACT CHECKER — ContactChecker_Progress.gs v1.2
 * 
 * 
 * Premium Sidebar Progress UI for Contact Checker Alerts
 * 
 * FEATURES:
 * - Real-time progress tracking with animated ring
 * - 8-stage pipeline visualization
 * - Live activity log with timestamps
 * - Final stats summary
 * - Consistent Safe Ship branding
 * - v1.2: Improved cancellation (checks after every email/Slack)
 * 
 **************************************************************/
/* ═══════════════════════════════════════════════════════════════
 * SHARED BRAND CONFIG - Matches SSCCP_AD_CONFIG.BRAND
 * ═══════════════════════════════════════════════════════════════ */

var UNIFIED_SIDEBAR_BRAND = {
  PRIMARY: "#8B1538",      // Burgundy
  PRIMARY_DARK: "#6B1028", // Dark burgundy
  GOLD: "#D4AF37",         // Gold accent
  DARK_BG: "#0f172a",      // Dark background
  CARD_BG: "#1e293b",      // Card background
  SUCCESS: "#10b981",      // Green
  WARNING: "#f59e0b",      // Yellow
  ERROR: "#ef4444",        // Red
  MUTED: "#94a3b8"         // Gray text
};


/* ═══════════════════════════════════════════════════════════════
 * SHARED CSS - Used by all sidebars
 * ═══════════════════════════════════════════════════════════════ */

function getUnifiedSidebarCSS_() {
  var brand = UNIFIED_SIDEBAR_BRAND;
  return `
    * { box-sizing: border-box; margin: 0; padding: 0; }
    
    :root {
      --primary: ${brand.PRIMARY};
      --primary-dark: ${brand.PRIMARY_DARK};
      --gold: ${brand.GOLD};
      --dark-bg: ${brand.DARK_BG};
      --card-bg: ${brand.CARD_BG};
      --success: ${brand.SUCCESS};
      --warning: ${brand.WARNING};
      --error: ${brand.ERROR};
      --muted: ${brand.MUTED};
    }
    
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: var(--dark-bg);
      color: #fff;
      font-size: 12px;
      min-height: 100vh;
    }
    
    /* Header - Burgundy gradient with gold border */
    .header {
      background: linear-gradient(135deg, var(--primary) 0%, var(--primary-dark) 100%);
      padding: 16px;
      text-align: center;
      border-bottom: 3px solid var(--gold);
    }
    .header .brand {
      font-size: 11px;
      color: var(--gold);
      letter-spacing: 2px;
      text-transform: uppercase;
      font-weight: 700;
    }
    .header .title {
      font-size: 18px;
      font-weight: 800;
      margin-top: 4px;
    }
    .header .subtitle {
      font-size: 10px;
      color: rgba(255,255,255,0.7);
      margin-top: 2px;
    }
    
    .content { padding: 12px; }
    
    /* Mode badge - Pill style */
    .mode-badge {
      display: inline-block;
      padding: 6px 12px;
      border-radius: 20px;
      font-size: 11px;
      font-weight: 700;
      margin-bottom: 12px;
    }
    .mode-badge.live { background: var(--success); }
    .mode-badge.test { background: var(--warning); color: #000; }
    .mode-badge.safe { background: #3b82f6; }
    .mode-badge.forward { background: #8b5cf6; }
    
    /* Ring progress */
    .ring-section { text-align: center; margin-bottom: 16px; }
    .ring-wrap { position: relative; width: 100px; height: 100px; margin: 0 auto; }
    .ring-svg { transform: rotate(-90deg); width: 100%; height: 100%; }
    .ring-bg { fill: none; stroke: var(--card-bg); stroke-width: 8; }
    .ring-fill {
      fill: none;
      stroke: url(#ringGrad);
      stroke-width: 8;
      stroke-linecap: round;
      stroke-dasharray: 345.4;
      stroke-dashoffset: 345.4;
      transition: stroke-dashoffset 0.5s ease;
    }
    .ring-center {
      position: absolute;
      top: 50%; left: 50%;
      transform: translate(-50%, -50%);
      text-align: center;
    }
    .ring-pct { font-size: 24px; font-weight: 800; color: var(--gold); }
    .ring-label { font-size: 9px; color: var(--muted); text-transform: uppercase; }
    .stage-name { font-size: 14px; font-weight: 700; margin-top: 8px; }
    .stage-status { font-size: 11px; color: var(--muted); }
    
    /* Section titles */
    .section-title {
      font-size: 10px;
      font-weight: 700;
      color: var(--gold);
      letter-spacing: 1px;
      text-transform: uppercase;
      margin-bottom: 8px;
      display: flex;
      align-items: center;
      gap: 6px;
    }
    
    /* Stats grid */
    .stats-grid {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 6px;
      margin-bottom: 12px;
    }
    .stats-grid-2 { grid-template-columns: repeat(2, 1fr); }
    .stat-box {
      background: var(--card-bg);
      border-radius: 8px;
      padding: 8px 4px;
      text-align: center;
    }
    .stat-val { font-size: 16px; font-weight: 800; }
    .stat-val.calls { color: #3b82f6; }
    .stat-val.sms { color: #10b981; }
    .stat-val.new { color: var(--gold); }
    .stat-val.time { color: #8b5cf6; }
    .stat-lbl { font-size: 8px; color: var(--muted); text-transform: uppercase; }
    
    /* Breakdown section */
    .breakdown {
      background: var(--card-bg);
      border-radius: 10px;
      padding: 10px;
      margin-bottom: 12px;
    }
    .breakdown-row {
      display: flex;
      justify-content: space-between;
      padding: 4px 0;
      border-bottom: 1px solid rgba(255,255,255,0.05);
    }
    .breakdown-row:last-child { border-bottom: none; }
    .breakdown-label { color: var(--muted); }
    .breakdown-value { font-weight: 700; }
    
    /* Pipeline stages */
    .pipeline {
      background: var(--card-bg);
      border-radius: 10px;
      padding: 10px;
      margin-bottom: 12px;
    }
    .stage {
      display: flex;
      align-items: center;
      gap: 8px;
      padding: 6px 0;
      border-bottom: 1px solid rgba(255,255,255,0.05);
    }
    .stage:last-child { border-bottom: none; }
    .stage-icon {
      width: 24px;
      height: 24px;
      border-radius: 50%;
      background: var(--card-bg);
      border: 2px solid var(--muted);
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 10px;
      flex-shrink: 0;
    }
    .stage.complete .stage-icon { background: var(--success); border-color: var(--success); }
    .stage.active .stage-icon { background: var(--gold); border-color: var(--gold); animation: pulse 1s infinite; }
    .stage.error .stage-icon { background: var(--error); border-color: var(--error); }
    .stage-info { flex: 1; min-width: 0; }
    .stage-name-text { font-size: 11px; font-weight: 600; }
    .stage-detail { font-size: 9px; color: var(--muted); }
    .stage-count {
      font-size: 10px;
      font-weight: 700;
      color: var(--success);
      background: rgba(16,185,129,0.2);
      padding: 2px 6px;
      border-radius: 6px;
      display: none;
    }
    
    @keyframes pulse {
      0%, 100% { opacity: 1; }
      50% { opacity: 0.5; }
    }
    
    /* Buttons */
    .btn-group { display: flex; flex-direction: column; gap: 8px; margin-bottom: 12px; }
    .btn {
      width: 100%;
      padding: 12px;
      border: none;
      border-radius: 10px;
      font-size: 12px;
      font-weight: 700;
      cursor: pointer;
      transition: all 0.2s;
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
    }
    .btn:disabled { opacity: 0.5; cursor: not-allowed; }
    .btn-live {
      background: linear-gradient(135deg, var(--primary), var(--primary-dark));
      color: #fff;
    }
    .btn-live:hover:not(:disabled) { transform: scale(1.02); }
    .btn-cancel { background: var(--error); color: #fff; display: none; }
    .btn-cancel:hover:not(:disabled) { background: #dc2626; }
    .btn-close {
      background: var(--card-bg);
      color: #fff;
      border: 1px solid rgba(255,255,255,0.2);
      display: none;
    }
    .btn-close:hover { background: rgba(255,255,255,0.1); }
    
    /* Action log */
    .log-section {
      background: var(--card-bg);
      border-radius: 10px;
      padding: 10px;
      max-height: 150px;
      overflow-y: auto;
    }
    .log-entry {
      display: flex;
      gap: 6px;
      padding: 3px 0;
      font-size: 10px;
      border-bottom: 1px solid rgba(255,255,255,0.03);
    }
    .log-entry:last-child { border-bottom: none; }
    .log-time { color: #64748b; font-family: monospace; min-width: 50px; flex-shrink: 0; }
    .log-msg { color: var(--muted); flex: 1; }
    .log-entry.success .log-msg { color: var(--success); }
    .log-entry.error .log-msg { color: var(--error); }
    .log-entry.warn .log-msg { color: var(--warning); }
    .log-entry.info .log-msg { color: #3b82f6; }
    
    /* Footer */
    .footer { text-align: center; padding: 8px; font-size: 8px; color: #475569; }
  `;
}

/* 
 * PROGRESS STORAGE (uses CacheService)
 * ═══════════════════════════════════════════════════════════════ */

var CC_PROGRESS_KEY = "CC_PROGRESS_V1";

function CC_setProgress_(data) {
  CacheService.getUserCache().put(CC_PROGRESS_KEY, JSON.stringify(data), 600);
}

function CC_getProgress() {
  var raw = CacheService.getUserCache().get(CC_PROGRESS_KEY);
  return raw ? JSON.parse(raw) : null;
}

function CC_clearProgress_() {
  CacheService.getUserCache().remove(CC_PROGRESS_KEY);
}

function CC_requestCancel() {
  var progress = CC_getProgress();
  if (progress) {
    progress.cancelRequested = true;
    CC_setProgress_(progress);
  }
}

function CC_isCancelled_() {
  var progress = CC_getProgress();
  return progress && progress.cancelRequested === true;
}

function CC_addLog_(msg, type) {
  var progress = CC_getProgress();
  if (progress) {
    progress.log = progress.log || [];
    progress.log.push({ 
      time: Utilities.formatDate(new Date(), "America/New_York", "HH:mm:ss"), 
      msg: msg, 
      type: type || "info" 
    });
    if (progress.log.length > 50) progress.log = progress.log.slice(-50);
    CC_setProgress_(progress);
  }
}

function CC_updateStage_(stageNum, state, status, percent, currentStage, mainStatus, count) {
  var progress = CC_getProgress();
  if (!progress) return;
  progress.percent = percent;
  progress.currentStage = currentStage;
  progress.status = mainStatus;
  progress["stage" + stageNum] = { state: state, status: status };
  if (count !== undefined) progress["stage" + stageNum].count = count;
  CC_setProgress_(progress);
}


/* 
 * MENU ENTRY POINTS (Show sidebar, user clicks Start)
 * ═══════════════════════════════════════════════════════════════ */

function CC_runUncontactedLeadsUI() {
  CC_showAlertSidebar_({
    reportType: "UNCONTACTED",
    title: "Uncontacted Leads",
    icon: "📌",
    description: "Leads that haven't been contacted yet",
    startFunction: "CC_startUncontactedLeads"
  });
}

function CC_runQuotedFollowUpUI() {
  CC_showAlertSidebar_({
    reportType: "QUOTED_FOLLOWUP",
    title: "Quoted Follow-Up",
    icon: "📋",
    description: "Quoted leads needing follow-up",
    startFunction: "CC_startQuotedFollowUp"
  });
}

function CC_runPriority1CallVmUI() {
  CC_showAlertSidebar_({
    reportType: "PRIORITY1_CALLVM",
    title: "Priority 1 Call/VM",
    icon: "🥅",
    description: "High priority leads requiring immediate action",
    startFunction: "CC_startPriority1CallVm"
  });
}

function CC_runSameDayTransfersUI() {
  CC_showAlertSidebar_({
    reportType: "SAME_DAY_TRANSFERS",
    title: "Same Day Transfers",
    icon: "🔄",
    description: "Transfers scheduled for today",
    startFunction: "CC_startSameDayTransfers"
  });
}

function CC_showAlertSidebar_(config) {
  CC_clearProgress_();
  
  var html = HtmlService.createHtmlOutput(getContactCheckerSidebarHtml_(config))
    .setTitle("🚢 " + config.title)
    .setWidth(380);
  
  SpreadsheetApp.getUi().showSidebar(html);
}


/* 
 * START FUNCTIONS (Called from sidebar)
 * ═══════════════════════════════════════════════════════════════ */

function CC_startUncontactedLeads() {
  return CC_runPipelineWithProgress_({
    reportType: "UNCONTACTED",
    title: "Uncontacted Leads Dashboard",
    sources: ["SMS", "CALL"]
  });
}

function CC_startQuotedFollowUp() {
  return CC_runPipelineWithProgress_({
    reportType: "QUOTED_FOLLOWUP",
    title: "Quoted Follow-Up Dashboard",
    sources: ["CONTACTED"]
  });
}

function CC_startPriority1CallVm() {
  return CC_runPipelineWithProgress_({
    reportType: "PRIORITY1_CALLVM",
    title: "Priority 1 Call/VM Dashboard",
    sources: ["P1CALL"]
  });
}

function CC_startSameDayTransfers() {
  return CC_runPipelineWithProgress_({
    reportType: "SAME_DAY_TRANSFERS",
    title: "Same Day Transfers Dashboard",
    sources: ["TRANSFERS"]
  });
}


/* 
 * PIPELINE WITH PROGRESS TRACKING
 * ═══════════════════════════════════════════════════════════════ */

function CC_runPipelineWithProgress_(ctx) {
  var startTime = new Date().getTime();
  var CFG = getConfig_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Initialize progress with 8 stages
  CC_setProgress_({
    percent: 0,
    currentStage: "Initializing...",
    status: "Starting " + ctx.reportType + "...",
    stage1: { state: "waiting", status: "Waiting..." },
    stage2: { state: "waiting", status: "Waiting..." },
    stage3: { state: "waiting", status: "Waiting..." },
    stage4: { state: "waiting", status: "Waiting..." },
    stage5: { state: "waiting", status: "Waiting..." },
    stage6: { state: "waiting", status: "Waiting..." },
    stage7: { state: "waiting", status: "Waiting..." },
    stage8: { state: "waiting", status: "Waiting..." },
    log: [{ time: Utilities.formatDate(new Date(), "America/New_York", "HH:mm:ss"), msg: "🚀 " + ctx.reportType + " started", type: "info" }],
    complete: false,
    startTime: startTime,
    stats: {}
  });
  
  Utilities.sleep(300);
  
  try {
    var run = SSCCP_startRun_(ctx.reportType);
    var now = new Date();
    var windowInfo = SSCCP_computeWindow_(CFG, now);
    windowInfo.label = SSCCP_getTrackerWindowLabel_(CFG, ctx);
    
    CC_addLog_("📋 Run ID: " + run.runId, "info");
    
    // 
    // STAGE 1: Gate Checks
    // 
    CC_updateStage_(1, "active", "Checking...", 5, "Gate Checks", "Verifying configuration...");
    CC_addLog_("🔒 Stage 1: Gate Checks", "info");
    
    SSCCP_ensureNotificationLog_(CFG);
    SSCCP_ensureRunLog_(CFG);
    
    var gate = SSCCP_runGateChecks_(CFG, ctx, function(){});
    if (!gate.ok) {
      CC_updateStage_(1, "error", "Failed!", 5, "Gate Check Failed", gate.summary);
      CC_addLog_("❌ Gate Check Failed: " + gate.summary, "error");
      
      SSCCP_logRunEvent_(CFG, run, "GATE_FAIL", gate.summary, gate.details);
      SSCCP_sendAdminDashboard_(CFG, run, ctx, windowInfo, "RED", {
        metrics: { totalLeadsScanned: 0, repsWithLeads: 0, managersSummarized: 0 },
        notes: [], exceptions: [{ title: "Gate Check Failure", items: [gate.summary].concat(gate.details || []) }],
        teamAgg: {}, performance: SSCCP_buildAdminPerformancePack_(CFG, { leads: [] }, {}, {}, {}, { duplicates: [], unassigned: [] }, { exceptions: [] }, { exceptions: [] }),
      });
      
      var progress = CC_getProgress();
      progress.complete = true;
      progress.error = true;
      progress.errorMsg = gate.summary;
      CC_setProgress_(progress);
      
      return { success: false, error: gate.summary };
    }
    
    CC_updateStage_(1, "complete", "Passed!", 10, "Gate Checks", "All checks passed");
    CC_addLog_("✅ Gate checks passed", "success");
    
    Utilities.sleep(200);
    
    // 
    // STAGE 2: Load Sales Roster
    // 
    CC_updateStage_(2, "active", "Loading...", 15, "Sales Roster", "Loading Sales_Roster...");
    CC_addLog_("👥 Stage 2: Loading Sales_Roster", "info");
    
    var roster = SSCCP_buildSalesRosterIndex_(CFG);
    var rosterCount = Object.keys(roster.byUsername).length;
    
    CC_updateStage_(2, "complete", rosterCount + " reps", 22, "Sales Roster", rosterCount + " sales reps loaded", rosterCount);
    CC_addLog_("✅ Loaded " + rosterCount + " sales reps", "success");
    
    Utilities.sleep(200);
    
    // 
    // STAGE 3: Load Team Roster
    // 
    CC_updateStage_(3, "active", "Loading...", 28, "Team Roster", "Loading Team_Roster...");
    CC_addLog_("💔 Stage 3: Loading Team_Roster", "info");
    
    var team = SSCCP_buildTeamRosterIndex_(CFG);
    var teamCount = team.managerHeaders ? team.managerHeaders.length : 0;
    
    CC_updateStage_(3, "complete", teamCount + " managers", 35, "Team Roster", teamCount + " managers loaded", teamCount);
    CC_addLog_("✅ Loaded " + teamCount + " managers", "success");
    
    Utilities.sleep(200);
    
    // 
    // STAGE 4: Load RingCentral Data
    // 
    CC_updateStage_(4, "active", "Loading...", 40, "RingCentral Data", "Building contact index...");
    CC_addLog_("📞 Stage 4: Loading RingCentral logs", "info");
    
    var rcIndex = {};
    var rcCount = 0;
    try {
      if (typeof RC_buildLookupIndex_ === "function") {
        rcIndex = RC_buildLookupIndex_();
        rcCount = Object.keys(rcIndex).length;
        CC_addLog_("✅ Indexed " + rcCount + " phone numbers", "success");
      } else {
        CC_addLog_("⭐️ RC enrichment not available", "skip");
      }
    } catch (e) {
      CC_addLog_("⚠️ RC load warning: " + e.message, "skip");
    }
    
    CC_updateStage_(4, rcCount > 0 ? "complete" : "skipped", rcCount > 0 ? rcCount + " phones" : "Skipped", 48, "RingCentral Data", rcCount > 0 ? rcCount + " phones indexed" : "Continuing without RC data", rcCount);
    
    Utilities.sleep(200);
    
    // 
    // STAGE 5: Scan Sources
    // 
    CC_updateStage_(5, "active", "Scanning...", 52, "Scanning Sources", "Reading tracker data...");
    CC_addLog_("📝 Stage 5: Scanning sources", "info");
    
    var scan = SSCCP_scanSources_(CFG, ctx, windowInfo, roster, function(){});
    
    if (!scan.ok) {
      CC_updateStage_(5, "error", "Failed!", 52, "Scan Failed", scan.summary);
      CC_addLog_("❌ Scan Failed: " + scan.summary, "error");
      
      SSCCP_logRunEvent_(CFG, run, "SCAN_FAIL", scan.summary, scan.details);
      SSCCP_sendAdminDashboard_(CFG, run, ctx, windowInfo, "RED", {
        metrics: { totalLeadsScanned: 0, repsWithLeads: 0, managersSummarized: 0 },
        notes: [], exceptions: [{ title: "Scan Failure", items: [scan.summary].concat(scan.details || []) }],
        teamAgg: {}, performance: SSCCP_buildAdminPerformancePack_(CFG, { leads: [] }, {}, roster, team, { duplicates: [], unassigned: [] }, { exceptions: [] }, { exceptions: [] }),
      });
      
      var progress = CC_getProgress();
      progress.complete = true;
      progress.error = true;
      progress.errorMsg = scan.summary;
      CC_setProgress_(progress);
      
      return { success: false, error: scan.summary };
    }
    
    var leadCount = scan.leads.length;
    CC_updateStage_(5, "complete", leadCount + " leads", 60, "Scanning Sources", leadCount + " leads found", leadCount);
    CC_addLog_("✅ Found " + leadCount + " leads", "success");
    
    Utilities.sleep(200);
    
    // 
    // STAGE 6: Enrich & Group
    // 
    CC_updateStage_(6, "active", "Processing...", 65, "Enriching Data", "Adding contact history...");
    CC_addLog_("🔗 Stage 6: Enriching & grouping", "info");
    
    try {
      if (typeof RC_enrichLeads_ === "function" && rcCount > 0) {
        RC_enrichLeads_(scan.leads, rcIndex);
        CC_addLog_("✅ Added RC contact history", "success");
      }
    } catch (e) {
      CC_addLog_("⚠️ RC enrichment warning: " + e.message, "skip");
    }
    
    var grouped = SSCCP_groupByRep_(scan.leads);
    var repCount = Object.keys(grouped).length;
    var integrity = SSCCP_computeIntegrity_(team, grouped);
    var mode = SSCCP_mode_(CFG);
    var stoplightBase = integrity.duplicates.length ? "RED" : (integrity.unassigned.length ? "YELLOW" : "GREEN");
    
    CC_updateStage_(6, "complete", repCount + " reps", 72, "Enriching Data", "Grouped to " + repCount + " reps", repCount);
    CC_addLog_("✅ Grouped to " + repCount + " reps", "success");
    
    if (integrity.duplicates.length) {
      CC_addLog_("⚠️ " + integrity.duplicates.length + " duplicates found", "skip");
    }
    if (integrity.unassigned.length) {
      CC_addLog_("⚠️ " + integrity.unassigned.length + " unassigned leads", "skip");
    }
    
    Utilities.sleep(200);
    
    // 
    // STAGE 7: Send Rep Alerts (with granular progress)
    // 
    CC_updateStage_(7, "active", "Starting...", 75, "Rep Alerts", "Preparing to send alerts...");
    CC_addLog_("📧 Stage 7: Sending Rep Alerts", "info");
    
    var repSend = CC_sendRepAlertsWithProgress_(CFG, run, ctx, windowInfo, grouped, roster, integrity, repCount);
    var repsSent = repSend.repsSent || 0;
    var emailsSent = repSend.emailsSent || 0;
    var slacksSent = repSend.slacksSent || 0;
    
    // Check if cancelled during rep alerts
    if (repSend.cancelled) {
      var elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(1);
      CC_updateStage_(7, "error", "Cancelled", 88, "Cancelled", "User cancelled operation");
      
      var progress = CC_getProgress();
      progress.complete = true;
      progress.cancelled = true;
      progress.stats = { leads: leadCount, reps: repCount, emails: emailsSent, slack: slacksSent, managers: 0, time: elapsed };
      CC_setProgress_(progress);
      
      return { success: false, cancelled: true, leads: leadCount, emails: emailsSent, time: elapsed };
    }
    
    CC_updateStage_(7, "complete", repsSent + " reps", 88, "Rep Alerts", emailsSent + " emails, " + slacksSent + " Slack", repsSent);
    CC_addLog_("✅ Sent to " + repsSent + " reps (" + emailsSent + " emails, " + slacksSent + " Slack)", "success");
    
    Utilities.sleep(200);
    
    // 
    // STAGE 8: Send Manager Reports (with granular progress)
    // 
    CC_updateStage_(8, "active", "Starting...", 89, "Manager Reports", "Preparing manager summaries...");
    CC_addLog_("📊 Stage 8: Sending Manager Reports", "info");
    
    var mgrSend = CC_sendManagerReportsWithProgress_(CFG, run, ctx, windowInfo, grouped, roster, team, integrity);
    var mgrsSent = mgrSend.managersSent || 0;
    
    // Check if cancelled during manager reports
    if (mgrSend.cancelled) {
      var elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(1);
      CC_updateStage_(8, "error", "Cancelled", 96, "Cancelled", "User cancelled operation");
      
      var progress = CC_getProgress();
      progress.complete = true;
      progress.cancelled = true;
      progress.stats = { leads: leadCount, reps: repCount, emails: emailsSent, slack: slacksSent, managers: mgrsSent, time: elapsed };
      CC_setProgress_(progress);
      
      return { success: false, cancelled: true, leads: leadCount, emails: emailsSent, managers: mgrsSent, time: elapsed };
    }
    
    CC_updateStage_(8, "complete", mgrsSent + " managers", 96, "Manager Reports", mgrsSent + " managers notified", mgrsSent);
    CC_addLog_("✅ Sent to " + mgrsSent + " managers", "success");
    
    // 
    // FINALIZE & SEND ADMIN DASHBOARD
    // 
    var finalStoplight = SSCCP_mergeStoplight_(stoplightBase, repSend, mgrSend);
    var teamAgg = SSCCP_buildTeamAgg_(grouped, roster, team, integrity);
    var notes = [].concat(scan.notes || []).concat(["Mode: " + mode.effectiveModeLabel]);
    
    var metrics = {
      totalLeadsScanned: scan.leads.length,
      repsWithLeads: Object.keys(grouped).length,
      managersSummarized: mgrSend.managersSent || 0,
    };
    
    SSCCP_sendAdminDashboard_(CFG, run, ctx, windowInfo, finalStoplight, {
      metrics: metrics, notes: notes,
      exceptions: SSCCP_buildExceptionBlocks_(integrity, scan, repSend, mgrSend),
      teamAgg: teamAgg,
      performance: SSCCP_buildAdminPerformancePack_(CFG, scan, grouped, roster, team, integrity, repSend, mgrSend),
    });
    
    SSCCP_logRunEvent_(CFG, run, "SUCCESS", ctx.reportType + " completed", [
      "Stoplight=" + finalStoplight, "Leads=" + metrics.totalLeadsScanned,
      "Reps=" + metrics.repsWithLeads, "Managers=" + metrics.managersSummarized,
    ]);
    
    // 
    // COMPLETE!
    // 
    var elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(1);
    
    var progress = CC_getProgress();
    progress.percent = 100;
    progress.currentStage = "Complete!";
    progress.status = ctx.reportType + " completed successfully!";
    progress.complete = true;
    progress.stoplight = finalStoplight;
    progress.stats = {
      leads: leadCount,
      reps: repCount,
      emails: emailsSent,
      slack: slacksSent,
      managers: mgrsSent,
      time: elapsed,
      stoplight: finalStoplight
    };
    progress.log.push({ 
      time: Utilities.formatDate(new Date(), "America/New_York", "HH:mm:ss"), 
      msg: "🎉 COMPLETE! " + finalStoplight + " — " + leadCount + " leads in " + elapsed + "s", 
      type: "success" 
    });
    CC_setProgress_(progress);
    
    return { 
      success: true, 
      stoplight: finalStoplight,
      leads: leadCount,
      reps: repCount,
      managers: mgrsSent,
      time: elapsed 
    };
    
  } catch (e) {
    Logger.log("Pipeline error: " + e);
    CC_addLog_("❌ ERROR: " + e.message, "error");
    
    var progress = CC_getProgress() || { log: [] };
    progress.percent = 0;
    progress.currentStage = "Error";
    progress.status = e.message;
    progress.complete = true;
    progress.error = true;
    progress.errorMsg = e.message;
    CC_setProgress_(progress);
    
    return { success: false, error: e.message };
  }
}


/* 
 * REP ALERTS WITH PROGRESS - Sends alerts with live updates
 * v1.2 - Improved cancellation checking
 * ═══════════════════════════════════════════════════════════════ */

function CC_sendRepAlertsWithProgress_(CFG, run, ctx, windowInfo, grouped, roster, integrity, totalReps) {
  var mode = SSCCP_mode_(CFG);
  var dedupeHours = CFG.DEDUPE_HOURS;
  var exceptions = [];
  var repsAlerted = 0;
  var emailsSent = 0;
  var slacksSent = 0;
  var skippedDedupe = 0;
  var skippedNoEmail = 0;

  CC_addLog_("⚙️ Dedupe window: " + dedupeHours + " hours", "info");

  var dupSet = new Set(integrity.duplicates || []);
  var unassignedSet = new Set(integrity.unassigned || []);
  var reps = Object.keys(grouped || {});
  var total = reps.length;

  for (var i = 0; i < reps.length; i++) {
    // Check for cancellation at START of each iteration
    if (CC_isCancelled_()) {
      CC_addLog_("⚠️ Cancelled by user at rep " + (i + 1) + "/" + total, "error");
      return { repsSent: repsAlerted, emailsSent: emailsSent, slacksSent: slacksSent, exceptions: exceptions, cancelled: true };
    }
    
    var username = reps[i];
    var leads = grouped[username] || [];
    
    // Update progress every rep
    var pct = Math.round(75 + (i / total) * 13);
    CC_updateStage_(7, "active", (i + 1) + "/" + total, pct, "Rep Alerts", "Processing " + username + "...");

    if (dupSet.has(username)) { exceptions.push({ title: "Duplicate Assignment (Skipped Rep)", items: [username] }); continue; }

    var profile = roster.byUsername[username];
    if (!profile || !profile.email) { 
      skippedNoEmail++;
      exceptions.push({ title: "Missing Sales_Roster Email", items: [username] }); 
      continue; 
    }

    var isUnassigned = unassignedSet.has(username);
    var intendedEmail = isUnassigned ? CFG.ADMIN_EMAIL : profile.email;
    var filtered = SSCCP_applyDedupe_(CFG, ctx.reportType, username, leads, dedupeHours);
    
    if (!filtered.length) { 
      skippedDedupe++;
      continue; 
    }
    
    CC_addLog_("📤 Sending to: " + username + " (" + filtered.length + " leads)", "info");

    var toEmail = mode.routeAllToAdmin ? CFG.ADMIN_EMAIL : intendedEmail;

    var payload = {
      reportType: ctx.reportType, reportTitle: ctx.title, runId: run.runId, window: windowInfo,
      rep: { username: username, name: profile.repName || username, email: intendedEmail },
      leads: filtered, mode: mode, routedBecauseUnassigned: isUnassigned,
      showPhone: true,
    };

    var subject = SSCCP_stopEmoji_("GREEN") + " Safe Ship — " + ctx.title + " — " + payload.rep.name + " (Run " + run.runId + ")";

    try {
      var html = SSCCP_renderRepEmailHtml_PREMIUM_(CFG, payload);
      MailApp.sendEmail({
        to: toEmail,
        subject: mode.routeAllToAdmin ? "[FORWARDED:" + intendedEmail + "] " + subject : subject,
        htmlBody: html,
        body: SSCCP_stripHtml_(html),
      });
      emailsSent++;

      SSCCP_logNotification_(CFG, {
        timestamp: new Date(), runId: run.runId, type: "Rep Alert (" + ctx.reportType + ")",
        route: mode.routeAllToAdmin ? "Forward-All/Admin" : (isUnassigned ? "Unassigned->Admin" : "Direct"),
        stoplight: "GREEN", username: username, repName: payload.rep.name, email: toEmail, manager: "",
        sourceSheets: payload.leads.map(function(x) { return x.source; }).filter(Boolean).join(", "),
        leadCount: payload.leads.length,
        jobNumbers: payload.leads.map(function(x) { return x.job; }).slice(0, 60).join(", "),
        emailSent: "YES", slackSent: "ATTEMPT", slackLookup: "ATTEMPT", error: "",
      });
    } catch (e) {
      exceptions.push({ title: "Email Send Error", items: [username + ": " + String(e)] });
      CC_addLog_("⚠️ Email failed: " + username, "error");
      SSCCP_logNotification_(CFG, {
        timestamp: new Date(), runId: run.runId, type: "Rep Alert (" + ctx.reportType + ")",
        route: "ERROR", stoplight: "RED", username: username, repName: "", email: toEmail, manager: "",
        sourceSheets: "", leadCount: filtered.length, jobNumbers: "",
        emailSent: "NO", slackSent: "NO", slackLookup: "NO", error: String(e),
      });
      continue;
    }

    // Check for cancellation AFTER email sent
    if (CC_isCancelled_()) {
      CC_addLog_("⚠️ Cancelled after email to " + username, "error");
      return { repsSent: repsAlerted, emailsSent: emailsSent, slacksSent: slacksSent, exceptions: exceptions, cancelled: true };
    }

    var slackRes = SSCCP_sendSlackDMToEmail_(CFG, toEmail, SSCCP_renderRepSlackText_(payload));
    if (slackRes.ok) { slacksSent++; }
    else { exceptions.push({ title: "Slack DM Failure", items: [username + " -> " + toEmail + ": " + slackRes.error] }); }

    // Check for cancellation AFTER slack sent
    if (CC_isCancelled_()) {
      CC_addLog_("⚠️ Cancelled after Slack to " + username, "error");
      return { repsSent: repsAlerted, emailsSent: emailsSent, slacksSent: slacksSent, exceptions: exceptions, cancelled: true };
    }

    repsAlerted++;
    Utilities.sleep(140);
  }

  if (skippedDedupe > 0) CC_addLog_("⭐️ Skipped (dedupe): " + skippedDedupe + " reps", "skip");
  if (skippedNoEmail > 0) CC_addLog_("⭐️ Skipped (no email): " + skippedNoEmail + " reps", "skip");

  return { repsSent: repsAlerted, emailsSent: emailsSent, slacksSent: slacksSent, exceptions: exceptions, cancelled: false };
}


/* 
 * MANAGER REPORTS WITH PROGRESS - Sends reports with live updates
 * v1.2 - Improved cancellation checking
 * ═══════════════════════════════════════════════════════════════ */

function CC_sendManagerReportsWithProgress_(CFG, run, ctx, windowInfo, grouped, roster, team, integrity) {
  var mode = SSCCP_mode_(CFG);
  var managerAgg = SSCCP_buildTeamAgg_(grouped, roster, team, integrity);
  var managers = Object.keys(managerAgg || {});
  managers.sort(function(a, b) { return (managerAgg[b].teamTotal || 0) - (managerAgg[a].teamTotal || 0); });

  var exceptions = [];
  var managersSent = 0;
  var total = managers.length;

  for (var i = 0; i < managers.length; i++) {
    // Check for cancellation at START of each iteration
    if (CC_isCancelled_()) {
      CC_addLog_("⚠️ Cancelled by user at manager " + (i + 1) + "/" + total, "error");
      return { managersSent: managersSent, exceptions: exceptions, cancelled: true };
    }
    
    var managerName = managers[i];
    var mp = managerAgg[managerName];
    
    // Update progress every manager
    var pct = Math.round(89 + (i / Math.max(total, 1)) * 7);
    CC_updateStage_(8, "active", (i + 1) + "/" + total, pct, "Manager Reports", "Sending to " + managerName + "...");
    CC_addLog_("📤 Manager " + (i + 1) + "/" + total + ": " + managerName, "info");

    var managerEmail = SSCCP_resolveManagerEmail_(roster, managerName);
    if (!managerEmail) { exceptions.push({ title: "Missing Manager Email (Sales_Roster)", items: [managerName] }); continue; }

    var toEmail = mode.routeAllToAdmin ? CFG.ADMIN_EMAIL : managerEmail;

    var payload = {
      reportType: ctx.reportType, reportTitle: ctx.title, runId: run.runId, window: windowInfo,
      manager: { name: managerName, email: managerEmail }, team: mp, mode: mode,
    };

    var subject = SSCCP_stopEmoji_("GREEN") + " Safe Ship — Team " + ctx.title + " — " + managerName + " (Run " + run.runId + ")";

    try {
      var html = SSCCP_renderManagerEmailHtml_(CFG, payload);
      MailApp.sendEmail({
        to: toEmail,
        subject: mode.routeAllToAdmin ? "[FORWARDED:" + managerEmail + "] " + subject : subject,
        htmlBody: html,
        body: SSCCP_stripHtml_(html),
      });

      SSCCP_logNotification_(CFG, {
        timestamp: new Date(), runId: run.runId, type: "Manager Report (" + ctx.reportType + ")",
        route: mode.routeAllToAdmin ? "Forward-All/Admin" : "Direct",
        stoplight: "GREEN", username: "", repName: "", email: toEmail, manager: managerName,
        sourceSheets: "", leadCount: mp.teamTotal || 0, jobNumbers: "",
        emailSent: "YES", slackSent: "ATTEMPT", slackLookup: "ATTEMPT", error: "",
      });
    } catch (e) {
      exceptions.push({ title: "Manager Email Send Error", items: [managerName + ": " + String(e)] });
      CC_addLog_("⚠️ Manager email failed: " + managerName, "error");
      continue;
    }

    // Check for cancellation AFTER email sent
    if (CC_isCancelled_()) {
      CC_addLog_("⚠️ Cancelled after email to " + managerName, "error");
      return { managersSent: managersSent, exceptions: exceptions, cancelled: true };
    }

    var slackRes = SSCCP_sendSlackDMToEmail_(CFG, toEmail, SSCCP_renderManagerSlackText_(payload));
    if (!slackRes.ok) { exceptions.push({ title: "Manager Slack DM Failure", items: [managerName + " -> " + toEmail + ": " + slackRes.error] }); }

    // Check for cancellation AFTER slack sent
    if (CC_isCancelled_()) {
      CC_addLog_("⚠️ Cancelled after Slack to " + managerName, "error");
      return { managersSent: managersSent, exceptions: exceptions, cancelled: true };
    }

    managersSent++;
    Utilities.sleep(180);
  }

  return { managersSent: managersSent, exceptions: exceptions, cancelled: false };
}


/* 
 * SIDEBAR HTML TEMPLATE
 * ═══════════════════════════════════════════════════════════════ */

/**************************************************************
 * 🎨 UNIFIED SIDEBAR THEME - All Sidebars Match Quick Sync Style
 * 
 * 
 * This file contains updated sidebar HTML functions that match
 * the Quick Sync dark theme style the user prefers.
 * 
 * SIDEBARS UPDATED:
 * 1. Today Sync (getTodaySyncSidebarHtml_)
 * 2. Smart Sync (getSmartSyncSidebarHtml_)  
 * 3. SMS Only (getSMSOnlySidebarHtml_)
 * 4. Calls Only (getCallsOnlySidebarHtml_)
 * 5. Contact Checker (getContactCheckerSidebarHtml_)
 * 6. Uncontacted Leads (SSCCP_openProgressSidebar_)
 * 
 * INSTALLATION:
 * Replace the corresponding functions in:
 * - RC_TodaySync.gs
 * - RingCentral_API (smart sync section)
 * - ContactChecker_Progress.gs
 * - Mainmenu.gs
 * 
 **************************************************************/




/* 
 * 📌 CONTACT CHECKER / UNCONTACTED LEADS SIDEBAR
 * Replace getContactCheckerSidebarHtml_() in ContactChecker_Progress.gs
 * Also replace SSCCP_openProgressSidebar_() HTML in Mainmenu.gs
 * ═══════════════════════════════════════════════════════════════ */

function getContactCheckerSidebarHtml_(config) {
  var icon = config.icon || "📌";
  var title = config.title || "Uncontacted Leads Alert";
  var description = config.description || "Leads that haven't been contacted";
  var startFunction = config.startFunction || "CC_startUncontactedLeads";
  var brand = UNIFIED_SIDEBAR_BRAND;
  
  return `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>${getUnifiedSidebarCSS_()}</style>
</head>
<body>
  <div class="header">
    <div class="brand">🚢 Safe Ship Contact Checker</div>
    <div class="title">${icon} ${title}</div>
    <div class="subtitle">${description}</div>
  </div>
  
  <div class="content">
    <div style="text-align: center;">
      <span class="mode-badge live" id="modeBadge">⏳ Loading...</span>
    </div>
    
    <div class="ring-section">
      <div class="ring-wrap">
        <svg class="ring-svg" viewBox="0 0 120 120">
          <defs>
            <linearGradient id="ringGrad" x1="0%" y1="0%" x2="100%" y2="0%">
              <stop offset="0%" stop-color="${brand.PRIMARY}"/>
              <stop offset="50%" stop-color="${brand.GOLD}"/>
              <stop offset="100%" stop-color="${brand.PRIMARY}"/>
            </linearGradient>
          </defs>
          <circle class="ring-bg" cx="60" cy="60" r="55"/>
          <circle class="ring-fill" id="ringFill" cx="60" cy="60" r="55"/>
        </svg>
        <div class="ring-center">
          <div class="ring-pct" id="ringPct">0%</div>
          <div class="ring-label">Complete</div>
        </div>
      </div>
      <div class="stage-name" id="mainStatus">Ready to start</div>
      <div class="stage-status" id="mainSubStatus">Click Start Run to begin</div>
    </div>
    
    <div class="section-title">📈 Lead Breakdown</div>
    <div class="stats-grid" id="priorityStats">
      <div class="stat-box"><div class="stat-val" id="statP0" style="color:#3b82f6">-</div><div class="stat-lbl">P0</div></div>
      <div class="stat-box"><div class="stat-val" id="statP1" style="color:#ef4444">-</div><div class="stat-lbl">P1</div></div>
      <div class="stat-box"><div class="stat-val" id="statP3" style="color:#f59e0b">-</div><div class="stat-lbl">P3</div></div>
      <div class="stat-box"><div class="stat-val" id="statP5" style="color:#8b5cf6">-</div><div class="stat-lbl">P5</div></div>
    </div>
    
    <div class="breakdown" id="repBreakdown">
      <div class="breakdown-row">
        <span class="breakdown-label">Total Reps:</span>
        <span class="breakdown-value" id="totalReps">-</span>
      </div>
      <div class="breakdown-row">
        <span class="breakdown-label">Total Leads:</span>
        <span class="breakdown-value" id="totalLeads">-</span>
      </div>
      <div class="breakdown-row">
        <span class="breakdown-label">Avg per Rep:</span>
        <span class="breakdown-value" id="avgPerRep">-</span>
      </div>
    </div>
    
    <div class="section-title">📋 Pipeline</div>
    <div class="pipeline" id="pipelineStages">
      <div class="stage" id="st1"><div class="stage-icon">🔒</div><div class="stage-info"><div class="stage-name-text">Gate Checks</div><div class="stage-detail" id="st1s">Waiting...</div></div><div class="stage-count" id="st1c"></div></div>
      <div class="stage" id="st2"><div class="stage-icon">👥</div><div class="stage-info"><div class="stage-name-text">Sales Roster</div><div class="stage-detail" id="st2s">Waiting...</div></div><div class="stage-count" id="st2c"></div></div>
      <div class="stage" id="st3"><div class="stage-icon">💔</div><div class="stage-info"><div class="stage-name-text">Team Roster</div><div class="stage-detail" id="st3s">Waiting...</div></div><div class="stage-count" id="st3c"></div></div>
      <div class="stage" id="st4"><div class="stage-icon">📞</div><div class="stage-info"><div class="stage-name-text">RingCentral</div><div class="stage-detail" id="st4s">Waiting...</div></div><div class="stage-count" id="st4c"></div></div>
      <div class="stage" id="st5"><div class="stage-icon">📝</div><div class="stage-info"><div class="stage-name-text">Scan Sources</div><div class="stage-detail" id="st5s">Waiting...</div></div><div class="stage-count" id="st5c"></div></div>
      <div class="stage" id="st6"><div class="stage-icon">🔗</div><div class="stage-info"><div class="stage-name-text">Enrich Leads</div><div class="stage-detail" id="st6s">Waiting...</div></div><div class="stage-count" id="st6c"></div></div>
      <div class="stage" id="st7"><div class="stage-icon">📧</div><div class="stage-info"><div class="stage-name-text">Send Alerts</div><div class="stage-detail" id="st7s">Waiting...</div></div><div class="stage-count" id="st7c"></div></div>
      <div class="stage" id="st8"><div class="stage-icon">📊</div><div class="stage-info"><div class="stage-name-text">Manager Reports</div><div class="stage-detail" id="st8s">Waiting...</div></div><div class="stage-count" id="st8c"></div></div>
    </div>
    
    <div class="btn-group">
      <button class="btn btn-live" id="startBtn">▶ Start Run</button>
      <button class="btn btn-cancel" id="cancelBtn">⛔ Cancel Run</button>
      <button class="btn btn-close" id="closeBtn">Close</button>
    </div>
    
    <div class="section-title">📝 Action Log</div>
    <div class="log-section" id="logSection">
      <div class="log-entry"><span class="log-time">--:--:--</span><span class="log-msg">Ready. Click "Start Run" to begin.</span></div>
    </div>
  </div>
  
  <div class="footer">Safe Ship Contact Checker Pro • v1.0</div>
  
  <script>
    var polling = null;
    var circ = 2 * Math.PI * 55;
    var ring = document.getElementById("ringFill");
    ring.style.strokeDasharray = circ;
    ring.style.strokeDashoffset = circ;
    
    google.script.run.withSuccessHandler(function(mode) {
      var badge = document.getElementById("modeBadge");
      if (mode && mode.testMode) { badge.className = "mode-badge test"; badge.textContent = "🧪 TEST MODE"; }
      else if (mode && mode.safeMode) { badge.className = "mode-badge safe"; badge.textContent = "🔒 SAFE MODE"; }
      else if (mode && mode.forwardAll) { badge.className = "mode-badge forward"; badge.textContent = "📨 FORWARD-ALL"; }
      else { badge.className = "mode-badge live"; badge.textContent = "🟢 LIVE MODE"; }
    }).withFailureHandler(function(e) {
      document.getElementById("modeBadge").className = "mode-badge live";
      document.getElementById("modeBadge").textContent = "🟢 LIVE MODE";
    }).CC_getMode();
    
    document.getElementById("startBtn").onclick = function() {
      this.style.display = "none";
      document.getElementById("cancelBtn").style.display = "block";
      document.getElementById("mainStatus").textContent = "Starting...";
      document.getElementById("mainSubStatus").textContent = "Initializing pipeline...";
      google.script.run.withSuccessHandler(function(r){}).withFailureHandler(function(e){alert("Error: "+e.message);document.getElementById("startBtn").style.display="block";document.getElementById("cancelBtn").style.display="none"}).${startFunction}();
      polling = setInterval(checkProgress, 600);
    };
    
    document.getElementById("cancelBtn").onclick = function() {
      if (confirm("Cancel the current operation?")) {
        this.disabled = true; this.textContent = "⏳ Cancelling...";
        google.script.run.CC_requestCancel();
      }
    };
    
    document.getElementById("closeBtn").onclick = function() { google.script.host.close(); };
    
    function checkProgress() {
      google.script.run.withSuccessHandler(updateUI).withFailureHandler(function(e){}).CC_getProgress();
    }
    
    function updateUI(d) {
      if (!d) return;
      var pct = d.percent || 0;
      ring.style.strokeDashoffset = circ - (pct / 100) * circ;
      document.getElementById("ringPct").textContent = pct + "%";
      document.getElementById("mainStatus").textContent = d.currentStage || "Processing...";
      document.getElementById("mainSubStatus").textContent = d.status || "";
      
      if (d.stats) {
        document.getElementById("totalReps").textContent = d.stats.reps || "-";
        document.getElementById("totalLeads").textContent = d.stats.leads || "-";
        var reps = d.stats.reps || 0, leads = d.stats.leads || 0;
        document.getElementById("avgPerRep").textContent = reps > 0 ? Math.round(leads / reps) : "-";
      }
      
      for (var i = 1; i <= 8; i++) {
        var s = d["stage" + i];
        if (s) {
          document.getElementById("st" + i).className = "stage " + (s.state || "");
          document.getElementById("st" + i + "s").textContent = s.status || "Waiting...";
          if (s.count != null && s.count > 0) {
            var countEl = document.getElementById("st" + i + "c");
            countEl.textContent = s.count.toLocaleString();
            countEl.style.display = "block";
          }
        }
      }
      
      if (d.log && d.log.length) {
        var html = "";
        d.log.forEach(function(l) {
          html += '<div class="log-entry ' + (l.type || "") + '"><span class="log-time">' + l.time + '</span><span class="log-msg">' + l.msg + '</span></div>';
        });
        document.getElementById("logSection").innerHTML = html;
        document.getElementById("logSection").scrollTop = 99999;
      }
      
      if (d.complete || d.cancelled) {
        clearInterval(polling);
        document.getElementById("startBtn").style.display = "none";
        document.getElementById("cancelBtn").style.display = "none";
        document.getElementById("closeBtn").style.display = "block";
      }
    }
  </script>
</body>
</html>`;
}

function CC_getMode() {
  try {
    var CFG = getConfig_();
    var mode = SSCCP_mode_(CFG);
    return {
      testMode: mode.testMode || false,
      safeMode: mode.safeMode || false,
      forwardAll: mode.forwardAll || false
    };
  } catch (e) {
    return { testMode: false, safeMode: false, forwardAll: false };
  }
}
/**************************************************************
 * 🚢 SAFE SHIP — MAIN MENU v3.1 (Complete + Test Mode + Sidebar UI)
 * File: MainMenu.gs
 *
 * 
 * THREE MENUS:
 *   1. 📧 Contact Checker — Run alerts with sidebar progress UI
 *   2. 📌 RingCentral API — Sync, authorize, diagnostics
 *   3. 🎯 GRANOT Alerts — Alert builder, quick alerts, reports
 *
 * v3.1 NEW FEATURES:
 *   - 🧪 Test Mode: First 5 leads to Admin only
 *   - Sidebar progress UI for all alert types
 *   - Working Cancel button during runs
 *   - LIVE confirmation when Safe Mode OFF + Forward-All OFF
 *   - Phone visibility toggle
 *   - Full GRANOT alert system with presets
 *   - All safe features and diagnostics
 **************************************************************/

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  let CFG;
  try {
    CFG = getConfig_();
  } catch (err) {
    ui.createMenu("🚀 Safe Ship Contact Checker")
      .addItem("❌ Config/Menu Error (Show Details)", "SSCCP_MENU_showMenuError")
      .addToUi();
    return;
  }

  const safeMode = !!CFG.SAFE_MODE;
  const forwardAll = !!CFG.FORWARD_ALL;
  const includePhone = SSCCP_MENU_getShowPhones_();
  
  // v3.1: Get Test Mode status
  const testMode = SSCCP_MENU_getTestMode_();

  // 
  // MENU 1: 📧 CONTACT CHECKER (with Sidebar Progress UI)
  // 
  ui.createMenu("📧 Contact Checker")
    
    // ▶ Run Alerts (TOP LEVEL SUBMENU)
    .addSubMenu(ui.createMenu("▶ Run Alerts")
      .addItem("📊 Alert Dashboard (NEW)", "SSCCP_openAlertDashboard")
      .addItem("📉 Uncontacted Leads (SMS + Call/VM)", "SSCCP_MENU_runUncontacted")
      .addItem("💰 Quoted Follow-Up (CONTACTED LEADS)", "SSCCP_MENU_runQuotedFollowUp")
      .addItem("📞 Priority 1 Call/VM", "SSCCP_MENU_runPriority1")
      .addItem("⚡ Same Day Transfers", "SSCCP_MENU_runSameDayTransfers"))
    
    .addSeparator()
    
    // 🧾 Pre-Flight (standalone)
    .addItem("🧾 Pre-Flight Manifest (Preview)", "SSCCP_previewManifest")
    
    .addSeparator()
    
    // 📤 Send Reports Only
    .addSubMenu(ui.createMenu("📤 Send Reports Only")
      .addItem("📧 Rep Reports Only", "SSCCP_MENU_sendRepReportsOnly")
      .addItem("👥 Manager Reports Only", "SSCCP_MENU_sendManagerReportsOnly")
      .addItem("📊 Admin Dashboard Only", "SSCCP_MENU_sendAdminDashboardOnly"))
    
    .addSeparator()
    
    // 🛡️ Safe Features — v3.1: Added Test Mode at TOP
    .addSubMenu(ui.createMenu("🛡️ Safe Features")
      .addItem("🧪 Toggle Test Mode (" + (testMode ? "ON - 5 leads to Admin" : "OFF") + ")", "SSCCP_MENU_toggleTestMode")
      .addSeparator()
      .addItem("🔒 Toggle Safe Mode (" + (safeMode ? "ON" : "OFF") + ")", "SSCCP_toggleSafeMode")
      .addItem("📨 Toggle Forward-All (" + (forwardAll ? "ON" : "OFF") + ")", "SSCCP_toggleForwardAll")
      .addSeparator()
      .addItem("📞 Toggle Phone # (" + (includePhone ? "Show" : "Hidden") + ")", "SSCCP_MENU_togglePhoneVisibility"))
    
    // ⚙️ Configuration
    .addSubMenu(ui.createMenu("⚙️ Configuration")
      .addItem("📌 Show Current Status", "SSCCP_showStatus")
      .addItem("✅ Run Gate Checks", "SSCCP_runGateChecksInteractive")
      .addSeparator()
      .addItem("📧 Test Admin Email", "SSCCP_testAdminEmail")
      .addItem("💬 Test Slack DM", "SSCCP_testSlackDMAdmin")
      .addItem("📧 Test Column Reading", "SSCCP_testColumnReading"))
    
    // 📜 Run History
    .addSubMenu(ui.createMenu("📜 Run History")
      .addItem("🔄 Open Notification_Log", "SSCCP_openNotificationLog")
      .addItem("🔄 Open Run_Log", "SSCCP_openRunLog")
      .addSeparator()
      .addItem("🧹 Clear Notification_Log", "SSCCP_clearNotificationLogConfirm"))
    
    .addSeparator()
    
    // 🧪 Audits & Health
    .addSubMenu(ui.createMenu("🧪 Audits & Health")
      .addItem("🔍 Data Validator (Full)", "DV_runFullValidation")
      .addItem("📊 RC Index Stats", "RC_testIndexStats")
      .addItem("🔍 Test RC Enrichment", "RC_testEnrichment"))
    
    .addToUi();

  // 
  // MENU 2: 📌 RINGCENTRAL API
  // 
  ui.createMenu("📌 RingCentral API")
    
    // Quick Actions
    .addItem("⚡ Quick Sync (Incremental)", "RC_API_quickSync")
    .addItem("📤 Today Sync (Full Export)", "RC_API_syncTodayOnly")
    .addItem("🔄 Smart Sync (All)", "RC_API_smartSyncAll")
    .addSeparator()
    
    // Manual Sync submenu
    .addSubMenu(ui.createMenu("📊 Manual Sync")
      .addItem("📱 Today's SMS Only", "RC_API_syncTodaySMSOnly")
      .addItem("📞 Today's Calls Only", "RC_API_syncTodayCallsOnly")
      .addSeparator()
      .addItem("📱 Force Sync Yesterday SMS", "RC_API_forceSyncYesterdaySMS")
      .addItem("📞 Force Sync Yesterday Calls", "RC_API_forceSyncYesterdayCalls"))
    
    .addSeparator()
    
    // Authorizationh
    .addItem("🔐 Authorize RingCentral", "RC_API_authorize")
    .addItem("📌 Test Connection", "RC_API_testConnection")
    .addItem("🚪 Logout / Reset", "RC_API_logout")
    .addSeparator()
    
    // Status
    .addItem("📋 Show Sync Status", "RC_API_showSyncStatus")
    .addItem("🔄 Reset Sync Cache", "RC_API_resetSyncCache")
    .addItem("⚙️ Check Setup", "RC_API_checkSetup")
    .addSeparator()
    
    // Diagnostics submenu
    .addSubMenu(ui.createMenu("🔧 Diagnostics")
      .addItem("📊 Test Index Stats", "RC_testIndexStats")
      .addItem("🔍 Test RC Enrichment", "RC_testEnrichment")
      .addItem("🔍 Debug SMS Matching", "RC_debugSMSMatching")
      .addItem("📱 View Contact History", "RC_viewSelectedContactHistory")
      .addItem("🔧 Test Phone Matching", "RC_testPhoneMatching")
      .addItem("⏱️ Debug Duration Parsing", "RC_debugDurationParsing")
      .addItem("🔢 Verify Outbound Counts", "RC_verifyOutboundCounts")
      .addItem("🔢 Quick Count Check", "RC_quickOutboundCheck"))
    
    .addToUi();

  // 
  // MENU 3: 🎯 GRANOT ALERTS
  // 
  ui.createMenu("🎯 GRANOT Alerts")
    
    // Main Alert Builder
    .addItem("🎯 Open Alert Builder", "GRANOT_openAlertBuilder")
    .addSeparator()
    
    // Quick Alerts submenu
    .addSubMenu(ui.createMenu("⚡ Quick Alerts")
      .addItem("🔴 Not Worked Leads", "GRANOT_quickAlert_NotWorked")
      .addItem("🔥 Hot Leads (Move ≤7 days)", "GRANOT_quickAlert_Hot")
      .addItem("💰 High Value Leads ($5K+)", "GRANOT_quickAlert_HighValue")
      .addItem("⏰ Stale Leads (3+ days, 0 contact)", "GRANOT_quickAlert_Stale")
      .addItem("🔵 Zero Calls Leads", "GRANOT_quickAlert_ZeroCalls"))
    
    // Reports submenu
    .addSubMenu(ui.createMenu("📊 Reports")
      .addItem("👥 Manager Summary", "GRANOT_report_ManagerSummary")
      .addItem("🏢 Team Reports", "GRANOT_report_TeamReports")
      .addItem("🏥 Admin Health Report", "GRANOT_report_AdminHealth")
      .addSeparator()
      .addItem("📤 Export Current View", "GRANOT_exportCurrentView"))
    
    // Tools submenu
    .addSubMenu(ui.createMenu("🔧 Tools")
      .addItem("📊 Show GRANOT Stats", "GRANOT_showStats")
      .addItem("🔍 Test Lead Lookup", "GRANOT_testLookup")
      .addItem("💾 Manage Filter Presets", "GRANOT_managePresets"))
    
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

/* 
 * v3.1: TEST MODE FUNCTIONS
 * ═══════════════════════════════════════════════════════════════ */

function SSCCP_MENU_getTestMode_() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty("SSCCP_TEST_MODE") === "true";
}

function SSCCP_MENU_toggleTestMode() {
  var props = PropertiesService.getScriptProperties();
  var cur = props.getProperty("SSCCP_TEST_MODE") === "true";
  props.setProperty("SSCCP_TEST_MODE", cur ? "false" : "true");
  
  var newState = !cur;
  SpreadsheetApp.getActiveSpreadsheet().toast(
    newState 
      ? "🧪 Test Mode ON: First 5 leads will be sent to Admin only (no rep/manager emails)" 
      : "Test Mode OFF: Full run will send to all reps and managers",
    "Safe Ship", 
    8
  );
}

/* 
 * v3.1: CANCELLATION SUPPORT
 * ═══════════════════════════════════════════════════════════════ */

function SSCCP_MENU_cancelCurrentRun() {
  var cache = CacheService.getScriptCache();
  cache.put("SSCCP_CANCELLED", "true", 300);
  SpreadsheetApp.getActiveSpreadsheet().toast("⛔ Cancel requested... will stop after current email", "Cancelling", 5);
}

/* 
 * CONFIRM WRAPPERS — With Sidebar Progress UI
 * Only confirms in TRUE live: Safe Mode OFF + Forward-All OFF + Test Mode OFF
 * ═══════════════════════════════════════════════════════════════ */

function SSCCP_MENU_runUncontacted() {
  if (!SSCCP_confirmIfLive_("Uncontacted Leads Alert (SMS TRACKER + CALL & VOICEMAIL TRACKER)")) return;
  
  // v3.1: Always use sidebar UI
  SSCCP_openProgressSidebar_("UNCONTACTED", "Uncontacted Leads Alert");
}

function SSCCP_MENU_runQuotedFollowUp() {
  if (!SSCCP_confirmIfLive_("Quoted Follow-Up Alert (CONTACTED LEADS)")) return;
  
  SSCCP_openProgressSidebar_("QUOTED_FOLLOWUP", "Quoted Follow-Up Alert");
}

function SSCCP_MENU_runPriority1() {
  if (!SSCCP_confirmIfLive_("Priority 1 Call/VM Alert (PRIORITY 1 CALL & VOICEMAIL TRACKER)")) return;
  
  SSCCP_openProgressSidebar_("PRIORITY1_CALLVM", "Priority 1 Call/VM Alert");
}

function SSCCP_MENU_runSameDayTransfers() {
  if (!SSCCP_confirmIfLive_("Same Day Transfers (SAME DAY TRANSFERS)")) return;
  
  SSCCP_openProgressSidebar_("SAME_DAY_TRANSFERS", "Same Day Transfers Alert");
}

/* 
 * v3.1: SIDEBAR PROGRESS UI
 * ═══════════════════════════════════════════════════════════════ */

/* 
 * 🎨 MAINMENU SIDEBAR - DARK THEME
 * 
 * 
 * REPLACE the SSCCP_openProgressSidebar_ function in Mainmenu.gs
 * (approximately lines 252-442)
 * 
 * This gives the Uncontacted Leads Alert sidebar the same dark
 * theme as Quick Sync.
 * 
 * ═══════════════════════════════════════════════════════════════ */

function SSCCP_openProgressSidebar_(reportType, title) {
  var testMode = SSCCP_MENU_getTestMode_();
  var CFG = getConfig_();
  
  // Get icon based on report type
  var icon = "📌";
  var desc = "Processing leads";
  if (reportType === "QUOTED_FOLLOWUP") { icon = "📋"; desc = "Quoted leads needing follow-up"; }
  else if (reportType === "PRIORITY1_CALLVM") { icon = "🔥"; desc = "High priority leads requiring action"; }
  else if (reportType === "SAME_DAY_TRANSFERS") { icon = "🔄"; desc = "Transfers scheduled for today"; }
  else if (reportType === "UNCONTACTED") { icon = "📌"; desc = "Leads that haven\'t been contacted"; }
  
  // Mode banner HTML
  var modeBannerHtml = '';
  if (testMode) {
    modeBannerHtml = '<div class="mode-box test">🧪 TEST MODE: First 5 leads → Admin only</div>';
  } else if (CFG.SAFE_MODE) {
    modeBannerHtml = '<div class="mode-box safe">🔒 SAFE MODE: All emails → Admin</div>';
  } else if (CFG.FORWARD_ALL) {
    modeBannerHtml = '<div class="mode-box forward">📨 FORWARD-ALL: Copies → Admin</div>';
  } else {
    modeBannerHtml = '<div class="mode-box live">🟢 LIVE MODE: Direct to reps/managers</div>';
  }

  var html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { 
          font-family: -apple-system, BlinkMacSystemFont, sans-serif; 
          background: #0f172a; 
          color: #e2e8f0; 
          min-height: 100vh; 
          padding: 20px; 
        }
        
        .header { text-align: center; margin-bottom: 16px; }
        .header h1 { font-size: 20px; color: #D4AF37; margin-bottom: 4px; }
        .header p { font-size: 11px; color: #64748b; }
        .header .run-id { font-size: 10px; color: #475569; margin-top: 4px; }
        
        .mode-box {
          border-radius: 8px;
          padding: 10px 12px;
          margin-bottom: 12px;
          font-size: 11px;
          border-left: 3px solid;
        }
        .mode-box.live { background: rgba(16,185,129,0.15); border-left-color: #10b981; color: #10b981; }
        .mode-box.test { background: rgba(245,158,11,0.15); border-left-color: #f59e0b; color: #f59e0b; }
        .mode-box.safe { background: rgba(59,130,246,0.15); border-left-color: #3b82f6; color: #3b82f6; }
        .mode-box.forward { background: rgba(245,158,11,0.15); border-left-color: #f59e0b; color: #f59e0b; }
        
        .progress-ring { width: 120px; height: 120px; margin: 0 auto 16px; position: relative; }
        .progress-ring svg { transform: rotate(-90deg); }
        .progress-ring .bg { fill: none; stroke: #334155; stroke-width: 8; }
        .progress-ring .fill { 
          fill: none; 
          stroke: url(#grad); 
          stroke-width: 8; 
          stroke-linecap: round; 
          stroke-dasharray: 314; 
          stroke-dashoffset: 314; 
          transition: stroke-dashoffset 0.3s; 
        }
        .progress-ring .center { 
          position: absolute; 
          top: 50%; 
          left: 50%; 
          transform: translate(-50%, -50%); 
          text-align: center; 
        }
        .progress-ring .pct { font-size: 28px; font-weight: 800; color: #fff; }
        .progress-ring .label { font-size: 10px; color: #64748b; }
        
        .stage-text { text-align: center; margin-bottom: 16px; }
        .stage-text .title { font-size: 14px; font-weight: 600; color: #fff; }
        .stage-text .sub { font-size: 11px; color: #64748b; }
        
        .stats { 
          display: grid; 
          grid-template-columns: 1fr 1fr; 
          gap: 10px; 
          margin-bottom: 16px; 
        }
        .stat { 
          background: #1e293b; 
          border-radius: 10px; 
          padding: 14px; 
          text-align: center; 
        }
        .stat .val { font-size: 26px; font-weight: 800; color: #fff; }
        .stat .lbl { font-size: 9px; color: #64748b; text-transform: uppercase; margin-top: 2px; }
        
        .log { 
          background: #1e293b; 
          border-radius: 12px; 
          padding: 12px; 
          max-height: 180px; 
          overflow-y: auto; 
          margin-bottom: 16px; 
        }
        .log-title { 
          font-size: 10px; 
          color: #D4AF37; 
          font-weight: 600; 
          margin-bottom: 10px; 
          text-transform: uppercase; 
          letter-spacing: 1px; 
        }
        .log-entry { 
          font-size: 10px; 
          padding: 5px 0; 
          border-bottom: 1px solid rgba(255,255,255,0.05); 
          display: flex; 
          gap: 8px; 
        }
        .log-entry:last-child { border-bottom: none; }
        .log-entry .time { color: #475569; font-family: monospace; min-width: 60px; }
        .log-entry .msg { color: #94a3b8; }
        .log-entry.success .msg { color: #10b981; }
        .log-entry.error .msg { color: #ef4444; }
        .log-entry.warn .msg { color: #f59e0b; }
        .log-entry.info .msg { color: #3b82f6; }
        
        .btn { 
          width: 100%; 
          padding: 14px; 
          border: none; 
          border-radius: 10px; 
          font-size: 14px; 
          font-weight: 700; 
          cursor: pointer; 
          margin-bottom: 8px; 
        }
        .btn-start { background: linear-gradient(135deg, #D4AF37, #F4D03F); color: #000; }
        .btn-start:hover { opacity: 0.9; }
        .btn-start:disabled { background: #334155; color: #64748b; cursor: not-allowed; }
        .btn-cancel { 
          background: rgba(239, 68, 68, 0.2); 
          color: #ef4444; 
          border: 1px solid rgba(239, 68, 68, 0.3); 
        }
        .btn-cancel:hover { background: rgba(239, 68, 68, 0.3); }
        .btn-cancel:disabled { opacity: 0.5; cursor: not-allowed; }
        
        .footer { text-align: center; font-size: 8px; color: #475569; margin-top: 16px; }
      </style>
    </head>
    <body>
      <div class="header">
        <h1>${icon} ${title}</h1>
        <p>${desc}</p>
        <div class="run-id">Run ID: <span id="runId">—</span></div>
      </div>
      
      ${modeBannerHtml}
      
      <div class="progress-ring">
        <svg viewBox="0 0 120 120" width="120" height="120">
          <defs>
            <linearGradient id="grad" x1="0%" y1="0%" x2="100%" y2="0%">
              <stop offset="0%" stop-color="#8B1538"/>
              <stop offset="100%" stop-color="#D4AF37"/>
            </linearGradient>
          </defs>
          <circle class="bg" cx="60" cy="60" r="50"/>
          <circle class="fill" id="ringFill" cx="60" cy="60" r="50"/>
        </svg>
        <div class="center">
          <div class="pct" id="pct">0%</div>
          <div class="label">Complete</div>
        </div>
      </div>
      
      <div class="stage-text">
        <div class="title" id="currentStep">Ready to start</div>
        <div class="sub" id="stepStatus">Click Start Run to begin</div>
      </div>
      
      <div class="stats">
        <div class="stat">
          <div class="val" id="leadsCount">0</div>
          <div class="lbl">Leads</div>
        </div>
        <div class="stat">
          <div class="val" id="emailsSent">0</div>
          <div class="lbl">Emails Sent</div>
        </div>
      </div>
      
      <button class="btn btn-start" id="btnRun" onclick="startRun()">
        ${icon} Start${testMode ? ' Test' : ''} Run
      </button>
      
      <button class="btn btn-cancel" id="btnCancel" onclick="cancelRun()" disabled>
        ⛔ Cancel Run
      </button>
      
      <div class="log" id="logContainer">
        <div class="log-title">📜 Activity Log</div>
        <div id="logList">
          <div class="log-entry info">
            <span class="time">--:--:--</span>
            <span class="msg">Ready. Click "Start Run" to begin.</span>
          </div>
        </div>
      </div>
      
      <div class="footer">SAFE SHIP CONTACT CHECKER PRO • v3.2</div>
      
      <script>
        var reportType = '${reportType}';
        var isRunning = false;
        var pollInterval = null;
        var circ = 2 * Math.PI * 50;
        var ring = document.getElementById('ringFill');
        ring.style.strokeDasharray = circ;
        ring.style.strokeDashoffset = circ;
        
        function log(msg, type) {
          var list = document.getElementById('logList');
          var time = new Date().toLocaleTimeString('en-US', {hour12: false});
          var entry = document.createElement('div');
          entry.className = 'log-entry ' + (type || '');
          entry.innerHTML = '<span class="time">' + time + '</span><span class="msg">' + msg + '</span>';
          list.appendChild(entry);
          document.getElementById('logContainer').scrollTop = document.getElementById('logContainer').scrollHeight;
        }
        
        function updateUI(data) {
          if (data.runId) document.getElementById('runId').textContent = data.runId;
          if (data.step) {
            document.getElementById('currentStep').textContent = data.step;
            document.getElementById('stepStatus').textContent = data.substep || '';
          }
          if (data.progress !== undefined) {
            var pct = Math.round(data.progress);
            document.getElementById('pct').textContent = pct + '%';
            ring.style.strokeDashoffset = circ - (pct / 100) * circ;
          }
          if (data.leads !== undefined) document.getElementById('leadsCount').textContent = data.leads;
          if (data.emails !== undefined) document.getElementById('emailsSent').textContent = data.emails;
        }
        
        function startRun() {
          isRunning = true;
          document.getElementById('btnRun').disabled = true;
          document.getElementById('btnCancel').disabled = false;
          log('🚀 Starting run...', 'info');
          updateUI({ step: 'Starting...', substep: 'Initializing pipeline...' });
          
          google.script.run
            .withSuccessHandler(function(result) {
              log('✅ Run initiated: ' + result.runId, 'success');
              updateUI({ runId: result.runId });
              startPolling();
            })
            .withFailureHandler(function(err) {
              log('❌ Failed to start: ' + err, 'error');
              isRunning = false;
              document.getElementById('btnRun').disabled = false;
              document.getElementById('btnCancel').disabled = true;
              updateUI({ step: 'Error', substep: err });
            })
            .SSCCP_SIDEBAR_startRun(reportType);
        }
        
        function startPolling() {
          pollInterval = setInterval(function() {
            google.script.run
              .withSuccessHandler(function(status) {
                updateUI(status);
                if (status.log) log(status.log, status.logType || 'info');
                
                if (status.complete || status.cancelled || status.error) {
                  stopPolling();
                  isRunning = false;
                  document.getElementById('btnRun').disabled = false;
                  document.getElementById('btnCancel').disabled = true;
                  
                  if (status.cancelled) {
                    log('⛔ Run cancelled by user', 'warn');
                    updateUI({ step: 'Cancelled', substep: 'Operation stopped' });
                  } else if (status.error) {
                    log('❌ Error: ' + status.error, 'error');
                    updateUI({ step: 'Error', substep: status.error });
                  } else {
                    log('🎉 Run completed successfully!', 'success');
                    updateUI({ step: 'Complete!', substep: 'All tasks finished', progress: 100 });
                  }
                }
              })
              .withFailureHandler(function(err) {
                log('⚠️ Poll error: ' + err, 'warn');
              })
              .SSCCP_SIDEBAR_getStatus();
          }, 1500);
        }
        
        function stopPolling() {
          if (pollInterval) {
            clearInterval(pollInterval);
            pollInterval = null;
          }
        }
        
        function cancelRun() {
          log('⏳ Requesting cancel...', 'warn');
          document.getElementById('btnCancel').disabled = true;
          google.script.run
            .withSuccessHandler(function() {
              log('📤 Cancel signal sent - will stop after current email', 'warn');
            })
            .SSCCP_SIDEBAR_cancelRun();
        }
      </script>
    </body>
    </html>
  `).setTitle(icon + " " + title).setWidth(380);
  
  SpreadsheetApp.getUi().showSidebar(html);
}

/* 
 * v3.1: SIDEBAR BACKEND FUNCTIONS
 * ═══════════════════════════════════════════════════════════════ */

function SSCCP_SIDEBAR_startRun(reportType) {
  var cache = CacheService.getScriptCache();
  
  // Clear any previous state
  cache.remove("SSCCP_CANCELLED");
  cache.remove("SSCCP_RUN_STATUS");
  
  var runId = reportType + "_" + Utilities.formatDate(new Date(), "America/New_York", "yyyyMMdd-HHmmss") + "_" + Math.random().toString(36).slice(2, 6).toUpperCase();
  
  // Store initial status
  cache.put("SSCCP_RUN_STATUS", JSON.stringify({
    runId: runId,
    reportType: reportType,
    step: "Initializing...",
    progress: 0,
    leads: 0,
    emails: 0,
    complete: false,
    cancelled: false,
    error: null,
    log: "Run started",
    logType: "info"
  }), 600);
  
  // Start the actual run in background (will update cache as it progresses)
  SSCCP_SIDEBAR_executeRun_(reportType, runId);
  
  return { runId: runId };
}

function SSCCP_SIDEBAR_getStatus() {
  var cache = CacheService.getScriptCache();
  var statusJson = cache.get("SSCCP_RUN_STATUS");
  
  if (!statusJson) {
    return { step: "No run in progress", complete: true };
  }
  
  return JSON.parse(statusJson);
}

function SSCCP_SIDEBAR_cancelRun() {
  var cache = CacheService.getScriptCache();
  cache.put("SSCCP_CANCELLED", "true", 300);
  
  // Update status
  var statusJson = cache.get("SSCCP_RUN_STATUS");
  if (statusJson) {
    var status = JSON.parse(statusJson);
    status.log = "Cancel requested...";
    status.logType = "warn";
    cache.put("SSCCP_RUN_STATUS", JSON.stringify(status), 600);
  }
}

function SSCCP_SIDEBAR_updateStatus_(updates) {
  var cache = CacheService.getScriptCache();
  var statusJson = cache.get("SSCCP_RUN_STATUS");
  
  if (statusJson) {
    var status = JSON.parse(statusJson);
    for (var key in updates) {
      status[key] = updates[key];
    }
    cache.put("SSCCP_RUN_STATUS", JSON.stringify(status), 600);
  }
}

function SSCCP_SIDEBAR_isCancelled_() {
  var cache = CacheService.getScriptCache();
  return cache.get("SSCCP_CANCELLED") === "true";
}

function SSCCP_SIDEBAR_executeRun_(reportType, runId) {
  var CFG = getConfig_();
  var testMode = SSCCP_MENU_getTestMode_();
  
  try {
    // Step 1: Load rosters
    SSCCP_SIDEBAR_updateStatus_({ step: "Loading Sales_Roster...", progress: 5, log: "Loading rosters", logType: "info" });
    var roster = SSCCP_buildSalesRosterIndex_(CFG);
    
    if (SSCCP_SIDEBAR_isCancelled_()) { SSCCP_SIDEBAR_updateStatus_({ cancelled: true }); return; }
    
    SSCCP_SIDEBAR_updateStatus_({ step: "Loading Team_Roster...", progress: 10 });
    var team = SSCCP_buildTeamRosterIndex_(CFG);
    
    if (SSCCP_SIDEBAR_isCancelled_()) { SSCCP_SIDEBAR_updateStatus_({ cancelled: true }); return; }
    
    // Step 2: Load RC data
    SSCCP_SIDEBAR_updateStatus_({ step: "Loading RingCentral data...", progress: 15, log: "Loading RC index", logType: "info" });
    var rcIndex = {};
    try {
      if (typeof RC_buildLookupIndex_ === "function") {
        rcIndex = RC_buildLookupIndex_();
      }
    } catch (e) {
      SSCCP_SIDEBAR_updateStatus_({ log: "RC data unavailable (continuing)", logType: "warn" });
    }
    
    if (SSCCP_SIDEBAR_isCancelled_()) { SSCCP_SIDEBAR_updateStatus_({ cancelled: true }); return; }
    
    // Step 3: Scan sources
    SSCCP_SIDEBAR_updateStatus_({ step: "Scanning lead sources...", progress: 25, log: "Scanning tracker sheets", logType: "info" });
    
    var ctx = { reportType: reportType, title: reportType, sources: SSCCP_getSourcesForType_(reportType) };
    var windowInfo = SSCCP_computeWindow_(CFG, new Date());
    var scan = SSCCP_scanSources_(CFG, ctx, windowInfo, roster, function(){});
    
    if (!scan.ok) {
      SSCCP_SIDEBAR_updateStatus_({ error: scan.summary, complete: true, log: "Scan failed: " + scan.summary, logType: "error" });
      return;
    }
    
    SSCCP_SIDEBAR_updateStatus_({ leads: scan.leads.length, log: "Found " + scan.leads.length + " leads", logType: "success" });
    
    if (SSCCP_SIDEBAR_isCancelled_()) { SSCCP_SIDEBAR_updateStatus_({ cancelled: true }); return; }
    
    // Step 4: Enrich with RC data
    SSCCP_SIDEBAR_updateStatus_({ step: "Enriching with RC data...", progress: 35 });
    try {
      if (typeof RC_enrichLeads_ === "function") {
        RC_enrichLeads_(scan.leads, rcIndex);
      }
    } catch (e) {
      SSCCP_SIDEBAR_updateStatus_({ log: "RC enrichment skipped", logType: "warn" });
    }
    
    // Step 5: Group by rep
    SSCCP_SIDEBAR_updateStatus_({ step: "Grouping by rep...", progress: 40 });
    var grouped = SSCCP_groupByRep_(scan.leads);
    var integrity = SSCCP_computeIntegrity_(team, grouped);
    var reps = Object.keys(grouped);
    
    SSCCP_SIDEBAR_updateStatus_({ log: "Grouped into " + reps.length + " reps", logType: "info" });
    
    // Step 6: Send rep alerts
    SSCCP_SIDEBAR_updateStatus_({ step: "Sending rep alerts...", progress: 45 });
    
    var mode = SSCCP_mode_(CFG);
    var emailsSent = 0;
    var testLeadsProcessed = 0;
    var testLimit = 5;
    
    for (var i = 0; i < reps.length; i++) {
      if (SSCCP_SIDEBAR_isCancelled_()) { 
        SSCCP_SIDEBAR_updateStatus_({ cancelled: true, log: "Cancelled after " + emailsSent + " emails", logType: "warn" }); 
        return; 
      }
      
      // Test mode limit
      if (testMode && testLeadsProcessed >= testLimit) {
        SSCCP_SIDEBAR_updateStatus_({ log: "Test mode limit reached (" + testLimit + " leads)", logType: "info" });
        break;
      }
      
      var username = reps[i];
      var leads = grouped[username] || [];
      
      var profile = roster.byUsername[username];
      if (!profile || !profile.email) continue;
      
      // In test mode, limit leads and send to admin
      if (testMode) {
        leads = leads.slice(0, testLimit - testLeadsProcessed);
        testLeadsProcessed += leads.length;
      }
      
      var toEmail = testMode ? CFG.ADMIN_EMAIL : (mode.routeAllToAdmin ? CFG.ADMIN_EMAIL : profile.email);
      
      var progress = 45 + Math.round((i / reps.length) * 40);
      SSCCP_SIDEBAR_updateStatus_({ 
        step: "Emailing: " + (profile.repName || username), 
        progress: progress,
        log: "Sending to " + (testMode ? "Admin (test)" : profile.repName || username),
        logType: "info"
      });
      
      try {
        // Build and send email
        var payload = {
          reportType: reportType,
          reportTitle: ctx.title,
          runId: runId,
          window: windowInfo,
          rep: { username: username, name: profile.repName || username, email: profile.email },
          leads: leads,
          mode: mode,
          showPhone: true,
          isTestMode: testMode
        };
        
        var html = SSCCP_renderRepEmailHtml_PREMIUM_(CFG, payload);
        var subject = (testMode ? "🧪 [TEST] " : "") + SSCCP_stopEmoji_("GREEN") + " Safe Ship — " + ctx.title + " — " + payload.rep.name;
        
        MailApp.sendEmail({
          to: toEmail,
          subject: subject,
          htmlBody: html,
          body: SSCCP_stripHtml_(html)
        });
        
        emailsSent++;
        SSCCP_SIDEBAR_updateStatus_({ emails: emailsSent });
        
        // Skip Slack in test mode
        if (!testMode) {
          SSCCP_sendSlackDMToEmail_(CFG, toEmail, SSCCP_renderRepSlackText_(payload));
        }
        
        Utilities.sleep(150);
        
      } catch (e) {
        SSCCP_SIDEBAR_updateStatus_({ log: "Email error for " + username + ": " + e, logType: "error" });
      }
    }
    
    // Step 7: Manager reports (skip in test mode)
    if (!testMode && !SSCCP_SIDEBAR_isCancelled_()) {
      SSCCP_SIDEBAR_updateStatus_({ step: "Sending manager reports...", progress: 90, log: "Starting manager reports", logType: "info" });
      
      var managerAgg = SSCCP_buildTeamAgg_(grouped, roster, team, integrity);
      var managers = Object.keys(managerAgg);
      
      for (var m = 0; m < managers.length; m++) {
        if (SSCCP_SIDEBAR_isCancelled_()) { 
          SSCCP_SIDEBAR_updateStatus_({ cancelled: true }); 
          return; 
        }
        
        var managerName = managers[m];
        var managerEmail = SSCCP_resolveManagerEmail_(roster, managerName);
        if (!managerEmail) continue;
        
        try {
          var mgrPayload = {
            reportType: reportType,
            reportTitle: ctx.title,
            runId: runId,
            window: windowInfo,
            manager: { name: managerName, email: managerEmail },
            team: managerAgg[managerName],
            mode: mode
          };
          
          var mgrHtml = SSCCP_renderManagerEmailHtml_(CFG, mgrPayload);
          var mgrSubject = SSCCP_stopEmoji_("GREEN") + " Safe Ship — Team " + ctx.title + " — " + managerName;
          var mgrToEmail = mode.routeAllToAdmin ? CFG.ADMIN_EMAIL : managerEmail;
          
          MailApp.sendEmail({
            to: mgrToEmail,
            subject: mgrSubject,
            htmlBody: mgrHtml,
            body: SSCCP_stripHtml_(mgrHtml)
          });
          
          emailsSent++;
          SSCCP_SIDEBAR_updateStatus_({ emails: emailsSent, log: "Manager: " + managerName, logType: "info" });
          
          SSCCP_sendSlackDMToEmail_(CFG, mgrToEmail, SSCCP_renderManagerSlackText_(mgrPayload));
          Utilities.sleep(180);
          
        } catch (e) {
          SSCCP_SIDEBAR_updateStatus_({ log: "Manager error: " + e, logType: "error" });
        }
      }
    }
    
    // Complete
    SSCCP_SIDEBAR_updateStatus_({ 
      step: testMode ? "Test run complete!" : "Run complete!",
      progress: 100, 
      complete: true,
      log: "✅ Finished: " + emailsSent + " emails sent",
      logType: "success"
    });
    
  } catch (e) {
    SSCCP_SIDEBAR_updateStatus_({ 
      error: String(e),
      complete: true,
      log: "Fatal error: " + e,
      logType: "error"
    });
  }
}

function SSCCP_getSourcesForType_(reportType) {
  switch (reportType) {
    case "UNCONTACTED": return ["SMS", "CALL"];
    case "QUOTED_FOLLOWUP": return ["CONTACTED"];
    case "PRIORITY1_CALLVM": return ["P1CALL"];
    case "SAME_DAY_TRANSFERS": return ["TRANSFERS"];
    default: return [];
  }
}

/* 
 * SEND REPORTS ONLY — Skip certain pipeline steps
 * ═══════════════════════════════════════════════════════════════ */

function SSCCP_MENU_sendRepReportsOnly() {
  if (!SSCCP_confirmIfLive_("Send Rep Reports Only (No Manager Reports)")) return;
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Running Rep Reports Only...", "Contact Checker", 5);
  
  // If you have a dedicated function for this, call it
  // Otherwise, we can add this functionality
  if (typeof SSCCP_runRepReportsOnly_ === "function") {
    SSCCP_runRepReportsOnly_();
  } else {
    SpreadsheetApp.getUi().alert("Rep Reports Only function not yet implemented.\n\nUse the full Run Alerts for now.");
  }
}

function SSCCP_MENU_sendManagerReportsOnly() {
  if (!SSCCP_confirmIfLive_("Send Manager Reports Only (No Rep Alerts)")) return;
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Running Manager Reports Only...", "Contact Checker", 5);
  
  if (typeof SSCCP_runManagerReportsOnly_ === "function") {
    SSCCP_runManagerReportsOnly_();
  } else {
    SpreadsheetApp.getUi().alert("Manager Reports Only function not yet implemented.\n\nUse the full Run Alerts for now.");
  }
}

function SSCCP_MENU_sendAdminDashboardOnly() {
  SpreadsheetApp.getActiveSpreadsheet().toast("Running Admin Dashboard Only...", "Contact Checker", 5);
  
  if (typeof SSCCP_runAdminDashboardOnly_ === "function") {
    SSCCP_runAdminDashboardOnly_();
  } else {
    SpreadsheetApp.getUi().alert("Admin Dashboard Only function not yet implemented.\n\nUse the full Run Alerts for now.");
  }
}

/* 
 * LIVE MODE CONFIRMATION
 * ═══════════════════════════════════════════════════════════════ */

function SSCCP_confirmIfLive_(actionName) {
  const CFG = getConfig_();
  
  // v3.1: No confirmation needed in test mode
  if (SSCCP_MENU_getTestMode_()) return true;
  
  // If Safe Mode or Forward-All is ON, no confirmation needed
  if (CFG.SAFE_MODE || CFG.FORWARD_ALL) return true;

  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "⚠️ LIVE MODE CONFIRMATION",
    "You are about to run:\n\n" +
      actionName +
      "\n\n" +
      "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
      "🔴 Safe Mode is OFF\n" +
      "🔴 Forward-All is OFF\n" +
      "🔴 Test Mode is OFF\n" +
      "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
      "This will send LIVE alerts to reps and managers.\n\n" +
      "Proceed?",
    ui.ButtonSet.YES_NO
  );
  return resp === ui.Button.YES;
}

/* 
 * PHONE VISIBILITY TOGGLE
 * ═══════════════════════════════════════════════════════════════ */

function SSCCP_MENU_togglePhoneVisibility() {
  const props = PropertiesService.getScriptProperties();

  // Migrate legacy key if present
  const legacy = props.getProperty("SSCCP_INCLUDE_PHONE");
  if (legacy !== null && legacy !== "") {
    props.setProperty("SSCCP_SHOW_REP_PHONES", legacy === "true" ? "true" : "false");
    props.deleteProperty("SSCCP_INCLUDE_PHONE");
  }

  const cur = props.getProperty("SSCCP_SHOW_REP_PHONES") === "true";
  const next = !cur;
  props.setProperty("SSCCP_SHOW_REP_PHONES", next ? "true" : "false");

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Phone # in Rep Alerts is now " + (next ? "ON (Show Full Numbers)" : "OFF (Masked ***-***-1234)"),
    "Safe Ship",
    6
  );
}

function SSCCP_MENU_getShowPhones_() {
  const props = PropertiesService.getScriptProperties();

  const v = props.getProperty("SSCCP_SHOW_REP_PHONES");
  if (v === "true") return true;
  if (v === "false") return false;

  // Fallback to legacy key
  return props.getProperty("SSCCP_INCLUDE_PHONE") === "true";
}

/* 
 * GRANOT ALERT FUNCTIONS — Quick Alerts with Presets
 * ═══════════════════════════════════════════════════════════════ */

function GRANOT_quickAlert_NotWorked() {
  if (typeof GRANOT_runAlert === "function") {
    GRANOT_runAlert({ workStatus: "NOT_WORKED", sendMode: "BY_REP" });
  } else {
    SpreadsheetApp.getUi().alert("GRANOT_runAlert function not found.\n\nMake sure GRANOT_Alerts.gs is in the project.");
  }
}

function GRANOT_quickAlert_Hot() {
  if (typeof GRANOT_runAlert === "function") {
    GRANOT_runAlert({ moveProximityDays: 7, sendMode: "BY_REP" });
  } else {
    SpreadsheetApp.getUi().alert("GRANOT_runAlert function not found.");
  }
}

function GRANOT_quickAlert_HighValue() {
  if (typeof GRANOT_runAlert === "function") {
    GRANOT_runAlert({ minEstTotal: 5000, sendMode: "BY_REP" });
  } else {
    SpreadsheetApp.getUi().alert("GRANOT_runAlert function not found.");
  }
}

function GRANOT_quickAlert_Stale() {
  if (typeof GRANOT_runAlert === "function") {
    GRANOT_runAlert({ minLeadAgeDays: 3, maxTotalContacts: 0, sendMode: "BY_REP" });
  } else {
    SpreadsheetApp.getUi().alert("GRANOT_runAlert function not found.");
  }
}

function GRANOT_quickAlert_ZeroCalls() {
  if (typeof GRANOT_runAlert === "function") {
    GRANOT_runAlert({ maxCalls: 0, sendMode: "BY_REP" });
  } else {
    SpreadsheetApp.getUi().alert("GRANOT_runAlert function not found.");
  }
}

/* 
 * GRANOT REPORTS
 * ═══════════════════════════════════════════════════════════════ */

function GRANOT_report_ManagerSummary() {
  if (typeof GRANOT_runAlert === "function") {
    GRANOT_runAlert({ sendMode: "SUMMARY" });
  } else {
    SpreadsheetApp.getUi().alert("GRANOT_runAlert function not found.");
  }
}

function GRANOT_report_TeamReports() {
  if (typeof GRANOT_runAlert === "function") {
    GRANOT_runAlert({ sendMode: "BY_TEAM" });
  } else {
    SpreadsheetApp.getUi().alert("GRANOT_runAlert function not found.");
  }
}

function GRANOT_report_AdminHealth() {
  if (typeof GRANOT_runAlert === "function") {
    GRANOT_runAlert({ sendMode: "ADMIN" });
  } else {
    SpreadsheetApp.getUi().alert("GRANOT_runAlert function not found.");
  }
}

function GRANOT_exportCurrentView() {
  if (typeof GRANOT_exportData === "function") {
    GRANOT_exportData();
  } else {
    SpreadsheetApp.getUi().alert("GRANOT_exportData function not found.");
  }
}

/* 
 * GRANOT TOOLS
 * ═══════════════════════════════════════════════════════════════ */

function GRANOT_openAlertBuilder() {
  if (typeof GRANOT_showAlertBuilder === "function") {
    GRANOT_showAlertBuilder();
  } else {
    SpreadsheetApp.getUi().alert("GRANOT Alert Builder not found.\n\nMake sure GRANOT_Alerts.gs is in the project.");
  }
}

function GRANOT_showStats() {
  if (typeof GRANOT_showDataStats === "function") {
    GRANOT_showDataStats();
  } else {
    SpreadsheetApp.getUi().alert("GRANOT Stats function not found.");
  }
}

function GRANOT_testLookup() {
  if (typeof GRANOT_testLeadLookup === "function") {
    GRANOT_testLeadLookup();
  } else {
    SpreadsheetApp.getUi().alert("GRANOT Test Lookup function not found.");
  }
}

function GRANOT_managePresets() {
  if (typeof GRANOT_showPresetManager === "function") {
    GRANOT_showPresetManager();
  } else {
    SpreadsheetApp.getUi().alert("GRANOT Preset Manager not found.");
  }
}

/* 
 * SLACK TEST
 * ═══════════════════════════════════════════════════════════════ */

function SSCCP_testSlackDMAdmin() {
  const CFG = getConfig_();
  
  if (!CFG.ADMIN_EMAIL) {
    SpreadsheetApp.getUi().alert("ADMIN_EMAIL is not set in Script Properties.");
    return;
  }
  
  if (!CFG.SLACK.BOT_TOKEN) {
    SpreadsheetApp.getUi().alert("SLACK_BOT_TOKEN is not set in Script Properties.");
    return;
  }
  
  const result = SSCCP_sendSlackDMToEmail_(CFG, CFG.ADMIN_EMAIL, "🧪 Test message from Safe Ship Contact Checker\n\nIf you see this, Slack DM is working!");
  
  if (result.ok) {
    SpreadsheetApp.getActiveSpreadsheet().toast("✅ Slack DM sent to " + CFG.ADMIN_EMAIL, "Slack Test", 6);
  } else {
    SpreadsheetApp.getUi().alert("❌ Slack DM failed:\n\n" + result.error);
  }
}

/* 
 * CLEAR NOTIFICATION LOG (with confirmation)
 * ═══════════════════════════════════════════════════════════════ */

function SSCCP_clearNotificationLogConfirm() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "⚠️ Clear Notification Log?",
    "This will delete ALL rows in Notification_Log (except header).\n\nThis cannot be undone.\n\nProceed?",
    ui.ButtonSet.YES_NO
  );
  
  if (resp !== ui.Button.YES) return;
  
  const CFG = getConfig_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEETS.NOTIFICATION_LOG);
  
  if (!sh) {
    ui.alert("Notification_Log sheet not found.");
    return;
  }
  
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.deleteRows(2, lastRow - 1);
  }
  
  ss.toast("✅ Notification_Log cleared.", "Safe Ship", 5);
}

/* 
 * FALLBACK ERROR VIEWER
 * ═══════════════════════════════════════════════════════════════ */

function SSCCP_MENU_showMenuError() {
  try {
    getConfig_();
    SpreadsheetApp.getUi().alert(
      "Config loaded OK now.\n\nReload the spreadsheet (Ctrl+Shift+R) to restore the full menu."
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      "Menu/Config error:\n\n" + String(e && e.stack ? e.stack : e)
    );
  }
}

/* 
 * FORCE SHOW MENU (Debug)
 * ═══════════════════════════════════════════════════════════════ */

function SSCCP_MENU_FORCE_SHOW() {
  SpreadsheetApp.getUi()
    .createMenu("🚀 Safe Ship (TEST)")
    .addItem("✅ Menu is working", "SSCCP_showStatus")
    .addToUi();
}

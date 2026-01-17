/**************************************************************
 * [SHIP] SAFE SHIP CONTACT CHECKER -- Config.gs
 * Contact Activity Rendering + Intelligent Process Direction
 *
 * CRITICAL RULE:
 * - NO top-level executable code that can throw.
 * - Everything must happen inside functions.
 * - getConfig_() is defined in Code.gs - do NOT redefine here
 *
 * v2.7 -- INTELLIGENT PROCESS DIRECTION
 * Process Direction now analyzes ACTUAL contact activity:
 *   - If no SMS AND no calls -> "SMS & CALL/VM"
 *   - If calls made but no SMS -> "SEND SMS" (missing action)
 *   - If SMS sent but no calls -> "CALL/VM" (missing action)
 *   - If both done -> Based on tracker type (follow-up action)
 **************************************************************/

function renderLeadContactActivity_v2_(lead) {
  var rc = lead.rc || {};
  var html = '';
  
  // Alert: Lead answered through text
  if (rc.hasReplied) {
    html += '<div style="background:linear-gradient(135deg,#10b981,#059669);color:#fff;padding:8px 14px;border-radius:6px;margin:10px 18px 0;display:flex;align-items:center;gap:8px;font-size:12px;">';
    html += '<span style="font-size:16px;">[MSG]</span>';
    html += '<span><strong>This Lead Answered Through Text!</strong> Check SMS conversation for their response</span></div>';
  }
  
  // Alert: Wrong Priority Risk (call > 4 minutes)
  if (rc.hasLongCall) {
    html += '<div style="background:linear-gradient(135deg,#f59e0b,#d97706);color:#000;padding:8px 14px;border-radius:6px;margin:10px 18px 0;display:flex;align-items:center;gap:8px;font-size:12px;">';
    html += '<span style="font-size:16px;">[!]</span>';
    html += '<span><strong>Wrong Priority Risk - Review File</strong> (Call duration: ' + (rc.longestCallEverDisplay || rc.longestCallDisplay || "4+ min") + ')</span></div>';
  }
  
  html += '<div style="padding:14px 18px;">';
  
  // Stats Row (Yesterday -> Today)
  html += '<div style="font-size:10px;color:#64748b;font-weight:600;margin-bottom:8px;">[CHART] CONTACT ACTIVITY (YESTERDAY -> TODAY)</div>';
  html += '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;">';
  html += renderStatBubble_v2_("SMS SENT", rc.smsYesterday || 0, rc.smsToday || 0, "[MOBILE]", "#dbeafe", "#1e40af");
  html += renderStatBubble_v2_("CALLS", rc.callsYesterday || 0, rc.callsToday || 0, "[PHONE]", "#dcfce7", "#166534");
  html += renderStatBubble_v2_("VOICEMAILS", rc.vmYesterday || 0, rc.vmToday || 0, "📥", "#fef3c7", "#92400e");
  html += '<div style="background:#f1f5f9;border-radius:6px;padding:6px 10px;text-align:center;min-width:70px;">';
  html += '<div style="font-size:9px;color:#64748b;font-weight:600;">LONGEST</div>';
  html += '<div style="font-size:14px;font-weight:800;color:#334155;">' + (rc.longestCallDisplay || "0:00") + '</div></div>';
  html += '</div>';
  
  // 
  // BUILD CALLS HTML (with IN badge for inbound)
  // 
  var callsHtml = '';
  var callBg = (rc.last5Calls && rc.last5Calls.length > 0) ? "#dcfce7" : "#fee2e2";
  var callColor = (rc.last5Calls && rc.last5Calls.length > 0) ? "#166534" : "#991b1b";
  
  if (rc.last5Calls && rc.last5Calls.length > 0) {
    for (var i = 0; i < rc.last5Calls.length; i++) {
      var call = rc.last5Calls[i];
      
      // IN badge for inbound calls (cyan)
      var inBadge = '';
      if (call.isInbound) {
        inBadge = '<span style="display:inline-block;background:#06b6d4;color:#fff;font-size:9px;font-weight:700;padding:1px 4px;border-radius:3px;margin-right:4px;">IN</span>';
      }
      
      // VM badge for voicemails (orange) - only for outbound
      var vmBadge = '';
      if (call.isVM && !call.isInbound) {
        vmBadge = '<span style="display:inline-block;background:#f59e0b;color:#fff;font-size:9px;font-weight:700;padding:1px 4px;border-radius:3px;margin-left:4px;">VM</span>';
      }
      
      callsHtml += '<div style="font-size:10px;color:' + callColor + ';padding:1px 0;white-space:nowrap;">';
      callsHtml += inBadge + call.timeAgo + ' <span style="opacity:0.7;">(' + call.timeFormatted + ')</span> ' + vmBadge + '<strong>' + call.durationDisplay + '</strong></div>';
    }
  } else {
    callsHtml = '<div style="font-size:10px;color:' + callColor + ';font-weight:600;">Never</div>';
  }
  
  // 
  // BUILD SMS HTML
  // 
  var smsHtml = '';
  var smsBg = (rc.last5SMS && rc.last5SMS.length > 0) ? "#dbeafe" : "#fee2e2";
  var smsColor = (rc.last5SMS && rc.last5SMS.length > 0) ? "#1e40af" : "#991b1b";
  
  if (rc.last5SMS && rc.last5SMS.length > 0) {
    for (var j = 0; j < rc.last5SMS.length; j++) {
      var sms = rc.last5SMS[j];
      smsHtml += '<div style="font-size:10px;color:' + smsColor + ';padding:1px 0;white-space:nowrap;">';
      smsHtml += sms.timeAgo + ' <span style="opacity:0.7;">(' + sms.timeFormatted + ')</span> <span style="color:#10b981;">[x]</span></div>';
    }
  } else {
    smsHtml = '<div style="font-size:10px;color:' + smsColor + ';font-weight:600;">Never</div>';
  }
  
  // 
  // BUILD PROCESS DIRECTION (INTELLIGENT - based on actual activity)
  // 
  var processDirection = getProcessDirection_(lead.trackerType, rc);
  
  // 
  // RENDER THREE BOXES: CALLS | SMS | PROCESS DIRECTION
  // 
  html += '<div style="display:flex;gap:8px;flex-wrap:wrap;">';
  
  // CALLS BOX (green)
  html += '<div style="background:' + callBg + ';border-radius:6px;padding:8px 10px;flex:1;min-width:120px;">';
  html += '<div style="font-size:9px;color:' + callColor + ';font-weight:700;margin-bottom:4px;">[PHONE] LAST 5 CALLS</div>';
  html += callsHtml;
  html += '</div>';
  
  // SMS BOX (blue)
  html += '<div style="background:' + smsBg + ';border-radius:6px;padding:8px 10px;flex:1;min-width:120px;">';
  html += '<div style="font-size:9px;color:' + smsColor + ';font-weight:700;margin-bottom:4px;">[MSG] LAST 5 SMS</div>';
  html += smsHtml;
  html += '</div>';
  
  // PROCESS DIRECTION BOX (purple/action)
  html += '<div style="background:' + processDirection.bgColor + ';border-radius:6px;padding:8px 10px;flex:1;min-width:120px;border:2px solid ' + processDirection.borderColor + ';">';
  html += '<div style="font-size:9px;color:' + processDirection.textColor + ';font-weight:700;margin-bottom:4px;">🎯 PROCESS DIRECTION</div>';
  html += '<div style="font-size:12px;color:' + processDirection.textColor + ';font-weight:800;margin-bottom:4px;">' + processDirection.icon + ' ' + processDirection.action + '</div>';
  html += '<div style="font-size:9px;color:' + processDirection.textColor + ';opacity:0.8;">' + processDirection.reason + '</div>';
  html += '</div>';
  
  html += '</div>'; // end flex container
  html += '</div>'; // end padding container
  return html;
}

/**
 * 🎯 Get Process Direction based on ACTUAL contact activity + tracker type
 * 
 * @param {string} trackerType - The source tracker (SMS, CALL/VM, etc.)
 * @param {object} rc - RingCentral enrichment data with contact history
 * @returns {object} Process direction config with action, icon, colors, reason
 */
function getProcessDirection_(trackerType, rc) {
  var type = String(trackerType || "").toUpperCase();
  rc = rc || {};
  
  // Calculate totals from RC data
  var totalSMS = (rc.smsYesterday || 0) + (rc.smsToday || 0);
  var totalCalls = (rc.callsYesterday || 0) + (rc.callsToday || 0);
  var hasSMS = totalSMS > 0 || (rc.last5SMS && rc.last5SMS.length > 0);
  var hasCalls = totalCalls > 0 || (rc.last5Calls && rc.last5Calls.length > 0);
  
  // 
  // PRIORITY 1: If lead replied, always show "CHECK REPLY"
  // 
  if (rc.hasReplied) {
    return {
      action: "CHECK REPLY",
      icon: "[MSG][GOAL]",
      reason: "Lead responded via text!",
      source: "Reply Detected",
      bgColor: "#DCFCE7",
      borderColor: "#10B981",
      textColor: "#166534"
    };
  }
  
  // 
  // PRIORITY 2: Determine what's MISSING based on activity
  // 
  
  // Neither SMS nor Calls made -> Need BOTH
  if (!hasSMS && !hasCalls) {
    return {
      action: "SMS & CALL/VM",
      icon: "[MOBILE][PHONE]",
      reason: "No contact attempts yet",
      source: "Activity Analysis",
      bgColor: "#FEE2E2",
      borderColor: "#EF4444",
      textColor: "#991B1B"
    };
  }
  
  // Has calls but NO SMS -> Need SMS
  if (hasCalls && !hasSMS) {
    return {
      action: "SEND SMS",
      icon: "[MOBILE]",
      reason: "Called but no SMS sent",
      source: "Activity Analysis",
      bgColor: "#DBEAFE",
      borderColor: "#3B82F6",
      textColor: "#1E40AF"
    };
  }
  
  // Has SMS but NO calls -> Need calls
  if (hasSMS && !hasCalls) {
    return {
      action: "CALL/VM",
      icon: "[PHONE]",
      reason: "SMS sent but no calls",
      source: "Activity Analysis",
      bgColor: "#DCFCE7",
      borderColor: "#22C55E",
      textColor: "#166534"
    };
  }
  
  // 
  // PRIORITY 3: Both SMS and Calls made -> Show follow-up based on tracker
  // 
  
  // Check if more attempts needed today
  var callsToday = rc.callsToday || 0;
  var smsToday = rc.smsToday || 0;
  
  // If no activity TODAY, encourage follow-up
  if (callsToday === 0 && smsToday === 0) {
    return {
      action: "FOLLOW UP TODAY",
      icon: "🔄",
      reason: "No contact today yet",
      source: "Activity Analysis",
      bgColor: "#FEF3C7",
      borderColor: "#F59E0B",
      textColor: "#92400E"
    };
  }
  
  // 
  // PRIORITY 4: Active today - show tracker-based direction
  // 
  
  // SAME DAY TRANSFERS - urgent
  if (type.indexOf("TRANSFER") !== -1 || type.indexOf("SAME DAY") !== -1) {
    return {
      action: "URGENT TRANSFER",
      icon: "⚡",
      reason: "Same Day Transfer",
      source: "Transfer Tracker",
      bgColor: "#F3E8FF",
      borderColor: "#A855F7",
      textColor: "#7C3AED"
    };
  }
  
  // PRIORITY 1 CALL/VM
  if (type.indexOf("PRIORITY") !== -1) {
    return {
      action: "PRIORITY CALL",
      icon: "[GOAL][PHONE]",
      reason: "Priority 1 Lead",
      source: "Priority 1 Tracker",
      bgColor: "#FEE2E2",
      borderColor: "#EF4444",
      textColor: "#991B1B"
    };
  }
  
  // QUOTED FOLLOW-UP / CONTACTED LEADS
  if (type.indexOf("QUOTED") !== -1 || type.indexOf("CONTACTED") !== -1 || type.indexOf("FOLLOW") !== -1) {
    return {
      action: "CLOSE THE DEAL",
      icon: "💰",
      reason: "Already quoted - follow up",
      source: "Quoted Tracker",
      bgColor: "#FEF3C7",
      borderColor: "#F59E0B",
      textColor: "#92400E"
    };
  }
  
  // SMS TRACKER - continue SMS cadence
  if (type === "SMS") {
    return {
      action: "CONTINUE SMS",
      icon: "[MOBILE][x]",
      reason: "Maintain SMS cadence",
      source: "SMS Tracker",
      bgColor: "#DBEAFE",
      borderColor: "#3B82F6",
      textColor: "#1E40AF"
    };
  }
  
  // CALL & VOICEMAIL TRACKER - continue call cadence
  if (type === "CALL/VM" || type === "CALL" || type === "VOICEMAIL") {
    return {
      action: "CONTINUE CALLS",
      icon: "[PHONE][x]",
      reason: "Maintain call cadence",
      source: "Call & VM Tracker",
      bgColor: "#DCFCE7",
      borderColor: "#22C55E",
      textColor: "#166534"
    };
  }
  
  // DEFAULT - general follow up
  return {
    action: "KEEP WORKING",
    icon: "[OK]",
    reason: "Active lead - continue",
      source: "General",
    bgColor: "#F1F5F9",
    borderColor: "#94A3B8",
    textColor: "#475569"
  };
}

function renderStatBubble_v2_(label, yesterday, today, icon, bgColor, textColor) {
  var html = '<div style="background:' + bgColor + ';border-radius:6px;padding:6px 10px;text-align:center;min-width:80px;">';
  html += '<div style="font-size:9px;color:' + textColor + ';font-weight:600;margin-bottom:2px;">' + icon + ' ' + label + '</div>';
  html += '<div style="display:flex;align-items:center;justify-content:center;gap:3px;">';
  html += '<span style="font-size:13px;font-weight:800;color:' + textColor + ';">' + yesterday + '</span>';
  html += getActivityDot_v2_(yesterday);
  html += '<span style="font-size:11px;color:#64748b;">-></span>';
  html += '<span style="font-size:13px;font-weight:800;color:' + textColor + ';">' + today + '</span>';
  html += getActivityDot_v2_(today);
  html += '</div></div>';
  return html;
}

function getActivityDot_v2_(count) {
  var color = count === 0 ? "#ef4444" : (count >= 3 ? "#3b82f6" : "#f59e0b");
  return '<span style="width:5px;height:5px;border-radius:50%;background:' + color + ';display:inline-block;margin-left:1px;"></span>';
}

// NOTE: getConfig_() is defined in Code.gs - DO NOT add a duplicate here
// The broken SSCCP_getConfig() bridge has been REMOVED to fix the menu error
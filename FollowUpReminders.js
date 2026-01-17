/**************************************************************
 * FollowUpReminders.gs (Compatibility)
 * Keeps legacy entry points so buttons/menu never break.
 **************************************************************/

function sendFollowUpReminders() {
  SSCCP_runQuotedFollowUpAlerts();
}

// Optional legacy names (safe)
function runQuotedContactedLeadAlerts() {
  SSCCP_runQuotedFollowUpAlerts();
}

function runPriority1CallVmAlerts() {
  SSCCP_runPriority1CallVmAlerts();
}

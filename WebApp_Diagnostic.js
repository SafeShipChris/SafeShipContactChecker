/**
 * 🔧 DIAGNOSTIC FUNCTION - Run this in Apps Script to test what's working
 * 
 * Steps:
 * 1. Open Apps Script Editor
 * 2. Select this function from dropdown
 * 3. Click Run
 * 4. Check the Execution Log for results
 */
function WEBAPP_DIAGNOSTIC_TEST() {
  Logger.log('🔍 Starting diagnostic test...');
  
  var results = {
    timestamp: new Date().toISOString(),
    tests: []
  };
  
  // Test 1: Security context
  try {
    var context = WebApp_Security_getRepContext();
    results.tests.push({
      name: 'Security Context',
      status: 'OK',
      data: {
        authorized: context.authorized,
        username: context.username,
        isAdmin: context.isAdmin
      }
    });
    Logger.log('✅ Security: ' + JSON.stringify(context));
  } catch (e) {
    results.tests.push({ name: 'Security Context', status: 'FAILED', error: e.message });
    Logger.log('❌ Security FAILED: ' + e.message);
  }
  
  // Test 2: Templates
  try {
    var templates = WebApp_Templates_getAll();
    results.tests.push({
      name: 'Templates',
      status: 'OK',
      count: templates.length
    });
    Logger.log('✅ Templates: ' + templates.length + ' loaded');
  } catch (e) {
    results.tests.push({ name: 'Templates', status: 'FAILED', error: e.message });
    Logger.log('❌ Templates FAILED: ' + e.message);
  }
  
  // Test 3: Dashboard Stats
  try {
    var stats = WebApp_Data_getDashboardStats('TEST', true);
    results.tests.push({
      name: 'Dashboard Stats',
      status: 'OK',
      data: stats
    });
    Logger.log('✅ Stats: ' + JSON.stringify(stats));
  } catch (e) {
    results.tests.push({ name: 'Dashboard Stats', status: 'FAILED', error: e.message });
    Logger.log('❌ Stats FAILED: ' + e.message);
  }
  
  // Test 4: Leads fetch
  try {
    var leads = WebApp_Data_getRepLeads('TEST', { tab: 'p0_sms', page: 1, pageSize: 5 }, true);
    results.tests.push({
      name: 'Leads Fetch',
      status: 'OK',
      total: leads.total,
      returned: leads.leads.length
    });
    Logger.log('✅ Leads: ' + leads.total + ' total, ' + leads.leads.length + ' returned');
  } catch (e) {
    results.tests.push({ name: 'Leads Fetch', status: 'FAILED', error: e.message });
    Logger.log('❌ Leads FAILED: ' + e.message);
  }
  
  // Test 5: Audit
  try {
    var audit = WebApp_Audit_getRepSummary('TEST');
    results.tests.push({
      name: 'Audit Summary',
      status: 'OK',
      data: audit
    });
    Logger.log('✅ Audit: ' + JSON.stringify(audit));
  } catch (e) {
    results.tests.push({ name: 'Audit Summary', status: 'FAILED', error: e.message });
    Logger.log('❌ Audit FAILED: ' + e.message);
  }
  
  // Test 6: Full initial load
  try {
    var initial = WEBAPP_getInitialLoad();
    results.tests.push({
      name: 'Initial Load',
      status: initial.authorized ? 'OK' : 'AUTH_FAILED',
      authorized: initial.authorized,
      errors: initial.errors,
      loadTimeMs: initial.loadTimeMs
    });
    Logger.log('✅ Initial Load: ' + initial.loadTimeMs + 'ms');
    if (initial.errors && initial.errors.length > 0) {
      Logger.log('   ⚠️ Errors: ' + initial.errors.join(', '));
    }
  } catch (e) {
    results.tests.push({ name: 'Initial Load', status: 'FAILED', error: e.message });
    Logger.log('❌ Initial Load FAILED: ' + e.message);
  }
  
  Logger.log('\n📋 SUMMARY:');
  Logger.log(JSON.stringify(results, null, 2));
  
  return results;
}
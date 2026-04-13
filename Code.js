const API_ENDPOINT = 'https://thefinuportal-backend-1014598876589.europe-west1.run.app/';

var UserEmail = '';
var UserSpreadsheet = null;

try {
  UserEmail = Session.getActiveUser().getEmail();
  // In time-based triggers, getActiveUser().getEmail() returns '' for consumer accounts.
  // Fall back to the email stored during interactive setup.
  if (!UserEmail) {
    UserEmail = PropertiesService.getUserProperties().getProperty('USER_EMAIL') || '';
  }
  UserSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
} catch(e) {
  // Fails silently in trigger/non-interactive contexts
  try {
    if (!UserEmail) {
      UserEmail = PropertiesService.getUserProperties().getProperty('USER_EMAIL') || '';
    }
  } catch(e2) {}
}

const USER_START_HERE_SHEET = 'Start Here';
const USER_DATA_SHEET = 'Data';
const USER_BALANCE_HISTORY_SHEET = 'Balance History';
const USER_ACCOUNTS_SHEET = 'Accounts';
const USER_DEFINITION_SHEET = 'Definition';
const USER_TRANSACTIONS_SHEET = 'Transactions';
const USER_INVESTMENTS_SHEET = 'Investments';
const USER_CATEGORIES_SHEET = 'Categories';
const USER_RECONCILE_SHEET = 'Reconcile';
const USER_PLAID_SHEET = 'Plaid Data';
const USER_NET_WORTH_SHEET = 'Net Worth';
const USER_JOINT_NET_WORTH_SHEET = 'Joint Net Worth';
const USER_MONTHLY_BUDGET_SHEET = 'Monthly Budget';
const USER_JOINT_MONTHLY_BUDGET_SHEET= 'Joint Monthly Budget';
const USER_YEARLY_BUDGET_SHEET = 'Yearly Budget';
const USER_JOINT_YEARLY_BUDGET_SHEET = 'Joint Yearly Budget';
const USER_BUDGET_MAKER_SHEET = 'Budget Maker';


const installableTriggers = {
  onEdit: 'onEditHandler',
  onOpen: 'onOpenHandler',
  dailyAutoSync: 'runThefinUPlaidAutoSync'
};

var _appUserId = null;

function getAppUserId() {
  if (!_appUserId) {
    _appUserId = { client_user_id: generateRandomNumber() };
  }
  return _appUserId;
}

// Configuration
const RESTAPI_CONFIG = {
  API_BASE_URL: API_ENDPOINT,
  TIMEOUT: 30000 // 30 seconds
};

/**
 * Simple trigger — runs every time the spreadsheet opens.
 * Google Marketplace requirement: onOpen must ALWAYS create the add-on menu
 * and must NOT call services that require authorization in AuthMode.NONE.
 *
 * Heavy operations (templates, triggers, auto-sync) are handled by:
 *   - The Setup Wizard (first-time setup)
 *   - The installable onOpenHandler() trigger (recovery on subsequent opens)
 */
function onOpen(e) {
  try {
    // ALWAYS create the add-on menu — works in every AuthMode
    SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Open', 'showSidebar')
      .addToUi();
  } catch (menuErr) {
    Logger.log('onOpen menu error: ' + menuErr.toString());
  }

  // Persist email when we have authorization (LIMITED or FULL)
  try {
    var authMode = e && e.authMode;
    if (authMode !== ScriptApp.AuthMode.NONE && UserEmail) {
      PropertiesService.getUserProperties().setProperty('USER_EMAIL', UserEmail);
    }
  } catch (err) {
    Logger.log('onOpen email persist error: ' + err.toString());
  }
}

/**
 * Runs once when the add-on is installed from the Marketplace.
 * Google Marketplace requirement: must call onOpen(e) to create the menu.
 * All heavy setup (templates, triggers, auto-sync) happens in the Setup Wizard.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Installable onOpen trigger handler (runs in AuthMode.FULL).
 * Only performs recovery/maintenance work if the setup wizard has been completed.
 * This ensures we never run heavy operations before the user finishes onboarding.
 *
 * On each open (post-setup):
 *   1. Recover any missing template sheets
 *   2. Self-heal missing installable triggers
 *   3. Re-apply all sheet formulas in correct dependency order
 *   4. Run a throttled auto-sync (if auto-sync enabled, max once per hour)
 */
function onOpenHandler(e) {
  try {
    // Skip all work if setup wizard hasn't been completed yet
    if (!isSetupCompleted()) {
      return;
    }

    showSyncToast_('Checking your sheets and data...');

    // 1. Recover any missing template sheets
    installTemplateInitialSetup();

    // 2. Self-heal missing triggers
    ensureTriggersExist();

    // 3. Re-apply all formulas so cross-references stay correct
    reApplyAllFormulas();

    // 4. Run auto-sync if enabled (throttled, locked, subscription-checked)
    var syncEnabled = PropertiesService.getUserProperties().getProperty('AUTO_SYNC_STATUS') === 'true';
    if (syncEnabled) {
      throttledSync();
    } else {
      showSyncToast_('Sheets are up to date.');
    }
  } catch (err) {
    Logger.log('onOpenHandler error: ' + err.toString());
    showSyncToast_('An error occurred during initialization. Please reload.');
  }
}

/**
 * Runs auto-sync directly inside the installable onOpenHandler trigger.
 * Installable triggers get up to 30 minutes of execution time, which is
 * sufficient for runThefinUPlaidAutoSync (25-min safety limit).
 *
 * Guards:
 *   - Throttled: skips if last sync was less than 1 hour ago
 *   - Locked: prevents concurrent syncs from overlapping (e.g. multiple tabs)
 *   - Validates subscription before running
 *
 * User feedback:
 *   - Toast notifications for sync start/progress/completion
 *   - Sheet protection during sync to prevent conflicting edits
 */
function throttledSync() {
  var userProperties = PropertiesService.getUserProperties();
  var lastSync = userProperties.getProperty('LAST_SYNC_TIMESTAMP');
  var now = new Date().getTime();
  var oneHour = 60 * 60 * 1000;

  // Throttle: skip if synced recently
  if (lastSync && (now - parseInt(lastSync)) < oneHour) {
    Logger.log('[throttledSync] Skipped — last sync was ' + Math.round((now - parseInt(lastSync)) / 60000) + ' min ago.');
    showSyncToast_('Sync skipped — last sync was ' + Math.round((now - parseInt(lastSync)) / 60000) + ' min ago.');
    return;
  }

  // Validate subscription once
  var session = validateUserSession();
  if (!session.result || !session.result.data || session.result.data.isSubscribed !== true) {
    Logger.log('[throttledSync] Skipped — user not subscribed.');
    return;
  }

  // Lock: prevent concurrent syncs from multiple tabs
  var lock = LockService.getUserLock();
  if (!lock.tryLock(5000)) {
    Logger.log('[throttledSync] Skipped — another sync is already running.');
    return;
  }

  try {
    // Notify user and lock sheets
    showSyncToast_('Syncing your financial data... Please wait.');
    protectSheetsForSync_();
    userProperties.setProperty('SYNC_IN_PROGRESS', 'true');

    runThefinUPlaidAutoSync();
    userProperties.setProperty('LAST_SYNC_TIMESTAMP', new Date().getTime().toString());

    // Notify completion
    showSyncToast_('Sync completed successfully. Your data is up to date.');
  } catch (e) {
    Logger.log('[throttledSync] error: ' + e.toString());
    showSyncToast_('Sync encountered an error. Please try again from the sidebar.');
  } finally {
    // Always unlock sheets and clear status
    unprotectSheetsAfterSync_();
    userProperties.deleteProperty('SYNC_IN_PROGRESS');
    lock.releaseLock();
  }
}

/**
 * Shows a toast notification to the user. Fails silently in non-interactive contexts
 * (e.g. time-based triggers where there is no active spreadsheet UI).
 */
function showSyncToast_(message) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'TheFinU Auto-Sync', -1);
  } catch (e) {
    // Toast unavailable in non-interactive trigger contexts
  }
}

/**
 * Protects all data sheets with a warning during sync to prevent user edits
 * that could conflict with incoming data. Uses warning-only protection so
 * the user sees a "are you sure?" prompt rather than being fully locked out.
 * Each protection is tagged with a description so we can find and remove them later.
 */
function protectSheetsForSync_() {
  try {
    var ss = UserSpreadsheet;
    var sheetsToProtect = syncRequiredSheets();
    sheetsToProtect.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        var protection = sheet.protect().setDescription('THEFINU_SYNC_LOCK');
        protection.setWarningOnly(true);
      }
    });
  } catch (e) {
    Logger.log('protectSheetsForSync_ error: ' + e.toString());
  }
}

/**
 * Removes the sync-time protections added by protectSheetsForSync_().
 * Only removes protections tagged with our description — leaves user-created protections intact.
 */
function unprotectSheetsAfterSync_() {
  try {
    var ss = UserSpreadsheet;
    var sheetsToUnprotect = syncRequiredSheets();
    sheetsToUnprotect.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
        protections.forEach(function(protection) {
          if (protection.getDescription() === 'THEFINU_SYNC_LOCK') {
            protection.remove();
          }
        });
      }
    });
  } catch (e) {
    Logger.log('unprotectSheetsAfterSync_ error: ' + e.toString());
  }
}

/**
 * Re-applies all formulas across base and budget sheets in the correct dependency order.
 * Safe to call on every open — formulas are idempotent.
 */
function reApplyAllFormulas() {
  try {
    // Base formulas first (Transactions + Definition) — budget sheets depend on these
    var baseSheets = appBaseFormulaTemplates();
    baseSheets.forEach(function(sheetName) {
      if (UserSpreadsheet.getSheetByName(sheetName)) {
        reApplyFormulaToSpreadsheet(sheetName);
      }
    });

    SpreadsheetApp.flush();

    // Budget sheet formulas (depend on Definition being ready)
    var budgetSheets = appBudgetFormulaTemplates();
    budgetSheets.forEach(function(sheetName) {
      if (UserSpreadsheet.getSheetByName(sheetName)) {
        reApplyFormulaToSpreadsheet(sheetName);
      }
    });

    // Re-apply Definition so cross-references to budget sheets resolve
    if (UserSpreadsheet.getSheetByName(USER_DEFINITION_SHEET)) {
      reApplyFormulaToSpreadsheet(USER_DEFINITION_SHEET);
    }
  } catch (e) {
    Logger.log('reApplyAllFormulas error: ' + e.toString());
  }
}

function onEditHandler(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  var column = e.range.getColumn();

  if (sheetName === USER_ACCOUNTS_SHEET) {
    if (column === 6 || column === 8 || column === 10 || column === 11) {
      populateNetWorth();
      populateJointNetWorth();
    }
  } else if (sheetName === USER_BALANCE_HISTORY_SHEET) {
    if (column === 2 || column === 5) {
      var row = e.range.getRow();
      var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (rowData && rowData[1] !== '' && rowData[4] !== '') {
        populateNetWorth();
        populateJointNetWorth();
      }
    }
  }
}


/**
 * Validates that required sheets exist in the user's spreadsheet.
 * @param {string[]} sheetNames - Array of sheet name constants to check.
 * @returns {{ valid: boolean, missing: string[] }}
 */
function validateRequiredSheets(sheetNames){
  var spreadsheet = UserSpreadsheet;
  var missing = [];
  sheetNames.forEach(function(name){
    if(!spreadsheet.getSheetByName(name)){
      missing.push(name);
    }
  });
  return { valid: missing.length === 0, missing: missing };
}

/**
 * Returns sheets required for data sync operations (auto-sync & link import).
 */
function syncRequiredSheets(){
  return [
    USER_TRANSACTIONS_SHEET,
    USER_INVESTMENTS_SHEET,
    USER_BALANCE_HISTORY_SHEET,
    USER_ACCOUNTS_SHEET,
    USER_DEFINITION_SHEET
  ];
}

/**
 * Returns sheets required for report generation.
 */
function reportRequiredSheets(){
  return [
    USER_TRANSACTIONS_SHEET,
    USER_BALANCE_HISTORY_SHEET,
    USER_ACCOUNTS_SHEET,
    USER_CATEGORIES_SHEET,
    USER_DEFINITION_SHEET,
    USER_NET_WORTH_SHEET,
    USER_JOINT_NET_WORTH_SHEET,
    USER_MONTHLY_BUDGET_SHEET,
    USER_JOINT_MONTHLY_BUDGET_SHEET,
    USER_YEARLY_BUDGET_SHEET,
    USER_JOINT_YEARLY_BUDGET_SHEET
  ];
}

/**
 * Sends an email notification to the current user.
 * @param {string} subject
 * @param {string} body - Plain text body.
 */
function sendUserNotification(subject, body){
  try{
    var email = Session.getEffectiveUser().getEmail();
    if(email){
      MailApp.sendEmail(email, subject, body);
    }
  }catch(e){
    Logger.log('sendUserNotification error: ' + e.toString());
  }
}

function appBaseTemplates(){
  return [
    USER_START_HERE_SHEET,
    USER_CATEGORIES_SHEET,
    USER_DATA_SHEET,
    USER_BALANCE_HISTORY_SHEET,
    USER_ACCOUNTS_SHEET,
    USER_TRANSACTIONS_SHEET,
    USER_INVESTMENTS_SHEET,
    USER_NET_WORTH_SHEET,
    USER_JOINT_NET_WORTH_SHEET,
    USER_RECONCILE_SHEET,
    USER_MONTHLY_BUDGET_SHEET,
    USER_JOINT_MONTHLY_BUDGET_SHEET,
    USER_YEARLY_BUDGET_SHEET,
    USER_JOINT_YEARLY_BUDGET_SHEET,
    USER_DEFINITION_SHEET,
    USER_BUDGET_MAKER_SHEET,
  ];
}

function appBudgetFormulaTemplates(){
  return [
    USER_MONTHLY_BUDGET_SHEET,
    USER_JOINT_MONTHLY_BUDGET_SHEET,
    USER_YEARLY_BUDGET_SHEET,
    USER_JOINT_YEARLY_BUDGET_SHEET,
    USER_BUDGET_MAKER_SHEET
  ];
}

function appBaseFormulaTemplates(){
  return [
    USER_TRANSACTIONS_SHEET,
    USER_DEFINITION_SHEET
  ];
}

/**
 * Opens the sidebar UI.
 */
/**
 * Opens the sidebar UI.
 * Routes to the dashboard (if setup is complete and subscription is active)
 * or the setup wizard (if not).
 */
function showSidebar() {
  var userValidation = validateUserSession();
  var isSubscribed = userValidation.result && userValidation.result.data && userValidation.result.data.isSubscribed === true;

  if (isSubscribed) {
    // Mark subscription progress
    try {
      markSetupStepCompleted('subscription', { status: 'active', activatedAt: new Date().toISOString() });
    } catch (e) {
      Logger.log('Error marking subscription step completed: ' + e.toString());
    }
    // Show dashboard if all setup steps are complete, otherwise show wizard
    if (isSetupCompleted()) {
      showUserDashboardSidebar();
    } else {
      showSetupWizardSidebar();
    }
  } else {
    // Only clean up if the user previously completed setup (i.e. subscription expired).
    // Don't clean up on first-ever open when the wizard hasn't run yet.
    var progress = getSetupWizardProgress();
    var wasSetUp = progress.subscription && progress.subscription.status === 'active';
    if (wasSetUp) {
      try {
        cleanupExpiredSubscription();
      } catch (e) {
        Logger.log('cleanupExpiredSubscription error: ' + e.toString());
      }
    }
    showSetupWizardSidebar();
  }
}

function showSetupWizardSidebar(){
  const template = HtmlService.createTemplateFromFile('Index');
  template.wizardContent = showSetupWizardTemplate();
  const html = template.evaluate()
      .setTitle('ThefinU')
      .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);
}

function showUserDashboardSidebar(){
  const html = HtmlService.createTemplateFromFile('UserIndex')
    .evaluate()
    .setTitle('ThefinU')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function generateRandomNumber() {
  const min = 10000; // Minimum 5-digit number
  const max = 99999; // Maximum 5-digit number
  let number = Math.floor(Math.random() * (max - min + 1)) + min;
  return number.toString();
}

function getTodayDate(){
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
  return today;
}

function getTodayDateTime(){
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
  return today;
}

function getDateTime() {
  var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM dd, yyyy hh:mm a");
  return currentDate;
}

function createStripeSession(){
  try{
    const email = Session.getActiveUser().getEmail();
    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

    const apiUrl = API_ENDPOINT + 'api/payment/create-checkout-session';

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: getAuthHeaders(),
      payload: JSON.stringify({
        email: email,
        spreadsheetId: spreadsheetId
      }),
      muteHttpExceptions: true
    };

    const request = UrlFetchApp.fetch(apiUrl, options);
    const json = JSON.parse(request.getContentText());

    if (request.getResponseCode() === 200 && json.url) {
      return {
        success: true,
        checkoutUrl: json.url
      };
    }

    return {
      success: false,
      error: json.error || 'Failed to create checkout session'
    };
  }catch(e){
    return {
      success: false,
      error: e.toString()
    };
  }
}

function showSetupWizardTemplate(){
  try{
    const response = getAppSettings();
    if( response.success === true ){
      const template = HtmlService.createTemplateFromFile('SetupWizard');
      template.message = response.result.appInstruction;
      template.progress = getSetupWizardProgress();
      return template.evaluate().getContent();
    }else{
      const template = HtmlService.createTemplateFromFile('Error');
      template.message = "Something went wrong, please try again.";
      return template.evaluate().getContent();
    }
  }catch(error){
    const template = HtmlService.createTemplateFromFile('Error');
    template.message = "Something went wrong, please try again.";
    return template.evaluate().getContent();
  }
}

function showUserDashboardTemplate() {
  var template = HtmlService.createTemplateFromFile('UserDashboard');
  var userProps = PropertiesService.getUserProperties();
  template.isAutoSyncEnabled = userProps.getProperty('AUTO_SYNC_STATUS') === 'true';

  // Pass subscription cancellation state to template
  template.cancelAtPeriodEnd = false;
  template.currentPeriodEnd = '';
  try {
    var session = validateUserSession();
    if (session.success && session.result && session.result.data) {
      template.cancelAtPeriodEnd = session.result.data.cancelAtPeriodEnd === true;
      template.currentPeriodEnd = session.result.data.currentPeriodEnd || '';
    }
  } catch (e) {
    Logger.log('showUserDashboardTemplate subscription check error: ' + e.toString());
  }

  return template.evaluate().getContent();
}

/**
 * Generates an error message UI string.
 */
function getErrorUI(errorMessage) {
  const template = HtmlService.createTemplateFromFile('Error');
  template.message = errorMessage || "An unexpected error occurred while processing your request.";
  return template.evaluate().getContent();
}

function getTemplateBlockUI(file){
  const template = HtmlService.createTemplateFromFile(file);
  return template.evaluate().getContent();
}

/**
 * Setup wizard progress helpers
 * Stored in User Properties under key: SETUP_WIZARD_PROGRESS_{spreadsheetId}
 */
function getSetupProgressKey_(){
  try{
    var id = SpreadsheetApp.getActiveSpreadsheet().getId();
    return 'SETUP_WIZARD_PROGRESS_' + id;
  }catch(e){
    return 'SETUP_WIZARD_PROGRESS';
  }
}

function getSetupWizardProgress(){
  try{
    const userProps = PropertiesService.getUserProperties();
    const key = getSetupProgressKey_();
    const raw = userProps.getProperty(key);
    if( raw ){
      return JSON.parse(raw);
    }
  }catch(e){
    Logger.log('getSetupWizardProgress parse error: ' + e.toString());
  }
  return {
    subscription: { status: 'pending', startedAt: null, activatedAt: null },
    template: { status: 'pending', installedAt: null },
    config: { status: 'pending', autoSync: null, savedAt: null },
    completedSteps: [],
    updatedAt: null
  };
}

function setSetupWizardProgress(progressObj){
  try{
    const userProps = PropertiesService.getUserProperties();
    const key = getSetupProgressKey_();
    progressObj.updatedAt = new Date().toISOString();
    userProps.setProperty(key, JSON.stringify(progressObj));
    return true;
  }catch(e){
    Logger.log('setSetupWizardProgress error: ' + e.toString());
    return false;
  }
}

function markSetupStepCompleted(stepName, meta){
  try{
    const progress = getSetupWizardProgress();
    meta = meta || {};
    switch(stepName){
      case 'subscription':
        progress.subscription.status = meta.status || 'active';
        progress.subscription.startedAt = progress.subscription.startedAt || meta.startedAt || new Date().toISOString();
        progress.subscription.activatedAt = meta.activatedAt || (progress.subscription.status === 'active' ? new Date().toISOString() : null);
        break;
      case 'template':
        progress.template.status = meta.status || 'completed';
        progress.template.installedAt = meta.installedAt || new Date().toISOString();
        break;
      case 'config':
        progress.config.status = meta.status || 'completed';
        progress.config.autoSync = (typeof meta.autoSync !== 'undefined') ? meta.autoSync : progress.config.autoSync;
        progress.config.savedAt = meta.savedAt || new Date().toISOString();
        break;
      default:
        // noop
        break;
    }
    // update completedSteps array
    const idx = progress.completedSteps.indexOf(stepName);
    if( idx === -1 && (progress[stepName] && progress[stepName].status && progress[stepName].status !== 'pending') ){
      progress.completedSteps.push(stepName);
    }
    setSetupWizardProgress(progress);
    return progress;
  }catch(e){
    Logger.log('markSetupStepCompleted error: ' + e.toString());
    return null;
  }
}

function clearSubscriptionProgress(){
  try{
    const progress = getSetupWizardProgress();
    progress.subscription = { status: 'pending', startedAt: null, activatedAt: null };
    if( Array.isArray(progress.completedSteps) ){
      const idx = progress.completedSteps.indexOf('subscription');
      if( idx !== -1 ) progress.completedSteps.splice(idx, 1);
    }
    setSetupWizardProgress(progress);
    return progress;
  }catch(e){
    Logger.log('clearSubscriptionProgress error: ' + e.toString());
    return null;
  }
}

function isSetupCompleted(){
  const progress = getSetupWizardProgress();
  // required steps: subscription, template, config
  return (progress.subscription && progress.subscription.status === 'active') &&
         (progress.template && progress.template.status === 'completed') &&
         (progress.config && progress.config.status === 'completed');
}

/**
 * Shows a confirmation dialog and schedules subscription cancellation at end of billing period.
 * The user retains full access until the period ends. Cleanup happens when subscription actually expires.
 */
function cancelUserSubscription() {
  try {
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert('Cancel Subscription', 'Are you sure you want to cancel? Your subscription will remain active until the end of your current billing period.', ui.ButtonSet.YES_NO);
    if (result == ui.Button.YES) {
      var response = confirmCancelUserSubscription();
      if (response.success === true) {
        clearAppSettingsCache();
        var userSession = validateUserSession();
        var endDate = '';
        var cancelAtPeriodEnd = false;
        if (userSession.success && userSession.result && userSession.result.data) {
          endDate = userSession.result.data.currentPeriodEnd || '';
          cancelAtPeriodEnd = userSession.result.data.cancelAtPeriodEnd === true;
        }
        return { success: true, periodEnd: endDate, cancelAtPeriodEnd: cancelAtPeriodEnd };
      }
      return { success: false, message: 'Unable to cancel subscription. Please try again.' };
    }
    return { success: false, message: 'Cancellation was not confirmed.' };
  } catch (e) {
    Logger.log('cancelUserSubscription error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Cleans up user data when subscription has fully expired.
 * Called from showSidebar when a previously active subscription is no longer valid.
 */
function cleanupExpiredSubscription() {
  try {
    // Remove all add-on triggers
    ScriptApp.getProjectTriggers().forEach(function(trigger) {
      ScriptApp.deleteTrigger(trigger);
    });
    deleteAllSheetsAndRecreate();
    clearAllUserProperties();
    clearAppSettingsCache();
    clearSubscriptionProgress();
  } catch (e) {
    Logger.log('cleanupExpiredSubscription error: ' + e.toString());
  }
}

/**
 * Server method used by client polling to check subscription state.
 * Returns structured JSON: { success: boolean, subscribed: boolean, data: { ... } }
 */
function verifySubscriptionStatus(){
  try{
    const response = validateUserSession();
    if( response && response.success === true && response.result && response.result.data && response.result.data.isSubscribed === true ){
      // mark subscription step completed
      const meta = {
        status: 'active',
        activatedAt: new Date().toISOString()
      };
      markSetupStepCompleted('subscription', meta);
      
      return { success: true, subscribed: true, data: response.result.data };
    }else{
      return { success: true, subscribed: false, data: response.result ? response.result.data : null };
    }
  }catch(e){
    Logger.log('verifySubscriptionStatus error: ' + e.toString());
    return { success: false, subscribed: false, error: e.toString() };
  }
}

/**
 * Copies any missing template sheets from the source spreadsheet into the user's spreadsheet.
 * Then applies all formulas in the correct dependency order and populates reports.
 */
function installTemplateInitialSetup() {
  try {
    var response = getAppSettings();
    if (response.success !== true) {
      return { success: false, message: 'Unable to load app settings. Please try again.' };
    }

    var spreadsheetTemplateUrl = response.result.spreadsheetTemplateUrl;
    var sourceSpreadsheet = SpreadsheetApp.openByUrl(spreadsheetTemplateUrl);
    var userSpreadsheet = UserSpreadsheet;
    var requiredSheets = appBaseTemplates();
    var sourceSheets = sourceSpreadsheet.getSheets();
    var requiredSet = new Set(requiredSheets.map(function(name) { return name.toLowerCase(); }));
    var sheetsInstalled = 0;

    sourceSheets.forEach(function(sourceSheet) {
      var sheetName = sourceSheet.getName();
      if (!requiredSet.has(sheetName.toLowerCase())) return;
      if (!userSpreadsheet.getSheetByName(sheetName)) {
        var copiedSheet = sourceSheet.copyTo(userSpreadsheet);
        copiedSheet.setName(sheetName);
        sheetsInstalled++;
      }
    });

    // Apply all formulas in correct dependency order
    reApplyAllFormulas();

    populateNetWorth();
    populateJointNetWorth();

    // Activate Start Here sheet (guard: sheet may not exist yet in edge cases)
    var startSheet = userSpreadsheet.getSheetByName(USER_START_HERE_SHEET);
    if (startSheet) {
      startSheet.activate();
    }
    SpreadsheetApp.flush();

    // Mark template step completed
    try {
      markSetupStepCompleted('template', { status: 'completed', installedAt: new Date().toISOString() });
    } catch (e) {
      Logger.log('Error marking template step completed: ' + e.toString());
    }

    return { success: true, message: 'Template(s) installed successfully' };
  } catch (error) {
    Logger.log('installTemplateInitialSetup error: ' + error.message);
    return { success: false, message: 'Something went wrong, please try again.' };
  }
}

/**
 * Saves user configuration (auto-sync preference) and ensures all triggers are correctly set up.
 * Sets the AUTO_SYNC_STATUS property first, then calls setupInstallableTrigger() once
 * which reads the property to decide whether to create the daily trigger.
 */
function saveConfiguration(data) {
  try {
    // 1. Persist the auto-sync preference BEFORE setting up triggers
    PropertiesService.getUserProperties().setProperty('AUTO_SYNC_STATUS', data.autoSync === true ? 'true' : 'false');

    // 2. Rebuild all triggers (reads AUTO_SYNC_STATUS internally)
    setupInstallableTrigger();

    // 3. Mark config step completed
    try {
      markSetupStepCompleted('config', { status: 'completed', autoSync: data.autoSync, savedAt: new Date().toISOString() });
    } catch (e) {
      Logger.log('markSetupStepCompleted error in saveConfiguration: ' + e.toString());
    }
    return {
      success: true,
      message: 'Saved Successfully'
    };
  } catch (error) {
    Logger.log('Error while saveConfiguration: ' + error.message);
    return {
      success: false,
      message: 'Something went wrong, please try again.'
    };
  }
}

/**
 * Creates all installable triggers from scratch.
 * Removes any existing duplicates first, then creates onEdit, onOpen,
 * and (conditionally) the daily auto-sync trigger.
 * Called during onInstall and saveConfiguration.
 */
function setupInstallableTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Remove all existing installable triggers managed by this add-on
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    var handler = trigger.getHandlerFunction();
    if (handler === installableTriggers.onOpen ||
        handler === installableTriggers.onEdit ||
        handler === installableTriggers.dailyAutoSync) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 2. Create the installable onEdit trigger
  ScriptApp.newTrigger(installableTriggers.onEdit)
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // 3. Create the installable onOpen trigger
  ScriptApp.newTrigger(installableTriggers.onOpen)
    .forSpreadsheet(ss)
    .onOpen()
    .create();

  // 4. Create the daily auto-sync trigger ONLY if auto-sync is enabled
  var syncEnabled = PropertiesService.getUserProperties().getProperty('AUTO_SYNC_STATUS') === 'true';
  if (syncEnabled) {
    ScriptApp.newTrigger(installableTriggers.dailyAutoSync)
      .timeBased()
      .everyDays(1)
      .atHour(6)
      .create();
  }
}

/**
 * Self-healing trigger check. Verifies that all required installable triggers exist
 * and recreates any that are missing. Safe to call on every open.
 * Unlike setupInstallableTrigger(), this does NOT delete existing triggers —
 * it only adds missing ones to avoid unnecessary trigger churn.
 */
function ensureTriggersExist() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var triggers = ScriptApp.getProjectTriggers();
    var existingHandlers = {};
    triggers.forEach(function(trigger) {
      existingHandlers[trigger.getHandlerFunction()] = true;
    });

    // Ensure onEdit trigger
    if (!existingHandlers[installableTriggers.onEdit]) {
      ScriptApp.newTrigger(installableTriggers.onEdit)
        .forSpreadsheet(ss)
        .onEdit()
        .create();
      Logger.log('[ensureTriggersExist] Recreated missing onEdit trigger');
    }

    // Ensure onOpen trigger
    if (!existingHandlers[installableTriggers.onOpen]) {
      ScriptApp.newTrigger(installableTriggers.onOpen)
        .forSpreadsheet(ss)
        .onOpen()
        .create();
      Logger.log('[ensureTriggersExist] Recreated missing onOpen trigger');
    }

    // Ensure dailyAutoSync trigger (only if enabled)
    var syncEnabled = PropertiesService.getUserProperties().getProperty('AUTO_SYNC_STATUS') === 'true';
    if (syncEnabled && !existingHandlers[installableTriggers.dailyAutoSync]) {
      ScriptApp.newTrigger(installableTriggers.dailyAutoSync)
        .timeBased()
        .everyDays(1)
        .atHour(6)
        .create();
      Logger.log('[ensureTriggersExist] Recreated missing dailyAutoSync trigger');
    }
  } catch (e) {
    Logger.log('ensureTriggersExist error: ' + e.toString());
  }
}

/**
 * Toggles the daily auto-sync on or off.
 * Persists the preference and manages only the dailyAutoSync trigger
 * without disturbing onEdit/onOpen triggers.
 */
function toggleAutoSyncSetting(status) {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    if (status === true) {
      // Remove existing auto-sync triggers to avoid duplicates
      triggers.forEach(function(trigger) {
        if (trigger.getHandlerFunction() === installableTriggers.dailyAutoSync) {
          ScriptApp.deleteTrigger(trigger);
        }
      });

      // Create a trigger that fires daily at 6 AM user's local time
      ScriptApp.newTrigger(installableTriggers.dailyAutoSync)
        .timeBased()
        .everyDays(1)
        .atHour(6)
        .create();

      PropertiesService.getUserProperties().setProperty('AUTO_SYNC_STATUS', 'true');
    } else {
      triggers.forEach(function(trigger) {
        if (trigger.getHandlerFunction() === installableTriggers.dailyAutoSync) {
          ScriptApp.deleteTrigger(trigger);
        }
      });
      PropertiesService.getUserProperties().setProperty('AUTO_SYNC_STATUS', 'false');
    }
    return {
      success: true,
      message: 'Auto-Sync status updated successfully.'
    };
  } catch (error) {
    Logger.log('Error while toggleAutoSyncSetting: ' + error.message);
    return {
      success: false,
      message: 'Something went wrong, please try again.'
    };
  }
}

/**
 * Full system repair — called from the Settings "Repair System" button.
 * Reinstalls missing templates, rebuilds all triggers, re-applies formulas,
 * and populates reports. Safe to run on live accounts — only adds what's missing,
 * never deletes user data.
 */
function repairSystem() {
  var steps = [];
  try {
    // 1. Reinstall any missing template sheets
    var templateResult = installTemplateInitialSetup();
    steps.push('Templates: ' + (templateResult.success ? 'OK' : 'FAILED'));

    // 2. Rebuild all installable triggers (clean slate)
    setupInstallableTrigger();
    steps.push('Triggers: OK');

    // 3. Re-apply all formulas in correct dependency order
    reApplyAllFormulas();
    steps.push('Formulas: OK');

    // 4. Populate reports
    populateNetWorth();
    populateJointNetWorth();
    steps.push('Reports: OK');

    SpreadsheetApp.flush();

    return {
      success: true,
      message: 'System repair completed successfully.',
      details: steps
    };
  } catch (e) {
    Logger.log('repairSystem error: ' + e.toString());
    steps.push('Error: ' + e.message);
    return {
      success: false,
      message: 'Repair encountered an error: ' + e.message,
      details: steps
    };
  }
}

function getConnectedPlaidAccountsTemplate(){
  const response = getAppPlaidConnectedAccounts();
  if( response.success === true ){
    const template = HtmlService.createTemplateFromFile('AccountListCard');
    template.accounts = response.result;
    return template.evaluate().getContent();
  }else{
    const template = HtmlService.createTemplateFromFile('Error');
    template.message = "No data available right now, please try again!";
    return template.evaluate().getContent();
  }
}

function getAccountDetailsTemplate( accountId ){
  const response = getAppPlaidAccountById(accountId);
  if( response.success === true ){
    const template = HtmlService.createTemplateFromFile('AccountDetailCard');
    template.account = response.result;
    return template.evaluate().getContent();
  }else{
    const template = HtmlService.createTemplateFromFile('Error');
    template.message = "No data available right now, please try again!";
    return template.evaluate().getContent();
  }
}

function getRenameAccountTemplate( accountId, accountName ){
  const template = HtmlService.createTemplateFromFile('RenameAccountCard');
  template.accountId = accountId;
  template.currentName = accountName;
  return template.evaluate().getContent();
}

function runThefinUPlaidAutoSync(){
  var syncStartTime = new Date().getTime();
  // 25-minute safety limit (GAS max is 30 min)
  var MAX_RUNTIME_MS = 25 * 60 * 1000;

  try{
    Logger.log('[AUTO-SYNC] Started at ' + new Date().toISOString());
    Logger.log('[AUTO-SYNC] UserEmail: ' + (UserEmail || '(empty)'));

    if (!UserEmail) {
      Logger.log('[AUTO-SYNC] ABORTED — No user email available.');
      sendUserNotification(
        'TheFinU Auto-Sync Failed — Email Not Available',
        'Hi,\n\nYour daily auto-sync could not run because your email could not be determined.\n\nPlease open the TheFinU add-on sidebar once to fix this. Auto-sync will resume on the next scheduled run.\n\nBest,\nTheFinU'
      );
      return false;
    }

    let isSyncEnabled = PropertiesService.getUserProperties().getProperty("AUTO_SYNC_STATUS");
    if( isSyncEnabled !== 'true' ){
      Logger.log('[AUTO-SYNC] Sync is disabled, skipping.');
      return false;
    }

    // Validate required sheets before syncing
    var sheetCheck = validateRequiredSheets(appBaseTemplates());
    if(!sheetCheck.valid){
      var missingList = sheetCheck.missing.join(', ');
      Logger.log('[AUTO-SYNC] Aborted — missing sheets: ' + missingList);
      sendUserNotification(
        'TheFinU Auto-Sync Failed — Missing Sheets',
        'Hi,\n\nYour daily auto-sync could not run because the following required sheets are missing from your spreadsheet:\n\n' +
        missingList +
        '\n\nTo fix this, open TheFinU add-on sidebar, go to Settings, and click "Reset Templates" to restore the missing sheets.\n\nOnce restored, auto-sync will resume on the next scheduled run.\n\nBest,\nTheFinU'
      );
      return false;
    }

    const response = getAppPlaidConnectedAccounts();
    Logger.log('[AUTO-SYNC] getAppPlaidConnectedAccounts success: ' + response.success);
    var syncedCount = 0;
    var failedAccounts = [];
    if( response.success === true ){
      let accounts = response.result;
      Logger.log('[AUTO-SYNC] Total accounts: ' + (accounts ? accounts.length : 0));
      if( accounts && accounts.length > 0 ){
        for (var ai = 0; ai < accounts.length; ai++) {
          var account = accounts[ai];
          // Safety timeout check before each account
          if (new Date().getTime() - syncStartTime > MAX_RUNTIME_MS) {
            Logger.log('[AUTO-SYNC] Approaching 30-min limit, stopping after ' + syncedCount + ' accounts.');
            break;
          }
          // Only sync accounts that are linked and active
          if (account.is_linked !== true || account.status !== true) {
            Logger.log('[AUTO-SYNC] Skipped account ' + account.account_id + ' (is_linked: ' + account.is_linked + ', status: ' + account.status + ')');
            continue;
          }
          var account_id = account.account_id;
          Logger.log('[AUTO-SYNC] Syncing account: ' + account_id + ' (is_update: ' + account.is_update + ')');
          try {
            updateTransactionSheet(account_id);
            Logger.log('[AUTO-SYNC] Transactions synced for: ' + account_id);
          } catch (txErr) {
            Logger.log('[AUTO-SYNC] Transaction sync FAILED for ' + account_id + ': ' + txErr.message);
            failedAccounts.push(account_id + ' (transactions)');
          }
          try {
            var support_response = checkItemProductSupport(account_id, 'investments');
            if (support_response === true) {
              updateInvestmentSheet(account_id);
              Logger.log('[AUTO-SYNC] Investments synced for: ' + account_id);
            }
          } catch (invErr) {
            Logger.log('[AUTO-SYNC] Investment sync FAILED for ' + account_id + ': ' + invErr.message);
            failedAccounts.push(account_id + ' (investments)');
          }
          try {
            updateAccountBalanceHistory(account_id);
            Logger.log('[AUTO-SYNC] Balance history synced for: ' + account_id);
          } catch (balErr) {
            Logger.log('[AUTO-SYNC] Balance sync FAILED for ' + account_id + ': ' + balErr.message);
            failedAccounts.push(account_id + ' (balance)');
          }
          // Reset the is_update flag
          updateAppAccountDetailById(account_id, { is_update: false });
          syncedCount++;
          Logger.log('[AUTO-SYNC] Account ' + account_id + ' completed (' + syncedCount + '/' + accounts.length + ')');
        }

        // Refresh formulas and budget dropdowns if any accounts were synced
        if (syncedCount > 0) {
          Logger.log('[AUTO-SYNC] Reapplying formulas...');
          reApplyAllFormulas();
        }

        populateNetWorth();
        populateJointNetWorth();
      } else {
        Logger.log('[AUTO-SYNC] No accounts found for this user.');
      }
    } else {
      Logger.log('[AUTO-SYNC] Failed to fetch accounts. Response: ' + JSON.stringify(response));
    }

    var elapsed = Math.round((new Date().getTime() - syncStartTime) / 1000);
    Logger.log('[AUTO-SYNC] Completed. Synced ' + syncedCount + ' accounts in ' + elapsed + 's. Failures: ' + failedAccounts.length);
    var notifBody = 'Hi,\n\nYour daily auto-sync completed.\n\n' +
      'Accounts synced: ' + syncedCount + '\n' +
      'Time taken: ' + elapsed + 's\n' +
      'Date: ' + new Date().toLocaleString();
    if (failedAccounts.length > 0) {
      notifBody += '\n\nFailed: ' + failedAccounts.join(', ');
    }
    notifBody += '\n\nBest,\nTheFinU';
    sendUserNotification(
      syncedCount > 0 ? 'TheFinU Auto-Sync Completed' : 'TheFinU Auto-Sync — No Data Synced',
      notifBody
    );

    return true;
  }catch(error){
    Logger.log('[AUTO-SYNC] ERROR: ' + error.message);
    sendUserNotification(
      'TheFinU Auto-Sync Failed',
      'Hi,\n\nYour daily auto-sync encountered an error:\n\n' +
      error.message +
      '\n\nPlease open TheFinU add-on sidebar and try a manual sync. If the problem persists, go to Settings and click "Reset Templates".\n\nBest,\nTheFinU'
    );
    return false;
  }
}

/**
 * Backend functions called from the UI
 */
function validateSheetsForSync(){
  var sheetCheck = validateRequiredSheets(appBaseTemplates());
  return sheetCheck;
}

function linkNewAccount() {
  var sheetCheck = validateRequiredSheets(appBaseTemplates());
  if(!sheetCheck.valid){
    Logger.log('Missing sheets detected before linking: ' + sheetCheck.missing.join(', ') + '. Auto-installing...');
    var installResult = installTemplateInitialSetup();
    if(!installResult.success){
      SpreadsheetApp.getUi().alert('Failed to install missing sheets. Please try again or go to Settings and click "Reset Templates".');
      return;
    }
    var recheck = validateRequiredSheets(appBaseTemplates());
    if(!recheck.valid){
      SpreadsheetApp.getUi().alert('Cannot link account — the following required sheets are still missing: ' + recheck.missing.join(', ') + '.\n\nPlease go to Settings and click "Reset Templates" to restore them.');
      return;
    }
    // Populate net worth reports with any existing data on freshly installed sheets
    populateNetWorth();
    populateJointNetWorth();
    SpreadsheetApp.flush();
  }
  const html = HtmlService.createHtmlOutputFromFile('ConnectPlaidAccount').setWidth(450).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Connect Plaid Account");
}

function updateAccountName(accountId, newName) {

  try{
    const response = updateAppAccountDetailById(accountId,{name: newName});
    changeAccountNameOnBalanceHistorySheet(accountId, newName);
    changeAccountNameOnTransactionSheet(accountId, newName);
    changeAccountNameOnInvestmentSheet(accountId, newName);
    populateNetWorth();
    populateJointNetWorth();
    if( response.success === true ){
      return {
        success: true,
        message: 'Account name updated'
      };
    }else{
      return {
        success: false,
        message: 'Something went wrong, please try again.'
      };
    }
  }catch(error){
    Logger.log(`Error while updateAccountName: ${error.message}`);
    return {
      success: false,
      message: "Something went wrong, please try again."
    }
  }
}

function removeAccountFromList(accountId) {
  try{
    var result = SpreadsheetApp.getUi().alert(
      'Remove Account',
      'Do you want to remove the account from the list? If once removed, you cannot see the account from the list. please confirm',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (result == SpreadsheetApp.getUi().Button.YES) {
      const response = updateAppAccountDetailById(accountId,{status: false});
      if( response.success === true ){
        return {
          success: true,
          message: 'Account removed'
        };
      }else{
        return {
          success: false,
          message: 'Something went wrong, please try again.'
        };
      }
    }
  }catch(error){
    Logger.log(`Error while removeAccountFromList: ${error.message}`);
    return {
      success: false,
      message: "Something went wrong, please try again."
    }
  }
}

function confirmLinkAccountToTemplate( accountId ){
  // Show an HTML modal for confirmation instead of using SpreadsheetApp.getUi().alert
  const accountName = getPlaidAccountNameByAccountId(accountId);
  const template = HtmlService.createTemplateFromFile('ConfirmLinkAccount');
  template.accountId = accountId;
  template.accountName = accountName || '';
  const ui = template.evaluate().setWidth(480).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Import Historical Data');
  return true;
}

/**
 * Called by the ConfirmLinkAccount modal when the user confirms.
 * This prepares TASK_STATUS and opens the LinkImportRunner modal.
 */
function confirmLinkAccountToTemplateConfirmed(accountId){
  const userProperties = PropertiesService.getUserProperties();
  try{
    // Install missing sheets before opening the import runner
    var sheetCheck = validateRequiredSheets(appBaseTemplates());
    if(!sheetCheck.valid){
      Logger.log('Missing sheets before import: ' + sheetCheck.missing.join(', ') + '. Auto-installing...');
      var installResult = installTemplateInitialSetup();
      if(!installResult.success){
        return { success: false, error: 'Failed to install missing sheets. Please try again or go to Settings and click "Reset Templates".' };
      }
      // Populate net worth reports with any existing data on freshly installed sheets
      populateNetWorth();
      populateJointNetWorth();
      SpreadsheetApp.flush();
    }

    userProperties.setProperty('TASK_STATUS','READY');
    userProperties.setProperty('LINK_ACCOUNT_ID', accountId);
    // Open the import runner modal
    const html = HtmlService.createTemplateFromFile('LinkImportRunner');
    html.accountId = accountId;
    const ui = html.evaluate().setWidth(480).setHeight(360);
    SpreadsheetApp.getUi().showModalDialog(ui, 'Importing Historical Data');
    return { success: true };
  }catch(e){
    userProperties.setProperty('TASK_STATUS','ERROR: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Finalize linking for an account after data insertion completed by client.
 * Performs template install, balance updates, sheet linking, formulas and net worth population.
 */
function finalizeLink(accountId){
  const lock = LockService.getUserLock();
  lock.waitLock(30000);
  const props = PropertiesService.getUserProperties();
  try{
    if(!accountId) return { success: false, message: 'Missing accountId' };

    var sheetCheck = validateRequiredSheets(appBaseTemplates());
    if(!sheetCheck.valid){
      Logger.log('Missing sheets detected during finalizeLink: ' + sheetCheck.missing.join(', ') + '. Auto-installing...');
      var installResult = installTemplateInitialSetup();
      if(!installResult.success){
        return { success: false, message: 'Failed to install missing sheets. Please try again or go to Settings and click "Reset Templates".' };
      }
      var recheck = validateRequiredSheets(appBaseTemplates());
      if(!recheck.valid){
        return { success: false, message: 'Missing required sheets: ' + recheck.missing.join(', ') + '. Please go to Settings and click "Reset Templates" to restore them.' };
      }
    }

    updateAccountBalanceHistory(accountId);
    linkAccountsSheetData(accountId);
    updateAppAccountDetailById(accountId, {
      is_linked: true,
      status: true,
      updates: false,
      linked_date: getTodayDateTime()
    });

    // Re-apply all formulas so dropdowns, year/month lists, and cross-references update
    reApplyAllFormulas();

    populateNetWorth();
    populateJointNetWorth();
    SpreadsheetApp.flush();

    props.setProperty('TASK_STATUS','COMPLETED');
    props.deleteProperty('LINK_ACCOUNT_ID');

    return { success: true };
  }catch(e){
    props.setProperty('TASK_STATUS','ERROR: ' + e.toString());
    return { success: false, error: e.toString() };
  }finally{
    lock.releaseLock();
  }
}

function processUnlinkAccountTask() {
  const lock = LockService.getUserLock();
  lock.waitLock(30000);

  const props = PropertiesService.getUserProperties();
  const accountId = props.getProperty('UNLINK_ACCOUNT_ID');

  if (!accountId) return;

  try {

    // --- DATA CLEANUP ---
    clearTransactionsData(accountId);
    clearInvestmentsData(accountId);
    clearBalanceHistoryData(accountId);
    clearAccountData(accountId);
    updateAppAccountDetailById(accountId, {
      is_linked: false,
      status: true,
      updates: false,
      linked_date: getTodayDateTime()
    });

    populateNetWorth();
    populateJointNetWorth();
    populateMonthlyBudget();
    populateJointMonthlyBudget();
    populateYearlyBudget();
    populateJointYearlyBudget();

    let definitionSheet = UserSpreadsheet.getSheetByName(USER_DEFINITION_SHEET);
    let transactionSheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
    if( definitionSheet && definitionSheet.getRange("I3").getValue() <= 1 ){  
      definitionSheet.getRange("R2").setValue('2024'); // reset to default year
      definitionSheet.getRange("S2").setValue('1/1/2024'); // clear monthly period list
      definitionSheet.getRange("V1").setValue('2024'); // reset total income
      definitionSheet.getRange("X1").setValue('2024'); // reset total joint income
      transactionSheet.getRange("P1").setValue('Period');
      transactionSheet.getRange("O1").setValue('Type');
      transactionSheet.getRange("N1").setValue('Group');
    }

    SpreadsheetApp.flush();
    props.setProperty('TASK_STATUS', 'COMPLETED');
    props.deleteProperty('UNLINK_ACCOUNT_ID');
    return { success: true };
  } catch (e) {
    props.setProperty('TASK_STATUS', 'ERROR: ' + e.message);
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function resetTaskStatus(){
  const props = PropertiesService.getUserProperties();
  props.deleteProperty('TASK_STATUS');
  props.deleteProperty('LINK_ACCOUNT_ID');
}


function confirmUnlinkAccountFromTemplate(accountId){
  // Show an HTML modal for confirmation instead of using SpreadsheetApp.getUi().alert
  const accountName = getPlaidAccountNameByAccountId(accountId);
  const template = HtmlService.createTemplateFromFile('ConfirmUnlinkAccount');
  template.accountId = accountId;
  template.accountName = accountName || '';
  const ui = template.evaluate().setWidth(480).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Remove existing data?');
  return true;
}

/**
 * Called by the ConfirmUnlinkAccount modal when the user confirms.
 * This prepares TASK_STATUS and opens the UnlinkRunner modal.
 */
function confirmUnlinkAccountFromTemplateConfirmed(accountId){
  const userProperties = PropertiesService.getUserProperties();
  try{
    userProperties.setProperty('TASK_STATUS','READY');
    userProperties.setProperty('UNLINK_ACCOUNT_ID', accountId);
    // Open the unlink runner modal
    const html = HtmlService.createTemplateFromFile('UnlinkRunner');
    html.accountId = accountId;
    const ui = html.evaluate().setWidth(480).setHeight(320);
    SpreadsheetApp.getUi().showModalDialog(ui, 'Unlink Account');
    return { success: true };
  }catch(e){
    userProperties.setProperty('TASK_STATUS','ERROR: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function reApplyFormulaToSpreadsheet(item){
  const spreadsheet = UserSpreadsheet;
  switch(item){
    case USER_TRANSACTIONS_SHEET:
      if (spreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET)) {
        spreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET).getRange("P1").setFormula('=ARRAYFORMULA({"Period";EoMonth(Indirect("B2:B"&Definition!I3),-1)+1})'); // Set the formula
        spreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET).getRange("O1").setFormula('=ARRAYFORMULA({"Type";iferror(vlookup(INDIRECT("d2:d"&Definition!I3),Indirect(Definition!P2),Definition!C7,0),"Expense")})');
        spreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET).getRange("N1").setFormula('=ARRAYFORMULA({"Group";iferror(vlookup(INDIRECT("d2:d"&Definition!I3),Indirect(Definition!P2),Definition!C6,0),"NotGrouped")})');
      }
      break;
    case USER_DEFINITION_SHEET:
      if (spreadsheet.getSheetByName(USER_DEFINITION_SHEET)) {
        // Compute years and periods directly from Transactions (avoids INDIRECT timing issues)
        let defYears = getUniqueTransactionYears_(spreadsheet);
        let defPeriods = getUniqueTransactionPeriods_(spreadsheet);
        writeYearsToDefinition_(spreadsheet, defYears);
        writePeriodsToDefinition_(spreadsheet, defPeriods);
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("V1").setFormula("='Yearly Budget'!E2");
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("X1").setFormula("='Joint Yearly Budget'!D2");
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AC2").setFormula("='Yearly Budget'!E2");
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AC4").setFormula("='Monthly Budget'!C2");
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AD2").setFormula("='Joint Yearly Budget'!D2");
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AD4").setFormula("='Joint Monthly Budget'!C2");
      }
    break;
    case USER_BUDGET_MAKER_SHEET:
      if (spreadsheet.getSheetByName(USER_BUDGET_MAKER_SHEET)) {
        spreadsheet.getSheetByName(USER_BUDGET_MAKER_SHEET).getRange("D7").setFormula('=ARRAYFORMULA(if(Indirect("$E$7:$E$"&Definition!M13)="","",round(Indirect("$E$7:$E$"&Definition!M13)/12,2)))'); // Set the formula
        spreadsheet.getSheetByName(USER_BUDGET_MAKER_SHEET).getRange("F7").setFormula('=ARRAYFORMULA(if(Indirect("J$7:$J$"&Definition!M13)=0,"",Indirect("J$7:$J$"&Definition!M13)*$J$2/$J$49))'); // Set the formula
        spreadsheet.getSheetByName(USER_BUDGET_MAKER_SHEET).getRange("G7").setFormula('=ArrayFormula(If(Indirect("F$7:$F$"&Definition!M13)="","",Indirect("E$7:$E$"&Definition!M13)-Indirect("F$7:$F$"&Definition!M13)))'); // Set the formula
      }
    break;
    case USER_MONTHLY_BUDGET_SHEET:
      if (spreadsheet.getSheetByName(USER_MONTHLY_BUDGET_SHEET)) {
        let cell = spreadsheet.getSheetByName(USER_MONTHLY_BUDGET_SHEET).getRange("C2");
        cell.clearDataValidations();
        cell.clearContent();
        // Read transaction dates directly from Transactions sheet (source of truth)
        let periodValues = getUniqueTransactionPeriods_(spreadsheet);
        if (periodValues.length > 0) {
          // Write dates to Definition S column for formula references
          writePeriodsToDefinition_(spreadsheet, periodValues);
          SpreadsheetApp.flush();
          // Build "MMM yyyy" string list for the dropdown so it shows "Jan 2026" not datetime
          var tz = spreadsheet.getSpreadsheetTimeZone();
          var periodLabels = periodValues.map(function(d) {
            return Utilities.formatDate(d, tz, "MMM yyyy");
          });
          var rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(periodLabels, true)
            .setAllowInvalid(false)
            .build();
          cell.setDataValidation(rule);
          cell.setNumberFormat("MMM yyyy");
          cell.setValue(periodLabels[periodLabels.length - 1]); // Default to most recent month
        }
        spreadsheet.getSheetByName(USER_MONTHLY_BUDGET_SHEET).getRange("E4").setFormula('=Definition!AC24');
        spreadsheet.getSheetByName(USER_MONTHLY_BUDGET_SHEET).getRange("B5").setFormula('=Definition!AC5');
      }
    break;
    case USER_JOINT_MONTHLY_BUDGET_SHEET:
      if (spreadsheet.getSheetByName(USER_JOINT_MONTHLY_BUDGET_SHEET)) {
        let cell = spreadsheet.getSheetByName(USER_JOINT_MONTHLY_BUDGET_SHEET).getRange("C2");
        cell.clearDataValidations();
        cell.clearContent();
        let periodValues = getUniqueTransactionPeriods_(spreadsheet);
        if (periodValues.length > 0) {
          writePeriodsToDefinition_(spreadsheet, periodValues);
          SpreadsheetApp.flush();
          var tz = spreadsheet.getSpreadsheetTimeZone();
          var periodLabels = periodValues.map(function(d) {
            return Utilities.formatDate(d, tz, "MMM yyyy");
          });
          var rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(periodLabels, true)
            .setAllowInvalid(false)
            .build();
          cell.setDataValidation(rule);
          cell.setNumberFormat("MMM yyyy");
          cell.setValue(periodLabels[periodLabels.length - 1]);
        }
        spreadsheet.getSheetByName(USER_JOINT_MONTHLY_BUDGET_SHEET).getRange("E4").setFormula('=Definition!AD24');
        spreadsheet.getSheetByName(USER_JOINT_MONTHLY_BUDGET_SHEET).getRange("B5").setFormula('=Definition!AD5');
      }
    break;
    case USER_YEARLY_BUDGET_SHEET:
      if (spreadsheet.getSheetByName(USER_YEARLY_BUDGET_SHEET)) {
        let cell = spreadsheet.getSheetByName(USER_YEARLY_BUDGET_SHEET).getRange("E2");
        cell.clearDataValidations();
        cell.clearContent();
        // Read transaction years directly from Transactions sheet (source of truth)
        let yearValues = getUniqueTransactionYears_(spreadsheet);
        if (yearValues.length > 0) {
          // Write years to Definition R column so data validation range works
          writeYearsToDefinition_(spreadsheet, yearValues);
          SpreadsheetApp.flush();
          let yearRange = spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("R2:R" + (yearValues.length + 1));
          var rule = SpreadsheetApp.newDataValidation()
            .requireValueInRange(yearRange, true)
            .setAllowInvalid(false)
            .build();
          cell.setDataValidation(rule);
          cell.setValue(yearValues[yearValues.length - 1]); // Default to most recent year
        }
        spreadsheet.getSheetByName(USER_YEARLY_BUDGET_SHEET).getRange("B6:D6").setFormula('=Definition!AC3');
      }
    break;
    case USER_JOINT_YEARLY_BUDGET_SHEET:
      if (spreadsheet.getSheetByName(USER_JOINT_YEARLY_BUDGET_SHEET)) {
        let cell = spreadsheet.getSheetByName(USER_JOINT_YEARLY_BUDGET_SHEET).getRange("D2");
        cell.clearDataValidations();
        cell.clearContent();
        let yearValues = getUniqueTransactionYears_(spreadsheet);
        if (yearValues.length > 0) {
          writeYearsToDefinition_(spreadsheet, yearValues);
          SpreadsheetApp.flush();
          let yearRange = spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("R2:R" + (yearValues.length + 1));
          var rule = SpreadsheetApp.newDataValidation()
            .requireValueInRange(yearRange, true)
            .setAllowInvalid(false)
            .build();
          cell.setDataValidation(rule);
          cell.setValue(yearValues[yearValues.length - 1]);
        }
        spreadsheet.getSheetByName(USER_JOINT_YEARLY_BUDGET_SHEET).getRange("B5:C5").setFormula('=Definition!AD3');
      }
    break;
  }
}

/**
 * Parses a cell value into a Date object.
 * Handles: Date objects from getValues(), string dates like "2024-01-15" from Plaid,
 * and numeric serial dates from Google Sheets.
 * Returns null if the value cannot be parsed into a valid date.
 */
function parseTransactionDate_(val) {
  if (!val) return null;
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }
  if (typeof val === 'string') {
    // Plaid returns "YYYY-MM-DD" format
    var parts = val.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (parts) {
      var d = new Date(parseInt(parts[1], 10), parseInt(parts[2], 10) - 1, parseInt(parts[3], 10));
      return isNaN(d.getTime()) ? null : d;
    }
    // Try other date formats (MM/DD/YYYY, etc.)
    var parsed = new Date(val);
    return isNaN(parsed.getTime()) ? null : parsed;
  }
  if (typeof val === 'number' && val > 0) {
    // Google Sheets serial date number (days since Dec 30, 1899)
    var d = new Date(1899, 11, 30 + val);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

/**
 * Reads transaction dates from the Transactions sheet and returns unique years sorted ascending.
 * Handles both Date objects and string dates ("2024-01-15" from Plaid).
 */
function getUniqueTransactionYears_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var dates = sheet.getRange("B2:B" + lastRow).getValues().flat();
  var yearSet = {};
  var maxYear = new Date().getFullYear() + 1;
  dates.forEach(function(val) {
    var d = parseTransactionDate_(val);
    if (d) {
      var y = d.getFullYear();
      if (y >= 1900 && y <= maxYear) yearSet[y] = true;
    }
  });
  return Object.keys(yearSet).map(Number).sort(function(a, b) { return a - b; });
}

/**
 * Reads transaction dates from the Transactions sheet and returns unique month-start dates sorted ascending.
 * Handles both Date objects and string dates ("2024-01-15" from Plaid).
 */
function getUniqueTransactionPeriods_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var dates = sheet.getRange("B2:B" + lastRow).getValues().flat();
  var periodSet = {};
  var maxYear = new Date().getFullYear() + 1;
  dates.forEach(function(val) {
    var d = parseTransactionDate_(val);
    if (d && d.getFullYear() >= 1900 && d.getFullYear() <= maxYear) {
      var key = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
      if (!periodSet[key]) {
        periodSet[key] = new Date(d.getFullYear(), d.getMonth(), 1);
      }
    }
  });
  return Object.keys(periodSet).sort().map(function(k) { return periodSet[k]; });
}

/**
 * Writes computed year values into Definition R column so data validation ranges stay valid.
 */
function writeYearsToDefinition_(spreadsheet, yearValues) {
  var defSheet = spreadsheet.getSheetByName(USER_DEFINITION_SHEET);
  if (!defSheet) return;
  // Clear old year values in R column (R2 onwards)
  var clearRange = defSheet.getRange("R2:R1000");
  clearRange.clearContent();
  // Write new values
  if (yearValues.length > 0) {
    var data = yearValues.map(function(y) { return [y]; });
    var range = defSheet.getRange(2, 18, data.length, 1); // Column R = 18
    range.setValues(data);
    range.setNumberFormat("0"); // Plain number format for years
  }
}

/**
 * Writes computed period dates into Definition S column so data validation ranges stay valid.
 */
function writePeriodsToDefinition_(spreadsheet, periodValues) {
  var defSheet = spreadsheet.getSheetByName(USER_DEFINITION_SHEET);
  if (!defSheet) return;
  // Clear old period values in S column (S2 onwards)
  var clearRange = defSheet.getRange("S2:S1000");
  clearRange.clearContent();
  // Set format BEFORE writing values so Sheets stores clean dates without time drift
  if (periodValues.length > 0) {
    var data = periodValues.map(function(d) { return [d]; });
    var range = defSheet.getRange(2, 19, data.length, 1); // Column S = 19
    range.setNumberFormat("MMM yyyy");
    range.setValues(data);
  }
}

function hideSheetByName( sheetName ){
  const sheet = UserSpreadsheet.getSheetByName(sheetName);
  if(sheet){
    sheet.hideSheet();
  }
}

function protectSheetByName(sheetName){
  const sheet = UserSpreadsheet.getSheetByName(sheetName);
  if (sheet) {
    const protection = sheet.protect();
    protection.setDescription(`Protected: ${sheetName}`);
    protection.setWarningOnly(false); // Prevent editing
  }
}

function getPlaidAccountNameByAccountId(account_id){
  try{
    const response = getAppPlaidAccountById(account_id);
    //Logger.log( JSON.stringify(response, null, 2) );
    if( response.success === true ){
      return response.result.name;
    }else{
      return null;
    }
  }catch(error){
    Logger.log('getPlaidAccountNameByAccountId error: ' + error.message);
    return null;
  }
}

function checkItemProductSupport( account_id, product){

  try{
    const response = getAppPlaidAccountById(account_id);
    if( response.success === true ){
      let item = getPlaidItem( response.result.access_token );
      if( item ){
        let products = item.products;
        if (products.includes(product)) {
          return true;
        }
        return false;
      }
    }else{
      return false;
    }
  }catch(error){
    Logger.log(`Error while checkItemProductSupport: ${error.message}`);
    return false;
  }
}

function updatePlaidAccountIDOnSheets( institution_id, accounts ){
  const response = getAppPlaidConnectedAccounts();
  if( response.success === true ){
    let dbAccounts = response.result;
    if( dbAccounts.length > 0 && accounts.length > 0 ){
      accounts.forEach( function(account){
        dbAccounts.find( item => {
          if( item.account_name === account.name && item.mask === account.mask && item.institution_id === institution_id ) {
            updateAccountIdOnBalanceHistorySheet(item.account_id, account.id);
            updateAccountIdOnTransactionSheet(item.account_id, account.id);
            updateAccountIdOnInvestmentSheet(item.account_id, account.id);
          }
        });
      });
    }
  }
}

function getAccountDataByAccountId(accountId){
  const sheet = UserSpreadsheet.getSheetByName(USER_ACCOUNTS_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let row = 1; row < data.length; row++) {
    if( data[row].includes(accountId) ){
      return JSON.parse( JSON.stringify( data[row] ) );
    }
  }
  return null;
}

/**
 * Step 4: Polling function called by the UI to check the flag status
 */
function checkTaskStatus() {
  const status = PropertiesService.getUserProperties().getProperty('TASK_STATUS');
  return status;
}

function showAddBalanceHistoryFormTemplate(){
  const template = HtmlService.createTemplateFromFile('AddBalanceHistoryCard');
  return template.evaluate().getContent();
}

function showExistingAccountFormFieldsTemplate(){
  let sheet = UserSpreadsheet.getSheetByName(USER_ACCOUNTS_SHEET);
  let lastrow = sheet.getLastRow() + 1;
  let accounts = [];

  // Account Name in column B & Account ID in last column
  // break if both value is empty
  for( let i = 1; i <= lastrow; i++ ){
    let accountName = sheet.getRange(i + 1, 2).getValue();
    let accountId = sheet.getRange(i + 1, sheet.getLastColumn()).getValue();
    if( accountName === '' && accountId === '' ){
      break;
    }
    if( accountName !== '' && accountId !== '' ){
      accounts.push({
        name: accountName,
        id: accountId
      });
    }
  }
  const template = HtmlService.createTemplateFromFile('ExistingAccountFields');
  template.accounts = accounts;
  return template.evaluate().getContent();
}

function showManualAccountFormFieldsTemplate(){
  const template = HtmlService.createTemplateFromFile('ManualAccountFormFields');
  return template.evaluate().getContent();
}

function submitBalanceHistoryFormData(formData){
  try{
    let response = addManualAccountBalanceHistoryData( formData );
    return response;
  }
  catch(error){
    Logger.log(`Error while submitBalanceHistoryFormData: ${error.message}`);
    return {
      success: false,
      message: "Something went wrong, please try again."
    }
  }
}

function addManualAccountBalanceHistoryData( data ){
  try{
    let sheet = UserSpreadsheet.getSheetByName(USER_BALANCE_HISTORY_SHEET);
    let lastrow = sheet.getLastRow() + 1;
    // change date format to mm/dd/yyyy
    data.balanceDate = formatDateToMMDDYYYY( data.balanceDate );
    let accountName = '';
    let accountId = '';
    let accountNumber = '';
    if( data.accountType === 'existing' ){
      let accountData = getAccountDataByAccountId( data.accountName );
      if( accountData ){
        accountName = accountData[1];
        accountId = data.accountName;
        accountNumber = accountData[2];
      }
    }else{
      accountName = data.accountName;
      accountId = generateUniqueId();
    }
    var rowValues = [
      '',                     // col 1 (empty)
      data.balanceDate,       // col 2 - Date
      accountName,            // col 3 - Account Name
      accountNumber,          // col 4 - Account Number
      data.accountBalance,    // col 5 - Balance
      '',                     // col 6 (empty)
      accountId,              // col 7 - Account ID
      getTodayDateTime()      // col 8 - Date & Time
    ];
    var cell = sheet.getRange(lastrow, 1, 1, rowValues.length);
    cell.setValues([rowValues])
        .setFontSize(9)
        .setFontFamily("Comfortaa")
        .setFontColor("#000000")
        .setFontWeight("bold");
    if( data.accountType === 'manual' ){
      let accountSheet = UserSpreadsheet.getSheetByName(USER_ACCOUNTS_SHEET);
      let accountLastRow = accountSheet.getLastRow() + 1;
       for( let i = 1; i <= accountLastRow; i++ ){
        if( accountSheet.getRange(i + 1, 12).getValue() === ''){
          accountSheet.getRange(i + 1, 12).setValue(accountId); // account id
          break;
        }
      }
    }
    const range = sheet.getDataRange(); 
    range.sort({ column: 2, ascending: false });
    populateNetWorth();
    populateJointNetWorth();
    return {
      success: true,
      message: "Balance history added successfully."
    }
  }catch(error){
    Logger.log(`Error while addManualAccountBalanceHistoryData: ${error.message}`);
    return {
      success: false,
      message: "Something went wrong, please try again."
    }
  }
}

function generateUniqueId(){
  let timestamp = new Date().getTime().toString(36);
  let randomNum = Math.floor(Math.random() * 1e8).toString(36);
  return timestamp + randomNum;
}

function formatDateToMMDDYYYY(dateString){
  let date = new Date(dateString);
  // Use "getUTC" methods to prevent timezone shifting
  let month = (date.getUTCMonth() + 1).toString().padStart(2, '0');
  let day = date.getUTCDate().toString().padStart(2, '0');
  let year = date.getUTCFullYear();
  return `${month}/${day}/${year}`;
}


/**
 * Clears every single key-value pair in the UserProperties store.
 * Warning: This is irreversible.
 */
function clearAllUserProperties() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    
    // Get keys before deleting (for logging purposes)
    const keys = userProperties.getKeys();
    
    // Perform the wipe
    userProperties.deleteAllProperties();
    
    return {
      success: true,
      message: `Successfully cleared ${keys.length} data points. The app has been reset.`
    };
  } catch (e) {
    return {
      success: false,
      message: "Failed to clear data: " + e.message
    };
  }
}


function deleteAllSheetsAndRecreate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // Create a temporary sheet
  const tempSheet = ss.insertSheet('Temp');

  // Delete all existing sheets
  sheets.forEach(sheet => ss.deleteSheet(sheet));

  // Rename temp sheet
  tempSheet.setName('Sheet1');
}


// External API helpers removed — using client-driven paginated fetch and server-side safe inserts.

function batchInsertRows(type, rows){
  if(!rows || rows.length === 0) return true;
  var batchSize = 100;
  for(var i = 0; i < rows.length; i += batchSize){
    var slice = rows.slice(i, i + batchSize);
    if(type === 'transactions'){
      insertTransactionsData(slice);
    }else if(type === 'investments'){
      insertInvestmentsData(slice);
    }
  }
  return true;
}

// Helper: build a normalized header -> index map from a headers array
function buildHeaderIndexMapFromArray(headers){
  const map = {};
  headers.forEach(function(h, i){
    const k = (h || '').toString().trim().toLowerCase();
    map[k] = i;
  });
  return map;
}

function findHeaderIndexByKeywords(map, keywords){
  keywords = Array.isArray(keywords) ? keywords : [keywords];
  for(const k in map){
    const ok = keywords.every(function(kw){ return k.indexOf(kw) !== -1; });
    if(ok) return map[k];
  }
  return -1;
}

// ---------------------
// Client-driven paginated fetch and safe insert helpers
// ---------------------

/**
 * Optimized version of prepareLinkPayloadPage to avoid "Limit Exceeded" errors
 * by stripping unnecessary Plaid metadata before returning to the UI.
 */
function prepareLinkPayloadPage(accountId, next_cursor) {
  try {
    // 1. Fetch raw data from Plaid
    var transactions = getPlaidTransactionSyncData(accountId, next_cursor || '');
    
    if (!transactions || transactions.request_id == '') {
      return { success: false, error: 'no_data' };
    }

    // 2. Initialize the result object
    var result = {
      success: true,
      next_cursor: transactions.next_cursor || '',
      has_more: !!transactions.has_more,
      investments: []
    };

    // 3. Handle Investments (Only on first page to save memory/payload)
    if (!next_cursor) {
      var invRaw = getPlaidInvestmentsData(accountId);
      if (invRaw != null) {
        // Ensure formatPlaidInvestments also returns slim objects
        result.investments = formatPlaidInvestments(accountId, invRaw) || [];
      }
    }

    // 4. Cache account metadata to avoid redundant lookups
    const accountCache = {};
    const addedRaw = transactions.added || [];
    const modifiedRaw = transactions.modified || [];
    
    // Identify unique account IDs in this batch
    const uniqueAids = [...new Set(addedRaw.concat(modifiedRaw).map(t => t.account_id))];
    
    uniqueAids.forEach(aid => {
      try {
        const accResp = getAppPlaidAccountById(aid);
        if (accResp && accResp.success && accResp.result) {
          accountCache[aid] = accResp.result;
        }
      } catch (e) { 
        Logger.log('Account Cache Error for ' + aid + ': ' + e.toString()); 
      }
    });

    /**
     * Helper function to SLIM DOWN the transaction object.
     * This is the part that fixes the "Limit Exceeded" error.
     */
    function slimEnrich(t) {
      if (!t) return null;
      const a = accountCache[t.account_id] || {};
      
      // We explicitly define ONLY the keys we need. 
      // This ignores large objects like t.location, t.payment_meta, etc.
      return {
        transaction_id: t.transaction_id,
        account_id: t.account_id,
        date: t.date,
        name: t.name,
        amount: t.amount,
        // Take only the first category string instead of the whole array
        category: (t.category && t.category.length > 0) ? t.category[0] : 'Uncategorized',
        pending: !!t.pending,
        // Enrich from local account cache
        account_name: a.name || t.account_name || '',
        mask: a.mask || t.mask || '',
        account_number: a.mask || a.account_number || '',
        institution: a.institution_name || ''
      };
    }

    // 5. Map the raw transactions into the slim versions
    result.added = addedRaw.map(slimEnrich).filter(t => t !== null);
    result.modified = modifiedRaw.map(slimEnrich).filter(t => t !== null);

    Logger.log('Payload Slimmed: Added=' + result.added.length + ', Modified=' + result.modified.length);

    return result;

  } catch (e) {
    Logger.log('prepareLinkPayloadPage critical error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function insertTransactionBatchSafe(txObjects){
  // txObjects: array of transaction objects with keys: transaction_id, pending_transaction_id, date, name, amount, pending, account_id, etc.
  if(!txObjects || txObjects.length === 0) return { success: true, inserted: 0, updated: 0 };

  const lock = LockService.getUserLock();
  lock.waitLock(30000);
  try{
    const sheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
    if(!sheet){
      return { success: false, error: 'The "' + USER_TRANSACTIONS_SHEET + '" sheet is missing. Please go to Settings and click "Reset Templates" to restore it.' };
    }
    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    const headerMap = buildHeaderIndexMapFromArray(headers);
    let txnIdColIdx = findHeaderIndexByKeywords(headerMap, ['transaction','id']);
    if(txnIdColIdx === -1) txnIdColIdx = findHeaderIndexByKeywords(headerMap, ['transaction']);
    const lastCol = headers.length;

    // Build map of existing transaction_id -> row (only if we found a column)
    const data = sheet.getDataRange().getValues();
    const existingMap = {};
    if(txnIdColIdx !== -1){
      for(let r = 1; r < data.length; r++){
        const tid = data[r][txnIdColIdx];
        if(tid && tid !== '') existingMap[tid] = r+1; // 1-based
      }
    }

    const newRows = [];
    let inserted = 0, updated = 0;

    // Identify columns that should never be overwritten on updates (user-edited data)
    var preserveColIndices = {};
    for (var h = 0; h < headers.length; h++) {
      var hKey = (headers[h] || '').toString().trim().toLowerCase();
      if (hKey.indexOf('assigned') !== -1 || hKey.indexOf('category') !== -1 || hKey.indexOf('owner') !== -1) {
        preserveColIndices[h] = true;
      }
    }

    txObjects.forEach(function(tx){
      const tid = tx.transaction_id || '';
      const transaction_status = tx.pending ? 'Pending' : '';
      const accName = tx.account_name || tx.account || '';
      const accMask = tx.account_number || tx.mask || '';
      const isUpdate = !!(tid && existingMap[tid]);

      const rowArr = [];
      // Build row array according to headers order (tolerant mapping by header keywords)
      for(let j=0;j<headers.length;j++){
        const hRaw = headers[j];
        const key = (hRaw || '').toString().trim().toLowerCase();
        let v = '';
        if(key.indexOf('not cleared') !== -1) v = '';
        else if(key === 'date' || key.indexOf('date') !== -1) v = tx.date || '';
        else if(key.indexOf('description') !== -1 || key.indexOf('desc') !== -1) v = tx.name || '';
        else if(key.indexOf('category') !== -1) v = '';
        else if(key === 'amount' || key.indexOf('amount') !== -1) v = tx.amount || 0;
        else if(key.indexOf('owner') !== -1) v = '';
        else if(key.indexOf('assigned') !== -1) v = '';
        else if(key.indexOf('account') !== -1 && key.indexOf('number') === -1 && key.indexOf('id') === -1) v = accName;
        else if(key.indexOf('transaction status') !== -1 || (key.indexOf('status') !== -1 && key.indexOf('transaction') !== -1)) v = transaction_status;
        else if(key.indexOf('account number') !== -1 || key.indexOf('account no') !== -1) v = accMask;
        else if(key.indexOf('account id') !== -1 || key === 'account id') v = tx.account_id || '';
        else if(key.indexOf('institution') !== -1) v = tx.institution || '';
        else if(key.indexOf('transaction id') !== -1 || (key.indexOf('transaction') !== -1 && key.indexOf('id') !== -1)) v = tid;
        else if(key.indexOf('group') !== -1) v = '';
        else if(key.indexOf('type') !== -1) v = '';
        else if(key.indexOf('period') !== -1) v = '';
        else v = '';
        rowArr.push(v);
      }

      if(isUpdate){
        const rowNum = existingMap[tid];
        // Preserve user-edited columns (Assigned, Category, Owner) by reading current values
        var currentRow = data[rowNum - 1]; // data is 0-indexed, rowNum is 1-indexed
        for (var ci in preserveColIndices) {
          rowArr[ci] = currentRow[ci];
        }
        sheet.getRange(rowNum,1,1,lastCol).setValues([rowArr])
          .setFontSize(9).setFontFamily('Comfortaa').setFontColor('#000000').setFontWeight('bold');
        updated++;
      }else{
        newRows.push(rowArr);
      }
    });

    if(newRows.length > 0){
      const lastrow = sheet.getLastRow() + 1;
      sheet.getRange(lastrow,1,newRows.length,headers.length).setValues(newRows)
        .setFontSize(9).setFontFamily('Comfortaa').setFontColor('#000000').setFontWeight('bold');
      inserted += newRows.length;
    }

    Logger.log('insertTransactionBatchSafe result: inserted=' + inserted + ' updated=' + updated + ' batchCount=' + txObjects.length);

    return { success: true, inserted: inserted, updated: updated };
  }catch(e){
    Logger.log('insertTransactionBatchSafe error: ' + e.toString());
    return { success: false, error: e.toString() };
  }finally{
    lock.releaseLock();
  }
}

function getInvestmentRowBySecurityId(securityId){
  const sheet = UserSpreadsheet.getSheetByName(USER_INVESTMENTS_SHEET);
  const data = sheet.getDataRange().getValues();
  for(let r = 0; r < data.length; r++){
    if(data[r].includes(securityId)) return r+1;
  }
  return null;
}

function insertInvestmentBatchSafe(invObjects){
  if(!invObjects || invObjects.length === 0) return { success: true, inserted: 0, updated: 0 };
  const lock = LockService.getUserLock();
  lock.waitLock(30000);
  try{
    const sheet = UserSpreadsheet.getSheetByName(USER_INVESTMENTS_SHEET);
    if(!sheet){
      return { success: false, error: 'The "' + USER_INVESTMENTS_SHEET + '" sheet is missing. Please go to Settings and click "Reset Templates" to restore it.' };
    }
    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    const headerMap = buildHeaderIndexMapFromArray(headers);
    const lastCol = headers.length;

    const newRows = [];
    let inserted = 0, updated = 0;

    invObjects.forEach(function(inv){
      const securityId = inv.security_id || '';
      const rowArr = [];
      for(let j=0;j<headers.length;j++){
        const hRaw = headers[j];
        const key = (hRaw||'').toString().trim().toLowerCase();
        let v = '';
        if(key.indexOf('name') !== -1) v = inv.account_name || inv.name || '';
        else if(key.indexOf('cusip') !== -1) v = inv.cusip || '';
        else if(key.indexOf('ticker') !== -1) v = inv.ticker || inv.ticker_symbol || '';
        else if(key.indexOf('price as of') !== -1 || key.indexOf('price as') !== -1) v = inv.price_as_of || '';
        else if(key === 'price' || key.indexOf('price') !== -1) v = inv.price || 0;
        else if(key.indexOf('quantity') !== -1) v = inv.quantity || 0;
        else if(key.indexOf('cost') !== -1 && key.indexOf('basis') !== -1) v = inv.cost_basis || 0;
        else if(key.indexOf('value') !== -1) v = inv.value || 0;
        else if(key.indexOf('account') !== -1 && key.indexOf('id') === -1) v = inv.account || '';
        else if(key.indexOf('security') !== -1 && key.indexOf('id') !== -1) v = securityId;
        else if(key.indexOf('account id') !== -1 || key === 'account id') v = inv.account_id || '';
        else v = '';
        rowArr.push(v);
      }

      const existingRow = getInvestmentRowBySecurityId(securityId);
      if(existingRow){
        sheet.getRange(existingRow,1,1,lastCol).setValues([rowArr])
          .setFontSize(9).setFontFamily('Comfortaa').setFontColor('#000000').setFontWeight('bold');
        updated++;
      }else{
        newRows.push(rowArr);
      }
    });

    if(newRows.length > 0){
      const lastrow = sheet.getLastRow() + 1;
      sheet.getRange(lastrow,1,newRows.length,headers.length).setValues(newRows)
        .setFontSize(9).setFontFamily('Comfortaa').setFontColor('#000000').setFontWeight('bold');
      inserted += newRows.length;
    }

    Logger.log('insertInvestmentBatchSafe result: inserted=' + inserted + ' updated=' + updated + ' batchCount=' + invObjects.length);

    return { success: true, inserted: inserted, updated: updated };
  }catch(e){
    Logger.log('insertInvestmentBatchSafe error: ' + e.toString());
    return { success: false, error: e.toString() };
  }finally{
    lock.releaseLock();
  }
}

function removeTransactionsRowIfAccountIDIsEmpty() {
  var sheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;
  var accountIdValue = sheet.getRange(2, 11).getValue();
  if (accountIdValue === '' || accountIdValue === null || accountIdValue === undefined) {
    sheet.deleteRow(2);
  }
}

function updatePlaidWebhook() {

  const appSettingsData = getAppSettings();

  if( appSettingsData.success === true ){
    const PLAID_CLIENT_ID = appSettingsData.result.plaidClientKey;
    const PLAID_SECRET    = appSettingsData.result.plaidSecretKey;
    const NEW_WEBHOOK_URL = appSettingsData.result.plaidWebhookUrl;
    const PLAID_API_ENDPOINT  = 'https://' + appSettingsData.result.plaidEnvironment + '.plaid.com/item/webhook/update';
    const ACCESS_TOKEN = '';
    
    // --- REQUEST ---
    const payload = {
      client_id: PLAID_CLIENT_ID,
      secret:    PLAID_SECRET,
      access_token: ACCESS_TOKEN,
      webhook: NEW_WEBHOOK_URL
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true  // Allows you to see error responses
    };

    const response = UrlFetchApp.fetch(PLAID_API_ENDPOINT, options);
    const responseCode = response.getResponseCode();
    const responseBody = JSON.parse(response.getContentText());

    // --- LOGGING ---
    Logger.log('HTTP Status: ' + responseCode);
    Logger.log('Response: ' + JSON.stringify(responseBody, null, 2));

    if (responseCode === 200) {
      Logger.log('✅ Webhook updated successfully!');
      Logger.log('New webhook: ' + responseBody.item.webhook);
    } else {
      Logger.log('❌ Error: ' + responseBody.error_message);
    }

    return responseBody;
  }  
}

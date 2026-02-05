const API_ENDPOINT = 'https://thefinu.stallioni.com/';

const UserEmail = Session.getActiveUser().getEmail();
const UserSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const UserSpreadsheetUrl = UserSpreadsheet.getUrl();
const UserSpreadsheetId = UserSpreadsheet.getId();

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

var APP_USER_ID = {
  client_user_id: generateRandomNumber(),
};

// Configuration
const RESTAPI_CONFIG = {
  API_BASE_URL: API_ENDPOINT,
  SCRIPT_ID: 'AKfycbyqPA2eaEAgwNGdVyOAbIeq3_h74nGaujmk80lomXVrErl-948LuTWr9F3rRxPEUX_mhA',
  TIMEOUT: 30000 // 30 seconds
};

const WEBAPP_WEBHOOK = RESTAPI_CONFIG.API_BASE_URL +'plaidwebhook';

/**
 * Creates the menu when the spreadsheet opens.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Open', 'showSidebar')
    //.addItem('Generate Reports', 'startGenerationOfReports')
    .addToUi();
}

/**
 * Automatically runs on installation.
 */
function onInstall(e) {
  onOpen(e);
}

function onChange(e){
  Logger.log(JSON.stringify(e, null, 2));
}

function handleAddonEdit(e){
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  const column = range.getColumn();
  const row = range.getRow();
  const newValue = e.value;
  const oldValue = e.oldValue;
  // Get row data
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  if( sheet.getName() === USER_ACCOUNTS_SHEET ){
    if( column === 6 || column === 8 || column === 10 || column === 11 ){
      if( rowData && rowData[5] !=='' && rowData[7] !=='' && rowData[9] !== '' ){
        populateNetWorth();
        populateJointNetWorth();
      }
    }
  }
  if( sheet.getName() === USER_BALANCE_HISTORY_SHEET ){
    if( column === 2 || column === 5 ){
      if( rowData && rowData[1] !=='' && rowData[4] !==''){
        populateNetWorth();
        populateJointNetWorth();
      }
    }
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
    USER_DEFINITION_SHEET,
    //USER_MONTHLY_BUDGET_SHEET,
    //USER_JOINT_MONTHLY_BUDGET_SHEET,
    //USER_YEARLY_BUDGET_SHEET,
    //USER_JOINT_YEARLY_BUDGET_SHEET,
    //USER_BUDGET_MAKER_SHEET,
  ];
}

function appFeaturedTemplates(){
  return [
    USER_MONTHLY_BUDGET_SHEET,
    USER_JOINT_MONTHLY_BUDGET_SHEET,
    USER_YEARLY_BUDGET_SHEET,
    USER_JOINT_YEARLY_BUDGET_SHEET,
    USER_BUDGET_MAKER_SHEET
  ];
}

/**
 * Opens the sidebar UI.
 */
function showSidebar() {
  let userValidation = validateUserSession();
  if( userValidation.result && userValidation.result.data.isSubscribed === true ){
    // mark subscription progress
    try{ markSetupStepCompleted('subscription', { status: 'active', activatedAt: new Date().toISOString() }); }catch(e){}
    // only show dashboard if all setup steps are complete
    if( isSetupCompleted() ){
      showUserDashboardSidebar();
    }else{
      showSetupWizardSidebar();
    }
  }else{
    try{ clearSubscriptionProgress(); }catch(e){}
    showSetupWizardSidebar();
  }
}

function showSetupWizardSidebar(){
  const html = HtmlService.createTemplateFromFile('Index')
      .evaluate()
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

    const PRICE_ID = 'price_1SnHxFBKorklj30OWLWvqJcP';

    const response = getAppSettings();

    if( response.success === true ){
      
      const stripeKey = response.result.stripeSecretKey;
      const email = Session.getActiveUser().getEmail();
      const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

      let url = 'https://api.stripe.com/v1/checkout/sessions';

      var payload =
        'mode=subscription' +
        '&customer_email='+ email + 
        '&success_url=' + encodeURIComponent(API_ENDPOINT+'success?session_id={CHECKOUT_SESSION_ID}&spreadsheet_id='+spreadsheetId) +
        '&cancel_url=' + encodeURIComponent(API_ENDPOINT+'cancel?spreadsheet_id='+spreadsheetId) +
        '&line_items[0][price]=' + PRICE_ID +
        '&line_items[0][quantity]=1';

      var options = {
        method: 'post',
        contentType: 'application/x-www-form-urlencoded',
        payload: payload,
        headers: {
          Authorization: 'Bearer ' + stripeKey
        },
        muteHttpExceptions: true
      };
      
      var request = UrlFetchApp.fetch(url, options);
      const json = JSON.parse(request.getContentText());

      return {
        success: true,
        checkoutUrl: json.url
      };
    }
   
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

function showUserDashboardTemplate(){
  const template = HtmlService.createTemplateFromFile('UserDashboard');
  template.isAutoSyncEnabled = PropertiesService.getUserProperties().getProperty("AUTO_SYNC_STATUS") === 'true' ? true : false;
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
 * Stored in User Properties under key: SETUP_WIZARD_PROGRESS
 */
function getSetupWizardProgress(){
  try{
    const userProps = PropertiesService.getUserProperties();
    const raw = userProps.getProperty('SETUP_WIZARD_PROGRESS');
    if( raw ){ 
      Logger.log( JSON.stringify(raw, null, 2) );
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
    progressObj.updatedAt = new Date().toISOString();
    userProps.setProperty('SETUP_WIZARD_PROGRESS', JSON.stringify(progressObj));
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
   * Shows a Google Sheets UI confirmation dialog for cancelling subscription.
   * Returns true if the user confirmed (YES), false otherwise.
   */
  function cancelUserSubscription(){
    try{
      const ui = SpreadsheetApp.getUi();
      const result = ui.alert('Cancel Subscription', 'Are you sure you want to cancel your subscription?', ui.ButtonSet.YES_NO);
      if (result == ui.Button.YES) {
        let response = confirmCancelUserSubscription();
        return response;
      }
    }catch(e){
      Logger.log('cancelUserSubscription error: ' + e.toString());
      return false;
    }
  }

  function confirmCancelUserSubscription(){
    clearAllUserProperties();
    return {
      status: true,
      message: 'Subscription cancelled successfully'
    };
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
    //try{ clearSubscriptionProgress(); }catch(e){}
    
  }catch(e){
    Logger.log('verifySubscriptionStatus error: ' + e.toString());
    return { success: false, subscribed: false, error: e.toString() };
  }
}

function checkUserSubscription(){
  const response = validateUserSession();
  if( response.success === true ){
    if( response.result && response.result.data.isSubscribed === true ){
      return getTemplateBlockUI('SubscriptionActivated');
    }else{
      return getTemplateBlockUI('SubscriptionTimeOut');
    }
  }else{
    return getTemplateBlockUI('SubscriptionTimeOut');
  }
}

function installTemplateInitialSetup(){
  try{
    const response = getAppSettings();
    if( response.success === true ){
      const spreadsheetTemplateUrl = response.result.spreadsheetTemplateUrl;
      let sourceSpreadsheet = SpreadsheetApp.openByUrl(spreadsheetTemplateUrl);
      const userSpreadsheet = UserSpreadsheet;
      let requiredSheets = appBaseTemplates();
      let sourceSheets = sourceSpreadsheet.getSheets();
      const requiredSet = new Set(requiredSheets.map(name => name.toLowerCase()));
      sourceSheets.forEach(function(sourceSheet) {
        let sheetName = sourceSheet.getName();
        if (!requiredSet.has(sheetName.toLowerCase())) return;
        if (!userSpreadsheet.getSheetByName(sheetName) ) {
          let copiedSheet = sourceSheet.copyTo(userSpreadsheet);
          copiedSheet.setName(sheetName);
          reApplyFormulaToSpreadsheet(sheetName);
        }
      });

      // Active Start Here sheet
      SpreadsheetApp.getActive().getSheetByName(USER_START_HERE_SHEET).activate();
      SpreadsheetApp.flush();
    }
    // mark template step completed
    try{ markSetupStepCompleted('template', { status: 'completed', installedAt: new Date().toISOString() }); }catch(e){}
    return {
      status: true,
      message: "Template(s) installed successfully"
    }
  }catch(error){
    Logger.log(`Error while installTemplateInitialSetup: ${error.message}`);
    return {
      status: false,
      message: "Something went wrong, please try again."
    }
  }
}

function saveConfiguration(data){
  try{
    const triggers = ScriptApp.getProjectTriggers();
    if( data.autoSync === true ){
      let functionToRun = 'runThefinUPlaidAutoSync';
      // Remove existing triggers for the target function
      triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === functionToRun) {
          ScriptApp.deleteTrigger(trigger);
        }
      });
      // Create a new daily time-based trigger
      ScriptApp.newTrigger(functionToRun)
        .timeBased()
        .everyDays(1)
        .atHour(6) // Set your preferred hour
        .create();

      PropertiesService.getUserProperties().setProperty("AUTO_SYNC_STATUS", true);
    }else{
      triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'runThefinUPlaidAutoSync') {
          ScriptApp.deleteTrigger(trigger);
        }
      });
      PropertiesService.getUserProperties().setProperty("AUTO_SYNC_STATUS", false);
    }

    setupInstallableTrigger();

    // mark config step completed
    try{ markSetupStepCompleted('config', { status: 'completed', autoSync: data.autoSync, savedAt: new Date().toISOString() }); }catch(e){}
    return {
      status: true,
      message: "Saved Successfully"
    }
  }catch(error){
    Logger.log(`Error while saveConfiguration: ${error.message}`);
    return {
      status: false,
      message: "Something went wrong, please try again."
    }
  }
}

function setupInstallableTrigger(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 1. Avoid duplicate triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'handleAddonEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  // 2. Create the installable onEdit trigger
  ScriptApp.newTrigger('handleAddonEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}

function toggleAutoSyncSetting( status ){
  try{
    const triggers = ScriptApp.getProjectTriggers();
    if( status === true ){
      let functionToRun = 'runThefinUPlaidAutoSync';
      // Remove existing triggers for the target function
      triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === functionToRun) {
          ScriptApp.deleteTrigger(trigger);
        }
      });
      // Create a new daily time-based trigger
      ScriptApp.newTrigger(functionToRun)
        .timeBased()
        .everyDays(1)
        .atHour(6) // Set your preferred hour
        .create();
      PropertiesService.getUserProperties().setProperty("AUTO_SYNC_STATUS", true);
    }else{
      triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'runThefinUPlaidAutoSync') {
          ScriptApp.deleteTrigger(trigger);
        }
      });
      PropertiesService.getUserProperties().setProperty("AUTO_SYNC_STATUS", false);
    }
    return {
      status: true,
      message: "Auto-Sync status updated successfully."
    }
  }catch(error){
    Logger.log(`Error while toggleAutoSyncSetting: ${error.message}`);
    return {
      status: false,
      message: "Something went wrong, please try again."
    }
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
    //Logger.log( JSON.stringify(response.result, null, 2) );
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
  try{
    let isSyncEnabled = PropertiesService.getUserProperties().getProperty("AUTO_SYNC_STATUS");
    if( isSyncEnabled !== 'true' ){
      return false;
    }
    const response = getAppPlaidConnectedAccounts();
    if( response.success === true ){
      let accounts = response.result;
      if( accounts.length > 0 ){
        accounts.forEach(function(account){
          if( account.is_update === true ){
            let account_id = account.account_id;
            updateTransactionSheet( account_id );
            let support_response = checkItemProductSupport( account_id, 'investments');
            if( support_response === true ){
              updateInvestmentSheet(account_id);
            }
            updateAccountBalanceHistory(account_id);
            updateAppAccountDetailById(
              account_id,
              {
                is_update: false
              }
            );
          }
        });
        populateNetWorth();
        populateJointNetWorth();
      }
    }
    MailApp.sendEmail(
      UserEmail,
      'Thefinu - Plaid Account(s) Sync.',
      'New updates are synced with connected plaid account(s).',
    );
    return true;
  }catch(error){
    MailApp.sendEmail(
      UserEmail,
      'Thefinu - Plaid Account(s) Sync.',
      'Something went wrong while trying to sync the plaid account(s).',
    );
    Logger.log("An error occurred:", JSON.stringify(error, null, 2));
    return false;
  }
}

function handleOnEdit(e){
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  const column = range.getColumn();
  const row = range.getRow();
  const newValue = e.value;
  const oldValue = e.oldValue;
  const editCell = range.getA1Notation();
  // Get row data
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const defSheet = UserSpreadsheet.getSheetByName(USER_DEFINITION_SHEET);

  if( sheet.getName() === USER_BALANCE_HISTORY_SHEET ){
    let dateCol = 2;
    let accountsCol = 3;
    let accountNumberCol = 4;
    let balanceCol = 5;
    let balanceIDCol = 6;
    let accountIDCol = 7;
    let dateTimeCol = 8;
    
    if( column === balanceCol ){
      let accountId = rowData[accountIDCol -1];
      // check if accountId is not empty & newValue is a number 
      if( accountId && accountId !== '' && !isNaN(newValue) ){
        let accountData = checkAccountBalanceByAccountId( accountId );
        if( accountData !== null ){
          let accountBalance = parseFloat( accountData[4] ) || 0;
          let newBalance = parseFloat( newValue ) || 0;
          if( accountBalance === newBalance ){
            let message = 'Generating reports will refresh all data and may take a few minutes to complete. Do you want to proceed?';
            confirmReportGenerationAlertMessage( message );
          }
        }
      }
    }
  }else if( sheet.getName() === USER_ACCOUNTS_SHEET ){
    let ownCol = defSheet.getRange('F11').getValue();
    let groupCol = defSheet.getRange('F6').getValue();
    let assetliabilityCol = defSheet.getRange('F7').getValue();
    let hideCol = defSheet.getRange('F8').getValue();
    // check if edited column is one of the above and only trigger report generation if all values are present

    if (column === ownCol || column === groupCol || column === assetliabilityCol || column === hideCol ) {
      if( rowData[ownCol -1] === '' || rowData[groupCol -1] === '' || rowData[assetliabilityCol -1] === ''){
        return;
      }
      let message = 'Generating reports will refresh all data and may take a few minutes to complete. Do you want to proceed?';
      confirmReportGenerationAlertMessage( message );
    }
  }else if( sheet.getName() === USER_TRANSACTIONS_SHEET ){
    let catCol = defSheet.getRange('I5').getValue();
    let ownCol = defSheet.getRange('I11').getValue();
    let amtCol = defSheet.getRange('I10').getValue();
    if (column === catCol || column === ownCol || column === amtCol ) {
      let message = 'Generating reports will refresh all data and may take a few minutes to complete. Do you want to proceed?';
      confirmReportGenerationAlertMessage( message );
    }
  }else if( sheet.getName() === USER_MONTHLY_BUDGET_SHEET ){
    if( editCell === 'C2' ){
      startGenerationOfMonthlyBudget();
    }
  }else if( sheet.getName() === USER_YEARLY_BUDGET_SHEET ){
    if( editCell === 'E2' ){
      startGenerationOfYearlyBudget();
    }
  }else if( sheet.getName() === USER_JOINT_MONTHLY_BUDGET_SHEET ){
    if( editCell === 'C2' ){
      startGenerationOfJointMonthlyBudget();
    }
  }else if( sheet.getName() === USER_JOINT_YEARLY_BUDGET_SHEET ){
    if( editCell === 'D2' ){
      startGenerationOfJointYearlyBudget();
    }
  }else if( sheet.getName() === USER_CATEGORIES_SHEET ){
  }else{
    return;
  }
}

/**
 * Backend functions called from the UI
 */
function linkNewAccount() {
  const html = HtmlService.createHtmlOutputFromFile('ConnectPlaidAccount').setWidth(450).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Connect Plaid Account");
}

function importTransactions(accountId) {
  Logger.log("Importing for: " + accountId);
  return "Transactions imported.";
}

function updateAccountName(accountId, newName) {

  try{
    const response = updateAppAccountDetailById(accountId,{name: newName});
    changeAccountNameOnBalanceHistorySheet(accountId, newName);
    changeAccountNameOnTransactionSheet(accountId, newName);
    changeAccountNameOnInvestmentSheet(accountId, newName);
    if( response.success === true ){
      return {
        status: true,
        message: 'Account name updated'
      };
    }else{
      return {
        status: false,
        message: 'Something went wrong, please try again.'
      };
    }
  }catch(error){
    Logger.log(`Error while updateAccountName: ${error.message}`);
    return {
      status: false,
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
          status: true,
          message: 'Account name updated'
        };
      }else{
        return {
          status: false,
          message: 'Something went wrong, please try again.'
        };
      }
    }
  }catch(error){
    Logger.log(`Error while updateAccountName: ${error.message}`);
    return {
      status: false,
      message: "Something went wrong, please try again."
    }
  }finally {
    SpreadsheetApp.getUi().alert('Account removed successfully.');
  }
}

function confirmLinkAccountToTemplate( accountId ){

  const userProperties = PropertiesService.getUserProperties();

  const accountName = getPlaidAccountNameByAccountId(accountId);

  if( accountName !== null ){
    var result = SpreadsheetApp.getUi().alert(
      'Insert historical data?',
      "New updates for '"+ accountName +"' will now sync with this spreadsheet. Do you wish to insert this account's historical data too? Clicking 'Yes' will add historical transactions and balances for '"+ accountName +"' to the current spreadsheet.", 
        SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    if (result == SpreadsheetApp.getUi().Button.YES) {
      SpreadsheetApp.getUi().alert('The process takes few minutes to be completed. Please wait do not change anything untill the process complete.');
      userProperties.setProperty('TASK_STATUS', 'PROCESSING');
      userProperties.setProperty('LINK_ACCOUNT_ID', accountId);
      linkAccountDataToSpreadsheet(accountId);
      return true;
    }else{
      SpreadsheetApp.getUi().alert('Something went wrong, please try again.');
      return false;
    }
  }else{
    SpreadsheetApp.getUi().alert('Something went wrong, please try again.');
    return false;
  }
}

function confirmUnlinkAccountFromTemplate(accountId){
  
  const userProperties = PropertiesService.getUserProperties();

  const accountName = getPlaidAccountNameByAccountId(accountId);

  var result = SpreadsheetApp.getUi().alert(
     'Remove existing data?',
     "'"+ accountName +"' will no longer sync with this spreadsheet. Do you wish to remove this account's existing data too? Clicking 'Yes' will remove transactions and balances for '"+ accountName +"' from the current spreadsheet.",
      SpreadsheetApp.getUi().ButtonSet.YES_NO);

  if (result == SpreadsheetApp.getUi().Button.YES) {
    SpreadsheetApp.getUi().alert('The process takes few minutes to be completed. Please wait do not change anything untill the process complete.');
    userProperties.setProperty('TASK_STATUS', 'PROCESSING');
    userProperties.setProperty('UNLINK_ACCOUNT_ID', accountId);
    unlinkAccountDataFromSpreadsheet(accountId);
    return true;
  }
}

/**
 * Step 2: The actual task executed by the trigger.
 */
function linkAccountDataToSpreadsheet(account_id) {

  try {
    // --- PERFORM YOUR TASK HERE (e.g., API calls, heavy data processing) ---
    Utilities.sleep(5000); // Simulating work
    linkTransactionSheet(account_id);
    let support_response = checkItemProductSupport( account_id, 'investments');
    if( support_response === true ){
      linkInvestmentSheet(account_id);
    }
    updateAccountBalanceHistory(account_id);
    linkAccountsSheetData(account_id);
    installFeaturedTemplates();
    updateAppAccountDetailById(
      account_id,
      {
        is_linked: true,
        status: true,
        updates: false,
        linked_date: getTodayDateTime()
      }
    );
    SpreadsheetApp.flush();
    //reApplyFormulaToSpreadsheet();
    // Update flag to 'COMPLETED'
    PropertiesService.getUserProperties().setProperty('TASK_STATUS', 'COMPLETED');
    PropertiesService.getUserProperties().setProperty('LINK_ACCOUNT_ID', '');
  } catch (err) {
    PropertiesService.getUserProperties().setProperty('TASK_STATUS', 'ERROR: ' + err.message);
  } finally {
    SpreadsheetApp.getUi().alert('Account linked successfully.');
  }

}

function unlinkAccountDataFromSpreadsheet(account_id){
  try {
    // --- PERFORM YOUR TASK HERE (e.g., API calls, heavy data processing) ---
    Utilities.sleep(5000); // Simulating work
    clearTransactionsData(account_id);
    clearInvestmentsData(account_id);
    clearBalanceHitoryData(account_id);
    clearAccountData(account_id);
    updateAppAccountDetailById(
      account_id,
      {
        is_linked: false,
        status: true,
        updates: false,
        linked_date: getTodayDateTime()
      }
    );
    // Update flag to 'COMPLETED'
    PropertiesService.getUserProperties().setProperty('TASK_STATUS', 'COMPLETED');
    PropertiesService.getUserProperties().setProperty('UNLINK_ACCOUNT_ID', '');
  } catch (err) {
    PropertiesService.getUserProperties().setProperty('TASK_STATUS', 'ERROR: ' + err.message);
  } finally {
    SpreadsheetApp.getUi().alert('Account unlinked successfully.');
  }
}

/**
 * Helper: Finds and deletes triggers by function name
 */
function deleteTriggerByFunction(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
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
    case USER_DEFINITION_SHEET:
      if (spreadsheet.getSheetByName(USER_DEFINITION_SHEET)) {
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("R2").setFormula('=ARRAYFORMULA(UNIQUE(YEAR(INDIRECT(P9))))'); // Set the formula
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("S2").setFormula('=sort(unique(ARRAYFORMULA(Date(Year(Indirect(P4)),MONTH(Indirect(P4)),1))),1,True)'); // Set the formula
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("V1").setFormula("='Yearly Budget'!E2"); // Set the formula
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("X1").setFormula("='Joint Yearly Budget'!D2"); // Set the formula
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AC2").setFormula("='Yearly Budget'!E2"); // Set the formula
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AC4").setFormula("='Monthly Budget'!C2"); // Set the formula
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AD2").setFormula("='Joint Yearly Budget'!D2"); // Set the formula
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AD4").setFormula("='Joint Monthly Budget'!C2"); // Set the formula
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AC10").setFormula("=IFNA(INDEX('Monthly Budget'!E:E,MATCH('Income','Monthly Budget'!B:B,0)),0)"); // Set the formula
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AC10").setFormula("=IFNA(INDEX('Monthly Budget'!E:E,MATCH('Income','Monthly Budget'!B:B,0)),0)"); // Set the formula
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AC11").setFormula("=IFNA(INDEX('Monthly Budget'!E:E,MATCH('Expense','Monthly Budget'!B:B,0)),0)"); // Set the formula
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AC11").setFormula("=IFNA(INDEX('Monthly Budget'!E:E,MATCH('Expense','Monthly Budget'!B:B,0)),0)"); // Set the formula
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AD10").setFormula("=IFNA(INDEX('Joint Monthly Budget'!F:F,MATCH('Income','Joint Monthly Budget'!B:B,0)+2),0)"); // Set the formula
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AD10").setFormula("=IFNA(INDEX('Joint Monthly Budget'!F:F,MATCH('Income','Joint Monthly Budget'!B:B,0)+2),0)"); // Set the formula
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AD11").setFormula("=IFNA(INDEX('Joint Monthly Budget'!F:F,MATCH('Expense','Joint Monthly Budget'!B:B,0)+2),0)"); // Set the formula
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
        // Get cell E2
        let cell = spreadsheet.getSheetByName(USER_MONTHLY_BUDGET_SHEET).getRange("C2");
        // Clear existing data validation and content
        cell.clearDataValidations();
        cell.clearContent();
        // Get the named range "Year"
        let periodRange = spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("S2:S1000");
        // Get values from the Year range and filter out empty/invalid values
        let periodValues = periodRange.getValues().flat().filter(function(value) {
          return value && (typeof value === 'string' || !isNaN(value));
        });
        if (periodValues.length === 0) {
          Logger.log("Error: No valid period values found in the 'Period' range (" + periodRange.getA1Notation() + ").");
          return;
        }
        // Create data validation rule using the full Year range
        var rule = SpreadsheetApp.newDataValidation()
          .requireValueInRange(periodRange, true) // Use full Year range
          .setAllowInvalid(false) // Reject invalid inputs
          .build();
        // Apply the data validation rule to E2
        cell.setDataValidation(rule);
        // Set the default value to the first valid value
        cell.setValue(periodValues[0]);
        spreadsheet.getSheetByName(USER_MONTHLY_BUDGET_SHEET).getRange("E4").setFormula('=Definition!AC24'); // Set the formula
      }
    break;
    case USER_JOINT_MONTHLY_BUDGET_SHEET:
      if (spreadsheet.getSheetByName(USER_JOINT_MONTHLY_BUDGET_SHEET)) {
      // Get cell E2
        let cell = spreadsheet.getSheetByName(USER_JOINT_MONTHLY_BUDGET_SHEET).getRange("C2");
        // Clear existing data validation and content
        cell.clearDataValidations();
        cell.clearContent();
        // Get the named range "Year"
        let periodRange = spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("S2:S1000");
        // Get values from the Year range and filter out empty/invalid values
        let periodValues = periodRange.getValues().flat().filter(function(value) {
          return value && (typeof value === 'string' || !isNaN(value));
        });
        if (periodValues.length === 0) {
          Logger.log("Error: No valid period values found in the 'Period' range (" + periodRange.getA1Notation() + ").");
          return;
        }
        // Create data validation rule using the full Year range
        var rule = SpreadsheetApp.newDataValidation()
          .requireValueInRange(periodRange, true) // Use full Year range
          .setAllowInvalid(false) // Reject invalid inputs
          .build();
        // Apply the data validation rule to E2
        cell.setDataValidation(rule);
        // Set the default value to the first valid value
        cell.setValue(periodValues[0]);
        spreadsheet.getSheetByName(USER_JOINT_MONTHLY_BUDGET_SHEET).getRange("E4").setFormula('=Definition!AD24'); // Set the formula
        spreadsheet.getSheetByName(USER_JOINT_MONTHLY_BUDGET_SHEET).getRange("B5").setFormula('=Definition!AD5'); // Set the formula
      }
    break;
    case USER_YEARLY_BUDGET_SHEET:
      if (spreadsheet.getSheetByName(USER_YEARLY_BUDGET_SHEET)) {
        // Get cell E2
        let cell = spreadsheet.getSheetByName(USER_YEARLY_BUDGET_SHEET).getRange("E2");
        // Clear existing data validation and content
        cell.clearDataValidations();
        cell.clearContent();
        // Get the named range "Year"
        let yearRange = spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("R2:R1000");
        // Get values from the Year range and filter out empty/invalid values
        let yearValues = yearRange.getValues().flat().filter(function(value) {
          return value && !isNaN(value) && String(value).match(/^\d{4}$/); // Ensure valid 4-digit years
        });
        if (yearValues.length === 0) {
          Logger.log("Error: No valid 4-digit year values found in the 'Year' range (" + yearRange.getA1Notation() + ").");
          return;
        }
        // Create data validation rule using the full Year range
        var rule = SpreadsheetApp.newDataValidation()
          .requireValueInRange(yearRange, true) // Use full Year range
          .setAllowInvalid(false) // Reject invalid inputs
          .build();
        // Apply the data validation rule to E2
        cell.setDataValidation(rule);
        // Set the default value to the first valid value
        cell.setValue(yearValues[0]);
        spreadsheet.getSheetByName(USER_YEARLY_BUDGET_SHEET).getRange("B6:D6").setFormula('=Definition!AC3'); // Set the formula

      }
    break;
    case USER_JOINT_YEARLY_BUDGET_SHEET:
      if (spreadsheet.getSheetByName(USER_JOINT_YEARLY_BUDGET_SHEET)) {
        // Get cell E2
        let cell = spreadsheet.getSheetByName(USER_JOINT_YEARLY_BUDGET_SHEET).getRange("D2");
        // Clear existing data validation and content
        cell.clearDataValidations();
        cell.clearContent();
        // Get the named range "Year"
        let yearRange = spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("R2:R1000");
        // Get values from the Year range and filter out empty/invalid values
        let yearValues = yearRange.getValues().flat().filter(function(value) {
          return value && !isNaN(value) && String(value).match(/^\d{4}$/); // Ensure valid 4-digit years
        });
        if (yearValues.length === 0) {
          Logger.log("Error: No valid 4-digit year values found in the 'Year' range (" + yearRange.getA1Notation() + ").");
          return;
        }
        // Create data validation rule using the full Year range
        var rule = SpreadsheetApp.newDataValidation()
          .requireValueInRange(yearRange, true) // Use full Year range
          .setAllowInvalid(false) // Reject invalid inputs
          .build();
        // Apply the data validation rule to E2
        cell.setDataValidation(rule);
        // Set the default value to the first valid value
        cell.setValue(yearValues[0]);
        spreadsheet.getSheetByName(USER_JOINT_YEARLY_BUDGET_SHEET).getRange("B5:C5").setFormula('=Definition!AD3'); // Set the formula
      }
    break;
  }
}


function installFeaturedTemplates(){

  let transactionSheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
  if( transactionSheet ){
    if( transactionSheet.getLastRow() > 2 ){
      const response = getAppSettings();
      if( response.success === true ){
        const spreadsheetTemplateUrl = response.result.spreadsheetTemplateUrl;
        let sourceSpreadsheet = SpreadsheetApp.openByUrl(spreadsheetTemplateUrl);
        const userSpreadsheet = UserSpreadsheet;
        let requiredSheets = appFeaturedTemplates();
        let sourceSheets = sourceSpreadsheet.getSheets();
        const requiredSet = new Set(requiredSheets.map(name => name.toLowerCase()));
        sourceSheets.forEach(function(sourceSheet) {
          let sheetName = sourceSheet.getName();
          if (!requiredSet.has(sheetName.toLowerCase())) return;
          if (!userSpreadsheet.getSheetByName(sheetName) ) {
            let copiedSheet = sourceSheet.copyTo(userSpreadsheet);
            copiedSheet.setName(sheetName);
            reApplyFormulaToSpreadsheet(sheetName);
          }
        });
        reApplyFormulaToSpreadsheet(USER_DEFINITION_SHEET);
      }
    }
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
    Logger.log(`Error while installTemplateInitialSetup: ${error.message}`);
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

function confirmReportGenerationAlertMessage(message){
  // Show alert message with yes or no options
  var result = SpreadsheetApp.getUi().alert(
    'Generate Reports?',
    message || 'Generating reports will refresh all data and may take a few minutes to complete. Do you want to proceed?',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  if (result == SpreadsheetApp.getUi().Button.YES) {
    SpreadsheetApp.getUi().alert('Report generation has started. Do not change anything until the process completes.');
    startReportGenerationTask();
  }
}

function checkAccountBalanceByAccountId(accountId){
  const sheet = UserSpreadsheet.getSheetByName(USER_ACCOUNTS_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let row = 1; row < data.length; row++) {
    if( data[row].includes(accountId) ){
      return JSON.parse( JSON.stringify( data[row] ) );
    }
  }
  return null;
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
      status: false,
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
    sheet.getRange(lastrow, 2).setValue(data.balanceDate).setFontSize(9)
        .setFontFamily("Comfortaa")
        .setFontColor("#000000")
        .setFontWeight("bold"); // Date
    sheet.getRange(lastrow, 3).setValue(accountName).setFontSize(9)
        .setFontFamily("Comfortaa")
        .setFontColor("#000000")
        .setFontWeight("bold"); // Account Name
    sheet.getRange(lastrow, 4).setValue(accountNumber).setFontSize(9)
        .setFontFamily("Comfortaa")
        .setFontColor("#000000")
        .setFontWeight("bold"); // Account Number
    sheet.getRange(lastrow, 5).setValue(data.accountBalance).setFontSize(9)
        .setFontFamily("Comfortaa")
        .setFontColor("#000000")
        .setFontWeight("bold"); // Balance
    sheet.getRange(lastrow, 7).setValue(accountId).setFontSize(9)
        .setFontFamily("Comfortaa")
        .setFontColor("#000000")
        .setFontWeight("bold"); // Account ID
    sheet.getRange(lastrow, 8).setValue(getTodayDateTime()).setFontSize(9)
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
      status: true,
      message: "Balance history added successfully."
    }
  }catch(error){
    Logger.log(`Error while addManualAccountBalanceHistoryData: ${error.message}`);
    return {
      status: false,
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
  let month = (date.getMonth() + 1).toString().padStart(2, '0');
  let day = date.getDate().toString().padStart(2, '0');
  let year = date.getFullYear();
  return `${month}/${day}/${year}`;
}

function test(){
  let sheet = UserSpreadsheet.getSheetByName(USER_ACCOUNTS_SHEET);
  let lastrow = sheet.getLastRow() + 1;
  for( let i = 1; i <= lastrow; i++ ){
    if( sheet.getRange(i + 1, 2).getValue() === ''){
      sheet.getRange(i + 1, 2).setValue(account_name);
      break;
    }
  }
  Logger.log("Done");
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

function getAccountNameByAccountId(account_id){
  try{
    const response = getAppPlaidAccountById(account_id);
    //Logger.log( JSON.stringify(response, null, 2) );
    if( response.success === true ){
      return response.result.name;
    }else{
      return null;
    }
  }catch(error){
    Logger.log(`Error while installTemplateInitialSetup: ${error.message}`);
    return null;
  }
}
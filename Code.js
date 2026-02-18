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
  const editCell = range.getA1Notation();
  // Get row data
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  if( sheet.getName() === USER_ACCOUNTS_SHEET ){
    if( column === 6 || column === 8 || column === 10 || column === 11 ){
      populateNetWorth();
      populateJointNetWorth();
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
  if( sheet.getName() === USER_MONTHLY_BUDGET_SHEET && editCell === 'C2' ){
    startGenerationOfMonthlyBudget();
  }
  if( sheet.getName() === USER_YEARLY_BUDGET_SHEET && editCell === 'E2' ){
    startGenerationOfYearlyBudget();
  }
  if( sheet.getName() === USER_JOINT_MONTHLY_BUDGET_SHEET && editCell === 'C2' ){
    startGenerationOfJointMonthlyBudget();
  }
  if( sheet.getName() === USER_JOINT_YEARLY_BUDGET_SHEET && editCell === 'D2' ){
    startGenerationOfJointYearlyBudget();
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
        if( response.success === true ){
          deleteAllSheetsAndRecreate();
          clearAllUserProperties();
          return true;
        }
        return false;
      }
    }catch(e){
      Logger.log('cancelUserSubscription error: ' + e.toString());
      return false;
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
    return true;
  }catch(error){
    Logger.log(`Error while runThefinUPlaidAutoSync: ${error.message}`);
    return false;
  }
}

function handleOnEdit(e){
  //if (!e || !e.range) return;
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

  Logger.log( "sheet: "+ sheet.getName() + ' cell: '+ editCell +'range: '+ range);
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
 * Starts the remote upload and scheduling process for linking account data.
 */
function startLinkingProcess(accountId){
  const userProperties = PropertiesService.getUserProperties();
  // avoid duplicate starts
  if( userProperties.getProperty('TASK_STATUS') === 'PROCESSING' ){
    return true;
  }

  try{
    // Gather transactions (added/modified)
    var addedTransactions = [];
    var modifiedTransactions = [];
    var has_more = false;
    var next_cursor = '';
    let transactions = getPlaidTransactionSyncData( accountId, next_cursor );
    if( transactions && transactions.request_id != '' ){
      has_more = transactions.has_more;
      next_cursor = transactions.next_cursor;
      addedTransactions.push(transactions.added || []);
      modifiedTransactions.push(transactions.modified || []);
      while( has_more === true ){
        let next_transactions = getPlaidTransactionSyncData( accountId, next_cursor);
        addedTransactions.push(next_transactions.added || []);
        modifiedTransactions.push(next_transactions.modified || []);
        next_cursor = next_transactions.next_cursor;
        has_more = next_transactions.has_more;
      }
    }
    var flatAdded = addedTransactions.reduce(function(acc, chunk){ return acc.concat(chunk || []); }, []);
    var flatModified = modifiedTransactions.reduce(function(acc, chunk){ return acc.concat(chunk || []); }, []);

    // Gather investments
    var investments = [];
    var invRaw = getPlaidInvestmentsData(accountId);
    if(invRaw != null){
      investments = formatPlaidInvestments(accountId, invRaw) || [];
    }

    var payload = {
      account_id: accountId,
      next_cursor: next_cursor || '',
      added: flatAdded,
      modified: flatModified,
      investments: investments
    };

    // Process payload locally (no external upload/fetch)
    try{
      userProperties.setProperty('TASK_STATUS','PROCESSING');
      userProperties.setProperty('LINK_ACCOUNT_ID', accountId);

      // Combine added and modified transactions for insertion
      var allTransactions = (payload.added || []).concat(payload.modified || []);
      if( Array.isArray(allTransactions) && allTransactions.length > 0 ){
        var txRows = [];
        let account_response = getAppPlaidAccountById( accountId );
        let account_data = account_response.result;
        allTransactions.forEach(function(transaction){
          let transaction_status = transaction.pending == true ? 'Pending' : '';
          txRows.push([
            '', // Not Cleared
            transaction.date || '',
            transaction.name || '',
            '',
            transaction.amount || 0,
            '',
            '',
            account_data ? account_data.name : '',
            transaction_status,
            account_data ? account_data.mask : '',
            transaction.account_id || accountId,
            account_data ? account_data.institution_name : '',
            transaction.transaction_id || '',
            '',
            '',
            ''
          ]);
        });
        batchInsertRows('transactions', txRows);
        sortingTransactionSheet();
      }

      if( Array.isArray(payload.investments) && payload.investments.length > 0 ){
        var invRows = [];
        payload.investments.forEach(function(inv){
          invRows.push([
            '',
            inv.account_name || inv.name || '',
            inv.cusip || '',
            inv.ticker || inv.ticker_symbol || '',
            inv.price_as_of || '',
            inv.price || 0,
            inv.quantity || 0,
            inv.cost_basis || 0,
            inv.value || 0,
            inv.account || '',
            inv.security_id || '',
            inv.account_id || accountId
          ]);
        });
        batchInsertRows('investments', invRows);
        sortingInvestmentSheet();
      }

      userProperties.setProperty('TASK_STATUS','COMPLETED');
      userProperties.deleteProperty('LINK_ACCOUNT_ID');
      userProperties.deleteProperty('LINK_PROCESS_TIMESTAMP');
      return true;
    }catch(e){
      userProperties.setProperty('TASK_STATUS','ERROR: ' + e.toString());
      Logger.log('startLinkingProcess insertion error: ' + e.toString());
      return false;
    }

  }catch(e){
    Logger.log('startLinkingProcess error: ' + e.toString());
    SpreadsheetApp.getUi().alert('Something went wrong.');
    return false;
  }
}

// processLinkAccountInterval removed: synchronous polling is used instead.

function processLinkAccountTask() {
  const lock = LockService.getUserLock();
  lock.waitLock(30000);

  const props = PropertiesService.getUserProperties();
  const accountId = props.getProperty('LINK_ACCOUNT_ID');

  if (!accountId) return;

  try {
    // --- PLAID IMPORT PIPELINE ---
    installFeaturedTemplates();
    linkTransactionSheet(accountId);
    if (checkItemProductSupport(accountId, 'investments')) {
      linkInvestmentSheet(accountId);
    }
    updateAccountBalanceHistory(accountId);
    linkAccountsSheetData(accountId);
    updateAppAccountDetailById(accountId, {
      is_linked: true,
      status: true,
      updates: false,
      linked_date: getTodayDateTime()
    });

    reApplyFormulaToSpreadsheet();

    populateNetWorth();
    populateJointNetWorth();

    SpreadsheetApp.flush();

    props.setProperty('TASK_STATUS', 'COMPLETED');
    props.deleteProperty('LINK_ACCOUNT_ID');

  } catch (e) {
    props.setProperty('TASK_STATUS', 'ERROR: ' + e.message);
  } finally {
    lock.releaseLock();
    cleanupTriggers_('processLinkAccountTask');
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

    installFeaturedTemplates();
    updateAccountBalanceHistory(accountId);
    linkAccountsSheetData(accountId);
    updateAppAccountDetailById(accountId, {
      is_linked: true,
      status: true,
      updates: false,
      linked_date: getTodayDateTime()
    });

    reApplyFormulaToSpreadsheet();
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

function resetTaskStatus(){
  const props = PropertiesService.getUserProperties();
  props.deleteProperty('TASK_STATUS');
  props.deleteProperty('LINK_ACCOUNT_ID');
}

function cleanupTriggers_(handlerName) {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
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

function processUnlinkAccountTask(accountIdParam) {
  const lock = LockService.getUserLock();
  lock.waitLock(30000);

  const props = PropertiesService.getUserProperties();
  const accountId = accountIdParam || props.getProperty('UNLINK_ACCOUNT_ID');

  if (!accountId) return { success: false, error: 'missing_accountId' };

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

    SpreadsheetApp.flush();
    props.setProperty('TASK_STATUS', 'COMPLETED');
    props.deleteProperty('UNLINK_ACCOUNT_ID');
    cleanupTriggers_('processUnlinkAccountTask');
    return { success: true };
  } catch (e) {
    props.setProperty('TASK_STATUS', 'ERROR: ' + e.message);
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
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
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AC10").setFormula(`=IFNA(INDEX('Monthly Budget'!E:E,MATCH("Income",'Monthly Budget'!B:B,0)),0)`);
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AC11").setFormula(`=IFNA(INDEX('Monthly Budget'!E:E,MATCH("Expense",'Monthly Budget'!B:B,0)),0)`); // Set the formula
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AD10").setFormula(`=IFNA(INDEX('Joint Monthly Budget'!F:F,MATCH("Income",'Joint Monthly Budget'!B:B,0)+2),0)`); // Set the formula
        spreadsheet.getSheetByName(USER_DEFINITION_SHEET).getRange("AD11").setFormula(`=IFNA(INDEX('Joint Monthly Budget'!F:F,MATCH("Expense",'Joint Monthly Budget'!B:B,0)+2),0)`); // Set the formula
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

function ensureTrigger(handler) {
  const triggers = ScriptApp.getProjectTriggers();

  const existing = triggers.find(t => t.getHandlerFunction() === handler);
  if (existing) return;

  ScriptApp.newTrigger(handler)
    .timeBased()
    .after(1000)
    .create();
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

    txObjects.forEach(function(tx){
      const tid = tx.transaction_id || '';
      const transaction_status = tx.pending ? 'Pending' : '';
      const accName = tx.account_name || tx.account || '';
      const accMask = tx.account_number || tx.mask || '';

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

      if(tid && existingMap[tid]){
        const rowNum = existingMap[tid];
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
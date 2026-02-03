const DEFAULT_API_ENDPOINT = 'https://thefinu.stallioni.com/';
const API_ENDPOINT = getScriptProperty('API_ENDPOINT', DEFAULT_API_ENDPOINT);

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
  SCRIPT_ID: getScriptProperty('SCRIPT_ID', 'AKfycbyqPA2eaEAgwNGdVyOAbIeq3_h74nGaujmk80lomXVrErl-948LuTWr9F3rRxPEUX_mhA'),
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
    .addToUi();
}

/**
 * Automatically runs on installation.
 */
function onInstall(e) {
  onOpen(e);
}

function getScriptProperty(key, fallback) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  return value ? value : fallback;
}

function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

/*function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() === 'Balance History'){
    // Only column 3 (Accounts)
    if (e.range.getColumn() !== 3) return;
    // Ignore multi-cell edits (paste, autofill)
    if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;
    const newValue = String(e.value || '').trim();
    const oldValue = String(e.oldValue || '').trim();
    // No meaningful change
    if (!newValue && !oldValue) return;
    if (newValue === oldValue) return;
    // Update Date & Time column (last column)
    sheet.getRange(e.range.getRow(), sheet.getLastColumn()).setValue(getTodayDateTime());
    updateAccountsSheet(oldValue, newValue);
  }
}*/

function updateAccountsSheet(oldAccount, newAccount) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Accounts');
  const normalize = v => String(v || '').toLowerCase().trim();

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }
  const colB = sheet.getRange(2, 2, lastRow - 1, 1).getValues();

  let lastContentRow = 1;
  let oldRowIndex = null;
  let newExists = false;

  colB.forEach((row, i) => {
    if (row[0]) {
      lastContentRow = i + 2;

      const value = normalize(row[0]);
      if (oldAccount && value === normalize(oldAccount)) {
        oldRowIndex = i + 2;
      }
      if (newAccount && value === normalize(newAccount)) {
        newExists = true;
      }
    }
  });

  // 🔁 Replace old account with new account
  if (oldRowIndex && newAccount) {
    sheet.getRange(oldRowIndex, 2).setValue(newAccount);
    return;
  }

  // ➕ Add new account if missing
  if (newAccount && !newExists) {
    sheet.insertRowsAfter(lastContentRow, 1);
    sheet.getRange(lastContentRow + 1, 2).setValue(newAccount);
  }
}

/*function syncAccountsFromBalanceHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const balanceSheet = ss.getSheetByName('Balance History');
  const accountsSheet = ss.getSheetByName('Accounts');

  if (!balanceSheet || !accountsSheet) {
    throw new Error('Required sheets not found');
  }

  // Normalize helper
  const normalize = v => String(v).trim().toLowerCase();

  // 1️⃣ Read Balance History (Column C = Accounts)
  const balanceData = balanceSheet.getDataRange().getValues();
  const balanceAccounts = balanceData
    .slice(1)
    .map(r => r[2])
    .filter(Boolean)
    .map(normalize);

  if (!balanceAccounts.length) return;

  // Deduplicate Balance History accounts
  const balanceSet = new Set(balanceAccounts);

  // 2️⃣ Read Accounts sheet Column B
  const maxRows = accountsSheet.getMaxRows();
  const accountCol = accountsSheet
    .getRange(2, 2, maxRows - 1, 1)
    .getValues();

  const existingSet = new Set();
  let lastContentRow = 1;

  accountCol.forEach((row, i) => {
    if (row[0]) {
      existingSet.add(normalize(row[0]));
      lastContentRow = i + 2;
    }
  });

  // 3️⃣ Find truly new accounts
  const newAccounts = [...balanceSet].filter(
    acc => !existingSet.has(acc)
  );

  if (!newAccounts.length) return;

  // 4️⃣ Prepare rows (restore original case from Balance History)
  const originalMap = {};
  balanceData.slice(1).forEach(r => {
    if (r[2]) {
      const key = normalize(r[2]);
      if (!originalMap[key]) originalMap[key] = r[2].trim();
    }
  });

  const rowsToInsert = newAccounts.map(acc => [
    '', originalMap[acc], '', '', '', '', '', '', '', '', '', ''
  ]);

  const colCount = accountsSheet.getLastColumn();

  // 5️⃣ Insert rows after last real account
  accountsSheet.insertRowsAfter(lastContentRow, rowsToInsert.length);

  // 6️⃣ Write data
  const range = accountsSheet.getRange(
    lastContentRow + 1,
    1,
    rowsToInsert.length,
    colCount
  );

  range
    .setValues(rowsToInsert)
    .setFontSize(9)
    .setFontFamily('Comfortaa')
    .setFontWeight('bold')
    .setFontColor('#000000');
}*/

/*function onChange(e){
  Logger.log(JSON.stringify(e, null, 2));
}*/

function appBaseTemplates(){
  return [
    USER_START_HERE_SHEET,
    USER_CATEGORIES_SHEET,
    USER_DATA_SHEET,
    USER_BALANCE_HISTORY_SHEET,
    USER_ACCOUNTS_SHEET,
    USER_TRANSACTIONS_SHEET,
    USER_INVESTMENTS_SHEET,
  ];
}

function appFeaturedTemplates(){
  return [
    USER_RECONCILE_SHEET,
    USER_NET_WORTH_SHEET,
    USER_JOINT_NET_WORTH_SHEET,
    USER_MONTHLY_BUDGET_SHEET,
    USER_JOINT_MONTHLY_BUDGET_SHEET,
    USER_YEARLY_BUDGET_SHEET,
    USER_JOINT_YEARLY_BUDGET_SHEET,
    USER_BUDGET_MAKER_SHEET,
    USER_DEFINITION_SHEET
  ];
}

/**
 * Opens the sidebar UI.
 */
function showSidebar() {
  let userValidation = validateUserSession();
  if( userValidation.result && userValidation.result.data.isSubscribed === true ){
    installTemplateInitialSetup();
    showUserDashboardSidebar();
  }else{
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
  var today = Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy");
  return today;
}

function getTodayDateTime(){
  var today = Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy HH:mm:ss");
  return today;
}

function createStripeSession(){

  try{

    const PRICE_ID = 'price_1SnHxFBKorklj30OWLWvqJcP';

    const response = getAppSettings();

    if( response.success === true ){
      
      const stripeKey = response.result.stripeSecretKey;
      const email = Session.getActiveUser().getEmail();

      let url = 'https://api.stripe.com/v1/checkout/sessions';

      var payload =
        'mode=subscription' +
        '&customer_email='+ email + 
        '&success_url=' + encodeURIComponent(API_ENDPOINT+'success?session_id={CHECKOUT_SESSION_ID}') +
        '&cancel_url=' + encodeURIComponent(API_ENDPOINT+'cancel') +
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
      return JSON.parse(request.getContentText()).url;
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
        }
      });
    }
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

      PropertiesService.getUserProperties().setProperty("SYNC_ACTIVATED", "true");
    }
    if( data.sheetEditInstant === true ){
      triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'handleEdit') {
          ScriptApp.deleteTrigger(trigger);
        }
      });

      const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
      ScriptApp.newTrigger('handleEdit')
        .forSpreadsheet(spreadsheetId)
        .onEdit()
        .create();
    }
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

function getConnectedPlaidAccountsTemplate(){
  const response = getAppPlaidConnectedAccounts();
  if( response.success === true ){
    if( response.result.length > 0 ){
      const template = HtmlService.createTemplateFromFile('AccountListCard');
      template.accounts = response.result;
      return template.evaluate().getContent();
    }else{
      return false;
    }
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
      }
    }
    return true;
  }catch(error){
    Logger.log("An error occurred:", JSON.stringify(error, null, 2));
    return false;
  }
}

function handleEdit(e){
  const range = e.range;
  const sheet = range.getSheet();

  const defSheet = UserSpreadsheet.getSheetByName(USER_DEFINITION_SHEET);
  
  // Get the column that was edited
  const editedColumn = range.getColumn();

  const editCell = range.getA1Notation();

  let dropDown;

  try{

    switch( sheet.getName() ){

      case USER_TRANSACTIONS_SHEET:
        let TRAN_CAT_COL_REF = 'I5';
        let TRAN_OWN_COL_REF = 'I11';
        let TRAN_AMT_COL_REF = 'I10';

        // Read the column indices from the Definition sheet (1-based index)
        let catCol = defSheet.getRange(TRAN_CAT_COL_REF).getValue();
        let ownCol = defSheet.getRange(TRAN_OWN_COL_REF).getValue();
        let amtCol = defSheet.getRange(TRAN_AMT_COL_REF).getValue();
        /*const row = range.getRow();
        const col = range.getColumn();

        // 3. Get the formatting from the row above
        const numColumns = sheet.getLastColumn();
        const sourceRange = sheet.getRange(row - 1, 1, 1, numColumns);
        const targetRange = sheet.getRange(row, 1, 1, numColumns);

        // 4. Apply formatting and dropdowns to the current row
        // This will NOT overwrite the text the user just typed.
        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);*/

        if (editedColumn === catCol || editedColumn === ownCol || editedColumn === amtCol ) {
          regenerateAllReports();
        }

      break;
      case USER_MONTHLY_BUDGET_SHEET:
        dropDown = 'C2';
        //Logger.log(editCell);
        if ( editCell === dropDown && typeof populateMonthlyBudget === 'function' ) {
          populateMonthlyBudget();
          regenerateNetWorthReports();
        }
      break;
      case USER_JOINT_MONTHLY_BUDGET_SHEET:
        dropDown = 'C2';
        if ( editCell === dropDown && typeof populateJointMonthlyBudget === 'function' ) {
          populateJointMonthlyBudget();
          regenerateNetWorthReports();
        }
      break;
      case USER_YEARLY_BUDGET_SHEET:
        dropDown = 'E2';
        if ( editCell === dropDown && typeof populateYearlyBudget === 'function' ) {
          populateYearlyBudget();
          regenerateNetWorthReports();
        }
      break;
      case USER_JOINT_YEARLY_BUDGET_SHEET:
        dropDown = 'D2';
        if ( editCell === dropDown && typeof populateJointYearlyBudget === 'function' ) {
          populateJointYearlyBudget();
          regenerateNetWorthReports();
        }
      break;
      case USER_ACCOUNTS_SHEET:
        let balanceCol = defSheet.getRange('F10').getValue();
        let ownerCol = defSheet.getRange('F11').getValue();
        //Logger.log(ownerCol);
        let groupCol = defSheet.getRange('F6').getValue();
        let assliabCol = defSheet.getRange('F7').getValue();
        if (range.getRow() > sheet.getLastRow() && e.value) {
          populateNetWorth();
          populateJointNetWorth();
        }
        if( editedColumn === balanceCol || editedColumn === ownerCol || editedColumn === groupCol || editedColumn === assliabCol ){
          populateNetWorth();
          populateJointNetWorth();
        }
      break;
      case USER_CATEGORIES_SHEET:
        let typeCol = defSheet.getRange('C7').getValue();
        let hiddenCol = defSheet.getRange('C8').getValue();
        let name1Col = defSheet.getRange('C9').getValue();
        let name2Col = defSheet.getRange('C10').getValue();
        //if( editedColumn === typeCol || editedColumn === hiddenCol || editedColumn === name1Col || editedColumn === name2Col ){
          populateMonthlyBudget();
          populateJointMonthlyBudget();
          populateYearlyBudget();
          populateJointYearlyBudget();
        //}
      break;
    }
  }catch( error ){
    Logger.log("Error in onEdit trigger: " + error.toString());
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
  }catch(error){
    Logger.log(`Error while updateAccountName: ${error.message}`);
    return {
      status: false,
      message: "Something went wrong, please try again."
    }
  }
}

function confirmLinkAccountToTemplate( accountId ){

  const userProperties = PropertiesService.getUserProperties();

  const accountName = getAccountNameByAccountId(accountId);

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

  const accountName = getAccountNameByAccountId(accountId);

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
    installFeaturedTemplates();
    linkTransactionSheet(account_id);
    let support_response = checkItemProductSupport( account_id, 'investments');
    if( support_response === true ){
      linkInvestmentSheet(account_id);
    }
    updateAccountBalanceHistory(account_id);
    linkAccountsSheetData(account_id);
    updateAppAccountDetailById(
      account_id,
      {
        is_linked: true,
        status: true,
        updates: false,
        linked_date: getTodayDateTime()
      }
    );
    reApplyFormulaToSpreadsheet();
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

function reApplyFormulaToSpreadsheet(){
  let transactionSheet = UserSpreadsheet.getSheetByName('Transactions');
  if( transactionSheet ){
    if( transactionSheet.getLastRow() > 2 ){
      var spreadsheet = UserSpreadsheet;
      let formulaSheets = [ USER_DEFINITION_SHEET, USER_MONTHLY_BUDGET_SHEET, USER_JOINT_MONTHLY_BUDGET_SHEET, USER_YEARLY_BUDGET_SHEET, USER_JOINT_YEARLY_BUDGET_SHEET, USER_BUDGET_MAKER_SHEET, USER_TRANSACTIONS_SHEET ];
      formulaSheets.forEach( function(item){
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
      });
    }
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
          }

          if( sheetName === USER_DEFINITION_SHEET ){
            hideSheetByName(sheetName);
            protectSheetByName(sheetName);
          }
        });
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

/**
 * Step 4: Polling function called by the UI to check the flag status
 */
function checkTaskStatus() {
  const status = PropertiesService.getUserProperties().getProperty('TASK_STATUS');
  return status;
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
}

function resetAccountOwnerCell(accountName){
  if (accountName == "") {
    return;
  }
}

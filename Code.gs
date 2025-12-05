//User DATA
const UserEmail = Session.getActiveUser().getEmail();
const UserSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const UserSpreadsheetUrl = UserSpreadsheet.getUrl();
const UserSpreadsheetId = UserSpreadsheet.getId();

const APP_TITLE = "ThefinU";
const APP_TEMPLATES_SPREADSHEET = 'https://docs.google.com/spreadsheets/d/1KqN_Tirz58lw9rkoPSSzRN4wcYy233QweNz0zfLWLq8/edit?usp=sharing';

// Configuration
const RESTAPI_CONFIG = {
    API_BASE_URL: 'https://stallioni.com/php_rest_api.php',
    SCRIPT_ID: 'AKfycbyqPA2eaEAgwNGdVyOAbIeq3_h74nGaujmk80lomXVrErl-948LuTWr9F3rRxPEUX_mhA',
    TIMEOUT: 30000 // 30 seconds
};

const WEBAPP_WEBHOOK = RESTAPI_CONFIG.API_BASE_URL + '?action=webhook';

const API_CONFIG = {
    scriptId: 'AKfycbyqPA2eaEAgwNGdVyOAbIeq3_h74nGaujmk80lomXVrErl-948LuTWr9F3rRxPEUX_mhA', // Replace with your backend script ID
    //scriptId: 'AKfycbyJ9OMP9xhszurIN8d4INF6UnNGOvdEgIp8hSmFp4U',
    baseUrl: 'https://script.googleapis.com/v1/scripts/'
};

var APP_USER_ID = {
    client_user_id: generateRandomNumber(),
};

//const Admin_Email = "chrisannaelser@gmail.com";
const Admin_Email = UserEmail;
const APP_Email = 'anna@thefinu.com';

//SHEETS
const APP_SPREADSHEET_ID = '1KqN_Tirz58lw9rkoPSSzRN4wcYy233QweNz0zfLWLq8';
const APP_USER_DETAILS_SHEET = 'User Data';

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
const USER_JOINT_MONTHLY_BUDGET_SHEET = 'Joint Monthly Budget';
const USER_YEARLY_BUDGET_SHEET = 'Yearly Budget';
const USER_JOINT_YEARLY_BUDGET_SHEET = 'Joint Yearly Budget';
const USER_BUDGET_MAKER_SHEET = 'Budget Maker';

const APP_BASE_TEMPLATES = [
    USER_START_HERE_SHEET,
    USER_CATEGORIES_SHEET,
    USER_DATA_SHEET,
    USER_BALANCE_HISTORY_SHEET,
    USER_ACCOUNTS_SHEET,
    USER_TRANSACTIONS_SHEET,
    USER_INVESTMENTS_SHEET,
    USER_PLAID_SHEET,
];

const APP_FEATURED_TEMPLATES = [
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

const DEFUALT_TEMPLATES = [
    USER_DEFINITION_SHEET,
    USER_START_HERE_SHEET,
    USER_CATEGORIES_SHEET,
    USER_DATA_SHEET,
    USER_BALANCE_HISTORY_SHEET,
    USER_ACCOUNTS_SHEET,
    USER_TRANSACTIONS_SHEET,
    USER_INVESTMENTS_SHEET,
    USER_RECONCILE_SHEET,
    USER_PLAID_SHEET,
    USER_NET_WORTH_SHEET,
    USER_JOINT_NET_WORTH_SHEET,
    USER_MONTHLY_BUDGET_SHEET,
    USER_JOINT_MONTHLY_BUDGET_SHEET,
    USER_YEARLY_BUDGET_SHEET,
    USER_JOINT_YEARLY_BUDGET_SHEET,
    USER_BUDGET_MAKER_SHEET,
];

const PLAID_DATA_ACCOUNT_ID = 0;
const PLAID_DATA_ACCESS_TOKEN = 1;
const PLAID_DATA_ITEM_ID = 2;
const PLAID_DATA_INSTITUTION_ID = 3;
const PLAID_DATA_MASK = 4;
const PLAID_DATA_ACCOUNT_NAME = 5;
const PLAID_DATA_IS_LINKED = 6;
const PLAID_DATA_LINKED_DATE = 7;
const PLAID_DATA_NEXT_CURSOR = 8;
const PLAID_DATA_STATUS = 9;
const PLAID_UPDATES_AVAILABLE = 10;

function openAppHomeScreen(e) {
    const card = CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader()
            .setTitle("Welcome to ThefinU")
            .setSubtitle(getAppIntroductionContent())
        ).addSection(CardService.newCardSection()
            .addWidget(CardService.newTextButton()
                .setText("Get Start Here")
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("getStartApp"))))
        .build();
    return card;
}

function getStartApp() {
    activateDailyAutoUpdate();
    createOnEditTrigger();
    SpreadsheetApp.getUi().alert("App setup being progress. Please wait..");
    installBaseTemplatesIfNotExist();
    installFeaturedTemplates();
    reApplyFormula();
    SpreadsheetApp.getUi().alert("Thank you. App setup is completed.");
    onOpen();
    showSidebar();
}

function onOpen() {
    SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem('Open', 'showSidebar')
        .addItem('Activate Daily Update', 'activateDailyAutoUpdate')
        .addToUi();
}

function onInstall(e) {
    //activateDailyAutoUpdate();
}

function showSidebar() {
    if (checkAppLinkedAccounts() == true) {
        var template = HtmlService.createTemplateFromFile('main');
    } else {
        var template = HtmlService.createTemplateFromFile('index');
    }
    var html = template.evaluate();
    html.setTitle(APP_TITLE);
    SpreadsheetApp.getUi().showSidebar(html);
}

function installBaseTemplatesIfNotExist() {
    let existingSheets = UserSpreadsheet.getSheets().map(sheet => sheet.getName()); // Get all sheet names
    let requiredSheets = APP_BASE_TEMPLATES; // Add your required sheets here
    // Check if all required sheets exist
    var allSheetsExist = requiredSheets.every(sheetName => existingSheets.includes(sheetName));
    if (allSheetsExist === false) {
        //SpreadsheetApp.getUi().alert("Please wait required template(s) installing..");
        installBaseTemplate(APP_BASE_TEMPLATES);
        //SpreadsheetApp.getUi().alert("Templates installed successfully.");
        if (isActiveTemplate(USER_START_HERE_SHEET) === true) {
            SpreadsheetApp.getActive().getSheetByName(USER_START_HERE_SHEET).activate();
        }
    }
}

function installBaseTemplate(requiredSheets = []) {

    try {

        var destinationSpreadsheet = UserSpreadsheet;
        var sourceSpreadsheet = SpreadsheetApp.openByUrl(APP_TEMPLATES_SPREADSHEET);
        var sourceSheets = sourceSpreadsheet.getSheets();

        // Convert requiredSheets to a Set for faster lookup
        const requiredSet = new Set(requiredSheets.map(name => name.toLowerCase()));

        sourceSheets.forEach(function (sourceSheet) {

            var sheetName = sourceSheet.getName();

            // Only copy sheets that are in the requiredSheets array
            if (!requiredSet.has(sheetName.toLowerCase())) {
                //Logger.log(`Skipped non-required sheet: ${sheetName}`);
                return;
            }

            // Delete if sheet already exists in destination
            var existingSheet = destinationSpreadsheet.getSheetByName(sheetName);

            if (!existingSheet) {
                // Copy the sheet
                let copiedSheet = sourceSheet.copyTo(destinationSpreadsheet);
                // Rename copied sheet
                copiedSheet.setName(sheetName);

                if (sheetName === USER_PLAID_SHEET) {
                    hideSheetByName(sheetName);
                }
            }
            //Logger.log(`Successfully copied required sheet: ${sheetName}`);
        });

    } catch (error) {
        Logger.log(`Error in installBaseTemplate: ${error.message}`);
        throw new Error(`Failed to install template: ${error.message}`);
    }
}

function installFeaturedTemplates() {

    let transactionSheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
    if (transactionSheet) {
        if (transactionSheet.getLastRow() > 2) {
            var requiredSheets = APP_FEATURED_TEMPLATES;
            var destinationSpreadsheet = UserSpreadsheet;
            var sourceSpreadsheet = SpreadsheetApp.openByUrl(APP_TEMPLATES_SPREADSHEET);
            var sourceSheets = sourceSpreadsheet.getSheets();

            // Convert requiredSheets to a Set for faster lookup
            const requiredSet = new Set(requiredSheets.map(name => name.toLowerCase()));

            sourceSheets.forEach(function (sourceSheet) {
                var sheetName = sourceSheet.getName();
                // Only copy sheets that are in the requiredSheets array
                if (!requiredSet.has(sheetName.toLowerCase())) {
                    //Logger.log(`Skipped non-required sheet: ${sheetName}`);
                    return;
                }
                // Delete if sheet already exists in destination
                var existingSheet = destinationSpreadsheet.getSheetByName(sheetName);
                if (!existingSheet) {
                    // Copy the sheet
                    var copiedSheet = sourceSheet.copyTo(destinationSpreadsheet);
                    // Rename copied sheet
                    copiedSheet.setName(sheetName);
                }

                if (sheetName === USER_DEFINITION_SHEET) {
                    hideSheetByName(sheetName);
                    protectSheetByName(sheetName);
                }

            });

        }
    }
}

function getTodayDate() {
    var today = Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy");
    return today;
}

function getTodayDateTime() {
    var today = Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy HH:mm:ss");
    return today;
}

function setAppTaskType(type) {
    let documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('app_task_type', type);
}

function getAppTaskType() {
    let documentProperties = PropertiesService.getDocumentProperties();
    return documentProperties.getProperty('app_task_type');
}

function setAppTaskStatus(status) {
    let documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('app_task_status', status);
}

function getAppTaskStatus() {
    let documentProperties = PropertiesService.getDocumentProperties();
    return documentProperties.getProperty('app_task_status');
}

function setAppTaskAccountId(account_id) {
    let documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('app_task_account_id', account_id);
}

function getAppTaskAccountId() {
    let documentProperties = PropertiesService.getDocumentProperties();
    return documentProperties.getProperty('app_task_account_id');
}

function setAppCurrentAccountId(account_id) {
    let documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('app_current_account_id', account_id);
}

function getAppCurrentAccountId() {
    let documentProperties = PropertiesService.getDocumentProperties();
    return documentProperties.getProperty('app_current_account_id');
}

function setPlaidErrorHandler(data) {
    let documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('plaid_error_handler', data);
}

function getPlaidErrorHandler() {
    let documentProperties = PropertiesService.getDocumentProperties();
    return documentProperties.getProperty('plaid_error_handler');
}

function showAppFashAlertMessage(message) {
    return SpreadsheetApp.getUi().alert(message);
}

function generateRandomNumber() {
    const min = 10000; // Minimum 5-digit number
    const max = 99999; // Maximum 5-digit number
    let number = Math.floor(Math.random() * (max - min + 1)) + min;
    return number.toString();
}

function renderTemplate(menutab) {
    return HtmlService.createTemplateFromFile(menutab).evaluate().getContent();
}

function initPlaidConnect() {
    const html = HtmlService.createHtmlOutputFromFile('connectPlaid').setWidth(450).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, "Connect Plaid Account");
}

function connectPlaidAccount() {
    SpreadsheetApp.getUi().alert("Please wait theFinU connecting with plaid account.");
    return false;
}

function isActiveTemplate(sheetName) {
    let spreadsheet = UserSpreadsheet.getSheetByName(sheetName);
    if (spreadsheet != null) {
        return true;
    } else {
        return false;
    }
}

function hideSheetByName(sheetName) {
    const sheet = UserSpreadsheet.getSheetByName(sheetName);
    if (sheet) {
        sheet.hideSheet();
    }
}

/*function checkTemplateHidden(sheetName){
  let spreadsheet = UserSpreadsheet.getSheetByName(sheetName);
  if( spreadsheet != null ){
    return spreadsheet.isSheetHidden();
  }else{
    return false;
  }
}*/

function updateUserPlaidDataSheet(row, data) {
    let sheet = UserSpreadsheet.getSheetByName(USER_PLAID_SHEET);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var rowValues = [];

    for (var j = 0; j < headers.length; j++) {
        var header = headers[j];
        var value = "";
        switch (header) {
            case "Account ID":
                value = data.account_id != '' ? data.account_id : '';
                break;
            case "Access Token":
                value = data.access_token != '' ? data.access_token : '';
                break;
            case "Item ID":
                value = data.item_id != '' ? data.item_id : '';
                break;
            case "Institution ID":
                value = data.institution_id != '' ? data.institution_id : '';
                break;
            case "Mask":
                value = data.mask != '' ? data.mask : '';
                break;
            case "Account Name":
                value = data.account_name != '' ? data.account_name : '';
                break;
            /*case "Institution":
              value = data.institution !='' ? data.institution : '';
              break;*/
            case "Is Linked":
                value = data.linked != '' ? data.linked : false;
                break;
            case "Linked Date":
                value = data.linked_date != '' ? data.linked_date : '';
                break;
            case "Next Cursor":
                value = data.next_cursor != '' ? data.next_cursor : '';
                break;
            case "Status":
                value = data.status != '' ? data.status : false;
                break;
            case "Updates":
                value = data.updates != '' ? data.updates : false;
                break;
        }
        rowValues.push(value);
    }

    var cell = sheet.getRange(row, 1, 1, sheet.getLastColumn());
    cell.setValues([rowValues]);
    cell.setFontSize(9)
        .setFontFamily("Comfortaa")
        .setFontColor("#000000")
        .setFontWeight("bold");
}

function addUserPlaidDataSheet(data = null) {

    let sheet = UserSpreadsheet.getSheetByName(USER_PLAID_SHEET);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let lastrow = sheet.getLastRow() + 1;

    var rowValues = [];

    for (var j = 0; j < headers.length; j++) {
        var header = headers[j];
        var value = "";
        switch (header) {
            case "Account ID":
                value = data.account_id != '' ? data.account_id : '';
                break;
            case "Access Token":
                value = data.access_token != '' ? data.access_token : '';
                break;
            case "Item ID":
                value = data.item_id != '' ? data.item_id : '';
                break;
            case "Institution ID":
                value = data.institution_id != '' ? data.institution_id : '';
                break;
            case "Mask":
                value = data.mask != '' ? data.mask : '';
                break;
            case "Account Name":
                value = data.account_name != '' ? data.account_name : '';
                break;
            /*case "Institution":
              value = data.institution !='' ? data.institution : '';
              break;*/
            case "Is Linked":
                value = data.linked != '' ? data.linked : false;
                break;
            case "Linked Date":
                value = data.linked_date != '' ? data.linked_date : '';
                break;
            case "Next Cursor":
                value = data.next_cursor != '' ? data.next_cursor : '';
                break;
            case "Status":
                value = data.status != '' ? data.status : true;
                break;
            case "Updates":
                value = data.updates != '' ? data.updates : false;
                break;

        }
        rowValues.push(value);
    }

    var cell = sheet.getRange(lastrow, 1, 1, sheet.getLastColumn());
    cell.setValues([rowValues]);
    cell.setFontSize(9)
        .setFontFamily("Comfortaa")
        .setFontColor("#000000")
        .setFontWeight("bold");
}

function getPlaidSheetAccountRow(account_id) {

    let spreadsheet = UserSpreadsheet.getSheetByName(USER_PLAID_SHEET);
    var data = spreadsheet.getDataRange().getValues();

    // Loop through rows to find the value
    for (var row = 0; row < data.length; row++) {
        if (data[row].includes(account_id)) {
            return row + 1; // Return the row number (1-based index)
        }
    }

    return null;
}

function checkInstitutionExist(institution_id) {
    let spreadsheet = UserSpreadsheet.getSheetByName(USER_PLAID_SHEET);
    var data = spreadsheet.getDataRange().getValues();

    // Loop through rows to find the value
    for (var row = 0; row < data.length; row++) {
        if (data[row].includes(institution_id)) {
            return row + 1; // Return the row number (1-based index)
        }
    }

    return null;
}

function checkAppLinkedAccounts() {
    try {
        let spreadsheet = UserSpreadsheet.getSheetByName(USER_PLAID_SHEET);
        var data = spreadsheet.getDataRange().getValues();
        if (data.length > 1) {
            return true;
        } else {
            return false;
        }
    } catch (e) {
        Logger.log("Error:" + e);
        return false;
    }

}

function checkSyncStatus() {
    let spreadsheet = UserSpreadsheet.getSheetByName(USER_PLAID_SHEET);
    var data = spreadsheet.getDataRange().getValues();
    var status = false;
    if (data.length > 1) {
        data.splice(0, 1);
        data.forEach(function (item, index) {
            if (item[PLAID_UPDATES_AVAILABLE] == true) {
                status = item[PLAID_UPDATES_AVAILABLE];
            }
        });
    }

    return status;
}

function isTriggerComplete() {
    return getAppTaskStatus();
}

function installSpreadSheet(sheetName) {
    var source = SpreadsheetApp.openByUrl(APP_TEMPLATES_SPREADSHEET);
    var destination = UserSpreadsheet;
    var tempsheet = source.getSheetByName(sheetName);
    var copiedSheet = tempsheet.copyTo(destination);
    copiedSheet.setName(sheetName);
    return true;
}

function fetchPlaidLinkedAccounts() {

    let spreadsheet = UserSpreadsheet.getSheetByName(USER_PLAID_SHEET);
    var data = spreadsheet.getDataRange().getValues();
    //Removed sheet title from the collection
    data.splice(0, 1);

    var collection = [];

    if (data.length > 0) {
        data.forEach(function (item, index) {
            let account = getPlaidAccountsById(item[PLAID_DATA_ACCESS_TOKEN], item[PLAID_DATA_ACCOUNT_ID]);
            if (account.hasOwnProperty('error_code') == true) {
                SpreadsheetApp.getUi().alert("The plaid item needs to be reauthenticate.");
                setPlaidErrorHandler({
                    error_code: account.error_code,
                    account_id: item[PLAID_DATA_ACCOUNT_ID],
                    access_token: item[PLAID_DATA_ACCESS_TOKEN]
                });
                handlingItemRequiredLogin();
            } else {
                if (account?.accounts?.length > 0) {
                    account.accounts.every(function (user) {
                        collection.push({
                            'account_id': user.account_id,
                            'name': item[PLAID_DATA_ACCOUNT_NAME] != '' ? item[PLAID_DATA_ACCOUNT_NAME] : user.name,
                            'official_name': user.official_name,
                            'mask': user.mask,
                            'subtype': user.subtype,
                            'type': user.type,
                            'institution': {
                                'institution_id': item[PLAID_DATA_INSTITUTION_ID],
                                'name': account.item.institution_name
                            },
                            'balances': user.balances,
                            'is_linked': item[PLAID_DATA_IS_LINKED],
                            'status': item[PLAID_DATA_STATUS]
                        });
                    })
                }
            }
        });
        return collection;
    }
}

function getCurrentUserAccount(account_id) {

    var collection = [];

    if (getUserAccountPlaidDataByAccountId(account_id) != null) {
        let data = getUserAccountPlaidDataByAccountId(account_id);
        let account = getPlaidAccountsById(data[PLAID_DATA_ACCESS_TOKEN], data[PLAID_DATA_ACCOUNT_ID]);
        if (account.accounts.length > 0) {
            account.accounts.forEach(function (item) {
                collection.push({
                    'account_id': item.account_id,
                    'name': data[PLAID_DATA_ACCOUNT_NAME] != '' ? data[PLAID_DATA_ACCOUNT_NAME] : item.name,
                    'official_name': item.official_name,
                    'mask': item.mask,
                    'subtype': item.subtype,
                    'type': item.type,
                    'institution': {
                        'institution_id': data[PLAID_DATA_INSTITUTION_ID],
                        'name': account.item.institution_name
                    },
                    'balances': item.balances,
                    'is_linked': data[PLAID_DATA_IS_LINKED],
                    'status': data[PLAID_DATA_STATUS]
                });
            });
        }
    }

    return collection;
}

function getUserAccountPlaidDataByAccountId(account_id) {

    let spreadsheet = UserSpreadsheet.getSheetByName(USER_PLAID_SHEET);
    var data = spreadsheet.getDataRange().getValues();

    // Loop through rows to find the value
    for (var row = 0; row < data.length; row++) {
        if (data[row].includes(account_id)) {
            return data[row]; // Return the row number (1-based index)
        }
    }

    return null;
}

function getUserAccountPlaidDataByItemId(item_id) {

    let spreadsheet = UserSpreadsheet.getSheetByName(USER_PLAID_SHEET);
    var data = spreadsheet.getDataRange().getValues();
    var collection = [];
    // Loop through rows to find the value
    for (var row = 0; row < data.length; row++) {
        if (data[row].includes(item_id)) {
            collection.push(data[row]); // Return the row number (1-based index)
        }
    }
    if (collection.length > 0) {
        return collection;
    }

    return null;
}

function updateUserAccountName(account_name) {

    let account_id = getAppCurrentAccountId();
    let accountData = getCurrentUserAccount(account_id);

    if (accountData.length > 0) {
        accountData.forEach(function (account) {
            if (getPlaidSheetAccountRow(account.account_id) != null) {
                let row = getPlaidSheetAccountRow(account.account_id);
                var sheet_data = getUserAccountPlaidDataByAccountId(account.account_id);
                let data = {
                    'account_id': account.account_id,
                    'access_token': sheet_data[PLAID_DATA_ACCESS_TOKEN],
                    'item_id': sheet_data[PLAID_DATA_ITEM_ID],
                    'institution_id': sheet_data[PLAID_DATA_INSTITUTION_ID],
                    'mask': sheet_data[PLAID_DATA_MASK],
                    'account_name': account_name,
                    'linked': sheet_data[PLAID_DATA_IS_LINKED],
                    'linked_date': sheet_data[PLAID_DATA_LINKED_DATE],
                    'next_cursor': sheet_data[PLAID_DATA_NEXT_CURSOR],
                    'status': sheet_data[PLAID_DATA_STATUS],
                    'updates': sheet_data[PLAID_UPDATES_AVAILABLE]
                };
                updateUserPlaidDataSheet(row, data);
                changeNameOfAccountOnBalanceHistorySheet(account.account_id, account_name);
                changeNameOfAccountOnTransactionSheet(account.account_id, account_name);
                changeNameOfAccountOnInvestmentSheet(account.account_id, account_name);
                SpreadsheetApp.getUi().alert("Account Name Changed.");
            }
        });
        return true;
    }
    return false;
}

function updateUserAccountNextCursor(account_id, next_cursor) {

    let accountData = getCurrentUserAccount(account_id);

    if (accountData.length > 0) {
        accountData.forEach(function (account) {
            if (getPlaidSheetAccountRow(account.account_id) != null) {
                let row = getPlaidSheetAccountRow(account.account_id);
                var sheet_data = getUserAccountPlaidDataByAccountId(account.account_id);
                let data = {
                    'account_id': account.account_id,
                    'access_token': sheet_data[PLAID_DATA_ACCESS_TOKEN],
                    'item_id': sheet_data[PLAID_DATA_ITEM_ID],
                    'institution_id': sheet_data[PLAID_DATA_INSTITUTION_ID],
                    'mask': sheet_data[PLAID_DATA_MASK],
                    'account_name': sheet_data[PLAID_DATA_ACCOUNT_NAME],
                    'linked': sheet_data[PLAID_DATA_IS_LINKED],
                    'linked_date': getTodayDateTime(),
                    'next_cursor': next_cursor,
                    'status': sheet_data[PLAID_DATA_STATUS],
                    'updates': sheet_data[PLAID_UPDATES_AVAILABLE]
                };
                updateUserPlaidDataSheet(row, data);
            }
        });
    }
}

function confirmLinkAccount() {

    let account_id = getAppCurrentAccountId();
    let account_name = getUserAccountNameByAccountId(account_id);
    var result = SpreadsheetApp.getUi().alert(
        'Insert historical data?',
        "New updates for '" + account_name + "' will now sync with this spreadsheet. Do you wish to insert this account's historical data too? Clicking 'Yes' will add historical transactions and balances for '" + account_name + "' to the current spreadsheet.",
        SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    if (result == SpreadsheetApp.getUi().Button.YES) {
        SpreadsheetApp.getUi().alert('The process takes few minutes to be completed. Please wait do not change anything untill the process complete.');
        return true;
    } else {
        return false;
    }
}

function confirmUnlinkAccount() {

    let account_id = getAppCurrentAccountId();

    let account_name = getUserAccountNameByAccountId(account_id);

    var result = SpreadsheetApp.getUi().alert(
        'Remove existing data?',
        "'" + account_name + "' will no longer sync with this spreadsheet. Do you wish to remove this account's existing data too? Clicking 'Yes' will remove transactions and balances for '" + account_name + "' from the current spreadsheet.",
        SpreadsheetApp.getUi().ButtonSet.YES_NO);

    if (result == SpreadsheetApp.getUi().Button.YES) {
        SpreadsheetApp.getUi().alert('The process takes few minutes to be completed. Please wait do not change anything untill the process complete.');
        return true;
    }
}

function confirmRemoveLinkedAccount() {

    var result = SpreadsheetApp.getUi().alert(
        'Want to remove account?',
        "If you remove the account, the data of the account will lost.",
        SpreadsheetApp.getUi().ButtonSet.YES_NO);

    if (result == SpreadsheetApp.getUi().Button.YES) {
        SpreadsheetApp.getUi().alert('The process takes few minutes to be completed. Please wait do not change anything untill the process complete.');
        removelinkAccount();
        return true;
    }
}

function linkDatatoSheet() {
    let account_id = getAppCurrentAccountId();
    setAppTaskType('link');
    setAppTaskStatus('started');
    processAppTask(account_id);
    return true;
}


function confirmSyncPlaid() {

    var result = SpreadsheetApp.getUi().alert(
        'Want to sync account?',
        "The process takes few minutes to be completed. Please wait do not change anything untill the process complete.",
        SpreadsheetApp.getUi().ButtonSet.YES_NO);
    if (result == SpreadsheetApp.getUi().Button.YES) {
        setAppTaskType('update');
        setAppTaskStatus('started');
        return true;
    }
}

function syncPlaidAccounts() {
    try {
        let spreadsheet = UserSpreadsheet.getSheetByName(USER_PLAID_SHEET);
        var data = spreadsheet.getDataRange().getValues();
        data.splice(0, 1);
        // Loop through rows to find the value
        if (data.length > 0) {
            data.forEach(function (item, index) {
                if (item[PLAID_UPDATES_AVAILABLE] == true) {
                    let account_id = item[PLAID_DATA_ACCOUNT_ID];
                    updateTransactionSheet(account_id);
                    //updateBalanceHistoryLastTransaction(account_id);
                    let support_response = checkItemProductSupport(account_id, 'investments');
                    if (support_response === true) {
                        updateInvestmentSheet(account_id);
                        //updateInvestmentBalanceHistroy( account_id );
                    }
                    updateAccountBalanceHistory(account_id);
                    let row = getPlaidSheetAccountRow(account_id);
                    var sheet_data = getUserAccountPlaidDataByAccountId(account_id);
                    let row_data = {
                        'account_id': sheet_data[PLAID_DATA_ACCOUNT_ID],
                        'access_token': sheet_data[PLAID_DATA_ACCESS_TOKEN],
                        'item_id': sheet_data[PLAID_DATA_ITEM_ID],
                        'institution_id': sheet_data[PLAID_DATA_INSTITUTION_ID],
                        'mask': sheet_data[PLAID_DATA_MASK],
                        'account_name': sheet_data[PLAID_DATA_ACCOUNT_NAME],
                        'linked': sheet_data[PLAID_DATA_IS_LINKED],
                        'linked_date': getTodayDateTime(),
                        'next_cursor': sheet_data[PLAID_DATA_NEXT_CURSOR],
                        'status': sheet_data[PLAID_DATA_STATUS],
                        'updates': false
                    };
                    updateUserPlaidDataSheet(row, row_data);
                }
            });
        }
        return true;
    } catch (error) {
        MailApp.sendEmail(
            Admin_Email,
            'ThefinuBeta - Plaid sync Aborted.',
            'Something went wrong while trying to sync the plaid account(s).',
        );
        Logger.log("An error occurred:", JSON.stringify(error, null, 2));
        return false;
    }
}

function appTaskStatus(status) {

    if (status == 'completed') {
        let task = getAppTaskType();
        switch (task) {
            case 'link':
                setAppTaskStatus('');
                setAppTaskAccountId('');
                setAppTaskType('');
                MailApp.sendEmail(
                    Admin_Email,
                    'ThefinuBeta - Task Completed',
                    'Hi, Your plaid account(s) has been linked successfully. Please check the spreadsheet.',
                );
                SpreadsheetApp.getUi().alert('The task is completed and thank you for the patience.');
                break;
            case 'unlink':
                setAppTaskStatus('');
                setAppTaskAccountId('');
                setAppTaskType('');
                MailApp.sendEmail(
                    Admin_Email,
                    'ThefinuBeta - Task Completed',
                    'Hi, Your plaid account(s) has been unlinked successfully. Please check the spreadsheet.',
                );
                SpreadsheetApp.getUi().alert('The task is completed and thank you for the patience.');
                break;
            case 'remove':
                setAppTaskStatus('');
                setAppTaskAccountId('');
                setAppTaskType('');
                MailApp.sendEmail(
                    Admin_Email,
                    'ThefinuBeta - Task Completed',
                    'Hi, Your plaid account(s) has been removed successfully. Please check the spreadsheet.',
                );
                SpreadsheetApp.getUi().alert('The task is completed and thank you for the patience.');
                break;
            case 'update':
                setAppTaskStatus('');
                setAppTaskAccountId('');
                setAppTaskType('');
                MailApp.sendEmail(
                    Admin_Email,
                    'ThefinuBeta - Task Completed',
                    'Hi, Your plaid account(s) has been updated successfully. Please check the spreadsheet.',
                );
                SpreadsheetApp.getUi().alert('The task is completed and thank you for the patience.');
                break;
        }
    }
}

function unlinkDatatoSheet() {
    let account_id = getAppCurrentAccountId();
    setAppTaskAccountId(account_id);
    setAppTaskType('unlink');
    setAppTaskStatus('started');
    processAppTask(account_id);
    return true;
}

function removelinkAccount() {
    let account_id = getAppCurrentAccountId();
    setAppTaskAccountId(account_id);
    setAppTaskType('remove');
    setAppTaskStatus('started');
    processAppTask(account_id);
    return true;
}

function processAppTask(account_id = null) {
    let task = getAppTaskType();
    switch (task) {
        case 'link':
            linkTransactionSheet(account_id);
            let support_response = checkItemProductSupport(account_id, 'investments');
            if (support_response === true) {
                linkInvestmentSheet(account_id);
            }
            updateAccountBalanceHistory(account_id);
            linkUserAccount(account_id);
            installFeaturedTemplates();
            reApplyFormula();
            setAppTaskStatus('completed');
            break;
        case 'unlink':
            clearTransactionsData(account_id);
            clearInvestmentsData(account_id);
            clearBalanceHitoryData(account_id);
            clearAccountData(account_id);
            unlinkUserAccount(account_id);
            setAppTaskStatus('completed');
            break;
        case 'remove':
            clearTransactionsData(account_id);
            clearInvestmentsData(account_id);
            clearBalanceHitoryData(account_id);
            clearAccountData(account_id);
            removelinkUserAccount(account_id);
            setAppTaskStatus('completed');
            break;
        case 'update':
            syncPlaidAccounts();
            setAppTaskStatus('completed');
            break;
    }
}

function getUserAccessTokenByAccountId(account_id) {

    if (getUserAccountPlaidDataByAccountId(account_id) != null) {
        let data = getUserAccountPlaidDataByAccountId(account_id);
        return data[PLAID_DATA_ACCESS_TOKEN];
    }
    return null;
}

function getUserAccountNameByAccountId(account_id) {
    if (getUserAccountPlaidDataByAccountId(account_id) != null) {
        let data = getUserAccountPlaidDataByAccountId(account_id);
        return data[PLAID_DATA_ACCOUNT_NAME];
    }
    return null;
}

function updateUserPlaidTokens(token) {

    if (isActiveTemplate(USER_PLAID_SHEET) === true) {
        let response = getPlaidAccounts(token.access_token);
        if (response.request_id && response.request_id != '') {
            let accounts = response.accounts;
            let item = response.item;
            let institution_id = item.institution_id;
            let institution_name = item.institution_name;
            if (accounts.length > 0) {
                accounts.forEach(function (account) {
                    if (getUserAccountPlaidDataByAccountId(account.account_id) != null) {
                        let row = getPlaidSheetAccountRow(account.account_id);
                        var sheet_data = getUserAccountPlaidDataByAccountId(account.account_id);
                        let data = {
                            'account_id': account.account_id,
                            'access_token': token.access_token,
                            'item_id': token.item_id,
                            'institution_id': institution_id,
                            'mask': account.mask,
                            'account_name': sheet_data[PLAID_DATA_ACCOUNT_NAME],
                            'institution': institution_name,
                            'linked': sheet_data[PLAID_DATA_IS_LINKED],
                            'linked_date': getTodayDateTime(),
                            'next_cursor': sheet_data[PLAID_DATA_NEXT_CURSOR],
                            'status': sheet_data[PLAID_DATA_STATUS],
                            'updates': sheet_data[PLAID_UPDATES_AVAILABLE]
                        };
                        //Logger.log(data);
                        updateUserPlaidDataSheet(row, data);
                    } else {
                        let data = {
                            'account_id': account.account_id,
                            'access_token': token.access_token,
                            'item_id': token.item_id,
                            'institution_id': institution_id,
                            'mask': account.mask,
                            'account_name': account.name,
                            'institution': institution_name,
                            'linked': false,
                            'linked_date': '',
                            'next_cursor': '',
                            'status': false,
                            'updates': false,
                        };
                        //Logger.log(data);
                        addUserPlaidDataSheet(data);
                    }
                });
            }
            return true;
        }
    }
}

function linkUserAccount(account_id) {
    if (getPlaidSheetAccountRow(account_id) != null) {
        let row = getPlaidSheetAccountRow(account_id);
        var sheet_data = getUserAccountPlaidDataByAccountId(account_id);
        let data = {
            'account_id': sheet_data[PLAID_DATA_ACCOUNT_ID],
            'access_token': sheet_data[PLAID_DATA_ACCESS_TOKEN],
            'item_id': sheet_data[PLAID_DATA_ITEM_ID],
            'institution_id': sheet_data[PLAID_DATA_INSTITUTION_ID],
            'mask': sheet_data[PLAID_DATA_MASK],
            'account_name': sheet_data[PLAID_DATA_ACCOUNT_NAME],
            'linked': true,
            'linked_date': getTodayDateTime(),
            'next_cursor': sheet_data[PLAID_DATA_NEXT_CURSOR],
            'status': true,
            'updates': false,
        };
        updateUserPlaidDataSheet(row, data);
    }
}

function unlinkUserAccount(account_id) {
    if (getPlaidSheetAccountRow(account_id) != null) {
        let row = getPlaidSheetAccountRow(account_id);
        var sheet_data = getUserAccountPlaidDataByAccountId(account_id);
        let data = {
            'account_id': sheet_data[PLAID_DATA_ACCOUNT_ID],
            'access_token': sheet_data[PLAID_DATA_ACCESS_TOKEN],
            'item_id': sheet_data[PLAID_DATA_ITEM_ID],
            'institution_id': sheet_data[PLAID_DATA_INSTITUTION_ID],
            'mask': sheet_data[PLAID_DATA_MASK],
            'account_name': sheet_data[PLAID_DATA_ACCOUNT_NAME],
            'linked': false,
            'linked_date': getTodayDateTime(),
            'next_cursor': '',
            'status': true,
            'updates': false,
        };
        updateUserPlaidDataSheet(row, data);
    }
}

function removelinkUserAccount(account_id) {
    if (getPlaidSheetAccountRow(account_id) != null) {
        let row = getPlaidSheetAccountRow(account_id);
        const sheet = UserSpreadsheet.getSheetByName(USER_PLAID_SHEET);
        sheet.deleteRow(row);
    }
}

function setItemUpdateAvailable(item_id) {
    let sheetdata = getUserAccountPlaidDataByItemId(item_id);
    if (sheetdata != null && sheetdata.length > 0) {
        sheetdata.forEach(function (item) {
            let row = getPlaidSheetAccountRow(item[PLAID_DATA_ACCOUNT_ID]);
            let data = {
                'account_id': item[PLAID_DATA_ACCOUNT_ID],
                'access_token': item[PLAID_DATA_ACCESS_TOKEN],
                'item_id': item[PLAID_DATA_ITEM_ID],
                'institution_id': item[PLAID_DATA_INSTITUTION_ID],
                'mask': item[PLAID_DATA_MASK],
                'account_name': item[PLAID_DATA_ACCOUNT_NAME],
                'linked': item[PLAID_DATA_IS_LINKED],
                'linked_date': item[PLAID_DATA_LINKED_DATE],
                'next_cursor': item[PLAID_DATA_NEXT_CURSOR],
                'status': item[PLAID_DATA_STATUS],
                'updates': true
            };
            updateUserPlaidDataSheet(row, data);
        });
    }
}

function defaultTemplateList() {
    var title = '';
    var description = '';
    var slug = '';
    var data = [];
    DEFUALT_TEMPLATES.forEach(function (item, index) {
        switch (item) {
            case USER_CATEGORIES_SHEET:
                title = item;
                description = "A platform to establish categories, which you'll utilize to classify transactions, enabling you to accurately represent them within your budgetary framework.";
                slug = 'Categories';
                break;
            case USER_BALANCE_HISTORY_SHEET:
                title = item;
                description = 'Tracks daily account balances from linked accounts and maintains an ongoing history.';
                slug = 'BalanceHistory';
                break;
            case USER_ACCOUNTS_SHEET:
                title = item;
                description = 'Designate ownership, type, and grouping for your assets and liabilities, providing structured organization and clarity to your financial holdings.';
                slug = 'Accounts';
                break;
            case USER_TRANSACTIONS_SHEET:
                title = item;
                description = 'Displays transactions retrieved from linked bank accounts,offering the capability to assign ownership and categorize each transaction for better financial management.';
                slug = 'Transactions';
                break;
            case USER_INVESTMENTS_SHEET:
                title = item;
                description = 'Provides an organized overview of your linked investment holdings, presenting a structured layout for easy reference and management.';
                slug = 'Investments';
                break;
            case USER_RECONCILE_SHEET:
                title = item;
                description = 'Provides an organized overview of your linked investment holdings, presenting a structured layout for easy reference and management.';
                slug = 'Reconcile';
                break;
            case USER_MONTHLY_BUDGET_SHEET:
                title = item;
                description = "Strategic tool employed by individuals or a household to meticulously plan and oversee income and expenses within a specific month.";
                slug = 'MonthlyBudget';
                break;
            case USER_JOINT_MONTHLY_BUDGET_SHEET:
                title = item;
                description = "Strategic tool employed by individuals AND a household to meticulously plan and oversee income and expenses within a specific month.";
                slug = 'JointMonthlyBudget';
                break;
            case USER_YEARLY_BUDGET_SHEET:
                title = item;
                description = "Strategic tool employed by individuals or a household to meticulously plan and oversee income and expenses within a specific year.";
                slug = 'YearlyBudget';
                break;
            case USER_JOINT_YEARLY_BUDGET_SHEET:
                title = item;
                description = "Strategic tool employed by individuals AND a household to meticulously plan and oversee income and expenses within a specific year.";
                slug = 'JointYearlyBudget';
                break;
            case USER_NET_WORTH_SHEET:
                title = item;
                description = "Is a financial snapshot that provides a comprehensive view of an individual's or a household’s financial health.";
                slug = 'NetWorth';
                break;
            case USER_JOINT_NET_WORTH_SHEET:
                title = item;
                description = "Is a financial snapshot that provides a comprehensive view of an individual's AND a household's financial health.";
                slug = 'JointNetWorth';
                break;
            case USER_BUDGET_MAKER_SHEET:
                title = item;
                description = "A dynamic tool that showcases your income level in comparison to your peers and provides insights into their spending habits.";
                slug = 'BudgetMaker';
                break;
            case USER_DATA_SHEET:
                title = item;
                description = "Input fundamental details about yourself along with your financial aspirations and objectives.";
                slug = 'Data';
                break;
        }

        data.push({ title: title, description: description, slug: slug });

    });

    return data;
}


function installSheetConfirmation() {

    var result = SpreadsheetApp.getUi().alert(
        'Please confirm',
        'Are you sure you want to install the template?',
        SpreadsheetApp.getUi().ButtonSet.YES_NO);

    if (result == SpreadsheetApp.getUi().Button.YES) {
        return true;
    } else {
        return false;
    }
}

function removeSheetConfirmation() {

    var result = SpreadsheetApp.getUi().alert(
        'Please confirm',
        'Are you sure you want to remove the template?',
        SpreadsheetApp.getUi().ButtonSet.YES_NO);

    if (result == SpreadsheetApp.getUi().Button.YES) {
        return true;
    } else {
        return false;
    }

}

function renderTemplateInfo(sheet) {
    var template = HtmlService.createTemplateFromFile('sheetView');
    let list = defaultTemplateList();
    for (var i = 0; i < list.length; i++) {
        if (list[i].slug === sheet) {
            template.data = list[i];
        }
    }
    return template.evaluate().getContent();
}

function installTemplate(sheetName) {
    const spreadsheet = UserSpreadsheet;
    if (isActiveTemplate(sheetName) == true) {
        spreadsheet.getSheetByName(sheetName).showSheet();
    } else {
        installSpreadSheet(sheetName);
    }
    SpreadsheetApp.getUi().alert('Template installed successfully!');
    return true;
}

function removeTemplate(sheetName) {
    const spreadsheet = UserSpreadsheet;
    spreadsheet.getSheetByName(sheetName).hideSheet();
    //spreadsheet.deleteSheet(delsheet);
    SpreadsheetApp.getUi().alert('Template removed successfully!');
    return true;
}

function doPost(e) {
    try {
        var response = JSON.parse(e.postData.contents);
        Logger.log("doPost response:", JSON.stringify(response));
        let item_id = response.item_id;
        switch (response.webhook_code) {
            case 'SYNC_UPDATES_AVAILABLE':
                let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP_USER_DETAILS_SHEET);
                Logger.log("usersheet:", JSON.stringify(sheet));
                let data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
                Logger.log("data of sheet:", JSON.stringify(data));
                if (data.length > 0) {
                    data.forEach(function (item) {
                        let spreadsheetId = item[1];
                        let targetSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
                        let plaidSheet = targetSpreadsheet.getSheetByName(USER_PLAID_SHEET);
                        Logger.log("user plaid sheet:", JSON.stringify(plaidSheet));
                        if (plaidSheet.getLastRow() - 1 > 0) {
                            let plaidSheetData = plaidSheet.getRange(2, 1, plaidSheet.getLastRow() - 1, plaidSheet.getLastColumn()).getValues();
                            Logger.log("user plaid sheet data:", JSON.stringify(plaidSheetData));
                            if (plaidSheetData.length > 0) {
                                plaidSheetData.forEach(function (plaidItem) {
                                    if (plaidItem[2] === item_id && getUserAccountPlaidDataByItemId(item_id) != null) {
                                        setItemUpdateAvailable(item_id);
                                        MailApp.sendEmail(
                                            Admin_Email,
                                            'ThefinuBeta - Plaid Webhook Log',
                                            'Response: ' + JSON.stringify(response),
                                        );
                                    }
                                });
                            }
                        }
                    });
                }
                break;
        }
    } catch (error) {
        Logger.log("doPost Error: " + JSON.stringify(error));
        MailApp.sendEmail(
            Admin_Email,
            'ThefinuBeta - Plaid Webhook Error Log',
            'response: ' + JSON.stringify(error),
        );
    }
}

/*function itemWebhookUpdate(){

  let account_id = 'd0NAYAL04EiDOm9Y7mArIYXoBB5YYeC8eR69P';

  let access_token = getUserAccessTokenByAccountId(account_id);
  if( access_token != ''){
    let options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        client_id: PLAID_CLIENT_ID,
        secret: PLAID_SECRET,
        access_token: access_token,
        webhook: WEBAPP_WEBHOOK
      })
    };
    const PLAID_ITEM_WEBHOOK_ENDPOINT = 'https://'+ PLAID_ENV +'.plaid.com/item/webhook/update';
    let request = UrlFetchApp.fetch(PLAID_ITEM_WEBHOOK_ENDPOINT, options);
    let response = JSON.parse(request.getContentText());
    Logger.log(response);
    return response;
  }
}*/

function deleteAppTriggers() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

function checkItemProductSupport(account_id, product) {
    let access_token = getUserAccessTokenByAccountId(account_id);
    let item = getPlaidItem(access_token);
    if (item) {
        let products = item.products;
        if (products.includes(product)) {
            return true;
        }
    }
    return false;
}

function handlingItemRequiredLogin() {
    const html = HtmlService.createHtmlOutputFromFile('handlingItemRequiredLogin').setWidth(450).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, "Reauthentication");
}

function getItemErrorHandlerAccessToken() {
    let str = JSON.stringify(getPlaidErrorHandler());
    const match = str.match(/access_token=([^,}]+)/);
    return match ? match[1] : null;
}

function getPlaidSheetAccountData(row) {
    let spreadsheet = UserSpreadsheet.getSheetByName(USER_PLAID_SHEET);
    var lastColumn = spreadsheet.getLastColumn();
    var row = spreadsheet.getRange(row, 1, 1, lastColumn).getValues()[0];
    return row;
}

function checkPlaidSheetAccountExist(mask, institution_id) {

    let spreadsheet = UserSpreadsheet.getSheetByName(USER_PLAID_SHEET);
    var data = spreadsheet.getDataRange().getValues();

    // Loop through rows to find the value
    for (var row = 0; row < data.length; row++) {
        /*if( data[row][3] === institution_id && data[row][4] === mask ){
          return row + 1;
        }*/
        //Logger.log(data[row]);
        if (data[row][PLAID_DATA_MASK] === parseInt(mask) && data[row][PLAID_DATA_INSTITUTION_ID] === institution_id) {
            return row + 1; // Return the row number (1-based index)
        }
    }

    return false;
}

function activateDailyPlaidAutoUpdate() {

    try {
        checkUserDailyUpdateOfPlaidItem();
        let spreadsheet = UserSpreadsheet.getSheetByName(USER_PLAID_SHEET);
        var data = spreadsheet.getDataRange().getValues();
        data.splice(0, 1);
        let uniqueItems = [];
        // Loop through rows to find the value
        if (data.length > 0) {
            data.forEach(function (item, index) {
                if (item[PLAID_DATA_IS_LINKED] === true) {
                    let account_id = item[PLAID_DATA_ACCOUNT_ID];
                    updateTransactionSheet(account_id);
                    let support_response = checkItemProductSupport(account_id, 'investments');
                    if (support_response === true) {
                        updateInvestmentSheet(account_id);
                    }
                    updateAccountBalanceHistory(account_id);
                    let row = getPlaidSheetAccountRow(account_id);
                    var sheet_data = getUserAccountPlaidDataByAccountId(account_id);
                    let row_data = {
                        'account_id': sheet_data[PLAID_DATA_ACCOUNT_ID],
                        'access_token': sheet_data[PLAID_DATA_ACCESS_TOKEN],
                        'item_id': sheet_data[PLAID_DATA_ITEM_ID],
                        'institution_id': sheet_data[PLAID_DATA_INSTITUTION_ID],
                        'mask': sheet_data[PLAID_DATA_MASK],
                        'account_name': sheet_data[PLAID_DATA_ACCOUNT_NAME],
                        'linked': sheet_data[PLAID_DATA_IS_LINKED],
                        'linked_date': getTodayDateTime(),
                        'next_cursor': sheet_data[PLAID_DATA_NEXT_CURSOR],
                        'status': sheet_data[PLAID_DATA_STATUS],
                        'updates': false
                    };
                    updateUserPlaidDataSheet(row, row_data);
                    if (!uniqueItems.includes(sheet_data[PLAID_DATA_ITEM_ID])) {
                        // If the item is new (not found), push it to the uniqueItems array.
                        uniqueItems.push(sheet_data[PLAID_DATA_ITEM_ID]);
                    }
                }
            });

            if (uniqueItems.length > 0) {
                updatePlaidAccountStatus(uniqueItems);
            }

            MailApp.sendEmail(
                Admin_Email,
                'ThefinuBeta - Daily Task',
                'The daily task completed successfully on ' + getTodayDateTime(),
            );
        }
        return true;
    } catch (error) {
        MailApp.sendEmail(
            Admin_Email,
            'ThefinuBeta - Task Aborted ',
            'Something went wrong while trying to complete daily task.',
        );
        Logger.log("An error occurred:", JSON.stringify(error));
        return false;
    }
}

function activateDailyAutoUpdate() {

    try {
        const functionToRun = 'activateDailyPlaidAutoUpdate';
        const triggers = ScriptApp.getProjectTriggers();

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
        //SpreadsheetApp.getUi().alert("Daily update trigger refreshed successfully.");
    } catch (error) {
        Logger.log("Error while refreshing trigger: " + error.message);
        SpreadsheetApp.getUi().alert("Something went wrong, please try again.");
    }
}

function protectSheetByName(sheetName) {
    const sheet = UserSpreadsheet.getSheetByName(sheetName);
    if (sheet) {
        const protection = sheet.protect();
        protection.setDescription(`Protected: ${sheetName}`);
        protection.setWarningOnly(false); // Prevent editing
    }
}

function reApplyFormula() {
    let transactionSheet = UserSpreadsheet.getSheetByName('Transactions');
    if (transactionSheet) {
        if (transactionSheet.getLastRow() > 2) {
            var spreadsheet = UserSpreadsheet;
            let formulaSheets = [USER_DEFINITION_SHEET, USER_MONTHLY_BUDGET_SHEET, USER_JOINT_MONTHLY_BUDGET_SHEET, USER_YEARLY_BUDGET_SHEET, USER_JOINT_YEARLY_BUDGET_SHEET, USER_BUDGET_MAKER_SHEET, USER_TRANSACTIONS_SHEET];
            formulaSheets.forEach(function (item) {
                switch (item) {
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
                            let periodValues = periodRange.getValues().flat().filter(function (value) {
                                return value && (typeof value === 'string' || !isNaN(value));
                            });

                            if (periodValues.length === 0) {
                                Logger.log("Error: No valid period values found in the 'Period' range (" + periodRange.getA1Notation() + ").");
                                return;
                            }

                            // Log the filtered values for debugging
                            Logger.log("Valid period values: " + periodValues.join(", "));

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
                            let periodValues = periodRange.getValues().flat().filter(function (value) {
                                return value && (typeof value === 'string' || !isNaN(value));
                            });

                            if (periodValues.length === 0) {
                                Logger.log("Error: No valid period values found in the 'Period' range (" + periodRange.getA1Notation() + ").");
                                return;
                            }

                            // Log the filtered values for debugging
                            Logger.log("Valid period values: " + periodValues.join(", "));

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
                            let yearValues = yearRange.getValues().flat().filter(function (value) {
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
                            let yearValues = yearRange.getValues().flat().filter(function (value) {
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

function confirmPopulateReportToSheets() {

    try {
        populateNetWorth();
        populateJointNetWorth();
        populateMonthlyBudget();
        populateJointMonthlyBudget();
        populateYearlyBudget();
        populateJointYearlyBudget();
        return true;
    } catch (error) {
        Logger.log(JSON.stringify("generate report error: " + error));
        return false;
    }
}

function testFunction() {
    Logger.log(getAppIntroductionContent());
    return false;
    let uniqueItems = [];

    let plaidSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_PLAID_SHEET);
    if (plaidSheet.getLastRow() - 1 > 0) {
        let plaidSheetData = plaidSheet.getRange(2, 1, plaidSheet.getLastRow() - 1, plaidSheet.getLastColumn()).getValues();
        //Logger.log("user plaid sheet data:", JSON.stringify(plaidSheetData));
        if (plaidSheetData.length > 0) {
            plaidSheetData.forEach(function (plaidItem) {
                if (!uniqueItems.includes(plaidItem[2])) {
                    // If the item is new (not found), push it to the uniqueItems array.
                    uniqueItems.push(plaidItem[2]);
                }
            });
        }
    }

    if (uniqueItems.length > 0) {

        var payload = {
            item_ids: uniqueItems, // Array of item_ids
            spreadsheet_id: spreadsheetId
        };

        var options = {
            method: 'post',
            contentType: 'application/json',
            headers: {
                Authorization: 'Bearer ' + ScriptProperties.getProperty('APPS_SCRIPT_API_KEY')
            },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };

        try {
            var response = UrlFetchApp.fetch('https://your-server/updatePlaidAccountStatus', options);
            var responseCode = response.getResponseCode();
            var responseText = response.getContentText();
            Logger.log('updatePlaidAccountStatus Response Code: ' + responseCode);
            Logger.log('updatePlaidAccountStatus Response Body: ' + responseText);

            var result = JSON.parse(responseText);

            if (responseCode === 200 && result.success === true) {
                Logger.log('Status updated for item_ids: ' + itemIds.join(', '));
                return true;
            } else if (responseCode >= 400 && result.error) {
                throw new Error('API Error: ' + result.error + ' - ' + JSON.stringify(result.data));
            } else {
                throw new Error('Unexpected response: ' + responseText);
            }
        } catch (e) {
            Logger.log('Error updating status for item_ids ' + itemIds.join(', ') + ': ' + e);
            return false;
        }
    }
}

function getAppIntroductionContent() {
    let response = 'Link your bank account(s) with plaid and get account, transactions, investments and more features once account is connected. You can also manage monthly and yearly budgets basis and so more features you can use it by the app.';
    const source = SpreadsheetApp.openByUrl(APP_TEMPLATES_SPREADSHEET);
    if (source) {
        let sheet = source.getSheetByName('App Content');
        if (sheet) {
            let content = sheet.getRange('A2').getValue();
            if (content) return content.toString();
        }
    }
    return response;
}

/**
 * Automatically regenerates all reports when a Category, Group, or Type 
 * column is edited in the "Transactions" sheet.
 * * NOTE: For this script to work, you must have corresponding functions 
 * for all your reports (e.g., populateMonthlyBudget, populateNetWorth, etc.)
 */
function handleEdit(e) {

    const range = e.range;
    const sheet = range.getSheet();

    const defSheet = UserSpreadsheet.getSheetByName(USER_DEFINITION_SHEET);

    // Get the column that was edited
    const editedColumn = range.getColumn();

    const editCell = range.getA1Notation();

    let dropDown;

    try {

        switch (sheet.getName()) {

            case USER_TRANSACTIONS_SHEET:
                let TRAN_CAT_COL_REF = 'I5';
                let TRAN_OWN_COL_REF = 'I11';
                let TRAN_AMT_COL_REF = 'I10';

                // Read the column indices from the Definition sheet (1-based index)
                let catCol = defSheet.getRange(TRAN_CAT_COL_REF).getValue();
                let ownCol = defSheet.getRange(TRAN_OWN_COL_REF).getValue();
                let amtCol = defSheet.getRange(TRAN_AMT_COL_REF).getValue();

                if (editedColumn === catCol || editedColumn === ownCol || editedColumn === amtCol) {
                    regenerateAllReports();
                }

                break;
            case USER_MONTHLY_BUDGET_SHEET:
                dropDown = 'C2';
                Logger.log(editCell);
                if (editCell === dropDown && typeof populateMonthlyBudget === 'function') {
                    populateMonthlyBudget();
                    regenerateNetWorthReports();
                }
                break;
            case USER_JOINT_MONTHLY_BUDGET_SHEET:
                dropDown = 'C2';
                if (editCell === dropDown && typeof populateJointMonthlyBudget === 'function') {
                    populateJointMonthlyBudget();
                    regenerateNetWorthReports();
                }
                break;
            case USER_YEARLY_BUDGET_SHEET:
                dropDown = 'E2';
                if (editCell === dropDown && typeof populateYearlyBudget === 'function') {
                    populateYearlyBudget();
                    regenerateNetWorthReports();
                }
                break;
            case USER_JOINT_YEARLY_BUDGET_SHEET:
                dropDown = 'D2';
                if (editCell === dropDown && typeof populateJointYearlyBudget === 'function') {
                    populateJointYearlyBudget();
                    regenerateNetWorthReports();
                }
                break;
            case USER_ACCOUNTS_SHEET:
                let balanceCol = defSheet.getRange('F10').getValue();
                let ownerCol = defSheet.getRange('F11').getValue();
                let groupCol = defSheet.getRange('F6').getValue();
                let assliabCol = defSheet.getRange('F7').getValue();
                if (range.getRow() > sheet.getLastRow() && e.value) {
                    populateNetWorth();
                    populateJointNetWorth();
                }
                if (editedColumn === balanceCol || editedColumn === ownerCol || editedColumn === groupCol || editedColumn === assliabCol) {
                    populateNetWorth();
                    populateJointNetWorth();
                }
                break;
        }
    } catch (error) {
        Logger.log("Error in onEdit trigger: " + error.toString());
    }

}

/**
 * Creates the installable onEdit trigger.
 * Call this in onInstall() or a menu.
 */
function createOnEditTrigger() {
    // Delete any existing triggers to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'handleEdit') {
            ScriptApp.deleteTrigger(trigger);
        }
    });

    // FIXED: Add .forSpreadsheet() before .onEdit()
    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    ScriptApp.newTrigger('handleEdit')
        .forSpreadsheet(spreadsheetId)
        .onEdit()
        .create();

    console.log('onEdit installable trigger created successfully!');
}

function getDateTime() {
    var currentDate = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "MMM dd, yyyy hh:mm a");
    return currentDate;
}

/**
 * Master function that calls all individual report generation functions.
 * You MUST ensure these functions exist in your script project.
 */
function regenerateAllReports() {

    Logger.log("Starting full report regeneration...");

    // Monthly Reports
    if (typeof populateMonthlyBudget === 'function') {
        populateMonthlyBudget();
    }
    if (typeof populateJointMonthlyBudget === 'function') {
        populateJointMonthlyBudget();
    }

    // Yearly Reports
    if (typeof populateYearlyBudget === 'function') {
        populateYearlyBudget();
    }
    // This is the function we corrected earlier
    if (typeof populateJointYearlyBudget === 'function') {
        populateJointYearlyBudget();
    }

    // Net Worth Reports
    if (typeof populateNetWorth === 'function') {
        populateNetWorth();
    }
    if (typeof populateJointNetWorth === 'function') {
        populateJointNetWorth();
    }

    Logger.log("Report regeneration complete.");
}

function regenerateNetWorthReports() {
    // Net Worth Reports
    if (typeof populateNetWorth === 'function') {
        populateNetWorth();
    }
    if (typeof populateJointNetWorth === 'function') {
        populateJointNetWorth();
    }
}
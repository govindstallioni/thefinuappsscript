function linkInvestmentsSheetData( investmentsData ){
  var investmentCollection = [];
  if( investmentsData.length > 0 ){
    investmentsData.forEach( function(investment){
      investmentCollection.push([
        '', //
        investment.account_name, //Name
        investment.cusip, //CUSIP
        investment.ticker, //Ticker
        investment.price_as_of, //Price as Of
        investment.price, //Price
        investment.quantity, //Quantity
        investment.cost_basis, //Cost Basis
        investment.value, //Value
        investment.account, //Account
        investment.security_id, //Security ID
        investment.account_id, //Account ID
        getTodayDateTime() //Date & Time
      ]);
    });
  }

  if( investmentCollection.length > 0 ){
    insertInvestmentsData(investmentCollection);
  }
}

function linkInvestmentSheet( account_id ){
  var format_data = [];
  let invsetments = getPlaidInvestmentsData( account_id );
  if( invsetments != null && !invsetments.error ){
    format_data = formatPlaidInvestments( account_id, invsetments);
  }
  
  if( format_data.length > 0 ){
    linkInvestmentsSheetData(format_data);
  }

  sortingInvestmentSheet();

  return true;
}

function formatPlaidInvestments(account_id, investments){
  // Defensive checks: return empty array if investments payload isn't shaped as expected
  var collectionArr = [];
  if(!investments || typeof investments !== 'object') return collectionArr;
  var PlaidAccounts = Array.isArray(investments.accounts) ? investments.accounts : [];
  var PlaidAccountsHoldings = Array.isArray(investments.holdings) ? investments.holdings : [];
  var PlaidAccountsSecurities = Array.isArray(investments.securities) ? investments.securities : [];
  if(PlaidAccountsHoldings.length === 0 || PlaidAccountsSecurities.length === 0) return collectionArr;

  // Build a security_id → security lookup map to avoid nested loop
  var securityMap = {};
  PlaidAccountsSecurities.forEach(function(sec) {
    if (sec && sec.security_id) {
      securityMap[sec.security_id] = sec;
    }
  });

  // Fetch account name once (not per holding)
  var account_name = getPlaidAccountNameByAccountId(account_id) || '';
  var account = PlaidAccounts[0] || {};

  for (var i = 0; i < PlaidAccountsHoldings.length; i++) {
    var holding = PlaidAccountsHoldings[i] || {};
    var securities = securityMap[holding.security_id];
    if (!holding.security_id || !securities) continue;

    collectionArr.push({
      'security_id': holding.security_id || '',
      'account_id': account_id,
      'account': account_name,
      'account_number': account.mask || '',
      'cusip': securities.cusip || '',
      'ticker': securities.ticker_symbol || securities.ticker || '',
      'account_name': securities.name || '',
      'quantity': holding.quantity || 0,
      'cost_basis': holding.cost_basis || 0,
      'price_as_of': holding.institution_price_as_of || '',
      'price': holding.institution_price || 0,
      'value': holding.institution_value || 0,
      'type': securities.type || ''
    });
  }

  return collectionArr;
}

function insertInvestmentsData( collection ){

  var sheet = UserSpreadsheet.getSheetByName(USER_INVESTMENTS_SHEET);
  // Write headers
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  let lastrow = sheet.getLastRow() + 1;

  var cell = sheet.getRange(lastrow, 1, collection.length, headers.length);
  cell.setValues(collection);
  cell.setFontSize(9)
      .setFontFamily("Comfortaa")
      .setFontColor("#000000")
      .setFontWeight("bold");
}


function updateInvestmentSheet( account_id ){
  try{
    Logger.log('[INV-UPDATE] Starting for account: ' + account_id);
    Logger.log('[INV-UPDATE] Clearing existing investment data');
    clearInvestmentsData( account_id );
    Logger.log('[INV-UPDATE] Linking investment sheet data');
    linkInvestmentSheet( account_id );
    Logger.log('[INV-UPDATE] Completed for account: ' + account_id);
    return true;
  }catch(e){
    Logger.log('[INV-UPDATE] ERROR for account ' + account_id + ': ' + e.toString());
    Logger.log('[INV-UPDATE] Stack: ' + e.stack);
    return false;
  }
}

function sortingInvestmentSheet(){
  var sheet = UserSpreadsheet.getSheetByName(USER_INVESTMENTS_SHEET);
  const range = sheet.getDataRange(); 
  range.sort({ column: 5, ascending: false });
}

function clearInvestmentsData(account_id){

  const sheet = UserSpreadsheet.getSheetByName(USER_INVESTMENTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const columnIndexToCheck = 12;

  // Collect rows to delete (skip header row 0)
  var rowsToDelete = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][columnIndexToCheck - 1] === account_id) {
      rowsToDelete.push(i + 1);
    }
  }

  // Delete from bottom to top so indices stay valid
  for (var r = rowsToDelete.length - 1; r >= 0; r--) {
    sheet.deleteRow(rowsToDelete[r]);
  }

  sortingInvestmentSheet();

  return true;

}

function changeAccountNameOnInvestmentSheet( account_id, account_name ){
  let sheet = UserSpreadsheet.getSheetByName(USER_INVESTMENTS_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var row = 0; row < data.length; row++) {
    if (data[row].includes(account_id)) {
      sheet.getRange(row + 1 , 10).setValue(account_name); 
    }
  }
}

function updateAccountIdOnInvestmentSheet( account_id, new_id ){
  let sheet = UserSpreadsheet.getSheetByName(USER_INVESTMENTS_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var row = 0; row < data.length; row++) {
    if (data[row].includes(account_id)) {
      sheet.getRange(row + 1 , 12).setValue(new_id); 
    }
  }
}
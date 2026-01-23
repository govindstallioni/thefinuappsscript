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
        investment.account_id //Account ID
      ]);
    });
  }

  if( investmentCollection.length > 0 ){
    insertInvestmentsData(investmentCollection);
  }
}

function linkInvestmentSheet( account_id ){
  //let account_id = getUserCurrentAccountId();
  var format_data = [];
  let invsetments = getPlaidInvestmentsData( account_id );
  //Logger.log( JSON.stringify(invsetments, null, 2) );
  if( invsetments != null ){
    format_data = formatPlaidInvestments( account_id, invsetments);
  }
  
  if( format_data.length > 0 ){
    linkInvestmentsSheetData(format_data);
  }

  sortingInvestmentSheet();

  return true;
}

function formatPlaidInvestments(account_id, investments){

  var collectionArr = [];

  var PlaidAccounts = investments.accounts;
  var PlaidAccountsHoldings = investments.holdings;
  var PlaidAccountsSecurities = investments.securities;
  var account_name = '';

  for( var i = 0; i < PlaidAccountsHoldings.length; i++ ){
    let holding = PlaidAccountsHoldings[i];
    let account = PlaidAccounts[0];
    for ( var j = 0; j < PlaidAccountsSecurities.length; j++ ){
      let securities = PlaidAccountsSecurities[j];
      //Logger.log(securities.security_id);
      if( holding.security_id === securities.security_id ){
        account_name = getAccountNameByAccountId(account_id);
        collectionArr.push({
          'security_id': holding.security_id,
          'account_id' : holding.account_id,
          'account': account_name,
          'account_number': account.mask,
          'cusip': securities.cusip,
          'ticker': securities.ticker_symbol,
          'account_name': securities.name,
          'quantity': holding.quantity,
          'cost_basis': holding.cost_basis,
          'price_as_of': holding.institution_price_as_of,
          'price': holding.institution_price,
          'value': holding.institution_value,
          'type': securities.type
        });
      }
    }
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
  clearInvestmentsData( account_id );
  linkInvestmentSheet( account_id );
}

function sortingInvestmentSheet(){
  var sheet = UserSpreadsheet.getSheetByName(USER_INVESTMENTS_SHEET);
  const range = sheet.getDataRange(); 
  range.sort({ column: 5, ascending: false });
}

function clearInvestmentsData(account_id){

  const sheet = UserSpreadsheet.getSheetByName(USER_INVESTMENTS_SHEET);
  const range = sheet.getDataRange();
  const data = range.getValues();
  const columnIndexToCheck = 12; 

  // Array to store the rows to clear
  let rowsToClear = [];

  // Loop through the data to identify rows to clear
  for (let i = 0; i < data.length; i++) {
    if (data[i][columnIndexToCheck - 1] === account_id) {
      rowsToClear.push(i + 1); // Adjust for 1-based indexing
    }
  }

  // Clear the content of the identified rows in bulk
  if (rowsToClear.length > 0) {
    sheet.getRangeList(rowsToClear.map(row => `A${row}:L${row}`)).clearContent();
  }

  sortingInvestmentSheet();

  return true;

}

function changeAccountNameOnInvestmentSheet( account_id, account_name ){
  let sheet = UserSpreadsheet.getSheetByName(USER_INVESTMENTS_SHEET);
  var data = sheet.getDataRange().getValues();
  //Logger.log(data);
  for (var row = 0; row < data.length; row++) {
    if (data[row].includes(account_id)) {
      //Logger.log(row);
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
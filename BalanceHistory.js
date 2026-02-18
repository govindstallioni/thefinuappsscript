function linkBalanceHistorySheetData(account_id, lastDate){

  var lastTransactionDate = lastDate;

  const response = getAppPlaidAccountById(account_id);

  if( response.success === true ){

    let account = response.result;

    let sheet = UserSpreadsheet.getSheetByName(USER_BALANCE_HISTORY_SHEET);
    // Write headers
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    let rowValues = [];

    // Populate row values based on header indexes
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      // Map Plaid transaction data to the corresponding header
      var value = "";
      switch (header) {
        case "Accounts":
          value = account.name;
          break;
        case "Account Number":
          value = account.mask;
          break;
        case "Account ID":
          value = account_id;
          break;
        case "Date":
          value = lastTransactionDate;
          break;
        case "Balance":
          value = account.balance;
          break;
      }
      rowValues.push(value);
    }

    // Determine index of Account ID header (if present) to enforce idempotency
    const accountIdColIndex = headers.findIndex(h => h === 'Account ID');
    if (accountIdColIndex >= 0) {
      const values = sheet.getDataRange().getValues();
      for (let r = 1; r < values.length; r++) {
        if (values[r] && values[r][accountIdColIndex] === account_id) {
          return; // already have a balance history row for this account
        }
      }
    }

    // Append the new row and apply formatting
    sheet.appendRow(rowValues);
    const appendedRow = sheet.getLastRow();
    const cell = sheet.getRange(appendedRow, 1, 1, sheet.getLastColumn());
    cell.setFontSize(9)
      .setFontFamily("Comfortaa")
      .setFontColor("#000000")
      .setFontWeight("bold");
  }
}

function clearBalanceHistoryData(account_id){

  const sheet = UserSpreadsheet.getSheetByName(USER_BALANCE_HISTORY_SHEET);
  const range = sheet.getDataRange();
  const data = range.getValues();
  const columnIndexToCheck = 7; 

  // Array to store the rows to clear
  let rowsToClear = [];

  // Loop through the data to identify rows to clear
  for (let i = 0; i < data.length; i++) {
    if (data[i][columnIndexToCheck - 1] === account_id) {
      rowsToClear.push(i + 1); // Adjust for 1-based indexing
      //sheet.getRange( i+1, 1, 1, sheet.getLastColumn()).clearContent(); 
    }
  }

  // Clear the content of the identified rows in bulk
  if (rowsToClear.length > 0) {
    sheet.getRangeList(rowsToClear.map(row => `A${row}:H${row}`)).clearContent();
  }
  
  return true;
}

function updateAccountBalanceHistory( account_id ){

  try{

    var collectionArr = [];
    const response = getPlaidAccountBalance(account_id);
    //Logger.log( JSON.stringify(response, null, 2) );
    let accounts = response.accounts;
    accounts.forEach( function(account){
      let balances = account.balances;
      let accountRequest = getAppPlaidAccountById(account_id);
      if( accountRequest.success === true ){
        let account = accountRequest.result;
        updateAppAccountDetailById(account.account_id,{ balance: balances.current });
        collectionArr.push({ 'name': account.name, 'account_number': account.mask, 'date': getTodayDate(), 'balance': balances.current, 'account_id': account.account_id, 'request_id': response.request_id });
      }
    });

    if(collectionArr.length > 0 ){
      var sheet = UserSpreadsheet.getSheetByName(USER_BALANCE_HISTORY_SHEET);
      // Write headers
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      for (var i = 0; i < collectionArr.length; i++) {
        let account = collectionArr[i];
        // Create an array to hold the values for each column
        var rowValues = [];

        // Populate row values based on header indexes
        for (var j = 0; j < headers.length; j++) {
          var header = headers[j];
          // Map Plaid transaction data to the corresponding header
          var value = "";
          switch (header) {
            case "Accounts":
              value = account.name;
              break;
            case "Account Number":
              value = account.account_number;
              break;
            case "Account ID":
              value = account.account_id;
              break;
            case "Date":
              value = account.date;
              break;
            case "Balance":
              value = account.balance;
              break;
            case "Balance Update ID":
              value = account.request_id;
              break;
            case 'Date & Time':
              value = getTodayDateTime();
              break;
          }
          rowValues.push(value);
        }
        /*while( sheet.getRange(row, 7).getValue() === "" ){
          sheet.getRange(row, 1, 1, 7).setValues([rowValues]);
          row++;
          break;
        }*/
        let lastrow = sheet.getLastRow() + 1;
        for( var row = 1; row < lastrow; row++ ){
          if( sheet.getRange(lastrow, sheet.getLastColumn()).getValue() === '' ){
            var cell = sheet.getRange(lastrow, 1, 1, sheet.getLastColumn());
            cell.setValues([rowValues]);
            cell.setFontSize(9)
            .setFontFamily("Comfortaa")
            .setFontColor("#000000")
            .setFontWeight("bold");
            lastrow++;
            break;
          }
        }
      }
      const range = sheet.getDataRange(); 
      range.sort({ column: 2, ascending: false });
    }
    else {
      // No balances returned from Plaid — ensure at least one balance history row exists
      try {
        linkBalanceHistorySheetData(account_id, getTodayDate());
      } catch (e) {
        Logger.log('linkBalanceHistorySheetData error: ' + e);
      }
    }
  }catch(e){
    Logger.log("Error:"+ e);
  }
}

function changeAccountNameOnBalanceHistorySheet( account_id, account_name ){
  let sheet = UserSpreadsheet.getSheetByName(USER_BALANCE_HISTORY_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var row = 0; row < data.length; row++) {
    if (data[row].includes(account_id)) {
      sheet.getRange(row + 1 , 3).setValue(account_name); 
    }
  }
}

function updateAccountIdOnBalanceHistorySheet( account_id, new_id ){
  let sheet = UserSpreadsheet.getSheetByName(USER_BALANCE_HISTORY_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var row = 0; row < data.length; row++) {
    if (data[row].includes(account_id)) {
      sheet.getRange(row + 1 , 7).setValue(new_id); 
    }
  }
}

function sortingBalanceHistorySheet(){
  var sheet = UserSpreadsheet.getSheetByName(USER_BALANCE_HISTORY_SHEET);
  const range = sheet.getDataRange(); 
  range.sort({ column: 2, ascending: false });
}
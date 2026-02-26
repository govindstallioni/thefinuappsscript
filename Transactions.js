function linkPlaidAddedTransactions( transactions, account ){
  
  var collection = [];

  if( transactions.length > 0 ){
    transactions.forEach( function(newitem){
      if( newitem.length > 0 ){
        newitem.forEach( function(transaction){

          let transaction_status = transaction.pending == true ? 'Pending' : '';

          collection.push([
            '', //Not Cleared
            transaction.date, //Date
            transaction.name, //Description
            '', //Category
            transaction.amount, //Amount
            '', //Owner
            '', //Assigned
            account.name, //Account Name,
            transaction_status,
            account.mask, //Account Number
            transaction.account_id, //Account ID,
            account.institution_name, //Institution ID
            transaction.transaction_id, //Transaction ID
            '', //Group
            '', //Type
            '' //Period
          ]);
        });
      }
    });
  }
  
  if( collection.length > 0 ){
    insertTransactionsData(collection);
  }
}

function linkPlaidModifiedTransactions( transactions, account ){

  if( transactions.length > 0 ){
    transactions.forEach( function(newitem){
      if( newitem.length > 0 ){
        newitem.forEach( function(transaction){

          let transaction_status = transaction.pending == true ? 'Pending' : '';

          if( transaction.pending_transaction_id != null ){

            if( getTransactionRow( transaction.pending_transaction_id ) != null ){

              let data = {
                'account': account.name,
                'description': transaction.name,
                'amount': transaction.amount,
                'date': transaction.date,
                'account_number': account.mask,
                'account_id': transaction.account_id,
                'transaction_status': transaction_status,
                'institution': account.institution_name,
                'transaction_id': transaction.transaction_id
              };

              updateTransactionsData( transaction.pending_transaction_id, data);

            }

          }else{

            if( getTransactionRow(transaction.transaction_id ) != null ){
              let data = {
                'account': account.name,
                'description': transaction.name,
                'amount': transaction.amount,
                'date': transaction.date,
                'account_number': account.mask,
                'account_id': transaction.account_id,
                'transaction_status': transaction_status,
                'institution': account.institution_name,
                'transaction_id': transaction.transaction_id
              };
              updateTransactionsData( transaction.transaction_id, data);
            }
          }
          
        });
      }
    });
  }
}

function removeTransactionItem( transaction_id ){
  const sheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
  if( getTransactionRow(transaction_id) != null ){
    let row = getTransactionRow(transaction_id);
    sheet.deleteRow(row);
  }
}

function getTransactionRow( transaction_id, cachedData ){
  var data = cachedData;
  if (!data) {
    const sheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
    data = sheet.getDataRange().getValues();
  }

  // Loop through rows to find the value
  for (var row = 0; row < data.length; row++) {
    if (data[row].includes(transaction_id)) {
      return row + 1; // Return the row number (1-based index)
    }
  }

  return null;
}

/**
 * Builds a transaction ID -> row number lookup map from sheet data.
 * Transaction ID is in column 13 (index 12).
 */
function buildTransactionRowMap() {
  const sheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (var row = 1; row < data.length; row++) {
    var txId = data[row][12]; // Transaction ID column (index 12)
    if (txId && txId !== '') {
      map[txId] = row + 1; // 1-based row number
    }
  }
  return { data: data, map: map };
}

function clearTransactionsData( account_id ){
  const sheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
  const range = sheet.getDataRange();
  const data = range.getValues();
  const columnIndexToCheck = 11; 

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
    sheet.getRangeList(rowsToClear.map(row => `A${row}:M${row}`)).clearContent();
  }
  sortingTransactionSheet();
  //Logger.log(`Cleared content for ${rowsToClear.length} rows.`);
  return true;
}

function insertTransactionsData(collection){
  
  var sheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  let lastrow = sheet.getLastRow() + 1;

  var cell = sheet.getRange(lastrow, 1, collection.length, headers.length);
  cell.setValues(collection);
  cell.setFontSize(9)
      .setFontFamily("Comfortaa")
      .setFontColor("#000000")
      .setFontWeight("bold");
    
  return true;
      
}

function sortingTransactionSheet(){
  var sheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
  const range = sheet.getDataRange(); 
  range.sort({ column: 2, ascending: false });
}

function updateTransactionSheet(account_id){
  
  let response = getAppPlaidAccountById( account_id );
  
  if( response.success === true ){

    let result = response.result;
    
    let next_cursor = result.next_cursor;
    //fetch transactions first 500 data from plaid
    let transactions = getPlaidTransactionSyncData( account_id, next_cursor );

    let TransactionAccounts = transactions.accounts;
    
    if( TransactionAccounts.length > 0 ){

      let TransactionAdded = transactions.added;
      let TransactionModified = transactions.modified;

      TransactionAccounts.forEach(function(account){
    
        var newTransactionCollection = [];
        var accountNumber = result.mask;
        var accountName = getPlaidAccountNameByAccountId(account.account_id);
        var institutionName = result.institution_name;

        if( TransactionAdded.length > 0 ){

          TransactionAdded.forEach(function(transaction){

            let transaction_status = transaction.pending == true ? 'Pending' : '';
            
            if( transaction.pending_transaction_id != null ){
              if( getTransactionRow( transaction.pending_transaction_id ) != null ){
                let data = {
                  'account': accountName,
                  'description': transaction.name,
                  'amount': transaction.amount,
                  'date': transaction.date,
                  'account_number': accountNumber,
                  'account_id': transaction.account_id,
                  'transaction_status': transaction_status,
                  'institution': institutionName,
                  'transaction_id': transaction.transaction_id
                };
                //Logger.log(data);
                updateTransactionsData( transaction.pending_transaction_id, data);
              }
            }else{
              if( getTransactionRow(transaction.transaction_id ) != null ){
                let data = {
                  'account': accountName,
                  'description': transaction.name,
                  'amount': transaction.amount,
                  'date': transaction.date,
                  'account_number': accountNumber,
                  'account_id': transaction.account_id,
                  'transaction_status': transaction_status,
                  'institution': institutionName,
                  'transaction_id': transaction.transaction_id
                };
                //Logger.log(data);
                updateTransactionsData( transaction.transaction_id, data);
              }else{
                newTransactionCollection.push([
                  '', //Not Cleared
                  transaction.date, //Date
                  transaction.name, //Description
                  '', //Category
                  transaction.amount, //Amount
                  '', //Owner
                  '', //Assigned
                  accountName, //Account
                  transaction_status,
                  accountNumber, //Account Number
                  transaction.account_id, //Account ID,
                  institutionName, //Institution ID
                  transaction.transaction_id, //Transaction ID
                  '', //Group
                  '', //Type
                  '' //Period
                ]);
              }
            }

          });
        }

        if( newTransactionCollection.length > 0 ){
          insertTransactionsData(newTransactionCollection);
        }

        if( TransactionModified.length > 0 ){

          TransactionModified.forEach( function( transaction ){

            let transaction_status = transaction.pending == true ? 'Pending' : '';

            if( transaction.pending_transaction_id != null ){

              if( getTransactionRow( transaction.pending_transaction_id ) != null ){
                let data = {
                  'account': accountName,
                  'description': transaction.name,
                  'amount': transaction.amount,
                  'date': transaction.date,
                  'account_number': accountNumber,
                  'account_id': transaction.account_id,
                  'transaction_status': transaction_status,
                  'institution': institutionName,
                  'transaction_id': transaction.transaction_id
                };
                updateTransactionsData( transaction.pending_transaction_id, data);
              }

            }else{

              if( getTransactionRow(transaction.transaction_id ) != null ){
                let data = {
                  'account': accountName,
                  'description': transaction.name,
                  'amount': transaction.amount,
                  'date': transaction.date,
                  'account_number': accountNumber,
                  'account_id': transaction.account_id,
                  'transaction_status': transaction_status,
                  'institution': institutionName,
                  'transaction_id': transaction.transaction_id
                };
                updateTransactionsData( transaction.transaction_id, data);
              }

            }

          });

        }

      });

    }
    
    if( transactions.next_cursor != ''){
      updateAppAccountDetailById(account_id,{next_cursor: transactions.next_cursor});
    }
  
  }

  sortingTransactionSheet();
  
  return true;
}

function updateTransactionsData( transaction_id = null, transaction = null){

  if( getTransactionRow(transaction_id) != null ){

    const row = getTransactionRow(transaction_id);
    var sheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var rowValues = [];

    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      // Map Plaid transaction data to the corresponding header
      var value = "";
      switch (header) {
        case "Account":
          value = transaction.account;
          break;
        case "Transaction Status":
          value = transaction.transaction_status;
        break;
        case "Account Number":
          value = transaction.account_number;
          break;
        case "Amount":
          value = transaction.amount;
          break;
        case "Date":
          value = transaction.date;
          break;
        case "Description":
          value = transaction.description;
          break;
        case "Institution":
          value = transaction.institution;
          break;
        case "Transaction ID":
          value = transaction.transaction_id;
          break;
        case "Account ID":
          value = transaction.account_id;
          break;
        
      }
      rowValues.push(value);
    }

    if( rowValues.length > 0 ){
      var cell = sheet.getRange(row, 1, 1, 16);
        cell.setValues([rowValues]);
        cell.setFontSize(9)
        .setFontFamily("Comfortaa")
        .setFontColor("#000000")
        .setFontWeight("bold");
    }

  }
}


function getTransactionLastUpdateDate( account_id ){
  let spreadsheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
  var collections = spreadsheet.getDataRange().getValues();
  //Removed sheet title from the collection
  collections.splice(0, 1);
  var date = null;
  collections.some((row, index) => {
    //Logger.log(row[10]);
    if( row[10] === account_id ){
      if ( typeof row[8] === "string" && !row[8].toLowerCase().includes("pending") ) {
        date = row[1];
        return true; // Exit the loop when condition is met
      }
    }
    return false;
  });
  return date;
}

function changeAccountNameOnTransactionSheet( account_id, account_name ){
  let sheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
  var data = sheet.getDataRange().getValues();
  //Logger.log(data);
  for (var row = 0; row < data.length; row++) {
    if (data[row].includes(account_id)) {
      //Logger.log(row);
      sheet.getRange(row + 1 , 8).setValue(account_name); 
    }
  }
}

function updateAccountIdOnTransactionSheet( account_id, new_id ){
  let sheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var row = 0; row < data.length; row++) {
    if (data[row].includes(account_id)) {
      sheet.getRange(row + 1 , 11).setValue(new_id); 
    }
  }
}
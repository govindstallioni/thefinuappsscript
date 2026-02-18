function linkTransactionSheet( account_id ){
  var addedTransactions = [];
  var modifiedTransactions = [];
  var has_more = false;
  var next_cursor = '';
  //fetch transactions first 500 data from plaid
  let transactions = getPlaidTransactionSyncData( account_id, next_cursor );
  //check has_more and fetch next 500 transactions data by next_cursor from plaid and format the them.
  if( transactions.request_id != '' ){
    has_more = transactions.has_more;
    next_cursor = transactions.next_cursor;
    addedTransactions.push(transactions.added);
    modifiedTransactions.push(transactions.modified);
    while( has_more === true ){
      let next_transactions = getPlaidTransactionSyncData( account_id, next_cursor);
      addedTransactions.push(next_transactions.added);
      modifiedTransactions.push(next_transactions.modified);
      next_cursor = next_transactions.next_cursor;
      has_more = next_transactions.has_more;
    }
    if( has_more == false ){
      updateAppAccountDetailById(account_id,{next_cursor: next_cursor});
    }

    let account_response = getAppPlaidAccountById( account_id );
    let account_data = account_response.result;
  
    linkPlaidAddedTransactions(addedTransactions, account_data );
    linkPlaidModifiedTransactions(modifiedTransactions, account_data);
    sortingTransactionSheet();
  }
}
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
function getTransactionRow( transaction_id ){

  const sheet = getUserSpreadsheet().getSheetByName(USER_TRANSACTIONS_SHEET);
  var data = sheet.getDataRange().getValues();

  // Loop through rows to find the value
  for (var row = 0; row < data.length; row++) {
    if (data[row].includes(transaction_id)) {
      return row + 1; // Return the row number (1-based index)
    }
  }

  return null;
}

function clearTransactionsData( account_id ){
  const sheet = getUserSpreadsheet().getSheetByName(USER_TRANSACTIONS_SHEET);
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
  
  var sheet = getUserSpreadsheet().getSheetByName(USER_TRANSACTIONS_SHEET);
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
  var sheet = getUserSpreadsheet().getSheetByName(USER_TRANSACTIONS_SHEET);
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
        var accountName = getAccountNameByAccountId(account.account_id);
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
    var sheet = getUserSpreadsheet().getSheetByName(USER_TRANSACTIONS_SHEET);
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

function changeAccountNameOnTransactionSheet( account_id, account_name ){
  let sheet = getUserSpreadsheet().getSheetByName(USER_TRANSACTIONS_SHEET);
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
  let sheet = getUserSpreadsheet().getSheetByName(USER_TRANSACTIONS_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var row = 0; row < data.length; row++) {
    if (data[row].includes(account_id)) {
      sheet.getRange(row + 1 , 11).setValue(new_id); 
    }
  }
}
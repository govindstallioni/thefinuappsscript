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

function clearTransactionsData( account_id ){
  const sheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
  const data = sheet.getDataRange().getValues();
  const columnIndexToCheck = 11;

  // Collect rows to delete (bottom-up to preserve row indices)
  let rowsToDelete = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][columnIndexToCheck - 1] === account_id) {
      rowsToDelete.push(i + 1);
    }
  }

  // Delete rows from bottom to top so indices stay valid
  for (var r = rowsToDelete.length - 1; r >= 0; r--) {
    sheet.deleteRow(rowsToDelete[r]);
  }

  sortingTransactionSheet();
  return true;
}

/**
 * Removes transaction rows by transaction_id.
 * Plaid's /transactions/sync returns a `removed` array with { transaction_id } objects.
 * @param {Array} removedTransactions - Array of { transaction_id } objects from Plaid.
 */
function removeTransactionsFromSheet(removedTransactions){
  if(!removedTransactions || removedTransactions.length === 0) return;

  var sheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
  if(!sheet) return;

  var data = sheet.getDataRange().getValues();

  // Build a set of transaction_ids to remove
  var removeSet = {};
  removedTransactions.forEach(function(item){
    if(item.transaction_id) removeSet[item.transaction_id] = true;
  });

  // Find matching rows (skip header row 0)
  var rowsToDelete = [];
  for(var i = 1; i < data.length; i++){
    for(var c = 0; c < data[i].length; c++){
      if(removeSet[data[i][c]]){
        rowsToDelete.push(i + 1);
        break;
      }
    }
  }

  // Delete from bottom to top
  for(var r = rowsToDelete.length - 1; r >= 0; r--){
    sheet.deleteRow(rowsToDelete[r]);
  }
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
    var accountNumber = result.mask;
    var accountName = getPlaidAccountNameByAccountId(account_id);
    var institutionName = result.institution_name;

    let next_cursor = result.next_cursor;
    var has_more = true;

    // Paginate through all available pages
    while(has_more){
      let transactions = getPlaidTransactionSyncData( account_id, next_cursor );

      if(!transactions || transactions.error){
        break;
      }

      let TransactionAdded = transactions.added || [];
      let TransactionModified = transactions.modified || [];
      let TransactionRemoved = transactions.removed || [];

      // Track pending IDs that were updated via added[] so we don't delete them in removed[]
      var handledPendingIds = {};

      // --- 1. MODIFIED: update existing rows first ---
      if( TransactionModified.length > 0 ){
        TransactionModified.forEach(function(transaction){

          let transaction_status = transaction.pending == true ? 'Pending' : '';
          let txData = {
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

          if( transaction.pending_transaction_id != null && getTransactionRow( transaction.pending_transaction_id ) != null ){
            updateTransactionsData( transaction.pending_transaction_id, txData);
          }else if( getTransactionRow(transaction.transaction_id ) != null ){
            updateTransactionsData( transaction.transaction_id, txData);
          }

        });
      }

      // --- 2. ADDED: update pending rows or insert new ---
      var newTransactionCollection = [];

      if( TransactionAdded.length > 0 ){
        // Read fresh sheet data after modifications
        var sheetForLookup = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
        var freshData = sheetForLookup.getDataRange().getValues();

        TransactionAdded.forEach(function(transaction){

          let transaction_status = transaction.pending == true ? 'Pending' : '';
          let txData = {
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

          // Check if this finalized transaction replaces a pending one
          if( transaction.pending_transaction_id != null && getTransactionRow( transaction.pending_transaction_id, freshData ) != null ){
            // Update the pending row — preserves user's Category, Owner, Assigned
            updateTransactionsData( transaction.pending_transaction_id, txData);
            // Mark this pending ID so removed[] won't delete the row we just updated
            handledPendingIds[transaction.pending_transaction_id] = true;
          }else if( getTransactionRow(transaction.transaction_id, freshData ) != null ){
            // Transaction already exists — update it
            updateTransactionsData( transaction.transaction_id, txData);
          }else{
            // New transaction — insert
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
              transaction.account_id, //Account ID
              institutionName, //Institution
              transaction.transaction_id, //Transaction ID
              '', //Group
              '', //Type
              '' //Period
            ]);
          }

        });
      }

      if( newTransactionCollection.length > 0 ){
        insertTransactionsData(newTransactionCollection);
      }

      // --- 3. REMOVED: delete rows, but skip pending IDs already handled by added[] ---
      if( TransactionRemoved.length > 0 ){
        var filteredRemoved = TransactionRemoved.filter(function(item){
          return !handledPendingIds[item.transaction_id];
        });
        if( filteredRemoved.length > 0 ){
          removeTransactionsFromSheet(filteredRemoved);
        }
      }

      // Update cursor and check for more pages
      if( transactions.next_cursor && transactions.next_cursor !== ''){
        next_cursor = transactions.next_cursor;
        updateAppAccountDetailById(account_id, { next_cursor: transactions.next_cursor });
      }

      has_more = !!transactions.has_more;
    }

  }

  sortingTransactionSheet();

  return true;
}

function updateTransactionsData( transaction_id = null, transaction = null){

  var row = getTransactionRow(transaction_id);
  if( row == null ) return;

  var sheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Read existing row values to preserve user-entered data
  var existingValues = sheet.getRange(row, 1, 1, headers.length).getValues()[0];

  var rowValues = [];

  for (var j = 0; j < headers.length; j++) {
    var header = headers[j];
    var existing = existingValues[j];
    // Only overwrite columns that come from Plaid; preserve everything else
    switch (header) {
      case "Account":
        rowValues.push(transaction.account);
        break;
      case "Transaction Status":
        rowValues.push(transaction.transaction_status);
        break;
      case "Account Number":
        rowValues.push(transaction.account_number);
        break;
      case "Amount":
        rowValues.push(transaction.amount);
        break;
      case "Date":
        rowValues.push(transaction.date);
        break;
      case "Description":
        rowValues.push(transaction.description);
        break;
      case "Institution":
        rowValues.push(transaction.institution);
        break;
      case "Transaction ID":
        rowValues.push(transaction.transaction_id);
        break;
      case "Account ID":
        rowValues.push(transaction.account_id);
        break;
      default:
        // Preserve user-entered values (Category, Owner, Assigned, Not Cleared, Group, Type, Period)
        rowValues.push(existing !== undefined ? existing : '');
        break;
    }
  }

  if( rowValues.length > 0 ){
    var cell = sheet.getRange(row, 1, 1, headers.length);
      cell.setValues([rowValues]);
      cell.setFontSize(9)
      .setFontFamily("Comfortaa")
      .setFontColor("#000000")
      .setFontWeight("bold");
  }

}


function changeAccountNameOnTransactionSheet( account_id, account_name ){
  let sheet = UserSpreadsheet.getSheetByName(USER_TRANSACTIONS_SHEET);
  var data = sheet.getDataRange().getValues();
  for (var row = 0; row < data.length; row++) {
    if (data[row].includes(account_id)) {
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
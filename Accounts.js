function clearAccountData(account_id) {
  const sheet = getUserSpreadsheet().getSheetByName(USER_ACCOUNTS_SHEET);
  const range = sheet.getDataRange();
  const data = range.getValues();
  const columnIndexToCheck = 12; 

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
    sheet.getRangeList(rowsToClear.map(row => `A${row}:L${row}`)).clearContent();
    sortingBalanceHistorySheet();
    sortingAccountSheetByDate();
  }
}

function linkAccountsSheetData( account_id ){
  const sheet = getUserSpreadsheet().getSheetByName(USER_ACCOUNTS_SHEET);
  const ACCOUNT_ID_COL = 12; // column L (Account ID)

  // Read only the Account ID column starting from row 2
  const lastRow = Math.max( sheet.getLastRow(), 1 );
  if (lastRow < 2) {
    // sheet has only header or is empty — append a new row with account_id in col 12
    sheet.appendRow(new Array(sheet.getLastColumn()).fill(''));
    sheet.getRange(2, ACCOUNT_ID_COL).setValue(account_id);
    return;
  }

  const colVals = sheet.getRange(2, ACCOUNT_ID_COL, lastRow - 1, 1).getValues();

  // If account_id already present, do nothing
  for (let i = 0; i < colVals.length; i++) {
    if (colVals[i] && colVals[i][0] === account_id) return;
  }

  // Find first empty cell in Account ID column and set it, otherwise append
  for (let i = 0; i < colVals.length; i++) {
    if (!colVals[i] || colVals[i][0] === '') {
      sheet.getRange(2 + i, ACCOUNT_ID_COL).setValue(account_id);
      return;
    }
  }

  // No empty slot found — append a new row and set the Account ID cell only
  sheet.appendRow(new Array(sheet.getLastColumn()).fill(''));
  sheet.getRange(sheet.getLastRow(), ACCOUNT_ID_COL).setValue(account_id);
}

function sortingAccountSheetByDate(){
  const sheet = getUserSpreadsheet().getSheetByName(USER_ACCOUNTS_SHEET);

  const HEADER_ROW = 1;
  const START_ROW = 2;
  const DRIVER_COL = 12; // Account ID (Column L)
  const SORT_DATE_COL = 4; // Date column (optional)

  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return;

  const range = sheet.getRange(
    START_ROW,
    1,
    lastRow - HEADER_ROW,
    sheet.getLastColumn()
  );

  range.sort([
    { column: DRIVER_COL, ascending: false }, // content first
    { column: SORT_DATE_COL, ascending: false } // newest first (optional)
  ]);
}

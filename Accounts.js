function clearAccountData(account_id) {
  const sheet = UserSpreadsheet.getSheetByName(USER_ACCOUNTS_SHEET);
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
  let sheet = UserSpreadsheet.getSheetByName(USER_ACCOUNTS_SHEET);
  let lastrow = sheet.getLastRow() + 1;
  for( let i = 1; i <= lastrow; i++ ){
    if( sheet.getRange(i + 1, 12).getValue() === ''){
      sheet.getRange(i + 1, 12).setValue(account_id);
      break;
    }
  }
}

function sortingAccountSheetByDate(){
  const sheet = UserSpreadsheet.getSheetByName(USER_ACCOUNTS_SHEET);

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
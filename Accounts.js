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
  const sheet = UserSpreadsheet.getSheetByName(USER_ACCOUNTS_SHEET);
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

function updateAccountsSheet(oldAccount, newAccount) {
  const sheet = UserSpreadsheet.getSheetByName(USER_ACCOUNTS_SHEET);
  const normalize = v => v.toLowerCase().trim();

  const maxRows = sheet.getMaxRows();
  const colB = sheet.getRange(2, 2, maxRows - 1, 1).getValues();

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
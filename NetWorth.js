/**
 * Generates the Net Worth Statement report.
 * Standardized: Defaults to Liability unless "Asset" is explicitly found.
 */
function populateNetWorth() {
  var ss = SpreadsheetApp.getActive();
  var targetSheet = ss.getSheetByName("Net Worth");
  var defineSheet = ss.getSheetByName("Definition");
  var sourceSheet = ss.getSheetByName("Accounts");

  if (!targetSheet || !defineSheet || !sourceSheet) return;

  targetSheet.getRange('B3').setValue("⏳ Processing..");

  // --- 1. CLEANUP ---
  var startRow = 9;
  var lastRow = Math.max(targetSheet.getLastRow(), 200); 
  var clearRange = targetSheet.getRange('A' + startRow + ':H' + lastRow);

  clearRange.clear({ contentsOnly: true, formatOnly: false }); 
  clearRange.setBackground("#FFFFFF"); 
  clearRange.setFontColor(null);       

  // --- 2. DATA MAPPING ---
  var defValues = defineSheet.getRange("F5:F10").getValues();
  var colMap = { 
    acct: defValues[0][0] - 1, 
    grp: defValues[1][0] - 1, 
    al: defValues[2][0] - 1, 
    hide: defValues[3][0] - 1, 
    dt: defValues[4][0] - 1, 
    amt: defValues[5][0] - 1 
  };

  // --- 3. DATA PROCESSING ---
  var sourceLastRow = sourceSheet.getLastRow();
  if (sourceLastRow < 2) return;
  
  var maxCol = Math.max(colMap.acct, colMap.grp, colMap.al, colMap.hide, colMap.dt, colMap.amt) + 1;
  var sourceData = sourceSheet.getRange(2, 1, sourceLastRow - 1, maxCol).getValues();
  
  var tree = { 'Asset': {}, 'Liability': {} };
  var grandTotals = { 'Asset': 0, 'Liability': 0 };

  for (var i = 0; i < sourceData.length; i++) {
    var row = sourceData[i];
    if (!row[colMap.acct] || row[colMap.acct].toString().trim() === "") continue; 
    if (String(row[colMap.hide]).toLowerCase() == "yes") continue; 

    // FIXED PLACEMENT LOGIC: Default to Liability unless "Asset" is explicitly found
    var rawType = String(row[colMap.al]).trim().toLowerCase();
    var type = (rawType === "asset") ? "Asset" : "Liability"; 
    
    var group = row[colMap.grp] || "Uncategorized"; 
    var amt = row[colMap.amt];

    // Handle 0, empty, or "-" to show 0.00
    if (amt === "" || amt === "-" || amt === null || amt === undefined) {
      amt = 0;
    } else if (typeof amt !== 'number') {
      amt = parseFloat(String(amt).replace(/[^0-9.-]+/g, "")) || 0;
    }

    if (!tree[type][group]) {
      tree[type][group] = { items: [], total: 0 };
    }
    tree[type][group].items.push({ name: row[colMap.acct], date: row[colMap.dt], amt: amt });
    tree[type][group].total += amt;
    grandTotals[type] += amt;
  }

  // --- 4. BUILD OUTPUT ---
  var leftList = buildColumnStack(tree['Asset'], "ASSETS", grandTotals['Asset']);
  var rightList = buildColumnStack(tree['Liability'], "LIABILITIES", grandTotals['Liability']);
  
  var maxRows = Math.max(leftList.length, rightList.length);
  var outputGrid = [];
  for (var r = 0; r < maxRows; r++) {
    var rowData = [];
    rowData = rowData.concat((r < leftList.length) ? leftList[r] : ["", "", "", ""]);
    rowData = rowData.concat((r < rightList.length) ? rightList[r] : ["", "", "", ""]);
    outputGrid.push(rowData);
  }

  // --- 5. WRITE & FORMAT ---
  if (outputGrid.length > 0) {
    var writeRange = targetSheet.getRange(startRow, 1, outputGrid.length, 8);
    writeRange.setValues(outputGrid);
    
    var currencyFmt = '_($* #,##0.00_);_($* (#,##0.00)_);_($* 0.00_);_(@_)';
    targetSheet.getRange(startRow, 4, outputGrid.length, 1).setNumberFormat(currencyFmt); 
    targetSheet.getRange(startRow, 8, outputGrid.length, 1).setNumberFormat(currencyFmt); 
    targetSheet.getRange(startRow, 3, outputGrid.length, 1).setNumberFormat('M/d/yyyy');    
    targetSheet.getRange(startRow, 7, outputGrid.length, 1).setNumberFormat('M/d/yyyy');

    for (var r = 0; r < outputGrid.length; r++) {
      var markerA = outputGrid[r][0];
      var markerE = outputGrid[r][4];
      var sheetRow = r + startRow;

      if (markerA === "G") targetSheet.getRange(sheetRow, 2, 1, 3).setBackground("#6E8277").setFontColor("#FFFFFF");
      if (markerE === "G") targetSheet.getRange(sheetRow, 6, 1, 3).setBackground("#6E8277").setFontColor("#FFFFFF");
      if (markerA === "AL") targetSheet.getRange(sheetRow, 2, 1, 3).setBackground("#7D404A").setFontColor("#FFFFFF");
      if (markerE === "AL") targetSheet.getRange(sheetRow, 6, 1, 3).setBackground("#7D404A").setFontColor("#FFFFFF");
      if (outputGrid[r][1] === "Accounts") {
          targetSheet.getRange(sheetRow, 2, 1, 3).setFontColor("#FFFFFF").setBackground("#385350");
          targetSheet.getRange(sheetRow, 6, 1, 3).setFontColor("#FFFFFF").setBackground("#385350");
      }
    }
  }

  targetSheet.getRange("A:A").setBackground("#FFFFFF").setFontColor("#F9F9F9"); 
  targetSheet.getRange("E:E").setBackground("#FFFFFF").setFontColor("#F9F9F9");
  targetSheet.getRange('B3').setValue("Last updated on " + getDateTime());
}

function buildColumnStack(groupObj, title, grandTotal) {
  var stack = [];
  stack.push(["AL", title, "", grandTotal]);
  stack.push(["", "", "", ""]);
  stack.push(["", "Accounts", "Last Updated", "Amount"]);
  var sortedGroups = Object.keys(groupObj).sort();
  for (var i = 0; i < sortedGroups.length; i++) {
    var grpName = sortedGroups[i];
    var grpData = groupObj[grpName];
    stack.push(["", "", "", ""]); 
    stack.push(["G", grpName, "", grpData.total]); 
    grpData.items.sort((a, b) => a.name.localeCompare(b.name));
    for (var j = 0; j < grpData.items.length; j++) {
      var item = grpData.items[j];
      stack.push(["C", item.name, item.date, item.amt]);
    }
  }
  return stack;
}
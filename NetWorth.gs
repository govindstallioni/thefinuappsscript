/**
 * Generates the Net Worth Statement report, with fixed formatting to hide markers 
 * and prevent the collapse/narrowing issue.
 */
function populateNetWorth() {
  var ss = SpreadsheetApp.getActive();
  var targetSheet = ss.getSheetByName("Net Worth");
  var defineSheet = ss.getSheetByName("Definition");
  var sourceSheet = ss.getSheetByName("Accounts");

  // --- 1. CLEANUP ---
  var startRow = 9;
  var lastRow = Math.max(targetSheet.getLastRow(), 200); 
  var clearRange = targetSheet.getRange('A' + startRow + ':H' + lastRow);

  // Clear everything (content and formatting) in the report area
  clearRange.clear({ contentsOnly: true, formatOnly: false }); 
  clearRange.setBackground("#FFFFFF"); 
  clearRange.setFontColor(null);       

  // --- 2 & 3. DATA MAPPING AND PROCESSING (Skipped for brevity) ---
  
  var defValues = defineSheet.getRange("F5:F10").getValues();
  var colMap = { acct: defValues[0][0] - 1, grp: defValues[1][0] - 1, al: defValues[2][0] - 1, hide: defValues[3][0] - 1, dt: defValues[4][0] - 1, amt: defValues[5][0] - 1 };
  var sourceLastRow = sourceSheet.getLastRow();
  if (sourceLastRow < 2) return;
  var maxCol = Math.max(colMap.acct, colMap.grp, colMap.al, colMap.hide, colMap.dt, colMap.amt) + 1;
  var sourceData = sourceSheet.getRange(2, 1, sourceLastRow - 1, maxCol).getValues();
  
  var tree = { 'Asset': {}, 'Liability': {} };
  var grandTotals = { 'Asset': 0, 'Liability': 0 };

  for (var i = 0; i < sourceData.length; i++) {
    var row = sourceData[i];
    if (row.join("").trim() === "") continue;
    if (String(row[colMap.hide]).toLowerCase() == "yes") continue;

    var rawType = row[colMap.al];
    var type = (rawType == "Asset") ? "Asset" : "Liability"; 
    var group = row[colMap.grp] || "Uncategorized"; 
    var amt = row[colMap.amt];
    
    amt = (typeof amt === 'number' && !isNaN(amt)) ? amt : (typeof amt === 'string' ? parseFloat(String(amt).replace(/[^0-9.-]+/g, "")) : 0) || 0;
    if (amt === 0 && group === "Uncategorized") continue;

    if (!tree[type][group]) {
      tree[type][group] = { items: [], total: 0 };
    }
    tree[type][group].items.push({ name: row[colMap.acct], date: row[colMap.dt], amt: amt });
    tree[type][group].total += amt;
    grandTotals[type] += amt;
  }
  
  for (var type in tree) {
      for (var group in tree[type]) {
          if (tree[type][group].total === 0 && tree[type][group].items.length === 0) {
              delete tree[type][group];
          }
      }
  }

  // --- 4 & 5. BUILD AND MERGE STACKS (Skipped for brevity) ---
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

  // --- 6. WRITE DATA & APPLY FORMATTING ---
  if (outputGrid.length > 0) {
    var writeRange = targetSheet.getRange(startRow, 1, outputGrid.length, 8);
    writeRange.setValues(outputGrid);
    
    // Apply Currency/Date Formatting
    var currencyFmt = '_($* #,##0.00_);_($* (#,##0.00)_);_($* "-"_);_(@_)';
    var dateFmt = 'M/d/yyyy';
    targetSheet.getRange(startRow, 4, outputGrid.length, 1).setNumberFormat(currencyFmt); 
    targetSheet.getRange(startRow, 8, outputGrid.length, 1).setNumberFormat(currencyFmt); 
    targetSheet.getRange(startRow, 3, outputGrid.length, 1).setNumberFormat(dateFmt);    
    targetSheet.getRange(startRow, 7, outputGrid.length, 1).setNumberFormat(dateFmt);    
    
    // APPLY DIRECT BACKGROUND STYLING (The fix is in the coloring ranges)
    for (var r = 0; r < outputGrid.length; r++) {
      var markerA = outputGrid[r][0]; 
      var markerE = outputGrid[r][4];
      var sheetRow = r + startRow;
      
      // Color: #6E8277 (Groups) - Start at Col 2 (B)
      if (markerA === "G") {
        targetSheet.getRange(sheetRow, 2, 1, 3).setBackground("#6E8277").setFontColor("#FFFFFF"); 
        targetSheet.getRange(sheetRow, 1, 1, 1).setBackground("#FFFFFF").setFontColor("#FFFFFF");
      }
      if (markerE === "G") {
        targetSheet.getRange(sheetRow, 6, 1, 3).setBackground("#6E8277").setFontColor("#FFFFFF");
        targetSheet.getRange(sheetRow, 5, 1, 1).setBackground("#FFFFFF").setFontColor("#FFFFFF");
      }
      
      // Color: #7D404A (AL Headers) - Start at Col 2 (B)
      if (markerA === "AL") {
        targetSheet.getRange(sheetRow, 2, 1, 3).setBackground("#7D404A").setFontColor("#FFFFFF");
      }
      if (markerE === "AL") {
        targetSheet.getRange(sheetRow, 6, 1, 3).setBackground("#7D404A").setFontColor("#FFFFFF");
      }
      
      // Explicitly set Header Row (Row 11) background
      if (r === 2) {
          targetSheet.getRange(sheetRow, 2, 1, 3).setFontColor("#FFFFFF").setBackground("#385350");
          targetSheet.getRange(sheetRow, 6, 1, 3).setFontColor("#FFFFFF").setBackground("#385350");
      }
    }
  }

  // --- 7. FINAL COSMETICS (FIXED HIDING) ---
  
  // 🎯 FIX: Explicitly set the background of the marker columns to white, just in case.
  // This guarantees no dark color from the formatting loops bleeds in.
  targetSheet.getRange("A:A").setBackground("#FFFFFF"); 
  targetSheet.getRange("E:E").setBackground("#FFFFFF");

  // 🎯 FIX: Set font color to a near-white or transparent-like color (e.g., #F9F9F9) 
  // This ensures the marker text is invisible against the white background.
  targetSheet.getRange("A:A").setFontColor("#F9F9F9");
  targetSheet.getRange("E:E").setFontColor("#F9F9F9");

  targetSheet.getRange('B3').setValue("Last updated on " + getDateTime() );
  //targetSheet.getRange('B2').activate();
}

/**
 * Helper to build the 4-column stack remains unchanged.
 */
function buildColumnStack(groupObj, title, grandTotal) {
  
  var stack = [];
  
  // Row 9 (Index 0): Main Header Marker "AL"
  stack.push(["AL", title, "", grandTotal]);
  
  // Row 10 (Index 1): SPACER Row (Empty) - CRUCIAL for template spacing
  stack.push(["", "", "", ""]);
  
  // Row 11 (Index 2): Column Headers (Accounts, Last Updated, Amount)
  stack.push(["", "Accounts", "Last Updated", "Amount"]);

  var sortedGroups = Object.keys(groupObj).sort();

  for (var i = 0; i < sortedGroups.length; i++) {
    var grpName = sortedGroups[i];
    var grpData = groupObj[grpName];

    // 🛑 IMPROVEMENT 1: Add a break before the start of each Group
    stack.push(["", "", "", ""]); 
    
    // Group Header Row (Marker "G")
    stack.push(["G", grpName, "", grpData.total]);

    // Items
    grpData.items.sort(function(a,b){ return a.name.localeCompare(b.name); });
    
    for (var k = 0; k < grpData.items.length; k++) {
      var item = grpData.items[k];
      
      // 🛑 IMPROVEMENT 2: Add a break before the start of each Account Item 
      // (Only if it's not the first item, to prevent a double break after the Group Header)
      if (k > 0) {
          stack.push(["", "", "", ""]);
      }
      
      // Item Row (Marker "C")
      stack.push(["C", item.name, item.date, item.amt]);
    }
  }
  
  return stack;
}
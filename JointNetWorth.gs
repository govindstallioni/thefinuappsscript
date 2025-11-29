function populateJointNetWorth() {
  const ss = SpreadsheetApp.getActive();
  const targetSheet = ss.getSheetByName("Joint Net Worth");
  const defineSheet = ss.getSheetByName("Definition");
  const sourceSheet = ss.getSheetByName("Accounts");
  const startRow = 9;

  // --- 1. CLEANUP ---
  const lastRow = Math.max(targetSheet.getLastRow(), startRow); 
  const clearRange = targetSheet.getRange('A' + startRow + ':L' + lastRow);

  // Clear everything (content and formatting) in the report area (A9:L)
  clearRange.clear({ contentsOnly: true, formatOnly: false }); 
  clearRange.setBackground("#FFFFFF"); 
  clearRange.setFontColor(null);       

  // --- 2. GET DEFINITIONS & DATA ---
  
  // Column mapping (F5:F12)
  const defValues = defineSheet.getRange("F5:F12").getValues();
  const colMap = {
    acct: defValues[0][0] - 1, // B/H Column
    grp:  defValues[1][0] - 1, // B/H Column
    al:   defValues[2][0] - 1, // Marker Logic
    hide: defValues[3][0] - 1, // Filter Logic
    dt:   defValues[4][0] - 1, // C/I Column
    amt:  defValues[5][0] - 1, // F/L Column
    owner: defValues[6][0] - 1, // Owner column
    assigned: defValues[7][0] - 1 // Assigned amount column
  };

  // Owner Names (C11, C12)
  const name1 = defineSheet.getRange('C11').getValue() || "Owner 1";
  const name2 = defineSheet.getRange('C12').getValue() || "Owner 2";
    
  // Read all source data
  const sourceLastRow = sourceSheet.getLastRow();
  if (sourceLastRow < 2) return;
  const maxColIndex = Math.max(...Object.values(colMap)) + 1;
  const sourceData = sourceSheet.getRange(2, 1, sourceLastRow - 1, maxColIndex).getValues();
  
  // --- 3. PROCESS DATA INTO HIERARCHY AND CALCULATE TOTALS ---
  
  // Data structure: tree[Type][Group] = { items: [], total: { total: 0, name1: 0, name2: 0 } }
  let tree = { 'Asset': {}, 'Liability': {} };
  
  for (const row of sourceData) {
    if (row.join("").trim() === "" || String(row[colMap.hide]).toLowerCase() === "yes") continue;

    const type = (String(row[colMap.al]) === "Asset") ? "Asset" : "Liability"; 
    const group = row[colMap.grp] || "Uncategorized"; 
    const owner = String(row[colMap.owner]).trim();
    
    // Safety check and cleanup for amounts
    let amt = row[colMap.amt];
    amt = (typeof amt === 'number' && !isNaN(amt)) ? amt : (typeof amt === 'string' ? parseFloat(String(amt).replace(/[^0-9.-]+/g, "")) : 0) || 0;
    
    // Account for assigned amount based on owner
    let assignedAmt = row[colMap.assigned];
    assignedAmt = (typeof assignedAmt === 'number' && !isNaN(assignedAmt)) ? assignedAmt : (typeof assignedAmt === 'string' ? parseFloat(String(assignedAmt).replace(/[^0-9.-]+/g, "")) : 0) || 0;
    
    let name1Amt = 0;
    let name2Amt = 0;
    
    if (owner === name1) {
      name1Amt = assignedAmt;
    } else if (owner === name2) {
      name2Amt = assignedAmt;
    } else {
      // Default to split if Owner column is empty or invalid
      name1Amt = assignedAmt;
      name2Amt = assignedAmt;
    }

    if (amt === 0 && group === "Uncategorized") continue;

    if (!tree[type][group]) {
      tree[type][group] = { items: [], total: { total: 0, name1: 0, name2: 0 } };
    }

    tree[type][group].items.push({
      name: row[colMap.acct],
      date: row[colMap.dt],
      amt: amt,
      name1Amt: name1Amt,
      name2Amt: name2Amt
    });

    // Update totals
    tree[type][group].total.total += amt;
    tree[type][group].total.name1 += name1Amt;
    tree[type][group].total.name2 += name2Amt;
  }
  
  // --- 4. BUILD VISUAL STACKS & MERGE ---
  
  // Totals for headers (only used in the AL row)
  const assetGrandTotals = calculateGrandTotals(tree['Asset']);
  const liabilityGrandTotals = calculateGrandTotals(tree['Liability']);

  // Build the stacks using the optimized helper function
  const leftList = buildJointColumnStack(tree['Asset'], "ASSETS", name1, name2, assetGrandTotals);
  const rightList = buildJointColumnStack(tree['Liability'], "LIABILITIES", name1, name2, liabilityGrandTotals);

  const maxRows = Math.max(leftList.length, rightList.length);
  let outputGrid = [];

  for (let r = 0; r < maxRows; r++) {
    let rowData = [];
    // 6 columns for Left Side (A-F)
    rowData = rowData.concat((r < leftList.length) ? leftList[r] : ["", "", "", "", "", ""]);
    // 6 columns for Right Side (G-L)
    rowData = rowData.concat((r < rightList.length) ? rightList[r] : ["", "", "", "", "", ""]);
    outputGrid.push(rowData);
  }

  // --- 5. WRITE DATA & APPLY FORMATTING ---
  if (outputGrid.length > 0) {
    const writeRange = targetSheet.getRange(startRow, 1, outputGrid.length, 12);
    // Write all data in one go
    writeRange.setValues(outputGrid);
    
    // Define format
    const currencyFmt = '_($* #,##0.00_);_($* (#,##0.00)_);_($* "-"_);_(@_)';
    const dateFmt = 'M/d/yyyy';
    
    // Apply Currency/Date Formatting to all amount columns (D, E, F, J, K, L) and date columns (C, I)
    
    // Currency Columns: D, E, F, J, K, L
    targetSheet.getRange(startRow, 4, outputGrid.length, 3).setNumberFormat(currencyFmt); 
    targetSheet.getRange(startRow, 10, outputGrid.length, 3).setNumberFormat(currencyFmt); 
    
    // Date Columns: C, I
    targetSheet.getRange(startRow, 3, outputGrid.length, 1).setNumberFormat(dateFmt);    
    targetSheet.getRange(startRow, 9, outputGrid.length, 1).setNumberFormat(dateFmt);    
    
    // Apply Direct Background Styling based on markers
    const groupColor = "#6E8277"; // Group Header color
    const alColor = "#7D404A";    // AL Header color
    const headerRowColor = "#385350"; // Column Header color (Row 11)

    for (let r = 0; r < outputGrid.length; r++) {
      const markerA = outputGrid[r][0]; // Marker in Col A
      const markerG = outputGrid[r][6]; // Marker in Col G
      const sheetRow = r + startRow;
      
      // Color Ranges: B:F (Left) and H:L (Right)
      
      if (markerA === "G") {
        targetSheet.getRange(sheetRow, 2, 1, 5).setBackground(groupColor).setFontColor("#FFFFFF"); 
      }
      if (markerG === "G") {
        targetSheet.getRange(sheetRow, 8, 1, 5).setBackground(groupColor).setFontColor("#FFFFFF");
      }
      
      if (markerA === "AL") {
        targetSheet.getRange(sheetRow, 2, 1, 5).setBackground(alColor).setFontColor("#FFFFFF");
      }
      if (markerG === "AL") {
        targetSheet.getRange(sheetRow, 8, 1, 5).setBackground(alColor).setFontColor("#FFFFFF");
      }
      
      // Explicitly color the column header row (Row 11 is r=2)
      if (r === 2) {
          targetSheet.getRange(sheetRow, 2, 1, 5).setFontColor("#FFFFFF").setBackground(headerRowColor);
          targetSheet.getRange(sheetRow, 8, 1, 5).setFontColor("#FFFFFF").setBackground(headerRowColor);
      }
      
      // Ensure Item rows (marker "C") are white if needed (B-F and H-L)
      if (markerA === "C") {
          targetSheet.getRange(sheetRow, 2, 1, 5).setBackground("#FFFFFF").setFontColor("#000000"); 
      }
      if (markerG === "C") {
          targetSheet.getRange(sheetRow, 8, 1, 5).setBackground("#FFFFFF").setFontColor("#000000"); 
      }
    }
  }

  // --- 6. FINAL COSMETICS ---
  
  // Hide the Marker columns (A and G) by setting font color to white/near-white
  targetSheet.getRange("A:A").setBackground("#FFFFFF").setFontColor("#F9F9F9");
  targetSheet.getRange("G:G").setBackground("#FFFFFF").setFontColor("#F9F9F9");

  // Update Timestamp
  targetSheet.getRange('B3').setValue("Last updated on " + getDateTime());
  //targetSheet.getRange('B2').activate();
}

/**
 * Helper to calculate grand totals for Asset or Liability type
 */
function calculateGrandTotals(typeTree) {
    let total = 0, name1 = 0, name2 = 0;
    for (const group in typeTree) {
        total += typeTree[group].total.total;
        name1 += typeTree[group].total.name1;
        name2 += typeTree[group].total.name2;
    }
    return { total, name1, name2 };
}

/**
 * Helper to build the 6-column stack (Marker, Account/Group, Date, Name1 Amt, Name2 Amt, Total Amt).
 * Includes breaks before groups and accounts.
 */
function buildJointColumnStack(groupObj, title, name1, name2, grandTotals) {
  let stack = [];
  
  // Row 9 (Index 0): Main Header Marker "AL"
  stack.push(["AL", title, "", grandTotals.name1, grandTotals.name2, grandTotals.total]);
  
  // Row 10 (Index 1): SPACER Row (Empty) - CRUCIAL for template spacing
  stack.push(["", "", "", "", "", ""]);
  
  // Row 11 (Index 2): Column Headers 
  stack.push(["", "Accounts", "Last Updated", name1, name2, "Amount"]);

  const sortedGroups = Object.keys(groupObj).sort();

  for (const grpName of sortedGroups) {
    const grpData = groupObj[grpName];

    // 1. Break before the start of each Group
    stack.push(["", "", "", "", "", ""]); 
    
    // 2. Group Header Row (Marker "G")
    stack.push(["G", grpName, "", grpData.total.name1, grpData.total.name2, grpData.total.total]);

    // Items
    grpData.items.sort((a, b) => a.name.localeCompare(b.name));
    
    for (let k = 0; k < grpData.items.length; k++) {
      const item = grpData.items[k];
      
      // 3. Add a break before the start of each Account Item (if not the first item)
      if (k > 0) {
          stack.push(["", "", "", "", "", ""]);
      }
      
      // 4. Item Row (Marker "C")
      stack.push([
        "C", 
        item.name, 
        item.date, 
        item.name1Amt, 
        item.name2Amt, 
        item.amt
      ]);
    }
  }
  
  return stack;
}
/**
 * Generates the Joint Net Worth Statement report.
 * Standardized: Defaults to Liability unless "Asset" is explicitly found.
 */
function populateJointNetWorth() {
  const ss = SpreadsheetApp.getActive();
  const targetSheet = ss.getSheetByName("Joint Net Worth");
  const defineSheet = ss.getSheetByName("Definition");
  const sourceSheet = ss.getSheetByName("Accounts");
  const startRow = 9;

  if (!targetSheet || !defineSheet || !sourceSheet) return;

  targetSheet.getRange('B3').setValue("⏳ Processing..");

  // --- 1. CLEANUP ---
  const lastRow = Math.max(targetSheet.getLastRow(), startRow);
  const clearRange = targetSheet.getRange('A' + startRow + ':L' + lastRow);
  clearRange.clear({ contentsOnly: true, formatOnly: false }); 
  clearRange.setBackground("#FFFFFF"); 
  clearRange.setFontColor(null);

  // --- 2. GET DEFINITIONS ---
  const defValues = defineSheet.getRange("F5:F12").getValues();
  const colMap = {
    acct: defValues[0][0] - 1,
    grp:  defValues[1][0] - 1,
    al:   defValues[2][0] - 1,
    hide: defValues[3][0] - 1,
    dt:   defValues[4][0] - 1,
    amt:  defValues[5][0] - 1,
    owner: defValues[6][0] - 1,
    assigned: defValues[7][0] - 1
  };

  const name1 = defineSheet.getRange('C11').getValue() || "Owner 1";
  const name2 = defineSheet.getRange('C12').getValue() || "Owner 2";

  const sourceLastRow = sourceSheet.getLastRow();
  if (sourceLastRow < 2) return;
  
  const maxColIndex = Math.max(...Object.values(colMap)) + 1;
  const sourceData = sourceSheet.getRange(2, 1, sourceLastRow - 1, maxColIndex).getValues();

  // --- 3. PROCESS DATA ---
  let tree = { 'Asset': {}, 'Liability': {} };

  for (const row of sourceData) {
    if (!row[colMap.acct] || row.join("").trim() === "" || String(row[colMap.hide]).toLowerCase() === "yes") continue;

    // FIXED PLACEMENT LOGIC: Match NetWorth.gs behavior
    const rawType = String(row[colMap.al]).trim().toLowerCase();
    const type = (rawType === "asset") ? "Asset" : "Liability"; 
    
    const group = row[colMap.grp] || "Uncategorized"; 
    const owner = String(row[colMap.owner]).trim();

    let rawAmt = row[colMap.amt];
    let amt = (rawAmt === "" || rawAmt === "-" || rawAmt === null) ? 0 : 
              (typeof rawAmt === 'number' && !isNaN(rawAmt)) ? rawAmt : 
              (parseFloat(String(rawAmt).replace(/[^0-9.-]+/g, "")) || 0);
    
    let rawAssigned = row[colMap.assigned];
    let assignedAmt = (rawAssigned === "" || rawAssigned === "-" || rawAssigned === null) ? 0 : 
                      (typeof rawAssigned === 'number' && !isNaN(rawAssigned)) ? rawAssigned : 
                      (parseFloat(String(rawAssigned).replace(/[^0-9.-]+/g, "")) || 0);
    
    let name1Amt = 0, name2Amt = 0;
    if (owner === name1) { name1Amt = assignedAmt; } 
    else if (owner === name2) { name2Amt = assignedAmt; } 
    else { name1Amt = assignedAmt; name2Amt = assignedAmt; }

    if (!tree[type][group]) {
      tree[type][group] = { items: [], total: { total: 0, name1: 0, name2: 0 } };
    }

    tree[type][group].items.push({
      name: row[colMap.acct], date: row[colMap.dt], amt: amt, name1Amt: name1Amt, name2Amt: name2Amt
    });

    tree[type][group].total.total += amt;
    tree[type][group].total.name1 += name1Amt;
    tree[type][group].total.name2 += name2Amt;
  }
  
  // --- 4. BUILD VISUAL STACKS ---
  const assetGrandTotals = calculateGrandTotals(tree['Asset']);
  const liabilityGrandTotals = calculateGrandTotals(tree['Liability']);
  const leftList = buildJointColumnStack(tree['Asset'], "ASSETS", name1, name2, assetGrandTotals);
  const rightList = buildJointColumnStack(tree['Liability'], "LIABILITIES", name1, name2, liabilityGrandTotals);

  const maxRows = Math.max(leftList.length, rightList.length);
  let outputGrid = [];
  for (let r = 0; r < maxRows; r++) {
    let rowData = [];
    rowData = rowData.concat((r < leftList.length) ? leftList[r] : ["", "", "", "", "", ""]);
    rowData = rowData.concat((r < rightList.length) ? rightList[r] : ["", "", "", "", "", ""]);
    outputGrid.push(rowData);
  }

  // --- 5. WRITE & FORMAT ---
  if (outputGrid.length > 0) {
    const writeRange = targetSheet.getRange(startRow, 1, outputGrid.length, 12);
    writeRange.setValues(outputGrid);
    const currencyFmt = '_($* #,##0.00_);_($* (#,##0.00)_);_($* 0.00_);_(@_)';
    targetSheet.getRange(startRow, 4, outputGrid.length, 3).setNumberFormat(currencyFmt);
    targetSheet.getRange(startRow, 10, outputGrid.length, 3).setNumberFormat(currencyFmt); 
    targetSheet.getRange(startRow, 3, outputGrid.length, 1).setNumberFormat('M/d/yyyy');    
    targetSheet.getRange(startRow, 9, outputGrid.length, 1).setNumberFormat('M/d/yyyy');

    for (let r = 0; r < outputGrid.length; r++) {
      const markerA = outputGrid[r][0], markerG = outputGrid[r][6], sheetRow = r + startRow;
      if (markerA === "G") targetSheet.getRange(sheetRow, 2, 1, 5).setBackground("#6E8277").setFontColor("#FFFFFF");
      if (markerG === "G") targetSheet.getRange(sheetRow, 8, 1, 5).setBackground("#6E8277").setFontColor("#FFFFFF");
      if (markerA === "AL") targetSheet.getRange(sheetRow, 2, 1, 5).setBackground("#7D404A").setFontColor("#FFFFFF");
      if (markerG === "AL") targetSheet.getRange(sheetRow, 8, 1, 5).setBackground("#7D404A").setFontColor("#FFFFFF");
      if (r === 2) {
          targetSheet.getRange(sheetRow, 2, 1, 5).setFontColor("#FFFFFF").setBackground("#385350");
          targetSheet.getRange(sheetRow, 8, 1, 5).setFontColor("#FFFFFF").setBackground("#385350");
      }
      if (markerA === "C") targetSheet.getRange(sheetRow, 2, 1, 5).setBackground("#FFFFFF").setFontColor("#000000");
      if (markerG === "C") targetSheet.getRange(sheetRow, 8, 1, 5).setBackground("#FFFFFF").setFontColor("#000000");
    }
  }

  targetSheet.getRange("A:A").setBackground("#FFFFFF").setFontColor("#F9F9F9");
  targetSheet.getRange("G:G").setBackground("#FFFFFF").setFontColor("#F9F9F9");
  targetSheet.getRange('B3').setValue("Last updated on " + getDateTime());
}

function calculateGrandTotals(typeTree) {
    let total = 0, name1 = 0, name2 = 0;
    for (const group in typeTree) {
        total += typeTree[group].total.total;
        name1 += typeTree[group].total.name1;
        name2 += typeTree[group].total.name2;
    }
    return { total, name1, name2 };
}

function buildJointColumnStack(groupObj, title, name1, name2, grandTotals) {
  let stack = [];
  stack.push(["AL", title, "", grandTotals.name1, grandTotals.name2, grandTotals.total]);
  stack.push(["", "", "", "", "", ""]);
  stack.push(["", "Accounts", "Last Updated", name1, name2, "Amount"]);
  const sortedGroups = Object.keys(groupObj).sort();
  for (const grpName of sortedGroups) {
    const grpData = groupObj[grpName];
    stack.push(["", "", "", "", "", ""]);
    stack.push(["G", grpName, "", grpData.total.name1, grpData.total.name2, grpData.total.total]); 
    grpData.items.sort((a, b) => a.name.localeCompare(b.name));
    for (let k = 0; k < grpData.items.length; k++) {
      const item = grpData.items[k];
      stack.push(["C", item.name, item.date, item.name1Amt, item.name2Amt, item.amt]);
    }
  }
  return stack;
}
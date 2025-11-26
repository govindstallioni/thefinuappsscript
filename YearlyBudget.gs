/**
 * Optimized Yearly Budget Generator
 * Fixes:
 * - Corrects Actual Amount accumulation (Math.abs for Expenses).
 * - Adds Borders to report rows.
 * - Adds a Spacer row after the last category of each group.
 * - Preserves Hierarchy, Formatting, and Font (Comfortaa, 10px).
 */
function populateYearlyBudget() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outSheet = ss.getSheetByName("Yearly Budget");
  const defSheet = ss.getSheetByName("Definition");
  const catSheet = ss.getSheetByName("Categories");
  const tranSheet = ss.getSheetByName("Transactions");

  if (!outSheet || !defSheet || !catSheet || !tranSheet) {
    SpreadsheetApp.getUi().alert("Missing required sheets.");
    return;
  }

  // --- 1. CONFIGURATION & SETUP ---
  
  const configRaw = defSheet.getRange('C5:C12').getValues().flat();
  const tranConfigRaw = defSheet.getRange('I5:I12').getValues().flat();
  const monthColsRaw = defSheet.getRange('W2:W13').getValues().flat();
  
  const CONFIG = {
    CAT_COL: configRaw[0] - 1,
    GRP_COL: configRaw[1] - 1,
    TYP_COL: configRaw[2] - 1,
    HIDE_COL: configRaw[3] - 1,
    ALLOC1: configRaw[4] - 1,
    ALLOC2: configRaw[5] - 1,
    YEAR: defSheet.getRange('V1').getValue(),
    BUDGET_COLS: monthColsRaw.map(c => c - 1),
    
    TRAN_CAT: tranConfigRaw[0] - 1,
    TRAN_GRP: tranConfigRaw[1] - 1,
    TRAN_TYP: tranConfigRaw[2] - 1, // Type column in Transactions sheet
    TRAN_DATE: tranConfigRaw[4] - 1,
    TRAN_AMT: tranConfigRaw[5] - 1, // Amount column in Transactions sheet
  };

  // Clear Output Sheet
  const lastRow = outSheet.getLastRow();
  const lastCol = outSheet.getLastColumn();
  if (lastRow >= 10) {
    // Fix merged cells before clearing to avoid errors
    outSheet.getRange(10, 1, lastRow - 9, lastCol).breakApart();
    outSheet.getRange(10, 1, lastRow - 9, lastCol).clear({contentsOnly: true, formatOnly: true});
    outSheet.getRange(10, 1, lastRow - 9, lastCol).setBackground(null).setFontWeight(null).setFontStyle(null).setBorder(false, false, false, false, false, false);
  }
  outSheet.getRange("A:A").setFontColor("#FFFFFF");


  // --- 2. DATA PROCESSING (IN-MEMORY) ---

  const tree = {}; 
  const catMap = {}; 

  // A. Process Categories
  const catData = catSheet.getDataRange().getValues();
  
  for (let i = 1; i < catData.length; i++) { 
    const row = catData[i];
    const type = row[CONFIG.TYP_COL];
    const group = row[CONFIG.GRP_COL];
    const catName = row[CONFIG.CAT_COL];
    const hide = row[CONFIG.HIDE_COL];

    if (hide === "Hide" || !type || !group || !catName) continue;

    if (!tree[type]) tree[type] = {};
    if (!tree[type][group]) tree[type][group] = {};

    const catObj = {
      name: catName,
      group: group,
      type: type,
      budget: Array(12).fill(0),
      actual: Array(12).fill(0),
    };

    for (let m = 0; m < 12; m++) {
      const val = row[CONFIG.BUDGET_COLS[m]];
      catObj.budget[m] = (typeof val === 'number') ? val : 0;
    }

    tree[type][group][catName] = catObj;
    const key = `${catName}_${group}_${type}`;
    catMap[key] = catObj;
  }

  // B. Process Transactions
  const tranData = tranSheet.getDataRange().getValues();
  const targetYear = CONFIG.YEAR;

  for (let i = 1; i < tranData.length; i++) {
    const row = tranData[i];
    const date = row[CONFIG.TRAN_DATE];
    
    if (!date || new Date(date).getFullYear() != targetYear) continue;

    const catName = row[CONFIG.TRAN_CAT];
    const group = row[CONFIG.TRAN_GRP];
    const type = row[CONFIG.TRAN_TYP]; // e.g., "Income", "Expense"
    let amt = Number(row[CONFIG.TRAN_AMT]) || 0; // Raw amount from Transactions sheet
    const monthIdx = new Date(date).getMonth(); 

    const key = `${catName}_${group}_${type}`;
    
    if (catMap[key]) {
      // If the category type is NOT Income or Transfers, treat the amount as an Expense magnitude.
      if (type !== 'Income' && type !== 'Transfers') {
        amt = Math.abs(amt); 
      }
      
      catMap[key].actual[monthIdx] += amt;
    }
  }

  // --- 3. OUTPUT GENERATION ---

  const outputRows = [];
  const metaRows = []; 
  const grandTotal = { budget: 0, actual: 0 };

  const sortedTypes = Object.keys(tree).sort((a, b) => {
    if (a === "Income") return -1;
    if (b === "Income") return 1;
    return a.localeCompare(b);
  });

  sortedTypes.forEach(type => {
    const groups = tree[type];
    const typeTotals = { budget: Array(12).fill(0), actual: Array(12).fill(0) };
    
    const typeHeaderIndex = outputRows.length;
    outputRows.push(null); 
    metaRows.push("TYPE");

    const sortedGroups = Object.keys(groups).sort();

    sortedGroups.forEach(group => {
      const cats = groups[group];
      const groupTotals = { budget: Array(12).fill(0), actual: Array(12).fill(0) };

      const groupHeaderIndex = outputRows.length;
      outputRows.push(null); 
      metaRows.push("GROUP");

      const sortedCats = Object.keys(cats).sort();

      const parseAmount = (val) => {
          if (typeof val === 'number') return val;
          if (typeof val === 'string') return parseFloat(val.replace(/[^0-9.-]/g, '')) || 0;
          return 0;
      };

      sortedCats.forEach(catKey => {
        const cat = cats[catKey];
        const catRowData = buildRowData("C", cat.name, cat.budget, cat.actual, type);
        outputRows.push(catRowData);
        metaRows.push("CAT");

        for (let m = 0; m < 12; m++) {
          groupTotals.budget[m] += parseAmount(cat.budget[m]);
          groupTotals.actual[m] += parseAmount(cat.actual[m]);
        }
      });

      // Update Group Header Placeholder
      outputRows[groupHeaderIndex] = buildRowData("G", group, groupTotals.budget, groupTotals.actual, type);

      // --- ADD SPACER AFTER GROUP ---
      // Calculate width based on the previous row (Cat row)
      const rowWidth = outputRows[outputRows.length - 1].length;
      outputRows.push(Array(rowWidth).fill("")); // Add empty row
      metaRows.push("SPACER"); // Mark as spacer for formatting

      // Accumulate Type Totals
      for (let m = 0; m < 12; m++) {
        typeTotals.budget[m] += groupTotals.budget[m];
        typeTotals.actual[m] += groupTotals.actual[m];
      }
    });

    // Update Type Header Placeholder
    outputRows[typeHeaderIndex] = buildRowData("AL", type, typeTotals.budget, typeTotals.actual, type);

    grandTotal.budget += typeTotals.budget.reduce((a,b)=>a+b, 0);
    grandTotal.actual += typeTotals.actual.reduce((a,b)=>a+b, 0);
  });

  // --- 4. WRITE TO SHEET & APPLY FONTS ---
  
  if (outputRows.length > 0) {
    const numRows = outputRows.length;
    const numCols = outputRows[0].length;
    
    const range = outSheet.getRange(10, 1, numRows, numCols);
    range.setValues(outputRows);
    
    // Global Font
    range.setFontFamily("Comfortaa").setFontSize(10);
    
    // --- 5. HIGHLIGHTING, BORDERS & SPECIFIC FORMATTING ---
    const typeRanges = [];
    const groupRanges = [];
    const numFormatRanges = []; 
    const pctFormatRanges = []; 
    const borderRanges = []; // New array for borders

    const fmtCurrency = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"_);_(@_)';
    const fmtPercent = '0.00%';
    const colorType = '#e68e68'; 
    const colorGroup = '#fce5cd'; 
    const borderColor = '#000000';

    for (let i = 0; i < numRows; i++) {
      const rowNum = 10 + i;
      const rowType = metaRows[i];
      
      // Skip formatting for Spacer rows
      if (rowType === "SPACER") continue;

      const rowRange = `B${rowNum}:${columnToLetter(numCols)}${rowNum}`;
      
      // Collect range for Border (Type, Group, Cat)
      // Borders applied from Col B to End
      borderRanges.push(rowRange);

      if (rowType === "TYPE") {
        typeRanges.push(rowRange);
      } else if (rowType === "GROUP") {
        groupRanges.push(rowRange);
      }

      // Number Formatting (Annual)
      numFormatRanges.push(`C${rowNum}:E${rowNum}`);
      pctFormatRanges.push(`F${rowNum}`);

      // Monthly Formatting (Starts at H=8)
      for (let m=0; m<12; m++) {
        let startCol = 8 + (m*4); 
        let endCol = startCol + 2;
        let pctCol = startCol + 3;
        
        let startLet = columnToLetter(startCol);
        let endLet = columnToLetter(endCol);
        let pctLet = columnToLetter(pctCol);
        
        numFormatRanges.push(`${startLet}${rowNum}:${endLet}${rowNum}`);
        pctFormatRanges.push(`${pctLet}${rowNum}`);
      }
    }

    // Apply Formatting Batches
    if (borderRanges.length) {
        const rl = outSheet.getRangeList(borderRanges);
        // Apply solid black border to data rows
        rl.setBorder(true, true, true, true, true, true, borderColor, SpreadsheetApp.BorderStyle.SOLID);
    }

    if (typeRanges.length) {
      const rl = outSheet.getRangeList(typeRanges);
      rl.setBackground(colorType);
      rl.setFontWeight('bold');
    }
    if (groupRanges.length) {
      const rl = outSheet.getRangeList(groupRanges);
      rl.setBackground(colorGroup);
      rl.setFontWeight('bold');
    }
    if (numFormatRanges.length) outSheet.getRangeList(numFormatRanges).setNumberFormat(fmtCurrency);
    if (pctFormatRanges.length) outSheet.getRangeList(pctFormatRanges).setNumberFormat(fmtPercent);
    
    // Alignment
    outSheet.getRange(10, 3, numRows, numCols - 2).setHorizontalAlignment('right');
    outSheet.getRange(10, 2, numRows, 1).setHorizontalAlignment('left');
  }

  outSheet.getRange('D3').setValue(grandTotal.budget);
  outSheet.getRange('D4').setValue(grandTotal.actual);
  outSheet.getRange('B3').setValue("Last updated on " + new Date().toLocaleString());
}

/**
 * Helper to build a single output row array.
 */
function buildRowData(id, name, budgetArr, actualArr, type) {
  const row = [id, name];
  const isIncome = (type === "Income" || type === "Transfers");
  const mult = isIncome ? -1 : 1; 

  let annBudget = 0;
  let annActual = 0;

  for (let i=0; i<12; i++) {
    annBudget += budgetArr[i];
    annActual += actualArr[i];
  }

  // Annual Summary
  const annDiff = (annBudget - annActual) * mult;
  const annPct = (annBudget === 0) ? (annActual === 0 ? 0 : 1) : (annActual / annBudget);

  row.push(annBudget, annActual, annDiff, annPct);

  // Spacer Column G
  row.push(""); 

  // Monthly Data
  for (let i=0; i<12; i++) {
    const b = budgetArr[i];
    const a = actualArr[i];
    
    const d = (b - a) * mult; 
    const p = (b === 0) ? (a === 0 ? 0 : 1) : (a / b);
    
    row.push(b, a, d, p);
  }
  return row;
}

// Helper to convert column index (1-based) to letter
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
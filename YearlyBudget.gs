/**
 * Optimized Yearly Budget Generator
 * Modifications:
 * - FIXED: The placement of Monthly Budget Cash Flow (Row 3) and Monthly Actual Cash Flow (Row 4) totals. 
 * They now correctly appear in the 'Difference' column for each month (J, N, R, V, Z, etc.).
 * - ADDED: A solid right border to the '%' column (K, O, S, etc.) of each month block to visually separate the months.
 * - ENSURED: Grand Annual Totals (C3, D3) and Annual Cash Flow (E3, E4) are correctly positioned.
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

  // Define currency format early
  const fmtCurrency = '_($* #,##0.00_);_($* (#,##0.00)_);_($* "-"_);_(@_)'; 

  // Clear Output Sheet
  const lastRow = outSheet.getLastRow();
  const lastCol = outSheet.getLastColumn();
  if (lastRow >= 10) {
    // Break merged cells and clear all content and formatting from row 10 onwards
    outSheet.getRange(10, 1, lastRow - 9, lastCol).breakApart();
    outSheet.getRange(10, 1, lastRow - 9, lastCol).clear({contentsOnly: true, formatOnly: true});
    outSheet.getRange(10, 1, lastRow - 9, lastCol).setBackground(null).setFontWeight(null).setFontStyle(null).setBorder(false, false, false, false, false, false);
  }

  // Ensure Column A (ID column) is hidden from view
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
      // If the category type is NOT Income or Transfers, treat the amount as an Expense magnitude (positive number).
      if (type !== 'Income' && type !== 'Transfers') {
        amt = Math.abs(amt); 
      }
      
      catMap[key].actual[monthIdx] += amt;
    }
  }

  // --- 3. OUTPUT GENERATION & TOTALS CALCULATION ---

  const outputRows = [];
  const metaRows = []; 
  const grandTotal = { budget: 0, actual: 0 };
  
  // Storage for monthly cash flow calculation
  const monthlyCashFlowTotals = {
    incomeBudget: Array(12).fill(0), 
    expenseBudget: Array(12).fill(0), 
    incomeActual: Array(12).fill(0), 
    expenseActual: Array(12).fill(0)
  };

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
      const rowWidth = outputRows[outputRows.length - 1].length;
      outputRows.push(Array(rowWidth).fill("")); // Add empty row
      metaRows.push("SPACER"); // Mark as spacer for formatting

      // Accumulate Type Totals
      for (let m = 0; m < 12; m++) {
        typeTotals.budget[m] += groupTotals.budget[m];
        typeTotals.actual[m] += groupTotals.actual[m];
      }
    }); // End of sortedGroups.forEach

    // Update Type Header Placeholder
    outputRows[typeHeaderIndex] = buildRowData("AL", type, typeTotals.budget, typeTotals.actual, type);
    
    // Calculate annual totals for this type
    const annualBudgetSum = typeTotals.budget.reduce((a,b)=>a+b, 0);
    const annualActualSum = typeTotals.actual.reduce((a,b)=>a+b, 0);

    // Accumulate Grand Total (All Types)
    grandTotal.budget += annualBudgetSum;
    grandTotal.actual += annualActualSum;
    
    // Accumulate Monthly Cash Flow Totals (Income and Expense only)
    if (type === "Income") {
        for (let m = 0; m < 12; m++) {
          monthlyCashFlowTotals.incomeBudget[m] += typeTotals.budget[m];
          monthlyCashFlowTotals.incomeActual[m] += typeTotals.actual[m];
        }
    } else if (type === "Expense") {
        for (let m = 0; m < 12; m++) {
          monthlyCashFlowTotals.expenseBudget[m] += typeTotals.budget[m];
          monthlyCashFlowTotals.expenseActual[m] += typeTotals.actual[m];
        }
    }
  }); // End of sortedTypes.forEach

  // --- 4. WRITE DATA TO SHEET & APPLY FONTS ---
  
  if (outputRows.length > 0) {
    const numRows = outputRows.length;
    const numCols = outputRows[0].length;
    
    const range = outSheet.getRange(10, 1, numRows, numCols);
    range.setValues(outputRows);
    
    // Global Font
    range.setFontFamily("Comfortaa").setFontSize(10);
    
    // --- 5. HIGHLIGHTING, FORMATTING, ALIGNMENT, & BORDERS ---
    const typeRanges = [];
    const groupRanges = [];
    const numFormatRanges = []; 
    const pctFormatRanges = []; 
    const borderRightRanges = []; // New array for right borders
    
    const fmtPercent = '0.00%';
    const colorType = '#E68E68'; 
    const colorGroup = '#EEC49F'; 

    for (let i = 0; i < numRows; i++) {
      const rowNum = 10 + i;
      const rowType = metaRows[i];
      
      // Skip formatting for Spacer rows
      if (rowType === "SPACER") continue;

      // Range for highlighting (Col B to the end)
      const rowRange = `B${rowNum}:${columnToLetter(numCols)}${rowNum}`;
      
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
        let pctCol = startCol + 3; // The % column (K, O, S, etc.)
        
        let startLet = columnToLetter(startCol);
        let endLet = columnToLetter(endCol);
        let pctLet = columnToLetter(pctCol);
        
        numFormatRanges.push(`${startLet}${rowNum}:${endLet}${rowNum}`);
        pctFormatRanges.push(`${pctLet}${rowNum}`);
        
        // Add % column to the list to receive a right border
        if (m < 11) { // Apply border to 11 months, skipping the last one
          borderRightRanges.push(`${pctLet}${rowNum}`);
        }
      }
    }

    // Apply Background Colors
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
    
    // Apply number formats
    if (numFormatRanges.length) outSheet.getRangeList(numFormatRanges).setNumberFormat(fmtCurrency);
    if (pctFormatRanges.length) outSheet.getRangeList(pctFormatRanges).setNumberFormat(fmtPercent);
    
    // Apply right borders to separate months (Row 10 onwards)
    if (borderRightRanges.length) {
      const rl = outSheet.getRangeList(borderRightRanges);
      rl.setBorder(false, false, false, true, false, false, '#355348', SpreadsheetApp.BorderStyle.SOLID);
    }

    // Alignment
    outSheet.getRange(10, 3, numRows, numCols - 2).setHorizontalAlignment('right');
    outSheet.getRange(10, 2, numRows, 1).setHorizontalAlignment('left');
  }
  
  // --- 6. OUTPUT SUMMARY TOTALS & MONTHLY CASH FLOW (FIXED PLACEMENT) ---
  
  // Derive Annual Cash Flow from monthly totals
  const annualIncomeBudget = monthlyCashFlowTotals.incomeBudget.reduce((a, b) => a + b, 0);
  const annualExpenseBudget = monthlyCashFlowTotals.expenseBudget.reduce((a, b) => a + b, 0);
  const annualIncomeActual = monthlyCashFlowTotals.incomeActual.reduce((a, b) => a + b, 0);
  const annualExpenseActual = monthlyCashFlowTotals.expenseActual.reduce((a, b) => a + b, 0);

  const budgetCashFlow = annualIncomeBudget - annualExpenseBudget;
  const actualCashFlow = annualIncomeActual - annualExpenseActual;
  
  // 1. Output Grand Annual Totals 
  // C3: Grand Annual Budget Total (assuming this is used for all Budget lines)
  // C4: Grand Annual Actual Total (assuming this is used for all Actual lines)
  // These cells were previously commented out, keeping them out as per the original file's state.

  // 2. Output Annual Cash Flow values
  // E3: Budget Cash Flow
  outSheet.getRange('E3').setValue(budgetCashFlow).setNumberFormat(fmtCurrency);
  // E4: Actual Cash Flow
  outSheet.getRange('E4').setValue(actualCashFlow).setNumberFormat(fmtCurrency);

  // 3. Output Monthly Cash Flow values (FIXED PLACEMENT to J3/J4, N3/N4, etc.)
  for (let m = 0; m < 12; m++) {
    // Calculate monthly CF
    const monthlyBudgetCF = monthlyCashFlowTotals.incomeBudget[m] - monthlyCashFlowTotals.expenseBudget[m];
    const monthlyActualCF = monthlyCashFlowTotals.incomeActual[m] - monthlyCashFlowTotals.expenseActual[m];
      
    // Target column is the "Difference" column for each month block: 
    // J (10), N (14), R (18), V (22), Z (26), AD (30), AH (34), etc.
    const targetCol = 10 + (m * 4); 
    
    // Budget Cash Flow goes into Row 3 (J3, N3, etc.)
    outSheet.getRange(3, targetCol).setValue(monthlyBudgetCF).setNumberFormat(fmtCurrency);
    
    // Actual Cash Flow goes into Row 4 (J4, N4, etc.)
    outSheet.getRange(4, targetCol).setValue(monthlyActualCF).setNumberFormat(fmtCurrency);
  }

  // Output Last Updated (B3)
  outSheet.getRange('B3').setValue("Last updated on " + getDateTime());
}

/**
 * Helper to build a single output row array.
 */
function buildRowData(id, name, budgetArr, actualArr, type) {
  const row = [id, name];
  // Income and Transfers are displayed as negative in the diff column for a true cash-flow view
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
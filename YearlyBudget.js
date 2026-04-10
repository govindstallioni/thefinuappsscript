/**
 * Optimized Yearly Budget Generator
 * Transaction processing rewritten to match the proven MonthlyBudget pattern:
 * - Reads column mappings directly from Definition I5:I12
 * - Builds a visibleCategories set from Categories sheet
 * - Creates a simple actualMap keyed by category name
 * - Uses getDataRange() for all data reads
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
  const monthColsRaw = defSheet.getRange('W2:W13').getValues().flat();

  // Read year from E2 dropdown
  const e2Cell = outSheet.getRange('E2');
  const e2Display = String(e2Cell.getDisplayValue()).replace(/[^0-9]/g, '');
  const e2Val = e2Cell.getValue();
  let targetYear = parseInt(e2Display, 10);
  if (!(targetYear >= 1900 && targetYear <= 2100)) {
    targetYear = (e2Val instanceof Date) ? e2Val.getFullYear() : parseInt(String(e2Val), 10);
  }
  if (!(targetYear >= 1900 && targetYear <= 2100)) {
    targetYear = new Date().getFullYear();
  }

  // Use spreadsheet timezone for date extraction to avoid timezone mismatch
  const tz = ss.getSpreadsheetTimeZone();

  const CONFIG = {
    CAT_COL: configRaw[0] - 1,
    GRP_COL: configRaw[1] - 1,
    TYP_COL: configRaw[2] - 1,
    HIDE_COL: configRaw[3] - 1,
    BUDGET_COLS: monthColsRaw.map(c => c - 1),
  };

  outSheet.getRange('B3').setValue("⏳ Processing..");
  const fmtCurrency = '_($* #,##0.00_);_($* (#,##0.00)_);_($* 0.00_);_(@_)';

  // Clear Output Sheet (Rows 10+)
  const lastRow = outSheet.getLastRow();
  const lastCol = outSheet.getLastColumn();
  if (lastRow >= 10) {
    outSheet.getRange(10, 1, lastRow - 9, lastCol).breakApart();
    outSheet.getRange(10, 1, lastRow - 9, lastCol).clear({contentsOnly: true, formatOnly: true});
  }
  outSheet.getRange("A:A").setFontColor("#FFFFFF");

  // --- 2. DATA PROCESSING ---

  const tree = {};
  const visibleCategories = new Set();

  // A. Process Categories — build tree AND visibleCategories set (like MonthlyBudget)
  const catData = catSheet.getDataRange().getValues();
  for (let i = 1; i < catData.length; i++) {
    const row = catData[i];
    const type = String(row[CONFIG.TYP_COL] || "").trim();
    const group = String(row[CONFIG.GRP_COL] || "").trim();
    const catName = String(row[CONFIG.CAT_COL] || "").trim();
    if (row[CONFIG.HIDE_COL] === "Hide" || !type || !group || !catName) continue;

    visibleCategories.add(catName);

    if (!tree[type]) tree[type] = {};
    if (!tree[type][group]) tree[type][group] = {};

    const catObj = {
      name: catName, group: group, type: type,
      budget: Array(12).fill(0),
      actual: Array(12).fill(0),
    };

    for (let m = 0; m < 12; m++) {
      catObj.budget[m] = Number(row[CONFIG.BUDGET_COLS[m]]) || 0;
    }
    tree[type][group][catName] = catObj;
  }

  // B. Process Transactions — EXACT same approach as working MonthlyBudget
  //    Read column mappings DIRECTLY from Definition I5:I12
  const tranConfigRaw = defSheet.getRange('I5:I12').getValues().flat();
  const TRAN_COL = {
    DATE: tranConfigRaw[4] - 1,      // I9: Date column (0-based)
    CATEGORY: tranConfigRaw[0] - 1,  // I5: Category column (0-based)
    AMOUNT: tranConfigRaw[5] - 1,    // I10: Amount column (0-based)
  };

  // Read ALL transaction data (same as MonthlyBudget)
  const tranData = tranSheet.getDataRange().getValues();

  // Build a simple actualMap keyed by category name with per-month amounts
  const actualMap = {};
  let _totalRows = 0, _yearMatch = 0, _catMatch = 0;

  for (let i = 1; i < tranData.length; i++) {
    const row = tranData[i];
    const dateVal = row[TRAN_COL.DATE];
    if (!dateVal) continue;

    const d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(d.getTime())) continue;
    _totalRows++;

    // Extract year and month using spreadsheet timezone (avoids script timezone mismatch)
    const txnYear = parseInt(Utilities.formatDate(d, tz, 'yyyy'), 10);
    const txnMonth = parseInt(Utilities.formatDate(d, tz, 'M'), 10) - 1; // 0-based
    if (txnYear !== targetYear) continue;
    _yearMatch++;

    const transactionCategory = String(row[TRAN_COL.CATEGORY] || '').trim();
    if (!visibleCategories.has(transactionCategory)) continue;
    _catMatch++;

    // Robust amount parsing (handles string or number values)
    let rawAmt = row[TRAN_COL.AMOUNT];
    let amt = (typeof rawAmt === 'string') ?
      parseFloat(rawAmt.replace(/[$,]/g, '')) || 0 :
      Number(rawAmt) || 0;

    if (!actualMap[transactionCategory]) {
      actualMap[transactionCategory] = Array(12).fill(0);
    }
    actualMap[transactionCategory][txnMonth] += amt;
  }

  // C. Apply actualMap to the tree objects
  Object.keys(tree).forEach(type => {
    Object.keys(tree[type]).forEach(group => {
      Object.keys(tree[type][group]).forEach(catName => {
        const catObj = tree[type][group][catName];
        if (actualMap[catName]) {
          for (let m = 0; m < 12; m++) {
            let amt = actualMap[catName][m];
            // Expenses: use absolute value (Plaid amounts are positive for debits)
            if (catObj.type !== 'Income' && catObj.type !== 'Transfers') {
              amt = Math.abs(amt);
            }
            catObj.actual[m] = amt;
          }
        }
      });
    });
  });

  // --- 3. OUTPUT GENERATION ---

  const outputRows = [];
  const metaRows = [];
  const monthlyCashFlowTotals = {
    incomeBudget: Array(12).fill(0), expenseBudget: Array(12).fill(0),
    incomeActual: Array(12).fill(0), expenseActual: Array(12).fill(0)
  };

  const sortedTypes = Object.keys(tree).sort((a, b) =>
    a === "Income" ? -1 : b === "Income" ? 1 : a.localeCompare(b)
  );

  sortedTypes.forEach(type => {
    const typeTotals = { budget: Array(12).fill(0), actual: Array(12).fill(0) };
    const typeHeaderIndex = outputRows.length;
    outputRows.push(null); metaRows.push("TYPE");

    Object.keys(tree[type]).sort().forEach(group => {
      const groupTotals = { budget: Array(12).fill(0), actual: Array(12).fill(0) };
      const groupHeaderIndex = outputRows.length;
      outputRows.push(null); metaRows.push("GROUP");

      Object.keys(tree[type][group]).sort().forEach(catName => {
        const cat = tree[type][group][catName];
        outputRows.push(buildRowData("C", cat.name, cat.budget, cat.actual, type));
        metaRows.push("CAT");
        for (let m = 0; m < 12; m++) {
          groupTotals.budget[m] += cat.budget[m];
          groupTotals.actual[m] += cat.actual[m];
        }
      });

      outputRows[groupHeaderIndex] = buildRowData("G", group, groupTotals.budget, groupTotals.actual, type);
      outputRows.push(Array(outputRows[outputRows.length - 1].length).fill(""));
      metaRows.push("SPACER");

      for (let m = 0; m < 12; m++) {
        typeTotals.budget[m] += groupTotals.budget[m];
        typeTotals.actual[m] += groupTotals.actual[m];
      }
    });

    outputRows[typeHeaderIndex] = buildRowData("AL", type, typeTotals.budget, typeTotals.actual, type);

    if (type === "Income") {
      for (let m=0; m<12; m++) { monthlyCashFlowTotals.incomeBudget[m] += typeTotals.budget[m]; monthlyCashFlowTotals.incomeActual[m] += typeTotals.actual[m]; }
    } else if (type === "Expense") {
      for (let m=0; m<12; m++) { monthlyCashFlowTotals.expenseBudget[m] += typeTotals.budget[m]; monthlyCashFlowTotals.expenseActual[m] += typeTotals.actual[m]; }
    }
  });

  // --- 4. WRITE & FORMAT ---
  if (outputRows.length > 0) {
    const range = outSheet.getRange(10, 1, outputRows.length, outputRows[0].length);
    range.setValues(outputRows).setFontFamily("Comfortaa").setFontSize(10);

    const fmtPercent = '0.00%';
    const numFormatRanges = [], pctFormatRanges = [], typeRanges = [], groupRanges = [], borderRanges = [];

    for (let i = 0; i < outputRows.length; i++) {
      const rNum = 10 + i;
      if (metaRows[i] === "SPACER") continue;
      if (metaRows[i] === "TYPE") typeRanges.push(`B${rNum}:${columnToLetter(outputRows[0].length)}${rNum}`);
      if (metaRows[i] === "GROUP") groupRanges.push(`B${rNum}:${columnToLetter(outputRows[0].length)}${rNum}`);

      numFormatRanges.push(`C${rNum}:E${rNum}`);
      pctFormatRanges.push(`F${rNum}`);

      for (let m=0; m<12; m++) {
        let start = 8 + (m*4);
        numFormatRanges.push(`${columnToLetter(start)}${rNum}:${columnToLetter(start+2)}${rNum}`);
        pctFormatRanges.push(`${columnToLetter(start+3)}${rNum}`);
        if (m < 11) borderRanges.push(`${columnToLetter(start+3)}${rNum}`);
      }
    }

    if (typeRanges.length) outSheet.getRangeList(typeRanges).setBackground('#E68E68').setFontWeight('bold');
    if (groupRanges.length) outSheet.getRangeList(groupRanges).setBackground('#EEC49F').setFontWeight('bold');
    outSheet.getRangeList(numFormatRanges).setNumberFormat(fmtCurrency);
    outSheet.getRangeList(pctFormatRanges).setNumberFormat(fmtPercent);
    outSheet.getRangeList(borderRanges).setBorder(false, false, false, true, false, false, '#355348', SpreadsheetApp.BorderStyle.SOLID);
    outSheet.getRange(10, 3, outputRows.length, outputRows[0].length-2).setHorizontalAlignment('right');
  }

  // --- 5. CASH FLOW SUMMARY ---
  outSheet.getRange('E3').setValue(monthlyCashFlowTotals.incomeBudget.reduce((a,b)=>a+b,0) - monthlyCashFlowTotals.expenseBudget.reduce((a,b)=>a+b,0)).setNumberFormat(fmtCurrency);
  outSheet.getRange('E4').setValue(monthlyCashFlowTotals.incomeActual.reduce((a,b)=>a+b,0) - monthlyCashFlowTotals.expenseActual.reduce((a,b)=>a+b,0)).setNumberFormat(fmtCurrency);

  for (let m = 0; m < 12; m++) {
    const col = 10 + (m * 4);
    outSheet.getRange(3, col).setValue(monthlyCashFlowTotals.incomeBudget[m] - monthlyCashFlowTotals.expenseBudget[m]).setNumberFormat(fmtCurrency);
    outSheet.getRange(4, col).setValue(monthlyCashFlowTotals.incomeActual[m] - monthlyCashFlowTotals.expenseActual[m]).setNumberFormat(fmtCurrency);
  }

  outSheet.getRange('B3').setValue("Last updated on " + getDateTime() );
}

function buildRowData(id, name, budgetArr, actualArr, type) {
  const isIncome = (type === "Income" || type === "Transfers");
  const mult = isIncome ? -1 : 1;
  let annB = 0, annA = 0;
  for (let i=0; i<12; i++) { annB += budgetArr[i]; annA += actualArr[i]; }

  const row = [id, name, annB, annA || 0.0, (annB - annA) * mult, (annB === 0 ? (annA === 0 ? 0 : 1) : annA / annB), ""];
  for (let i=0; i<12; i++) {
    const b = budgetArr[i], a = actualArr[i] || 0.0;
    row.push(b, a, (b - a) * mult, (b === 0 ? (a === 0 ? 0 : 1) : a / b));
  }
  return row;
}

function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Optimized Yearly Budget Generator
 * Reviewed and Fixed:
 * - FIXED: Actual amounts now populate by using a robust numeric parser (removes $ and ,).
 * - FIXED: Data fetching logic improved with flexible key matching (handles missing Group/Type in Transactions).
 * - FIXED: Zero values explicitly forced to 0.00 to match the requested visual style.
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
    YEAR: defSheet.getRange('V1').getValue(),
    BUDGET_COLS: monthColsRaw.map(c => c - 1),
    
    TRAN_CAT: tranConfigRaw[0] - 1,
    TRAN_GRP: tranConfigRaw[1] - 1,
    TRAN_TYP: tranConfigRaw[2] - 1, 
    TRAN_DATE: tranConfigRaw[4] - 1,
    TRAN_AMT: tranConfigRaw[5] - 1, 
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
  const catMap = {}; 
  const fallbackCatMap = {}; // Maps Category Name -> Data Object (if Group/Type missing in Trans)

  // A. Process Categories
  const catData = catSheet.getDataRange().getValues();
  for (let i = 1; i < catData.length; i++) { 
    const row = catData[i];
    const type = String(row[CONFIG.TYP_COL] || "").trim();
    const group = String(row[CONFIG.GRP_COL] || "").trim();
    const catName = String(row[CONFIG.CAT_COL] || "").trim();
    if (row[CONFIG.HIDE_COL] === "Hide" || !type || !group || !catName) continue;

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

    const fullKey = `${catName}|${group}|${type}`.toUpperCase();
    catMap[fullKey] = catObj;
    fallbackCatMap[catName.toUpperCase()] = catObj;
    tree[type][group][catName] = catObj;
  }

  // B. Process Transactions
  const tranData = tranSheet.getDataRange().getValues();
  const targetYear = CONFIG.YEAR;

  for (let i = 1; i < tranData.length; i++) {
    const row = tranData[i];
    const dateVal = row[CONFIG.TRAN_DATE];
    if (!dateVal || !(dateVal instanceof Date) && isNaN(Date.parse(dateVal))) continue;
    
    const date = new Date(dateVal);
    if (date.getFullYear() != targetYear) continue;

    const tCat = String(row[CONFIG.TRAN_CAT] || "").trim().toUpperCase();
    const tGrp = String(row[CONFIG.TRAN_GRP] || "").trim().toUpperCase();
    const tTyp = String(row[CONFIG.TRAN_TYP] || "").trim().toUpperCase();
    
    // Robust Amount Parsing (Removes $ and ,)
    let rawAmt = row[CONFIG.TRAN_AMT];
    let amt = (typeof rawAmt === 'string') ? 
              parseFloat(rawAmt.replace(/[$,]/g, '')) || 0 : 
              Number(rawAmt) || 0;

    const fullKey = `${tCat}|${tGrp}|${tTyp}`;
    const targetObj = catMap[fullKey] || fallbackCatMap[tCat];

    if (targetObj) {
      if (targetObj.type !== 'Income' && targetObj.type !== 'Transfers') {
        amt = Math.abs(amt); 
      }
      targetObj.actual[date.getMonth()] += amt;
    }
  }

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

  outSheet.getRange('B3').setValue("Last updated on " + getDateTime());
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
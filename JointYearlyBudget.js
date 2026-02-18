/**
 * EXPERT OPTIMIZED Joint Yearly Budget Generator
 * * FIXES APPLIED:
 * 1. Populates "Actual Cash Flow" for both Master and Monthly columns.
 * 2. Updated Currency Format to show $0.00 instead of dashes for zero values.
 * 3. Maintains all existing robust key matching and year filtering.
 */
function populateJointYearlyBudget() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outSheet = ss.getSheetByName("Joint Yearly Budget"); 
  const defSheet = ss.getSheetByName("Definition");
  const catSheet = ss.getSheetByName("Categories");
  const tranSheet = ss.getSheetByName("Transactions");

  if (!outSheet || !defSheet || !catSheet || !tranSheet) {
    SpreadsheetApp.getUi().alert("Critical Error: One or more required sheets are missing.");
    return;
  }
  
  const borderColor = '#000000'; 
  const colorType = '#e68e68';
  const colorGroup = '#fce5cd';

  // --- 1. CONFIGURATION LOAD ---
  const configRaw = defSheet.getRange('C5:C12').getValues().flat();
  const tranConfigRaw = defSheet.getRange('I5:I12').getValues().flat();
  const monthColsRaw = defSheet.getRange('W2:W13').getValues().flat();
  const namesRaw = defSheet.getRange('C11:C12').getValues().flat();
  
  const year = outSheet.getRange('D2').getValue();

  outSheet.getRange('B3').setValue("⏳ Processing Actuals...");
  
  const NAME1 = String(namesRaw[0] || "").trim();
  const NAME2 = String(namesRaw[1] || "").trim();

  const CONFIG = {
    CAT_COL: configRaw[0] - 1,
    GRP_COL: configRaw[1] - 1,
    TYP_COL: configRaw[2] - 1,
    HIDE_COL: configRaw[3] - 1,
    ALLOC1: configRaw[4] - 1,
    ALLOC2: configRaw[5] - 1, 
    YEAR: year,
    NAME1: NAME1,
    NAME2: NAME2,
    BUDGET_COLS: monthColsRaw.map(col => col - 1),
    TRAN_CAT_COL: Number(tranConfigRaw[0]) - 1,
    TRAN_GRP_COL: Number(tranConfigRaw[1]) - 1,
    TRAN_TYP_COL: Number(tranConfigRaw[2]) - 1, 
    TRAN_DATE_COL: Number(tranConfigRaw[4]) - 1, 
    TRAN_AMT_COL: Number(tranConfigRaw[5]) - 1, 
    TRAN_OWNER_COL: Number(tranConfigRaw[6]) - 1,
    TRAN_ASSIGN_AMT_COL: Number(tranConfigRaw[7]) - 1
  };

  // --- 2. SHEET CLEANUP ---
  const lastRow = outSheet.getLastRow();
  const lastCol = outSheet.getLastColumn();
  const startRow = 9;

  if (lastRow >= startRow) {
    const rowsToClear = lastRow - startRow + 1;
    outSheet.getRange(startRow, 2, rowsToClear, 1).breakApart();
    const clearRange = outSheet.getRange(startRow, 1, rowsToClear, lastCol);
    clearRange.clear({contentsOnly: true, formatOnly: true});
  }
    
  // --- 3. DATA PROCESSING ---
  const tree = {};
  const catMap = {}; 
  const nameOnlyMap = {}; 

  const toNum = (val) => {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    const cleaned = String(val).replace(/[$,\s]/g, '');
    return parseFloat(cleaned) || 0;
  };

  const catData = catSheet.getDataRange().getValues();
  for (let i = 1; i < catData.length; i++) { 
    const row = catData[i];
    if (String(row[CONFIG.HIDE_COL]).trim() === "Hide" || !row[CONFIG.TYP_COL]) continue;

    const type = String(row[CONFIG.TYP_COL]).trim();
    const group = String(row[CONFIG.GRP_COL]).trim();
    const catName = String(row[CONFIG.CAT_COL]).trim();

    if (!tree[type]) tree[type] = {};
    if (!tree[type][group]) tree[type][group] = {};

    const alloc1 = Number(row[CONFIG.ALLOC1]) || 0;
    const alloc2 = Number(row[CONFIG.ALLOC2]) || 0;
    
    const catObj = {
      name: catName,
      budget1: Array(12).fill(0), budget2: Array(12).fill(0),
      actual1: Array(12).fill(0), actual2: Array(12).fill(0),
      alloc1: alloc1, alloc2: alloc2
    };

    for (let m = 0; m < 12; m++) {
      let val = toNum(row[CONFIG.BUDGET_COLS[m]]);
      catObj.budget1[m] = val * alloc1;
      catObj.budget2[m] = val * alloc2;
    }
    tree[type][group][catName] = catObj;

    const fullKey = (catName + "_" + group + "_" + type).toUpperCase();
    catMap[fullKey] = catObj;
    nameOnlyMap[catName.toUpperCase()] = catObj;
  }

  const tranData = tranSheet.getDataRange().getValues();
  for (let i = 1; i < tranData.length; i++) {
    const row = tranData[i];
    const tDate = row[CONFIG.TRAN_DATE_COL];
    if (!tDate) continue;
    
    const dateObj = new Date(tDate);
    if (isNaN(dateObj.getTime()) || dateObj.getFullYear() != CONFIG.YEAR) continue;

    const tCat = String(row[CONFIG.TRAN_CAT_COL] || "").trim();
    const tGrp = String(row[CONFIG.TRAN_GRP_COL] || "").trim();
    const tTyp = String(row[CONFIG.TRAN_TYP_COL] || "").trim();

    const fullLookupKey = (tCat + "_" + tGrp + "_" + tTyp).toUpperCase();
    const catEntry = catMap[fullLookupKey] || nameOnlyMap[tCat.toUpperCase()];
    
    if (catEntry) {
      let amt = toNum(row[CONFIG.TRAN_AMT_COL]);
      let assignAmt = toNum(row[CONFIG.TRAN_ASSIGN_AMT_COL]);
      if (tTyp !== 'Income' && tTyp !== 'Transfers') amt = Math.abs(amt);

      const ownerRaw = String(row[CONFIG.TRAN_OWNER_COL] || "").toLowerCase().trim();
      const monthIdx = dateObj.getMonth();
      const n1 = CONFIG.NAME1.toLowerCase();
      const n2 = CONFIG.NAME2.toLowerCase();
      
      let amt1 = 0, amt2 = 0;
      if (ownerRaw === n1) { amt1 = amt; } 
      else if (ownerRaw === n2) { amt2 = amt; } 
      else {
        if (assignAmt !== 0) { amt1 = assignAmt; amt2 = assignAmt; } 
        else { amt1 = amt * catEntry.alloc1; amt2 = amt * catEntry.alloc2; }
      }
      catEntry.actual1[monthIdx] += amt1;
      catEntry.actual2[monthIdx] += amt2;
    }
  }

  // --- 4. OUTPUT GENERATION ---
  const nameRows = [], mainDataRows = [], metaRows = [];
  const sortedTypes = Object.keys(tree).sort((a, b) => a === "Income" ? -1 : b === "Income" ? 1 : a.localeCompare(b));

  // Initialize for Cash Flow Calculation
  const monthlyAcf1 = Array(12).fill(0);
  const monthlyAcf2 = Array(12).fill(0);

  sortedTypes.forEach(type => {
    const isIncome = (type === "Income" || type === "Transfers");
    const mult = isIncome ? 1 : -1;
    const typeTotals = { b1: Array(12).fill(0), a1: Array(12).fill(0), b2: Array(12).fill(0), a2: Array(12).fill(0) };
    const typeIdx = mainDataRows.length;
    pushBlock(type, typeTotals, type, "TYPE", nameRows, mainDataRows, metaRows);

    Object.keys(tree[type]).sort().forEach(group => {
      const groupTotals = { b1: Array(12).fill(0), a1: Array(12).fill(0), b2: Array(12).fill(0), a2: Array(12).fill(0) };
      const groupIdx = mainDataRows.length;
      pushBlock(group, groupTotals, type, "GROUP", nameRows, mainDataRows, metaRows);

      Object.keys(tree[type][group]).sort().forEach(cName => {
        const c = tree[type][group][cName];
        pushBlock(c.name, {b1:c.budget1, a1:c.actual1, b2:c.budget2, a2:c.actual2}, type, "CAT", nameRows, mainDataRows, metaRows);
        for (let m=0; m<12; m++) {
          groupTotals.b1[m]+=c.budget1[m]; groupTotals.a1[m]+=c.actual1[m];
          groupTotals.b2[m]+=c.budget2[m]; groupTotals.a2[m]+=c.actual2[m];
          monthlyAcf1[m] += (c.actual1[m] * mult);
          monthlyAcf2[m] += (c.actual2[m] * mult);
        }
      });
      pushSpacer(mainDataRows[0].length, nameRows, mainDataRows, metaRows);
      updateBlock(mainDataRows, groupIdx, buildRows(group, groupTotals.b1, groupTotals.a1, groupTotals.b2, groupTotals.a2, type));

      for (let m=0; m<12; m++) {
        typeTotals.b1[m]+=groupTotals.b1[m]; typeTotals.a1[m]+=groupTotals.a1[m];
        typeTotals.b2[m]+=groupTotals.b2[m]; typeTotals.a2[m]+=groupTotals.a2[m];
      }
    });
    updateBlock(mainDataRows, typeIdx, buildRows(type, typeTotals.b1, typeTotals.a1, typeTotals.b2, typeTotals.a2, type));
  });

  // Write Cash Flow Row (Row 4)
  const yearlyAcf1 = safeSum(monthlyAcf1), yearlyAcf2 = safeSum(monthlyAcf2);
  const row4Update = [yearlyAcf1, yearlyAcf2, yearlyAcf1 + yearlyAcf2, ""]; 
  for (let m = 0; m < 12; m++) {
    row4Update.push(monthlyAcf1[m], monthlyAcf2[m], monthlyAcf1[m] + monthlyAcf2[m]);
  }
  outSheet.getRange(4, 4, 1, row4Update.length).setValues([row4Update]);

  // --- 5. FORMATTING & WRITING ---
  if (mainDataRows.length > 0) {
    const width = mainDataRows[0].length;
    const finalCol = colToLet(width + 3);

    outSheet.getRange(startRow, 2, mainDataRows.length, 1).setValues(nameRows.map(x => [x])).setFontFamily("Comfortaa").setFontSize(10);
    outSheet.getRange(startRow, 3, mainDataRows.length, width).setValues(mainDataRows).setFontFamily("Comfortaa").setFontSize(10);

    const cRanges = [`D4:${finalCol}4`], pRanges = [], mRanges = [], rBRanges = [], bBRanges = [];

    for (let i = 0; i < mainDataRows.length; i++) {
      const r = startRow + i;
      if (metaRows[i] === "SPACER") continue;
      mRanges.push(`B${r}:B${r+3}`);
      if (metaRows[i] !== "CAT") outSheet.getRange(r, 2, 4, width + 1).setBackground(metaRows[i] === "TYPE" ? colorType : colorGroup).setFontWeight('bold');
      bBRanges.push(`B${r+3}:${finalCol}${r+3}`);

      for (let offset = 0; offset < 4; offset++) {
        const row = r + offset;
        const dataRange = `D${row}:${finalCol}${row}`;
        if (offset === 3) pRanges.push(dataRange); else cRanges.push(dataRange);
        rBRanges.push(`D${row}`, `E${row}`, `F${row}`);
        for (let m=0; m<12; m++) rBRanges.push(`${colToLet(8+m*3)}${row}`, `${colToLet(9+m*3)}${row}`, `${colToLet(10+m*3)}${row}`);
      }
      i += 3;
    }
    
    // UPDATED: Standard Currency Format that shows $0.00 for zeros
    const fmtCurrency = '_($* #,##0.00_);_($* (#,##0.00)_);_($* 0.00_);_(@_)';
    
    if (cRanges.length) outSheet.getRangeList(cRanges).setNumberFormat(fmtCurrency);
    if (pRanges.length) outSheet.getRangeList(pRanges).setNumberFormat('0.00%');
    if (mRanges.length) outSheet.getRangeList(mRanges).getRanges().forEach(rng => rng.merge().setVerticalAlignment('top').setWrap(true));
    if (rBRanges.length) outSheet.getRangeList(rBRanges).setBorder(false, false, false, true, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);
    if (bBRanges.length) outSheet.getRangeList(bBRanges).setBorder(null, null, true, null, null, null, borderColor, SpreadsheetApp.BorderStyle.SOLID);
    
    outSheet.getRange(startRow, 4, mainDataRows.length, width-1).setHorizontalAlignment('right'); 
    outSheet.getRange(startRow, 3, mainDataRows.length, 1).setHorizontalAlignment('left');
  }
  outSheet.getRange('B3').setValue("Last updated: " + new Date().toLocaleString());
}

function pushBlock(name, d, type, tag, nR, dR, mR) {
  const rows = buildRows(name, d.b1, d.a1, d.b2, d.a2, type);
  rows.forEach(r => { nR.push(name); dR.push(r); mR.push(tag); });
}
function pushSpacer(w, nR, dR, mR) { nR.push(""); dR.push(Array(w).fill("")); mR.push("SPACER"); }
function updateBlock(dR, start, rows) { rows.forEach((r, i) => dR[start+i] = r); }

function buildRows(name, b1, a1, b2, a2, type) {
  const mult = (type === "Income" || type === "Transfers") ? -1 : 1;
  const sB1 = safeSum(b1), sB2 = safeSum(b2), sBT = sB1+sB2;
  const sA1 = safeSum(a1), sA2 = safeSum(a2), sAT = sA1+sA2;
  const r1 = ["Budget", sB1, sB2, sBT, ""], r2 = ["Actual", sA1, sA2, sAT, ""];
  const r3 = ["Diff", (sB1-sA1)*mult, (sB2-sA2)*mult, (sBT-sAT)*mult, ""];
  const r4 = ["%", safeDiv(sA1,sB1), safeDiv(sA2,sB2), safeDiv(sAT,sBT), ""];
  for (let m=0; m<12; m++) {
    const mb1=b1[m], ma1=a1[m], mb2=b2[m], ma2=a2[m];
    r1.push(mb1, mb2, mb1+mb2); r2.push(ma1, ma2, ma1+ma2);
    r3.push((mb1-ma1)*mult, (mb2-ma2)*mult, (mb1+mb2-(ma1+ma2))*mult);
    r4.push(safeDiv(ma1,mb1), safeDiv(ma2,mb2), safeDiv(ma1+ma2,mb1+mb2));
  }
  return [r1, r2, r3, r4];
}
function colToLet(c) {
  let l = '';
  while (c > 0) { let t = (c - 1) % 26; l = String.fromCharCode(t + 65) + l; c = (c - t - 1) / 26; }
  return l;
}
function safeSum(arr) { return arr.reduce((a, b) => a + b, 0); }
function safeDiv(n, d) { return d === 0 ? (n === 0 ? 0 : 1) : n / d; }
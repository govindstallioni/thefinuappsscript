/**
 * EXPERT OPTIMIZED Joint Yearly Budget Generator (Strict Column Mapping)
 * * COLUMN MAPPING:
 * - A: Empty (Identifiers removed)
 * - B: Categories
 * - C: Labels (Budget, Actual, Diff, %)
 * --- ANNUAL BLOCK ---
 * - D: Name 1
 * - E: Name 2
 * - F: Total
 * - G: Spacer (Empty)
 * --- MONTHLY BLOCKS (Jan to Dec) ---
 * - H, I, J: Jan (Name 1, Name 2, Total)
 * - K, L, M: Feb (Name 1, Name 2, Total)
 * - ...etc
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

  // --- 1. FAST CONFIGURATION LOAD ---
  const configRaw = defSheet.getRange('C5:C12').getValues().flat();
  const tranConfigRaw = defSheet.getRange('I5:I12').getValues().flat();
  const monthColsRaw = defSheet.getRange('W2:W13').getValues().flat();
  const namesRaw = defSheet.getRange('C11:C12').getValues().flat();
  const year = defSheet.getRange('V1').getValue();
  
  const NAME1 = namesRaw[0];
  const NAME2 = namesRaw[1];

  const CONFIG = {
    CAT_COL: configRaw[0] - 1, GRP_COL: configRaw[1] - 1, TYP_COL: configRaw[2] - 1, HIDE_COL: configRaw[3] - 1,
    ALLOC1: configRaw[4] - 1, ALLOC2: configRaw[5] - 1, 
    YEAR: year,
    NAME1: NAME1, NAME2: NAME2,
    BUDGET_COLS: monthColsRaw.map(c => c - 1),
    TRAN_CAT_COL: defSheet.getRange('I5').getValue() - 1,
    TRAN_GRP_COL: defSheet.getRange('I6').getValue() - 1,
    TRAN_TYP_COL: tranConfigRaw[2] - 1, 
    TRAN_DATE_COL: tranConfigRaw[4] - 1, 
    TRAN_AMT_COL: tranConfigRaw[5] - 1, 
    TRAN_OWNER_COL: tranConfigRaw[6] - 1,
    OWNER_MAP: { [String(NAME1).toLowerCase()]: 1, [String(NAME2).toLowerCase()]: 2 }
  };
    
  // --- 2. SHEET CLEANUP ---
  const lastRow = outSheet.getLastRow();
  const lastCol = outSheet.getLastColumn();
  const startRow = 9;

  if (lastRow >= startRow) {
    const rowsToClear = lastRow - startRow + 1;
    outSheet.getRange(startRow, 2, rowsToClear, 1).breakApart(); // Fix merged cells issue
    const clearRange = outSheet.getRange(startRow, 1, rowsToClear, lastCol);
    clearRange.clear({contentsOnly: true, formatOnly: true});
    clearRange.setBackground(null).setFontWeight(null).setBorder(false, false, false, false, false, false);
  }
    
  // --- 3. DATA PROCESSING (IN-MEMORY) ---
  const tree = {}; 
  const catMap = {}; 

  // A. Process Categories
  const catData = catSheet.getDataRange().getValues();
  for (let i = 1; i < catData.length; i++) { 
    const row = catData[i];
    if (row[CONFIG.HIDE_COL] === "Hide" || !row[CONFIG.TYP_COL]) continue;

    const type = row[CONFIG.TYP_COL];
    const group = row[CONFIG.GRP_COL];
    const catName = row[CONFIG.CAT_COL];
    
    if (!tree[type]) tree[type] = {};
    if (!tree[type][group]) tree[type][group] = {};

    const alloc1 = Number(row[CONFIG.ALLOC1]) || 0;
    const alloc2 = Number(row[CONFIG.ALLOC2]) || 0;
    
    const catObj = {
      name: catName, group: group, type: type,
      budget1: Array(12).fill(0), budget2: Array(12).fill(0),
      actual1: Array(12).fill(0), actual2: Array(12).fill(0),
    };

    for (let m = 0; m < 12; m++) {
      let val = row[CONFIG.BUDGET_COLS[m]];
      if (typeof val === 'string') val = parseFloat(val.replace(/[$,]/g, ''));
      if (isNaN(val)) val = 0;
      catObj.budget1[m] = val * alloc1;
      catObj.budget2[m] = val * alloc2;
    }
    tree[type][group][catName] = catObj;
    catMap[`${catName}_${group}_${type}`] = catObj;
  }

  // B. Process Transactions
  const tranData = tranSheet.getDataRange().getValues();
  for (let i = 1; i < tranData.length; i++) {
    const row = tranData[i];
    const tDate = row[CONFIG.TRAN_DATE_COL];
    if (!tDate || new Date(tDate).getFullYear() != CONFIG.YEAR) continue;

    const realTCat = row[CONFIG.TRAN_CAT_COL];
    const realTGrp = row[CONFIG.TRAN_GRP_COL];
    const type = row[CONFIG.TRAN_TYP_COL];
    const catEntry = catMap[`${realTCat}_${realTGrp}_${type}`];
    
    if (catEntry) {
      let amt = Number(row[CONFIG.TRAN_AMT_COL]) || 0;
      if (type !== 'Income' && type !== 'Transfers') amt = Math.abs(amt); 
      
      const ownerRaw = String(row[CONFIG.TRAN_OWNER_COL]).toLowerCase();
      const monthIdx = new Date(tDate).getMonth(); 
      let amt1 = 0, amt2 = 0;
      
      if (CONFIG.OWNER_MAP[ownerRaw] === 1) amt1 = amt;
      else if (CONFIG.OWNER_MAP[ownerRaw] === 2) amt2 = amt;
      else { amt1 = amt / 2; amt2 = amt / 2; } 
      
      catEntry.actual1[monthIdx] += amt1;
      catEntry.actual2[monthIdx] += amt2;
    }
  }

  // --- 4. OUTPUT GENERATION ---
  const nameRows = [], mainDataRows = [], metaRows = [];
  const grandTotal = { b1:0, a1:0, b2:0, a2:0 };

  const sortedTypes = Object.keys(tree).sort((a, b) => (a === "Income" ? -1 : (b === "Income" ? 1 : a.localeCompare(b))));

  sortedTypes.forEach(type => {
    const groups = tree[type];
    const typeTotals = { b1: Array(12).fill(0), a1: Array(12).fill(0), b2: Array(12).fill(0), a2: Array(12).fill(0) };
    
    const typeStartIdx = mainDataRows.length;
    const typeBlock = buildRowBlock(type, typeTotals.b1, typeTotals.a1, typeTotals.b2, typeTotals.a2, type);
    pushBlockData(typeBlock, type, "TYPE", nameRows, mainDataRows, metaRows);

    Object.keys(groups).sort().forEach(group => {
      const cats = groups[group];
      const groupTotals = { b1: Array(12).fill(0), a1: Array(12).fill(0), b2: Array(12).fill(0), a2: Array(12).fill(0) };

      const groupStartIdx = mainDataRows.length;
      const groupBlock = buildRowBlock(group, groupTotals.b1, groupTotals.a1, groupTotals.b2, groupTotals.a2, type);
      pushBlockData(groupBlock, group, "GROUP", nameRows, mainDataRows, metaRows);

      Object.keys(cats).sort().forEach(catKey => {
        const cat = cats[catKey];
        const catBlock = buildRowBlock(cat.name, cat.budget1, cat.actual1, cat.budget2, cat.actual2, type);
        pushBlockData(catBlock, cat.name, "CAT", nameRows, mainDataRows, metaRows);

        for (let m=0; m<12; m++) {
          groupTotals.b1[m] += cat.budget1[m]; groupTotals.a1[m] += cat.actual1[m];
          groupTotals.b2[m] += cat.budget2[m]; groupTotals.a2[m] += cat.actual2[m];
        }
      });

      // ADD SPACER ROW AFTER EACH GROUP (Requested Feature)
      // Note: dataWidth is needed here. We can get it from the last pushed row.
      const currentWidth = mainDataRows[mainDataRows.length - 1].length;
      pushSpacerRow(currentWidth, nameRows, mainDataRows, metaRows);

      const finalGroupBlock = buildRowBlock(group, groupTotals.b1, groupTotals.a1, groupTotals.b2, groupTotals.a2, type);
      updateBlockData(mainDataRows, groupStartIdx, finalGroupBlock);

      for (let m=0; m<12; m++) {
        typeTotals.b1[m] += groupTotals.b1[m]; typeTotals.a1[m] += groupTotals.a1[m];
        typeTotals.b2[m] += groupTotals.b2[m]; typeTotals.a2[m] += groupTotals.a2[m];
      }
    });

    const finalTypeBlock = buildRowBlock(type, typeTotals.b1, typeTotals.a1, typeTotals.b2, typeTotals.a2, type);
    updateBlockData(mainDataRows, typeStartIdx, finalTypeBlock);
    
    grandTotal.b1 += safeSum(typeTotals.b1); grandTotal.a1 += safeSum(typeTotals.a1);
    grandTotal.b2 += safeSum(typeTotals.b2); grandTotal.a2 += safeSum(typeTotals.a2);
  });

  // --- 5. WRITE & FORMAT DATA ---
  if (mainDataRows.length > 0) {
    const numRows = mainDataRows.length;
    const dataWidth = mainDataRows[0].length; 

    // Write Data to Sheet
    outSheet.getRange(startRow, 2, numRows, 1).setValues(nameRows.map(x => [x])).setFontFamily("Comfortaa").setFontSize(10).setHorizontalAlignment('left');
    outSheet.getRange(startRow, 3, numRows, dataWidth).setValues(mainDataRows).setFontFamily("Comfortaa").setFontSize(10);
    
    // --- ROW-BASED FORMATTING ---
    const currencyRanges = [];
    const percentRanges = [];
    const nameMergeRanges = [];
    
    const colorType = '#e68e68';
    const colorGroup = '#fce5cd'; 
    const borderColor = '#000000';

    // Updated Loop to handle Variable Row Heights (4 rows for data, 1 row for spacer)
    let i = 0;
    while (i < numRows) {
      const r = startRow + i;
      const meta = metaRows[i];

      // Handle Spacer Rows
      if (meta === "SPACER") {
        i++; // Move to next row
        continue;
      }

      // Handle Data Blocks (Type, Group, Cat) - 4 Rows
      if (meta === "TYPE" || meta === "GROUP" || meta === "CAT") {
        
        // 1. Borders for the Block (Req: Add Border)
        // Range includes Name (Col B) + Labels (Col C) + Data (Col D onwards)
        // Length: 1 (B) + 1 (C) + dataWidth = dataWidth + 2 columns
        const blockRange = outSheet.getRange(r, 2, 4, dataWidth + 1); 
        blockRange.setBorder(true, true, true, true, true, true, borderColor, SpreadsheetApp.BorderStyle.SOLID);

        // 2. Merge Name Column
        nameMergeRanges.push(`B${r}:B${r+3}`);

        // 3. Headers Background
        if (meta === "TYPE" || meta === "GROUP") {
          blockRange.setBackground(meta === "TYPE" ? colorType : colorGroup).setFontWeight('bold');
        }

        // 4. Number Formats (Iterate 4 rows of the block)
        for (let rowOffset = 0; rowOffset < 4; rowOffset++) {
           const currR = r + rowOffset;
           const rowType = rowOffset; // 0:Budget, 1:Actual, 2:Diff, 3:%
           const dataRowRange = `D${currR}:${colToLet(dataWidth + 2)}${currR}`;
           
           if (rowType === 3) percentRanges.push(dataRowRange);
           else currencyRanges.push(dataRowRange);
        }

        i += 4; // Jump 4 rows for next iteration
      }
    }

    // Batch Apply Formats
    const fmtCurrency = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"_);_(@_)';
    const fmtPercent = '0.00%';
    
    if (currencyRanges.length) outSheet.getRangeList(currencyRanges).setNumberFormat(fmtCurrency);
    if (percentRanges.length) outSheet.getRangeList(percentRanges).setNumberFormat(fmtPercent);
    if (nameMergeRanges.length) {
      const ranges = outSheet.getRangeList(nameMergeRanges).getRanges();
      ranges.forEach(rng => rng.merge().setVerticalAlignment('top').setWrap(true));
    }
    
    // Alignments
    outSheet.getRange(startRow, 4, numRows, dataWidth - 1).setHorizontalAlignment('right'); 
    outSheet.getRange(startRow, 3, numRows, 1).setHorizontalAlignment('left'); 
  }

  // Update Top Summary
  outSheet.getRange('C3').setValue(grandTotal.b1 + grandTotal.b2); 
  outSheet.getRange('C4').setValue(grandTotal.a1 + grandTotal.a2); 
  outSheet.getRange('B3').setValue("Last updated on " + new Date().toLocaleString());
}

// --- HELPERS ---

function pushBlockData(blockData, name, meta, nameRows, mainRows, metaRows) {
  for (let i=0; i<4; i++) {
    nameRows.push(name);
    mainRows.push(blockData[i]); 
    metaRows.push(meta);
  }
}

function pushSpacerRow(width, nameRows, mainRows, metaRows) {
  nameRows.push(""); // Empty Name
  mainRows.push(Array(width).fill("")); // Empty Data
  metaRows.push("SPACER");
}

function updateBlockData(mainRows, startIdx, blockData) {
  for (let i=0; i<4; i++) {
    mainRows[startIdx + i] = blockData[i]; 
  }
}

function buildRowBlock(name, b1Arr, a1Arr, b2Arr, a2Arr, type) {
  const isIncome = (type === "Income" || type === "Transfers");
  const mult = isIncome ? -1 : 1;
  const sum = safeSum;

  // --- ANNUAL CALCS ---
  const annB1 = sum(b1Arr), annA1 = sum(a1Arr);
  const annD1 = (annB1 - annA1) * mult;
  const annP1 = safeDiv(annA1, annB1);

  const annB2 = sum(b2Arr), annA2 = sum(a2Arr);
  const annD2 = (annB2 - annA2) * mult;
  const annP2 = safeDiv(annA2, annB2);

  const annBT = annB1 + annB2, annAT = annA1 + annA2;
  const annDT = (annBT - annAT) * mult;
  const annPT = safeDiv(annAT, annBT);

  // Row Arrays initialized with Labels (Col C)
  const r1 = ["Budget"], r2 = ["Actual"], r3 = ["Diff"], r4 = ["%"];
  
  // --- ANNUAL BLOCK PUSH ---
  // Col D: Name 1 (Budget, Actual, Diff, %)
  r1.push(annB1); r2.push(annA1); r3.push(annD1); r4.push(annP1);
  // Col E: Name 2 (Budget, Actual, Diff, %)
  r1.push(annB2); r2.push(annA2); r3.push(annD2); r4.push(annP2);
  // Col F: Total (Budget, Actual, Diff, %)
  r1.push(annBT); r2.push(annAT); r3.push(annDT); r4.push(annPT);
  // Col G: Spacer (Empty)
  r1.push(""); r2.push(""); r3.push(""); r4.push("");

  // --- MONTHLY BLOCKS PUSH ---
  for (let m=0; m<12; m++) {
    // Monthly Name 1
    const mb1 = b1Arr[m], ma1 = a1Arr[m];
    const md1 = (mb1 - ma1) * mult;
    const mp1 = safeDiv(ma1, mb1);

    // Monthly Name 2
    const mb2 = b2Arr[m], ma2 = a2Arr[m];
    const md2 = (mb2 - ma2) * mult;
    const mp2 = safeDiv(ma2, mb2);

    // Monthly Total
    const mbt = mb1 + mb2, mat = ma1 + ma2;
    const mdt = (mbt - mat) * mult;
    const mpt = safeDiv(mat, mbt);

    // Push Monthly Block (Name 1, Name 2, Total)
    r1.push(mb1, mb2, mbt);
    r2.push(ma1, ma2, mat);
    r3.push(md1, md2, mdt);
    r4.push(mp1, mp2, mpt);
  }
  
  return [r1, r2, r3, r4];
}

function colToLet(c) {
  let l = '';
  while (c > 0) {
    let t = (c - 1) % 26;
    l = String.fromCharCode(t + 65) + l;
    c = (c - t - 1) / 26;
  }
  return l;
}

function safeSum(arr) {
  return Array.isArray(arr) ? arr.reduce((a, b) => a + b, 0) : 0;
}

function safeDiv(num, den) {
  if (den === 0) return (num === 0 ? 0 : 1);
  return num / den;
}
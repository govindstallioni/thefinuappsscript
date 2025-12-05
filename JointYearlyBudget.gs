/**
 * EXPERT OPTIMIZED Joint Yearly Budget Generator (Strict Column Mapping)
 *
 * MODIFICATIONS:
 * 1. Store Allocation Ratios: The Alloc 1 and Alloc 2 percentages from the Categories sheet
 * are now stored on the in-memory category object (catObj).
 * 2. NEW Actual Allocation Logic: In 'B. Process Transactions', if a transaction's owner is
 * 'Joint', 'Household', or empty, the transaction amount is split using the category's
 * stored alloc1 and alloc2 ratios, not the previous 50/50 split.
 * 3. Ratio Redistribution Removed: The secondary ratio redistribution logic in buildRowBlock
 * has been removed. The 'Actual' amounts displayed now directly reflect the allocated
 * amounts calculated in the transaction processing loop.
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
  
  // --- COLOR AND CONSTANTS (Ensured highest scope within function to fix ReferenceError) ---
  const borderColor = '#000000'; 
  const colorType = '#e68e68';
  const colorGroup = '#fce5cd'; 

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
    TRAN_ASSIGN_AMT_COL: tranConfigRaw[7] - 1,
    // Note: OWNER_MAP is no longer strictly needed but kept for context.
    OWNER_MAP: { [String(NAME1).toLowerCase()]: 1, [String(NAME2).toLowerCase()]: 2 }
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
    clearRange.setBackground(null).setFontWeight(null).setBorder(false, false, false, false, false, false);
  }
    
  // --- 3. DATA PROCESSING (IN-MEMORY) ---
  const tree = {}; 
  const catMap = {}; 

  // Initialize monthly summary trackers (12 months = 12 elements)
  const summaryTotals = {
    budgetIncome1: Array(12).fill(0), actualIncome1: Array(12).fill(0),
    budgetExpense1: Array(12).fill(0), actualExpense1: Array(12).fill(0),
    budgetIncome2: Array(12).fill(0), actualIncome2: Array(12).fill(0),
    budgetExpense2: Array(12).fill(0), actualExpense2: Array(12).fill(0)
  };

  // A. Process Categories (UPDATED: Store Allocations)
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
      // STORE ALLOCATIONS HERE for use in the transaction loop
      alloc1: alloc1, 
      alloc2: alloc2
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

  // B. Process Transactions (Actuals are tracked based on new allocation rules)
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
      let assignAmt = Number(row[CONFIG.TRAN_ASSIGN_AMT_COL]) || 0;
      // Expense amounts are often negative in transaction sheets, but we track them as positive expenses here.
      if (type !== 'Income' && type !== 'Transfers') amt = Math.abs(amt); 
      
      const ownerRaw = String(row[CONFIG.TRAN_OWNER_COL]).toLowerCase();
      const monthIdx = new Date(tDate).getMonth(); 
      let amt1 = 0, amt2 = 0;
      
      // Get the category-specific allocations
      const categoryAlloc1 = catEntry.alloc1;
      const categoryAlloc2 = catEntry.alloc2;
      
      // === NEW ALLOCATION LOGIC ===
      const name1Lower = String(CONFIG.NAME1).toLowerCase();
      const name2Lower = String(CONFIG.NAME2).toLowerCase();
      
      if (ownerRaw === name1Lower) {
          // If Owner is Name 1: 100% to Name 1's Actual
          amt1 = amt;
      } else if (ownerRaw === name2Lower) {
          // If Owner is Name 2: 100% to Name 2's Actual
          amt2 = amt;
      } else if (ownerRaw === 'joint' || ownerRaw === 'household' ) {
          // If Owner is Joint, Household, or blank: Split by category allocation ratio
          //amt1 = amt * categoryAlloc1;
          //amt2 = amt * categoryAlloc2;
          amt1 = assignAmt;
          amt2 = assignAmt;
      } else {
          // Fallback for unrecognized owner (e.g., a third party, or misspelled name) - fall back to 50/50
          amt1 = amt / 2;
          amt2 = amt / 2;
      }
      // ==========================

      catEntry.actual1[monthIdx] += amt1;
      catEntry.actual2[monthIdx] += amt2;
    }
  }

  // --- 4. OUTPUT GENERATION ---
  const nameRows = [], mainDataRows = [], metaRows = [];
  
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

      // ADD SPACER ROW AFTER EACH GROUP
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
    
    // Aggregate data for the D3:F4 Summary and Monthly Summary calculations
    for (let m = 0; m < 12; m++) {
        if (type === "Income" || type === "Transfers") {
            summaryTotals.budgetIncome1[m] += typeTotals.b1[m];
            summaryTotals.actualIncome1[m] += typeTotals.a1[m];
            summaryTotals.budgetIncome2[m] += typeTotals.b2[m];
            summaryTotals.actualIncome2[m] += typeTotals.a2[m];
        } else { // Expense types
            summaryTotals.budgetExpense1[m] += typeTotals.b1[m];
            summaryTotals.actualExpense1[m] += typeTotals.a1[m];
            summaryTotals.budgetExpense2[m] += typeTotals.b2[m];
            summaryTotals.actualExpense2[m] += typeTotals.a2[m];
        }
    }
  });

  // --- 5. WRITE & FORMAT DATA ---
  if (mainDataRows.length > 0) {
    const numRows = mainDataRows.length;
    const dataWidth = mainDataRows[0].length; 
    const lastDataColIndex = dataWidth + 3; 
    const finalColLetterRevised = colToLet(lastDataColIndex);

    // Write Data to Sheet
    outSheet.getRange(startRow, 2, numRows, 1).setValues(nameRows.map(x => [x])).setFontFamily("Comfortaa").setFontSize(10).setHorizontalAlignment('left');
    outSheet.getRange(startRow, 3, numRows, dataWidth).setValues(mainDataRows).setFontFamily("Comfortaa").setFontSize(10);
    
    // --- ROW-BASED FORMATTING ---
    const currencyRanges = [];
    const percentRanges = [];
    const nameMergeRanges = [];
    const rightBorderRanges = []; 
    const bottomBorderRanges = []; 
    
    let i = 0;
    while (i < numRows) {
      const r = startRow + i;
      const meta = metaRows[i];

      // Handle Spacer Rows
      if (meta === "SPACER") {
        i++; 
        continue;
      }

      // Handle Data Blocks (Type, Group, Cat) - 4 Rows
      if (meta === "TYPE" || meta === "GROUP" || meta === "CAT") {
        
        // 1. Merge Name Column
        nameMergeRanges.push(`B${r}:B${r+3}`);

        // 2. Headers Background
        if (meta === "TYPE" || meta === "GROUP") {
          outSheet.getRange(r, 2, 4, dataWidth + 1).setBackground(meta === "TYPE" ? colorType : colorGroup).setFontWeight('bold');
        }

        // 3. Collect Bottom Border Range (Col B to end of data, on the 4th row)
        if (meta === "GROUP" || meta === "CAT") {
            const lastRowOfBlock = r + 3;
            // Range includes column B, C, and all data columns.
            const rangeStr = `B${lastRowOfBlock}:${finalColLetterRevised}${lastRowOfBlock}`;
            bottomBorderRanges.push(rangeStr);
        }

        // 4. Number Formats and Right Borders (Iterate 4 rows of the block)
        for (let rowOffset = 0; rowOffset < 4; rowOffset++) {
            const currR = r + rowOffset;
            const rowType = rowOffset; // 0:Budget, 1:Actual, 2:Diff, 3:%
            
            // Data range from D to finalColLetter
            const dataRowRange = `D${currR}:${finalColLetterRevised}${currR}`;
            
            if (rowType === 3) percentRanges.push(dataRowRange);
            else currencyRanges.push(dataRowRange);

            // COLLECT ALL COLUMN RIGHT BORDERS (N1, N2, Total columns)
            // Annual Block Borders (D, E, F)
            rightBorderRanges.push(`${colToLet(4)}${currR}`); // D (Name 1)
            rightBorderRanges.push(`${colToLet(5)}${currR}`); // E (Name 2)
            rightBorderRanges.push(`${colToLet(6)}${currR}`); // F (Total)
            
            // Monthly Block Borders (H, I, J | K, L, M | ... | AQ)
            for (let m = 0; m < 12; m++) {
                // Name 1 column (Col 8, 11, 14...)
                rightBorderRanges.push(`${colToLet(8 + m * 3)}${currR}`); 
                // Name 2 column (Col 9, 12, 15...)
                rightBorderRanges.push(`${colToLet(9 + m * 3)}${currR}`);
                // Total column (Col 10, 13, 16... up to AQ)
                rightBorderRanges.push(`${colToLet(10 + m * 3)}${currR}`);
            }
        }

        i += 4; // Jump 4 rows for next iteration
      }
    }

    // Batch Apply Formats
    const fmtCurrency = '_($* #,##0.00_);_($* (#,##0.00)_);_($* "-"_);_(@_)';
    const fmtPercent = '0.00%';
    
    if (currencyRanges.length) outSheet.getRangeList(currencyRanges).setNumberFormat(fmtCurrency);
    if (percentRanges.length) outSheet.getRangeList(percentRanges).setNumberFormat(fmtPercent);
    if (nameMergeRanges.length) {
      const ranges = outSheet.getRangeList(nameMergeRanges).getRanges();
      ranges.forEach(rng => rng.merge().setVerticalAlignment('top').setWrap(true));
    }

    // Apply the requested right borders (All column separators)
    if (rightBorderRanges.length) {
        const rl = outSheet.getRangeList(rightBorderRanges);
        // Right border only (false, false, false, true, false, false)
        rl.setBorder(false, false, false, true, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);
    }
    
    // CRITICAL FIX: Use iterative setBorder for horizontal lines to guarantee merged cell edge is drawn
    if (bottomBorderRanges.length) {
        const ranges = outSheet.getRangeList(bottomBorderRanges).getRanges();
        ranges.forEach(range => {
            // Apply bottom border only (Top, Left, Bottom, Right, Vertical, Horizontal)
            range.setBorder(null, null, true, null, null, null, borderColor, SpreadsheetApp.BorderStyle.SOLID);
        });
    }

    // APPLY FULL OUTER BORDER (Ensures the far-right border of AQ and the bottom of the final row)
    const finalRow = startRow + numRows - 1;
    const outerRange = outSheet.getRange(`B${startRow}:${finalColLetterRevised}${finalRow}`);
    outerRange.setBorder(true, true, true, true, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);

    // Alignments
    outSheet.getRange(startRow, 4, numRows, dataWidth - 1).setHorizontalAlignment('right'); 
    outSheet.getRange(startRow, 3, numRows, 1).setHorizontalAlignment('left'); 
  }

  // --- 6. WRITE SUMMARY CALCULATIONS (D3:F4 and Monthly H3:AS4) ---

  const monthlySummaryBudgetRow = [];
  const monthlySummaryActualRow = [];

  for (let m = 0; m < 12; m++) {
    // 1. Calculate Monthly Net Totals
    const netBudget1 = summaryTotals.budgetIncome1[m] - summaryTotals.budgetExpense1[m];
    const netActual1 = summaryTotals.actualIncome1[m] - summaryTotals.actualExpense1[m];

    const netBudget2 = summaryTotals.budgetIncome2[m] - summaryTotals.budgetExpense2[m];
    const netActual2 = summaryTotals.actualIncome2[m] - summaryTotals.actualExpense2[m];

    const netBudgetTotal = netBudget1 + netBudget2;
    const netActualTotal = netActual1 + netActual2;

    // 2. Populate Monthly Summary Rows (3 columns per month: N1, N2, Total)
    monthlySummaryBudgetRow.push(netBudget1, netBudget2, netBudgetTotal);
    monthlySummaryActualRow.push(netActual1, netActual2, netActualTotal);
  }
  
  // --- ANNUAL SUMMARY (D3:F4) ---
  const annNetBudget1 = safeSum(summaryTotals.budgetIncome1) - safeSum(summaryTotals.budgetExpense1);
  const annNetActual1 = safeSum(summaryTotals.actualIncome1) - safeSum(summaryTotals.actualExpense1);
  const annNetBudget2 = safeSum(summaryTotals.budgetIncome2) - safeSum(summaryTotals.budgetExpense2);
  const annNetActual2 = safeSum(summaryTotals.actualIncome2) - safeSum(summaryTotals.actualExpense2);

  const annNetBudgetTotal = annNetBudget1 + annNetBudget2;
  const annNetActualTotal = annNetActual1 + annNetActual2;

  // Write Annual Net Totals (D3:F4)
  outSheet.getRange('D3').setValue(annNetBudget1);
  outSheet.getRange('D4').setValue(annNetActual1);
  outSheet.getRange('E3').setValue(annNetBudget2);
  outSheet.getRange('E4').setValue(annNetActual2);
  outSheet.getRange('F3').setValue(annNetBudgetTotal);
  outSheet.getRange('F4').setValue(annNetActualTotal);

  // --- MONTHLY SUMMARY (H3:AS4, or equivalent) ---
  const monthlyDataWidth = monthlySummaryBudgetRow.length; // Should be 36 (12*3)
  let combinedSummaryRanges;

  if (monthlyDataWidth > 0) {
      // Write the 2x36 array starting at H3 (column index 8)
      const monthlySummaryRange = outSheet.getRange(3, 8, 2, monthlyDataWidth);
      monthlySummaryRange.setValues([
          monthlySummaryBudgetRow, 
          monthlySummaryActualRow
      ]);

      // Combine Annual and Monthly Ranges for Formatting
      combinedSummaryRanges = outSheet.getRangeList([
          outSheet.getRange('D3:F4').getA1Notation(), 
          monthlySummaryRange.getA1Notation()
      ]);
      
      // Add right borders to the monthly summary block for visual separation
      const monthlySummaryBorders = [];
      for(let m = 0; m < 12; m++) {
          const colN1 = 8 + m * 3; // Col H, K, N, ...
          const colN2 = 9 + m * 3; // Col I, L, O, ...
          const colTotal = 10 + m * 3; // Col J, M, P, ...
          
          // Apply border to Name 1 and Name 2 columns for both Budget (3) and Actual (4) rows
          monthlySummaryBorders.push(`${colToLet(colN1)}3`, `${colToLet(colN1)}4`); 
          monthlySummaryBorders.push(`${colToLet(colN2)}3`, `${colToLet(colN2)}4`);
          // Apply border to Total column
          monthlySummaryBorders.push(`${colToLet(colTotal)}3`, `${colToLet(colTotal)}4`);
      }

      if (monthlySummaryBorders.length) {
          const ml = outSheet.getRangeList(monthlySummaryBorders);
          // Right border only (false, false, false, true, false, false)
          ml.setBorder(false, false, false, true, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);
      }

  } else {
    // Only use the annual summary range if monthly is empty
    combinedSummaryRanges = outSheet.getRangeList([outSheet.getRange('D3:F4').getA1Notation()]);
  }

  // Apply Formatting to all Summary cells (D3:F4 and monthly)
  const fmtCurrency = '_($* #,##0.00_);_($* (#,##0.00)_);_($* "-"_);_(@_)';
  combinedSummaryRanges.setNumberFormat(fmtCurrency);
  combinedSummaryRanges.setFontWeight('bold');


  outSheet.getRange('B3').setValue("Last updated on " + getDateTime());
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

/**
 * Builds the 4 rows (Budget, Actual, Diff, %) for the Annual and Monthly columns.
 * **UPDATED:** The 'Actual' amounts (a1Arr, a2Arr) are now used directly as they
 * already reflect the category-based allocation from the transaction loop.
 */
function buildRowBlock(name, b1Arr, a1Arr, b2Arr, a2Arr, type) {
  const isIncome = (type === "Income" || type === "Transfers");
  // Multiplier flips the sign of 'Diff' for Income categories (positive Diff means bad)
  const mult = isIncome ? -1 : 1;
  const sum = safeSum;

  // --- 1. ANNUAL CALCS (Using Pre-Allocated Actuals) ---
  const annB1 = sum(b1Arr);
  const annB2 = sum(b2Arr);
  const annBT = annB1 + annB2;

  // Actuals are now the sums of the pre-allocated amounts from the transaction loop
  const annA1_display = sum(a1Arr); 
  const annA2_display = sum(a2Arr);
  const annAT_total = annA1_display + annA2_display; // Total Actual is the sum of the allocated parts

  // Differences and Percentages use the allocated actuals
  const annD1 = (annB1 - annA1_display) * mult;
  const annP1 = safeDiv(annA1_display, annB1);

  const annD2 = (annB2 - annA2_display) * mult;
  const annP2 = safeDiv(annA2_display, annB2);

  // Total differences and percentages use the total actual spend
  const annDT = (annBT - annAT_total) * mult;
  const annPT = safeDiv(annAT_total, annBT);

  // Row Arrays initialized with Labels (Col C)
  const r1 = ["Budget"], r2 = ["Actual"], r3 = ["Diff"], r4 = ["%"];
  
  // --- ANNUAL BLOCK PUSH (Cols D-F) ---
  r1.push(annB1); r2.push(annA1_display); r3.push(annD1); r4.push(annP1); // Use allocated actuals
  r1.push(annB2); r2.push(annA2_display); r3.push(annD2); r4.push(annP2); // Use allocated actuals
  r1.push(annBT); r2.push(annAT_total); r3.push(annDT); r4.push(annPT);
  // Col G: Spacer (Empty)
  r1.push(""); r2.push(""); r3.push(""); r4.push("");

  // --- 2. MONTHLY BLOCKS PUSH (Cols H onwards) ---
  for (let m=0; m<12; m++) {
    const mb1 = b1Arr[m];
    const mb2 = b2Arr[m];
    const mbt = mb1 + mb2;

    const ma1_display = a1Arr[m]; // Monthly allocated actual
    const ma2_display = a2Arr[m]; // Monthly allocated actual
    const mat_total = ma1_display + ma2_display; // Monthly Total Actual Spend

    // Monthly Name 1 
    const md1 = (mb1 - ma1_display) * mult;
    const mp1 = safeDiv(ma1_display, mb1);

    // Monthly Name 2 
    const md2 = (mb2 - ma2_display) * mult;
    const mp2 = safeDiv(ma2_display, mb2);

    // Monthly Total 
    const mdt = (mbt - mat_total) * mult;
    const mpt = safeDiv(mat_total, mbt);

    // Push Monthly Block (Name 1, Name 2, Total)
    r1.push(mb1, mb2, mbt);
    r2.push(ma1_display, ma2_display, mat_total); // Pushing allocated actuals directly
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
  if (den === 0) return (num === 0 ? 0 : 1); // Avoids division by zero, returns 100% (1) if num > 0
  return num / den;
}
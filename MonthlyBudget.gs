/**
 * @fileoverview Contains functions to populate the Monthly Budget sheet
 * based on configuration, category data, and transaction data.
 * FIX: Refactored data aggregation (Step 4) to correctly allow a single Group
 * (like 'Transfers') to be aggregated under multiple Types (e.g., Income and Expense)
 * if its constituent categories are split between those Types.
 * FIX: Removed all border generation commands.
 */

// --- Constants for Column Indices (Read from Control Sheet) ---
const COL = {
    CATEGORY: 0, // Placeholder, actual index read from C5
    GROUP: 1,    // Placeholder, actual index read from C6
    TYPE: 2,     // Placeholder, actual index read from C7
    HIDE: 3,     // Placeholder, actual index read from C8
    NAME1: 4,    // Placeholder, actual index read from C9 (Multiplier)
    NAME2: 5,    // Placeholder, actual index read from C10 (Multiplier)
    BUDGET: 6    // Placeholder, actual index read from W2:W12 (dynamic based on month)
};

/**
 * Main function to populate the Monthly Budget sheet.
 */
function populateMonthlyBudget() {

    // --- Setup and Sheet References ---
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const controlSheet = ss.getSheetByName('Definition');
    const catSheet = ss.getSheetByName("Categories");
    const tranSheet = ss.getSheetByName("Transactions");
    const outputSheet = ss.getSheetByName("Monthly Budget");

    if (!controlSheet || !catSheet || !tranSheet || !outputSheet) {
        SpreadsheetApp.getUi().alert("One or more required sheets (Definition, Categories, Transactions, Monthly Budget) are missing.");
        return;
    }

    // --- Performance Optimization: Read all Control Data in one go ---
    const budgetConfigRange = controlSheet.getRange('C3:C12').getValues().flat();
    const dateConfig = controlSheet.getRange('AC4').getValue();
    const monthColumns = controlSheet.getRange('W2:W12').getValues().flat();

    // Category Data Column Mappings (indices based on 0-based array)
    const catLstrw = budgetConfigRange[0]; // C3 (Last row of Categories)
    const catMappings = {
        categoryCol: budgetConfigRange[2] - 1, // C5
        groupCol: budgetConfigRange[3] - 1,    // C6
        typeCol: budgetConfigRange[4] - 1,     // C7
        hideCol: budgetConfigRange[5] - 1,     // C8
        name1IdCol: budgetConfigRange[7] - 1,  // C10 (Name 1 Multiplier Column)
        name2IdCol: budgetConfigRange[8] - 1,  // C11 (Name 2 Multiplier Column)
    };

    // --- Determine Target Month and Budget Column ---
    const dateValue = dateConfig instanceof Date ? dateConfig : new Date(dateConfig);
    const catMonthID = dateValue.getMonth() + 1;
    const monthIndex = catMonthID - 1;
    const catMonthBudgetCol = monthColumns[monthIndex] - 1; 

    // --- Clear Previous Data from Monthly Budget Sheet ---
    const lastRow = outputSheet.getLastRow();
    const lastCol = outputSheet.getLastColumn();
    let trgtShtLstrw = lastRow > 9 ? lastRow : 9;

    if (trgtShtLstrw >= 10) {
        const clearRange = outputSheet.getRange(10, 1, trgtShtLstrw - 9, lastCol);
        clearRange.breakApart();
        clearRange.clear({ contentsOnly: true, formatOnly: false, validationsOnly: false });
        // Also clear any previous borders
        clearRange.setBorder(false, false, false, false, false, false); 
        clearRange.setNumberFormat(null).setBackground(null).setFontColor(null);
    }
    outputSheet.getRange("A:A").setFontColor("#FFFFFF");

    // --- 1. Process Category Data (Budget Data) ---
    const catData = catSheet.getRange(1, 1, catLstrw, catSheet.getLastColumn()).getValues();
    const catAry = [];
    const visibleCategories = new Set(); // Store all visible category names for transaction filtering

    for (let l = 1; l < catLstrw; l++) {
        const row = catData[l];

        // 1. Check hideCol and skip categories marked to be hidden
        const hideValue = String(row[catMappings.hideCol] || '').trim().toLowerCase();
        if (hideValue === 'hide' || hideValue === 'yes' || hideValue === 'true') {
            continue; // Skip this category entirely from budget data
        }

        // Apply trimming to ensure consistent keys/names across sheets
        const categoryName = String(row[catMappings.categoryCol] || '').trim(); 
        const groupName = String(row[catMappings.groupCol] || '').trim();
        const typeName = String(row[catMappings.typeCol] || '').trim();

        if (categoryName) {
            visibleCategories.add(categoryName); // Add trimmed category to the visible set
        }

        // Parse multipliers and budget amount immediately to ensure they are numbers
        const name1Multiplier = Number(row[catMappings.name1IdCol] || 0);
        const name2Multiplier = Number(row[catMappings.name2IdCol] || 0);
        const budgetAmount = Number(row[catMonthBudgetCol] || 0);

        catAry.push([
            categoryName,                   // 0: Category (string, trimmed)
            groupName,                    // 1: Group (string, trimmed)
            typeName,                     // 2: Type (string, trimmed)
            row[catMappings.hideCol],     // 3: Hide (string, raw value)
            name1Multiplier,              // 4: Name1 Multiplier (number)
            name2Multiplier,              // 5: Name2 Multiplier (number)
            budgetAmount                  // 6: Budget (number)
        ]);
    }

    // --- 2. Process Transaction Data (Actual Data) ---
    const tranConfigRange = controlSheet.getRange('I3:I12').getValues().flat();
    const tanLstRw = tranConfigRange[0]; // Last row
    const tanLstCol = tranConfigRange[1]; // Last column
    
    // Fetch data starting from Col B (Index 2 in Sheet, Index 0 in JS Array)
    const tranData = tranSheet.getRange(2, 2, tanLstRw - 1, tanLstCol).getValues(); 
    
    // Adjusted Transaction Column Indices (relative to the fetched range, starting at B2)
    const TRAN_COL = {
        DATE: 0,
        CATEGORY: 2, 
        AMOUNT: 3,   
        OWNER: 4,
        ASSIGNED: 5,
        GROUPED_KEY: 12
    };

    const formatTransGroup = {}; // Used to collect transactions by their Grouped Key
    const targetMonth = dateValue.getMonth(); 
    const targetYear = dateValue.getFullYear();
    
    tranData.forEach(row => {
        const dateCell = row[TRAN_COL.DATE];
        let d;
        if (dateCell instanceof Date) d = dateCell;
        else if (typeof dateCell === 'string' && dateCell) d = new Date(dateCell);
        else return;
        
        if (d.getMonth() === targetMonth && d.getFullYear() === targetYear) {
            // Trim transaction category for consistent lookup against visibleCategories set
            const transactionCategory = String(row[TRAN_COL.CATEGORY] || '').trim();

            // CRITICAL: Skip transaction if its category is not visible (hidden)
            if (!visibleCategories.has(transactionCategory)) {
                return; 
            }
            
            // Trim grouped key for consistent grouping
            const groupedKey = String(row[TRAN_COL.GROUPED_KEY] || '').trim(); 
            const owner = row[TRAN_COL.OWNER]; 

            // Only process transactions linked to a group
            if (groupedKey !== 'Not Grouped' && groupedKey) {
                if (!formatTransGroup[groupedKey]) {
                    formatTransGroup[groupedKey] = [];
                }
                // Store raw amounts for visible categories only
                formatTransGroup[groupedKey].push({
                    'category': transactionCategory, // trimmed category name
                    'amount': row[TRAN_COL.AMOUNT] || 0,
                    'owner': owner || 'Joint' 
                });
            }
        }
    });
    
    // --- 3. Process All Visible Categories (Budget + Actual calculation) ---
    // Returns a flat list of fully processed category objects.
    const allProcessedCategories = aggregateCategoryData(catAry, formatTransGroup); 

    if (allProcessedCategories.length === 0) {
        Logger.log("No data formatted for output.");
        return;
    }

    // --- 4. Re-structure for Output: Type -> Group -> Categories (Handles split groups) ---
    // Structure: { TypeName: { GroupName: { budget: X, actual: Y, categories: [catObj, ...] } } }
    const structuredReport = {}; 
    
    allProcessedCategories.forEach(cat => {
        const type = cat.type || 'Other';
        const group = cat.group || 'Ungrouped';
        
        // Ensure the structure path exists
        if (!structuredReport[type]) structuredReport[type] = {};
        if (!structuredReport[type][group]) {
            structuredReport[type][group] = {
                budget: 0,
                actual: 0,
                categories: []
            };
        }
        
        // Aggregate Group Totals (This is the key fix: aggregation happens based on Type + Group)
        structuredReport[type][group].budget += cat.budget;
        structuredReport[type][group].actual += cat.actual;
        
        // Add Category Detail
        structuredReport[type][group].categories.push(cat);
    });

    // Final output generation loop
    const outputRows = [];
    const desiredOrder = ['Income', 'Expense'];
    
    // Determine the order of types to display
    const typeOrder = [];
    
    // 1. Add Income and Expense first
    desiredOrder.forEach(t => {
        if (structuredReport[t] && Object.keys(structuredReport[t]).length > 0) {
             typeOrder.push(t);
        }
    });

    // 2. Add remaining types (like 'Transfers', 'Loans', etc.) that were populated, in alphabetical order
    Object.keys(structuredReport)
        .filter(k => !desiredOrder.includes(k))
        .sort()
        .forEach(k => typeOrder.push(k));

    
    typeOrder.forEach(type => {
        const groups = structuredReport[type];
        
        // Calculate Type Totals
        let typeBudget = 0;
        let typeActual = 0;
        Object.keys(groups).forEach(groupName => {
            typeBudget += groups[groupName].budget;
            typeActual += groups[groupName].actual;
        });

        // Use 'Transfers' Type for cash flow logic (where actual > budget is usually bad for expenses)
        const isPositiveFlow = (type === 'Income' || type === 'Transfers');
        const typeSurdef = isPositiveFlow ? (typeActual - typeBudget) : (typeBudget - typeActual);
        const typePct = getItemPercentage(typeBudget, typeActual);

        // Add Type Header Row
        outputRows.push({
            data: [null, type, typeBudget, typeActual, typeSurdef, typePct],
            rowType: 'TYPE'
        });

        // Sort groups alphabetically for consistency
        const sortedGroups = Object.keys(groups).sort();

        // Add Group and Category Rows
        sortedGroups.forEach(groupName => {
            const item = groups[groupName]; // This item is the aggregated Group data for the current Type

            // Group Row
            const surdef = isPositiveFlow ? (Number(item.actual) - Number(item.budget)) : (Number(item.budget) - Number(item.actual));
            const percentage = getItemPercentage(Number(item.budget), Number(item.actual));
            
            outputRows.push({
                data: [null, groupName, item.budget, item.actual, surdef, percentage],
                rowType: 'GROUP'
            });
            
            // Category Rows (sort categories alphabetically)
            item.categories.sort((a, b) => a.category.localeCompare(b.category)).forEach(category => {
                const catSurdef = isPositiveFlow ? (Number(category.actual) - Number(category.budget)) : (Number(category.budget) - Number(category.actual));
                const catPercentage = getItemPercentage(Number(category.budget), Number(category.actual));
                outputRows.push({
                    data: [null, category.category, category.budget, category.actual, catSurdef, catPercentage],
                    rowType: 'CATEGORY'
                });
            });

            // Add Spacer Row (Empty Line) at the end of the group
            outputRows.push({
                data: [null, null, null, null, null, null],
                rowType: 'SPACER'
            });
        });
    });

    // --- 5. Write Data to Output Sheet ---
    const writeRange = outputSheet.getRange(10, 2, outputRows.length, 5); 
    const valuesToWrite = outputRows.map(r => r.data.slice(1)); 
    writeRange.setValues(valuesToWrite);
    
    // --- 6. Apply Formatting (Borders removed) ---
    const typeBg = '#e68e68'; 
    const groupBg = '#eec49f'; 
    const catBg = '#FFFFFF';   
    const numFormat = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"_);_(@_)';
    const pctFormat = '0.00%';
    
    let rangeStart = 10;
    for (const row of outputRows) {
        
        // Backgrounds
        let bg = catBg;
        if (row.rowType === 'TYPE') bg = typeBg;
        else if (row.rowType === 'GROUP') bg = groupBg;
        
        // Skip formatting logic for SPACER rows entirely (keep them white/empty)
        if (row.rowType === 'SPACER') {
            // Ensure no border is set on the row preceding the spacer
            outputSheet.getRange(rangeStart - 1, 2, 1, 5).setBorder(false, false, false, false, false, false);
            rangeStart++;
            continue;
        }

        const rowRange = outputSheet.getRange(rangeStart, 2, 1, 5); 
        rowRange.setBackground(bg);
        
        // Alignment and Format
        const nameCell = outputSheet.getRange(rangeStart, 2); 
        const dataRange = outputSheet.getRange(rangeStart, 3, 1, 4); 
        const pctRange = outputSheet.getRange(rangeStart, 6); 
        
        nameCell.setHorizontalAlignment("left");
        dataRange.setNumberFormat(numFormat).setHorizontalAlignment("right");
        
        if (row.data[5] === 'N/A') {
            pctRange.setNumberFormat('@').setValue('N/A');
        } else {
            pctRange.setNumberFormat(pctFormat);
        }

        // Bold Headers
        if (row.rowType === 'TYPE' || row.rowType === 'GROUP') {
             rowRange.setFontWeight('bold');
        }
        
        // Explicitly ensure no borders are added to this row
        rowRange.setBorder(false, false, false, false, false, false);


        rangeStart++;
    }

    // --- 7. Update Summary & Cleanup ---
    updateSheetSummary(outputSheet);
    //outputSheet.getRange('A1').activate();
}

/**
 * Helper to consolidate all visible categories, apply owner multipliers,
 * and calculate the final budget and actual amounts for each category item.
 * @param {Array<Array>} catAry - Array of visible categories from the Categories sheet.
 * @param {Object} formatTransGroup - Raw transaction data grouped by Grouped Key.
 * @returns {Array<Object>} Flat list of fully processed category objects.
 */
function aggregateCategoryData(catAry, formatTransGroup) {

    const parseAmount = (val) => {
        if (typeof val === 'number') return val;
        // Strip non-numeric characters except for leading/embedded signs and decimal points
        if (typeof val === 'string') return parseFloat(String(val).replace(/[^0-9.-]/g, '')) || 0;
        return 0;
    };

    // 1. Aggregate raw actual amounts from transactions to a Category-level map for easy lookup
    const actualMap = {}; 
    Object.keys(formatTransGroup).forEach(group => {
        formatTransGroup[group].forEach(item => {
            const cat = item.category || 'Unknown';
            const amt = parseAmount(item.amount); 
            // Key is the Category name
            actualMap[cat] = (actualMap[cat] || 0) + amt;
        });
    });

    const result = [];
    
    // 2. Iterate through visible categories from the Categories sheet (catAry)
    catAry.forEach(catRow => {
        const category = catRow[0]; 
        const group = catRow[1]; 
        const type = catRow[2]; 
        
        const name1Multiplier = catRow[4]; // Index 4: Name1 Multiplier (number)
        const name2Multiplier = catRow[5]; // Index 5: Name2 Multiplier (number)
        const combinedMultiplier = Number(name1Multiplier) + Number(name2Multiplier);
        
        const budgetAmount = parseAmount(catRow[6] || 0); // Index 6: Budget (number)
        
        // Get raw actual amount for this category (which is 0 if no visible transactions exist)
        const actualAmount = parseAmount(actualMap[category] || 0);
  
        // Push the fully calculated category object
        result.push({
            type: type,
            group: group,
            category: category,
            budget: budgetAmount, 
            actual: actualAmount 
        });
    });
    
    return result;
}


/**
 * Helper to update the summary section (Top of sheet)
 */
function updateSheetSummary(outputSheet) {
    let diffinc = 0;
    let diffexp = 0;
    
    const finalData = outputSheet.getDataRange().getValues();
    
    for (let i = 9; i < finalData.length; i++) { 
        const row = finalData[i];
        const label = row[1]; // Column B
        // Column D (index 3) contains the Actual amount for Type rows (B, C, D, E)
        if (label === "Income") diffinc = Number(row[3]) || diffinc; 
        if (label === "Expense") diffexp = Number(row[3]) || diffexp;
    }

    // Correct calculation for Over/Under based on actuals
    const overUnder = Number((diffinc - diffexp).toFixed(2)); // Income - Expense
    const summaryRange = outputSheet.getRange(5, 5);
    const statusRange = outputSheet.getRange(5, 6);

    if (overUnder === 0) {
        summaryRange.setValue('On Budget').setNumberFormat('@');
        statusRange.setValue('');
        summaryRange.setBackground('#D9EAD3');
    } else {
        const absOverUnder = Math.abs(overUnder);
        if (overUnder > 0) { // Positive means Actual Income > Actual Expense (Good)
            summaryRange.setValue(absOverUnder).setNumberFormat('$ #,##0.00');
            statusRange.setValue('Under Budget');
            summaryRange.setBackground('#D9EAD3');
        } else { // Negative means Actual Expense > Actual Income (Bad)
            summaryRange.setValue(absOverUnder).setNumberFormat('$ #,##0.00');
            statusRange.setValue('Over Budget');
            summaryRange.setBackground('#F4CCCC');
        }
    }
    
    const dt = getDateTime();
    outputSheet.getRange('B3').setValue("Last updated on " + dt);
}


function getItemPercentage(budget, actual) {
    budget = Number(budget);
    actual = Number(actual);
    if (!isFinite(budget) || !isFinite(actual)) return 'N/A';
    if (budget === 0) return actual === 0 ? 0 : 'N/A';
    return actual / budget;
}
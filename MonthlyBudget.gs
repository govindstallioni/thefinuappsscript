/**
 * @fileoverview Contains functions to populate the Monthly Budget sheet
 * based on configuration, category data, and transaction data.
 * Updated with Borders and Spacer rows.
 */

// --- Constants for Column Indices (Read from Control Sheet) ---
const COL = {
    CATEGORY: 0, // Placeholder, actual index read from C5
    GROUP: 1,    // Placeholder, actual index read from C6
    TYPE: 2,     // Placeholder, actual index read from C7
    HIDE: 3,     // Placeholder, actual index read from C8
    NAME1: 4,    // Placeholder, actual index read from C9
    NAME2: 5,    // Placeholder, actual index read from C10
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
    const catLstrw = budgetConfigRange[0]; // C3
    const catMappings = {
        categoryCol: budgetConfigRange[2] - 1, // C5
        groupCol: budgetConfigRange[3] - 1,    // C6
        typeCol: budgetConfigRange[4] - 1,     // C7
        hideCol: budgetConfigRange[5] - 1,     // C8
        name1IdCol: budgetConfigRange[7] - 1,  // C10
        name2IdCol: budgetConfigRange[8] - 1,  // C11
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
        // Use breakApart to prevent merge conflicts before clearing
        const clearRange = outputSheet.getRange(10, 1, trgtShtLstrw - 9, lastCol);
        clearRange.breakApart();
        clearRange.clear({ contentsOnly: true, formatOnly: false, validationsOnly: false });
        // Reset all formats including borders
        clearRange.setNumberFormat(null).setBackground(null).setFontColor(null).setBorder(false, false, false, false, false, false);
    }
    outputSheet.getRange("A:A").setFontColor("#FFFFFF");

    // --- 1. Process Category Data (Budget Data) ---
    const catData = catSheet.getRange(1, 1, catLstrw, catSheet.getLastColumn()).getValues();
    const catAry = []; 

    for (let l = 1; l < catLstrw; l++) {
        const row = catData[l];
        
        catAry.push([
            row[catMappings.categoryCol], // 0: Category
            row[catMappings.groupCol],    // 1: Group
            row[catMappings.typeCol],     // 2: Type
            row[catMappings.hideCol],     // 3: Hide
            row[catMappings.name1IdCol],  // 4: Name1
            row[catMappings.name2IdCol],  // 5: Name2
            row[catMonthBudgetCol]        // 6: Budget
        ]);
    }

    // Grouping by Group Name
    const formatCatGroup = {};
    catAry.forEach(item => {
        const key = item[1]; 
        if (key) { 
            if (!formatCatGroup[key]) {
                formatCatGroup[key] = [];
            }
            formatCatGroup[key].push(item);
        }
    });

    // --- 2. Process Transaction Data (Actual Data) ---
    const tranConfigRange = controlSheet.getRange('I3:I12').getValues().flat();
    const tanLstRw = tranConfigRange[0]; 
    const tanLstCol = tranConfigRange[1]; 
    
    // Fetch data starting from Col B (Index 2 in Sheet, Index 0 in JS Array)
    const tranData = tranSheet.getRange(2, 2, tanLstRw - 1, tanLstCol).getValues(); 
    
    const TRAN_COL = {
        DATE: 0,
        CATEGORY: 2, 
        AMOUNT: 3,   
        OWNER: 4,
        ASSIGNED: 5,
        GROUPED_KEY: 12 
    };

    const formatTransGroup = {};
    const targetMonth = dateValue.getMonth(); 
    const targetYear = dateValue.getFullYear();
    
    tranData.forEach(row => {
        const dateCell = row[TRAN_COL.DATE];
        let d;
        if (dateCell instanceof Date) d = dateCell;
        else if (typeof dateCell === 'string' && dateCell) d = new Date(dateCell);
        else return;
        
        if (d.getMonth() === targetMonth && d.getFullYear() === targetYear) {
            const groupedKey = row[TRAN_COL.GROUPED_KEY]; 
            const owner = row[TRAN_COL.OWNER]; 

            if (groupedKey !== 'Not Grouped' && owner !== 'Joint') {
                if (!formatTransGroup[groupedKey]) {
                    formatTransGroup[groupedKey] = [];
                }
                formatTransGroup[groupedKey].push({
                    'category': row[TRAN_COL.CATEGORY],
                    'amount': row[TRAN_COL.AMOUNT] || 0,
                    'assigned': row[TRAN_COL.ASSIGNED] || 0
                });
            }
        }
    });

    // --- 3. Consolidate Data ---
    const formatArrData = preFormatMonthlyReportData(formatCatGroup, formatTransGroup);

    if (formatArrData.length === 0) {
        Logger.log("No data formatted for output.");
        return;
    }

    // --- 4. Re-structure for Output: Order by Type (Income -> Expense) ---
    const groupedByType = {};
    groupedByType['Income'] = [];
    groupedByType['Expense'] = [];

    formatArrData.forEach(groupObj => {
        const type = groupObj.type || 'Other';
        if (!groupedByType[type]) {
            groupedByType[type] = [];
        }
        groupedByType[type].push(groupObj);
    });

    const typeOrder = ['Income', 'Expense', ...Object.keys(groupedByType).filter(k => k !== 'Income' && k !== 'Expense')];
    const outputRows = [];

    typeOrder.forEach(type => {
        const groups = groupedByType[type];
        if (!groups || groups.length === 0) return;

        // Calculate Type Totals
        let typeBudget = 0;
        let typeActual = 0;
        groups.forEach(g => {
            typeBudget += Number(g.budget);
            typeActual += Number(g.actual);
        });
        
        const isPositiveFlow = (type === 'Income' || type === 'Transfers');
        const typeSurdef = isPositiveFlow ? (typeActual - typeBudget) : (typeBudget - typeActual);
        const typePct = getItemPercentage(typeBudget, typeActual);

        // Add Type Header Row
        outputRows.push({
            data: [null, type, typeBudget, typeActual, typeSurdef, typePct],
            rowType: 'TYPE'
        });

        // Add Group and Category Rows
        groups.forEach(item => {
            // Group Row
            const surdef = isPositiveFlow ? (Number(item.actual) - Number(item.budget)) : (Number(item.budget) - Number(item.actual));
            const percentage = getItemPercentage(Number(item.budget), Number(item.actual));
            
            outputRows.push({
                data: [null, item.group, item.budget, item.actual, surdef, percentage],
                rowType: 'GROUP'
            });
            
            // Category Rows
            if (item.categories.length > 0) {
                item.categories.forEach(category => {
                    const catSurdef = isPositiveFlow ? (Number(category.amount) - Number(category.budget)) : (Number(category.budget) - Number(category.amount));
                    const catPercentage = getItemPercentage(Number(category.budget), Number(category.amount));
                    outputRows.push({
                        data: [null, category.category, category.budget, category.amount, catSurdef, catPercentage],
                        rowType: 'CATEGORY'
                    });
                });
            }

            // Add Spacer Row (Empty Line) at the end of the category list
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
    
    // --- 6. Apply Formatting ---
    const typeBg = '#e68e68'; 
    const groupBg = '#eec49f'; 
    const catBg = '#FFFFFF';   
    const numFormat = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"_);_(@_)';
    const pctFormat = '0.00%';
    const borderColor = '#000000';

    // Array to collect ranges for borders
    const borderRanges = [];

    let rangeStart = 10;
    for (const row of outputRows) {
        const rowRange = outputSheet.getRange(rangeStart, 2, 1, 5); 
        
        // If it is NOT a spacer, add to border list
        if (row.rowType !== 'SPACER') {
             // Convert range to A1 notation for range list (e.g., B10:F10)
             borderRanges.push(`B${rangeStart}:F${rangeStart}`);
        }

        // Backgrounds
        let bg = catBg;
        if (row.rowType === 'TYPE') bg = typeBg;
        else if (row.rowType === 'GROUP') bg = groupBg;
        
        // Skip formatting logic for SPACER rows entirely (keep them white/empty)
        if (row.rowType === 'SPACER') {
            rangeStart++;
            continue;
        }

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

        rangeStart++;
    }

    // Apply Borders in Batch
    if (borderRanges.length > 0) {
        const rangeList = outputSheet.getRangeList(borderRanges);
        rangeList.setBorder(true, true, true, true, true, true, borderColor, SpreadsheetApp.BorderStyle.SOLID);
    }

    // --- 7. Update Summary & Cleanup ---
    updateSheetSummary(outputSheet);
    outputSheet.getRange('A1').activate();
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
        const label = row[1]; 
        if (label === "Income") diffinc = Number(row[3]) || diffinc; 
        if (label === "Expense") diffexp = Number(row[3]) || diffexp;
    }
    
    const overUnder = Number((diffexp - diffinc).toFixed(2)); 
    const summaryRange = outputSheet.getRange(5, 5);
    const statusRange = outputSheet.getRange(5, 6);

    if (overUnder === 0) {
        summaryRange.setValue('On Budget').setNumberFormat('@');
        statusRange.setValue('');
    } else {
        const absOverUnder = Math.abs(overUnder);
        if (overUnder < 0) {
            summaryRange.setValue(absOverUnder).setNumberFormat('$ #,###');
            statusRange.setValue('UNDER this month');
        } else {
            summaryRange.setValue(absOverUnder).setNumberFormat('$ #,###');
            statusRange.setValue('OVER this month');
        }
    }
    
    const dt = getDateTime();
    outputSheet.getRange('B3').setValue("Last updated on " + dt);
}

function preFormatMonthlyReportData(catGroupArr, transGroupArr) {

    const parseAmount = (val) => {
        if (typeof val === 'number') return val;
        if (typeof val === 'string') return parseFloat(val.replace(/[^0-9.-]/g, '')) || 0;
        return 0;
    };

    const actualMap = {}; 
    Object.keys(transGroupArr).forEach(group => {
        actualMap[group] = {};
        transGroupArr[group].forEach(item => {
            const cat = item.category || 'Unknown';
            const amt = parseAmount(item.amount); 
            actualMap[group][cat] = (actualMap[group][cat] || 0) + amt;
        });
    });
    
    const result = [];

    Object.keys(catGroupArr).forEach(group => {
        const catList = catGroupArr[group];
        const groupType = catList.length > 0 ? catList[0][2] : 'Unknown';

        const groupObj = {
            group: group,
            type: groupType, 
            actual: 0,
            budget: 0,
            categories: []
        };

        let actualTotal = 0;
        let budgetTotal = 0;

        catList.forEach(catRow => {
            const category = catRow[0]; 
            const budgetAmount = parseAmount(catRow[6] || 0); 
            const actualAmount = parseAmount(actualMap[group]?.[category] || 0);
            
            groupObj.categories.push({
                category: category,
                budget: budgetAmount, 
                amount: actualAmount
            });
            
            actualTotal += actualAmount;
            budgetTotal += budgetAmount;
        });

        groupObj.actual = actualTotal;
        groupObj.budget = budgetTotal;
        result.push(groupObj);
    });
    
    return result;
}

function getItemPercentage(budget, actual) {
    budget = Number(budget);
    actual = Number(actual);
    if (!isFinite(budget) || !isFinite(actual)) return 'N/A';
    if (budget === 0) return actual === 0 ? 0 : 'N/A';
    return actual / budget;
}

function getDateTime() {
    return new Date().toLocaleString("en-US", {
        month: 'short', day: 'numeric', year: 'numeric',
        hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true
    });
}
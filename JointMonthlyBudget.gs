/**
 * @fileoverview Dedicated script for generating the Joint Monthly Budget report.
 * Updated to remove cell borders and hide categories marked in the Categories sheet,
 * and implement new logic for splitting 'Household'/'Joint' transactions based on category ratios.
 */

// ======================================================
//        MAIN CONTROLLER: JOINT MONTHLY BUDGET
// ======================================================

/**
 * The main function to populate the 'Joint Monthly Budget' sheet.
 * It coordinates configuration reading, data processing, and sheet rendering.
 */
function populateJointMonthlyBudget() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const outputSheet = ss.getSheetByName("Joint Monthly Budget");
	
	if (!outputSheet) { 
		SpreadsheetApp.getUi().alert("Sheet 'Joint Monthly Budget' missing. Please ensure the sheet exists."); 
		return; 
	}

	// 1. Get Configuration & Setup
	const config = getJointConfig(ss);
	
	// If configuration fails (missing sheet or invalid setup), stop execution.
	if (!config) return; 

	// 2. Clear & Prepare Sheet Headers
	prepareJointSheet(outputSheet, config);

	// 3. Process Data
	// 3a. Read categories and calculate split budgets and create a ratio map for transaction lookups
	const { grouped: groupedCategories, categoryRatioMap } = getJointCategoryData(config); 
	
	// 3b. Read transactions and calculate split actuals based on Owner/Assigned Amount and ratios
	const actualMap = getJointTransactionData(config, categoryRatioMap); 
	
	// 4. Build the Output Layout (4-row blocks for Type, Group, and Category + Spacers)
	const layout = buildJointLayout(groupedCategories, actualMap);

	// 5. Render to Sheet
	if (layout.rows.length > 0) {
		renderJointSheet(outputSheet, layout);
	}

	// 6. Update Summary Header (Over/Under Budget)
	updateJointSheetSummary(outputSheet);
	
	// Finish by activating cell A1
	//outputSheet.getRange('A1').activate();
}


// ======================================================
//        HELPER FUNCTIONS
// ======================================================

/**
 * 1. Reads all configuration from the Definition, Categories, and Transactions sheets.
 * Includes validating dynamic transaction column settings.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @return {object|null} The configuration object or null on failure.
 */
function getJointConfig(ss) {
	const controlSheet = ss.getSheetByName('Definition');
	const catSheet = ss.getSheetByName("Categories");
	const tranSheet = ss.getSheetByName("Transactions");

	if (!controlSheet || !catSheet || !tranSheet) {
		SpreadsheetApp.getUi().alert("Missing required sheets: 'Definition', 'Categories', or 'Transactions'.");
		return null;
	}

	// C3:C12 contains general budget settings (0-based indexing in array)
	const budgetConfigRange = controlSheet.getRange('C3:C12').getValues().flat();
	const dateConfig = controlSheet.getRange('AD4').getValue();
	const monthColumns = controlSheet.getRange('W2:W12').getValues().flat();
	
	// I3:I12 contains transaction settings: Last Row, Last Col, and custom column numbers
	const tranConfigRange = controlSheet.getRange('I3:I12').getValues().flat(); 

	const dateValue = dateConfig instanceof Date ? dateConfig : new Date(dateConfig);
	const monthIndex = dateValue.getMonth(); 
	
	// Extracting configuration values
	const tranLstrw = tranConfigRange[0]; // I3: Transaction Last Row
	const tranColEnd = tranConfigRange[1]; // I4: Transaction Last Column (End of data range)

	// Custom Mappings (1-based column indices)
	const tranCol = {
		DATE: 2, // Fixed to Column B (2) for reading the data range starting at B2
		AMOUNT: Number(tranConfigRange[7]),  // I10 (Main transaction amount column)
		CATEGORY: Number(tranConfigRange[2]), // I5
		OWNER: Number(tranConfigRange[8]),   // I11
		ASSIGNED_AMT: Number(tranConfigRange[9]) // I12 (Amount assigned to the specific owner)
	};

	// --- VALIDATION AND ERROR CHECKING ---
	const START_COL_INDEX = 1; // Always start reading from Column A (Date)
	
	const columnKeys = ['CATEGORY', 'AMOUNT', 'OWNER', 'ASSIGNED_AMT'];
	let isValid = true;
	let errorMessage = "Transaction column configuration is invalid. Please check 'Definition' sheet cells I5, I10-I12. ";
	
	columnKeys.forEach(key => {
		const colValue = Number(tranCol[key]);
		// Column must be a number >= 1 and <= Last Col (I4)
		if (!colValue || colValue < START_COL_INDEX || colValue > tranColEnd) {
			errorMessage += `Column '${key}' (Value: ${colValue}) must be a number >= ${START_COL_INDEX} and <= Last Col (${tranColEnd}). `;
			isValid = false;
		}
	});

	if (!isValid) {
		Logger.log(errorMessage);
		SpreadsheetApp.getUi().alert("Error: " + errorMessage);
		return null; 
	}
	// -------------------------------------

	// Return the final configuration object
	return {
		ss: ss,
		catSheet: catSheet,
		tranSheet: tranSheet,
		dateValue: dateValue,
		targetMonth: monthIndex,
		targetYear: dateValue.getFullYear(),
		
		// CRITICAL: Trim names for accurate matching against transaction data.
		name1: budgetConfigRange[8].toString().trim(), // C11
		name2: budgetConfigRange[9].toString().trim(), // C12
		
		// Category Columns
		catLstrw: budgetConfigRange[0], // C3
		catMonthBudgetCol: monthColumns[monthIndex] - 1, // 0-based column index for the budget amount
		
		// Transaction Configuration (I3-I12)
		tranLstrw: tranLstrw,
		tranColEnd: tranColEnd, 
		tranCol: tranCol, // 1-based indices for reading
		
		// Category Sheet Mappings (0-based column indices)
		mappings: {
			category: budgetConfigRange[2] - 1, // C5
			group: budgetConfigRange[3] - 1,    // C6
			type: budgetConfigRange[4] - 1,     // C7
			hide: budgetConfigRange[5] - 1,     // C8 <--- This is the hide column index
			name1Ratio: budgetConfigRange[6] - 1, // C9
			name2Ratio: budgetConfigRange[7] - 1  // C10
		}
	};
}

/**
 * 2. Clears the output sheet below the header and sets up partner names.
 * Fixes the breakApart error by ensuring the clearance range is large enough.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} outputSheet The sheet to prepare.
 * @param {object} config The configuration object.
 */
function prepareJointSheet(outputSheet, config) {
	const lastRow = outputSheet.getLastRow();
	const lastCol = outputSheet.getLastColumn();
	// Start clearing from row 9. Set a large fixed height (e.g., 2000 rows) 
	// to ensure all previous merged cells are fully encompassed.
	const clearHeight = lastRow >= 9 ? Math.max(lastRow - 8, 2000) : 2000; 
	
	// Start range from row 9, column 1
	const startRow = 9;
	
	// Only attempt to clear if there are rows present in the sheet
	if (lastRow >= startRow) {
		const range = outputSheet.getRange(startRow, 1, clearHeight, lastCol);
		
		// Attempt to break apart the merges first. Using a large range height
		// should prevent the 'must select all cells in a merged range' exception.
		try {
			 range.breakApart(); 
		} catch (e) {
			 Logger.log("Could not breakApart a range, continuing with clear: " + e.toString());
		}
		
		// Define the actual data range (or a large section of it) for explicit clearing
		const clearRange = outputSheet.getRange(startRow, 1, Math.max(lastRow - 8, 2000), lastCol);
		
		// Clear contents and formatting (including background, borders, number format)
		clearRange.clear({ contentsOnly: true, formatOnly: false });
		
		// *** FIX: Explicitly remove all borders using setBorder(false, false, ...) ***
		// This is necessary to ensure no border remnants remain, especially from previous formatting runs.
		clearRange.setBorder(false, false, false, false, false, false);
		
		// Explicitly clear number/background/font formats on the potentially larger range
		clearRange.setNumberFormat(null).setBackground(null).setFontColor(null);
	}
	
	// Set a white font color for column A (hidden helper column)
	outputSheet.getRange("A:A").setFontColor("#FFFFFF");
	
	// Set Correct Names in Header (D8 and E8)
	outputSheet.getRange('D8').setValue(config.name1);
	outputSheet.getRange('E8').setValue(config.name2);
}

/**
 * 3a. Reads Categories data, calculates the split budget based on ratios,
 * and groups the results by Type and Group, while skipping hidden categories.
 * Also returns a map of normalized ratios keyed by category name.
 * * @param {object} config The configuration object.
 * @return {object} An object containing { grouped: Nested Category Data, categoryRatioMap: Ratios by Category Name }.
 */
function getJointCategoryData(config) {
	// Read all category data (starting from row 1 to get headers if needed, but skip row 1 in loop)
	const data = config.catSheet.getRange(1, 1, config.catLstrw, config.catSheet.getLastColumn()).getValues();
	const catAry = [];
	const map = config.mappings;
	const categoryRatioMap = {}; // New map for ratios

	for (let l = 1; l < config.catLstrw; l++) {
		const row = data[l];
		const categoryName = row[map.category]?.toString().trim() || '';
		// Check if the hide column value is truthy (e.g., 'X', 'TRUE', 1, or just not empty)
		const isHidden = !!row[map.hide]; 

		// Apply point 2: Check hide column and do not process/show.
		if (!categoryName || isHidden) { 
			continue;
		}
		
		// Budget column index is 0-based based on the configuration
		const totalBudget = Number(row[config.catMonthBudgetCol]) || 0;
		
		// Ratios are 0-based column indices read from config
		const ratio1 = Number(row[map.name1Ratio]) || 0;
		const ratio2 = Number(row[map.name2Ratio]) || 0;
		
		// Ensure ratios sum to 1 before splitting. If not, normalize or assume 50/50 if both are 0.
		let normRatio1 = ratio1;
		let normRatio2 = ratio2;
		const ratioSum = ratio1 + ratio2;

		if (ratioSum > 0 && ratioSum !== 1) {
			normRatio1 = ratio1 / ratioSum;
			normRatio2 = ratio2 / ratioSum;
		} else if (ratioSum === 0) {
			// If both ratios are 0, we can assume a 50/50 split as a fallback
			normRatio1 = 0.5;
			normRatio2 = 0.5;
		}

		catAry.push({
			category: categoryName, // Use the trimmed name
			group: row[map.group],
			type: row[map.type],
			hide: isHidden, // Although we filter it out, keep for completeness
			budgetP1: totalBudget * normRatio1,
			budgetP2: totalBudget * normRatio2,
			budgetTotal: totalBudget
		});

		// Store the normalized ratios for transaction lookup
		categoryRatioMap[categoryName] = {
			ratio1: normRatio1,
			ratio2: normRatio2
		};
	}

	// Group the results by Type -> Group
	const grouped = {};
	catAry.forEach(item => {
		if (!grouped[item.type]) grouped[item.type] = {};
		if (!grouped[item.type][item.group]) grouped[item.type][item.group] = [];
		grouped[item.type][item.group].push(item);
	});
	
	return { grouped, categoryRatioMap }; // Return both structures
}

/**
 * 3b. Reads Transactions data for the target month/year and applies 
 * custom split logic for actual spending based on Owner and Assigned Amount,
 * incorporating the new logic for 'Household'/'Joint' transactions using category ratios.
 * * @param {object} config The configuration object.
 * @param {object} categoryRatioMap Map of category names to their split ratios.
 * @return {object} Map of category names to actual split spending {p1: amount, p2: amount}.
 */
function getJointTransactionData(config, categoryRatioMap) {
	const lastRw = config.tranLstrw;
	const lastCol = config.tranColEnd;
	const START_COL_INDEX = 1; // Column A
	const numColumns = lastCol - START_COL_INDEX + 1; 

	// Read all necessary data, starting from row 2
	const data = config.tranSheet.getRange(2, START_COL_INDEX, lastRw - 1, numColumns).getValues(); 
	
	const actualMap = {};

	// Calculate the 0-based index in the 'data' array for each required column 
	// by subtracting the starting column index (1 for Column A)
	const COL_TRAN = { 
		DATE: config.tranCol.DATE - START_COL_INDEX,
		CATEGORY: config.tranCol.CATEGORY - START_COL_INDEX, 
		AMOUNT: config.tranCol.AMOUNT - START_COL_INDEX, 
		OWNER: config.tranCol.OWNER - START_COL_INDEX,
		ASSIGNED_AMT: config.tranCol.ASSIGNED_AMT - START_COL_INDEX 
	};

	data.forEach(row => {
		const dateCell = row[COL_TRAN.DATE];
		// Basic date validation
		let d = (dateCell instanceof Date) ? dateCell : new Date(dateCell);
		if (isNaN(d.getTime())) return;

		// Filter by month and year
		if (d.getMonth() === config.targetMonth && d.getFullYear() === config.targetYear) {

			const parseAmount = (val) => {
				if (typeof val === 'number') return val;
				// Safely parse string amount
				if (typeof val === 'string') return parseFloat(val.toString().replace(/[^0-9.-]/g, '')) || 0;
				return 0;
			};

			const cat = (row[COL_TRAN.CATEGORY] || 'Unknown').toString().trim();
			// rawAmt is the value from the main AMOUNT column (I10 config)
			const rawAmt = parseAmount(row[COL_TRAN.AMOUNT] ) || 0;
			
			// Trim and check owner name for reliable matching
			const owner = row[COL_TRAN.OWNER] ? row[COL_TRAN.OWNER].toString().trim() : '';

			let amtP1 = 0;
			let amtP2 = 0;
			
			const name1 = config.name1; 
			const name2 = config.name2;
			
			// Assigned Amount is the value from the Assigned Amount column (I12 config)
			const assignedAmt = parseAmount( row[COL_TRAN.ASSIGNED_AMT] ) || 0;
			const totalAmt = Math.abs(rawAmt); // Absolute value of the transaction amount

			// --- NEW LOGIC: Household/Joint split by Category Ratios ---
			/*if (owner === 'Household' || owner === 'Joint') {
				const ratios = categoryRatioMap[cat];
				if (ratios) {
					// Calculate split based on category ratios using the main AMOUNT, as requested.
					// Note: Ratios are normalized (sum to 1) from getJointCategoryData.
					amtP1 = totalAmt * ratios.ratio1;
					amtP2 = totalAmt * ratios.ratio2;
				} else {
					// Fallback if category ratios aren't found for Household/Joint (e.g., category missing from Definitions)
					// Use existing 50/50 split logic for safety.
					amtP1 = totalAmt / 2;
					amtP2 = totalAmt / 2;
				}
			}*/

			// --- EXISTING LOGIC: Individual Owner uses Assigned Amount ---
			if (owner === name1) {
				// Owner is Name1: Name1 takes the Assigned Amount
				amtP1 = Math.abs(assignedAmt); 
				amtP2 = 0; 
			} else if (owner === name2) {
				// Owner is Name2: Name2 takes the Assigned Amount
				amtP2 = Math.abs(assignedAmt);
				amtP1 = 0;
			} else if(owner === 'Household' || owner === 'Joint'){
        // --- EXISTING LOGIC: Other/Blank Owner splits main AMOUNT 50/50 --- 
        amtP1 = Math.abs(assignedAmt);
        amtP2 = Math.abs(assignedAmt);
			}else{
        // Owner is blank, or some other value: Split the *main* AMOUNT column 50/50.
        amtP1 = totalAmt / 2;
        amtP2 = totalAmt / 2;
      }

			if (!actualMap[cat]) actualMap[cat] = { p1: 0, p2: 0 };
			
			actualMap[cat].p1 += amtP1;
			actualMap[cat].p2 += amtP2;
		}
	});
	
	return actualMap;
}

/**
 * 4. Constructs the output structure (rows, styles, merges) 
 * for rendering the report with summaries for Types and Groups.
 * Removed border logic.
 * @param {object} groupedData Category data grouped by Type and Group.
 * @param {object} actualMap Actual spending data per category.
 * @return {object} The layout structure.
 */
function buildJointLayout(groupedData, actualMap) {
	const rows = [];
	const styles = [];
	const merges = [];
	// Borders array removed as requested
	
	// Starting row in the sheet for the output data
	// We start rendering at sheet row 9, but our array index (r) starts at 0.
	let r = 0; 

	/**
	 * Helper to create a 4-row data block (Budget, Actual, Diff, %).
	 */
	function addDataBlock(name, budP1, budP2, budTot, actP1, actP2, actTot, rowIdx, isIncomeType, typeLabel) {
		// Diff Calculation
		const p1Diff = isIncomeType ? (actP1 - budP1) : (budP1 - actP1);
		const p2Diff = isIncomeType ? (actP2 - budP2) : (budP2 - actP2);
		const totDiff = isIncomeType ? (actTot - budTot) : (budTot - actTot);

		// Percentage Calculation
		const p1Pct = getJointItemPercentage(budP1, actP1);
		const p2Pct = getJointItemPercentage(budP2, actP2);
		const totPct = getJointItemPercentage(budTot, actTot);

		// Sheet Row Index (Sheet index starts at 1, header is 8 rows, so output starts at 9. 
		// Array index 0 corresponds to sheet row 9.)
		const sheetRow = 9 + rowIdx; 
		
		// Rows for the block (Cols A-F)
		rows.push([null, name, "Budget", budP1, budP2, budTot]);
		styles.push({ type: typeLabel + '_BUDGET', row: sheetRow });
		
		rows.push([null, null, "Actual", actP1, actP2, actTot]);
		styles.push({ type: typeLabel + '_ACTUAL', row: sheetRow + 1 });
		
		rows.push([null, null, "Diff", p1Diff, p2Diff, totDiff]);
		styles.push({ type: typeLabel + '_DIFF', row: sheetRow + 2 });
		
		rows.push([null, null, "%", p1Pct, p2Pct, totPct]);
		styles.push({ type: typeLabel + '_PERCENT', row: sheetRow + 3 });

		// Merge Name Column (Sheet Col B) across 4 rows.
		// Use sheetRow index for merges
		merges.push({row: sheetRow, col: 2, numRows: 4, numCols: 1});
		
		// Border push removed
	}

	// Ensure Income comes before Expense
	const typeOrder = ['Income', 'Expense', ...Object.keys(groupedData).filter(k => k !== 'Income' && k !== 'Expense')];

	typeOrder.forEach(type => {
		const groups = groupedData[type];
		if (!groups) return;

		const isInc = (type === 'Income'); // Determines the sign for Diff calculation

		// --- CALCULATE TYPE TOTALS ---
		let tBudP1=0, tBudP2=0, tBudTot=0;
		let tActP1=0, tActP2=0, tActTot=0;
		
		Object.keys(groups).forEach(gName => {
			groups[gName].forEach(c => {
				// Since categories were filtered in getJointCategoryData, we don't need to filter here.
				tBudP1 += c.budgetP1; tBudP2 += c.budgetP2; tBudTot += c.budgetTotal;
				tActP1 += (actualMap[c.category]?.p1 || 0);
				tActP2 += (actualMap[c.category]?.p2 || 0);
			});
		});
		tActTot = tActP1 + tActP2;

		// 1. Render TYPE Summary Block (Highest Level)
		addDataBlock(type, tBudP1, tBudP2, tBudTot, tActP1, tActP2, tActTot, r, isInc, 'TYPE');
		r += 4; // Advance array index by 4 rows

		// Process Groups
		Object.keys(groups).forEach(groupName => {
			const categories = groups[groupName];
			
			// --- CALCULATE GROUP TOTALS ---
			let gBudP1=0, gBudP2=0, gBudTot=0;
			let gActP1=0, gActP2=0, gActTot=0;
			
			categories.forEach(c => {
				gBudP1 += c.budgetP1; gBudP2 += c.budgetP2; gBudTot += c.budgetTotal;
				gActP1 += (actualMap[c.category]?.p1 || 0);
				gActP2 += (actualMap[c.category]?.p2 || 0);
			});
			gActTot = gActP1 + gActP2;

			// 2. Render GROUP Summary Block (Intermediate Level)
			addDataBlock(groupName, gBudP1, gBudP2, gBudTot, gActP1, gActP2, gActTot, r, isInc, 'GROUP');
			r += 4; // Advance array index by 4 rows

			// 3. Render CATEGORIES (Lowest Level)
			categories.forEach(cat => {
				const cActP1 = actualMap[cat.category]?.p1 || 0;
				const cActP2 = actualMap[cat.category]?.p2 || 0;
				const cActTot = cActP1 + cActP2;
				
				addDataBlock(cat.category, cat.budgetP1, cat.budgetP2, cat.budgetTotal, cActP1, cActP2, cActTot, r, isInc, 'CAT');
				r += 4; // Advance array index by 4 rows
			});

			// 4. ADD SPACER ROW after the last category of the group
			rows.push([null, null, null, null, null, null]); // Push empty row data
			r += 1; // Advance array index by 1 row
		});
	});

	// Removed 'borders' from the return object
	return { rows, styles, merges };
}

/**
 * 5. Writes the generated data and applies formatting (colors, number formats, merging).
 * Now includes border application for Group and Category summary rows.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The output sheet.
 * @param {object} layout The layout structure containing data and formatting instructions.
 */
function renderJointSheet(sheet, layout) {
	// Write data starting from row 9, column A (1)
	sheet.getRange(9, 1, layout.rows.length, 6).setValues(layout.rows);

	// Color definitions
	const C_TYPE_BLOCK = '#e68e68'; // Warm Orange/Brown for Type Summary
	const C_GRP_BLOCK = '#eec49f';  // Light Tan/Peach for Group Summary
	const C_CAT_BLOCK = '#FFFFFF';      
	const FMT_NUM = '_($* #,##0.00_);_($* (#,##0.00)_);_($* "-"_);_(@_)';
	const FMT_PCT = '0.00%';
	
	// Border definitions for the new logic
	const borderStyle = SpreadsheetApp.BorderStyle.SOLID; 
	const BORDER_COLOR = '#444444'; 
	
	// Apply Formatting (Styles)
	layout.styles.forEach(s => {
		
		let bg = C_CAT_BLOCK;
		
		// Determine background color based on block level
		if (s.type.includes('TYPE')) {
			bg = C_TYPE_BLOCK;
		} else if (s.type.includes('GROUP')) {
			bg = C_GRP_BLOCK;
		} else if (s.type.includes('CAT')) {
			bg = C_CAT_BLOCK;
		}
		
		// Apply background to the data columns (Col B to F)
		sheet.getRange(s.row, 2, 1, 5).setBackground(bg);

		// Alignment and Labels for all rows
		// Name (Col B) alignment: left
		sheet.getRange(s.row, 2).setHorizontalAlignment('left'); 
		// Label (Col C) and Numbers (Col D-F) alignment: right
		sheet.getRange(s.row, 3).setHorizontalAlignment('right'); 
		sheet.getRange(s.row, 4, 1, 3).setHorizontalAlignment('right'); 

		// Formats: Apply to numeric columns (Col D to F)
		if (s.type.includes('PERCENT')) {
			sheet.getRange(s.row, 4, 1, 3).setNumberFormat(FMT_PCT);
		} else {
			sheet.getRange(s.row, 4, 1, 3).setNumberFormat(FMT_NUM);
		}

		// Bold and Vertical Alignment for the main identifier row (Budget row)
		if (s.type.includes('BUDGET')) {
			// Bold the name cell text (Col B) for all levels and set top vertical alignment
			sheet.getRange(s.row, 2).setFontWeight('bold').setVerticalAlignment('top');
			
			// Bold the Label for the budget row (Col C)
			sheet.getRange(s.row, 3).setFontWeight('bold');
		} else {
			// Re-assert labels for Actual, Diff, and % rows in Col C 
			// The row data array might contain null in Col C due to anticipating the merge
			// but we must ensure the label text is there before merging.
			const labelRange = sheet.getRange(s.row, 3);
			if (s.type.includes('ACTUAL')) labelRange.setValue("Actual");
			if (s.type.includes('DIFF')) labelRange.setValue("Diff");
			if (s.type.includes('PERCENT')) labelRange.setValue("%");
		}
	});

	// Apply Merges
	// NOTE: This MUST happen AFTER setting the content of all cells, 
	// especially Col C (Budget, Actual, Diff, %)
	layout.merges.forEach(m => sheet.getRange(m.row, m.col, m.numRows, m.numCols).merge());

	// ------------------------------------------------------------------
	// FIX: Border Application Logic (Modified to handle merged cell in Col B)
	// ------------------------------------------------------------------
	layout.styles.forEach(s => {
		// Target the last row of the 4-row block (the PERCENT row)
		if (s.type.includes('_PERCENT')) {
			// Check if the block is a GROUP or a CAT (Category)
			const isGroupOrCategory = s.type.includes('GROUP') || s.type.includes('CAT');
			
			if (isGroupOrCategory) {
				
				// 1. FIX for Column B (merged cell): Apply border to the full 4-row merged range.
				// The merged range starts 3 rows above the current 's.row'.
				sheet.getRange(s.row - 3, 2, 4, 1).setBorder(
					null,  // Top
					null,  // Left
					true,  // BOTTOM 
					null,  // Right
					null,  // Vertical
					null,  // Horizontal
					BORDER_COLOR, 
					borderStyle
				);
				
				// 2. Apply border to Columns C-F (single row range)
				sheet.getRange(s.row, 3, 1, 4).setBorder(
					null,  // Top
					null,  // Left
					true,  // BOTTOM 
					null,  // Right
					null,  // Vertical
					null,  // Horizontal
					BORDER_COLOR, 
					borderStyle
				);
			}
		}
	});
	// ------------------------------------------------------------------
}

/**
 * 6. Calculates the overall total budget surplus or deficit and updates the summary 
 * cells (B3, E5, F5) in the header section of the sheet.
 * * IMPORTANT UPDATE: This now calculates the Total Surplus/Deficit based on 
 * Total Actual Income minus Total Actual Expense, as requested.
 * * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The output sheet.
 */
function updateJointSheetSummary(sheet) {
	const lastRow = sheet.getLastRow();
	if (lastRow < 9) return;
	
	// Read relevant columns: B (Type/Group/Category Name), C (Label), F (Total)
	// Read range starts at row 9, column B (2), covering 5 columns (B, C, D, E, F)
	const data = sheet.getRange(9, 2, lastRow - 8, 5).getValues(); 
	
	let totalActualIncome = 0;
	let totalActualExpense = 0;
	
	// Find the row corresponding to the 'Type' summary blocks and extract the Actual Total
	for (let i = 0; i < data.length; i++) {
		// data[] columns: [Col B (Name), Col C (Label), Col D, Col E, Col F (Total Amount)]
		const [colB, colC, , , colF] = data[i]; 
		
		// The top-level Type block is defined by the Name in Col B and the 'Budget' label in Col C.
		if (colB && colC === 'Budget') { 
			const typeName = colB.toString().trim();

			if (typeName === 'Income' || typeName === 'Expense') {
				// The next row (i+1) is the 'Actual' row for this Type Block.
				// This row contains the grand total actual in Col F (index 4 in the 5-column array).
				if (i + 1 < data.length) {
					const actualTotal = Number(data[i + 1][4]) || 0; 
					
					if (typeName === 'Income') {
						totalActualIncome = actualTotal;
					} else if (typeName === 'Expense') {
						totalActualExpense = actualTotal;
					}
				}
			}
		}
	}
	
	// Use variable names matching the user's requested formula structure for clarity:
	const diffinc = totalActualIncome;
	const diffexp = totalActualExpense;
	
	// Calculate Total Surplus/Deficit based on Actual Income - Actual Expense
	// Formula: Number((diffinc - diffexp).toFixed(2));
	const totalSurplus = Number((diffinc - diffexp).toFixed(2));

	const summaryRange = sheet.getRange(5, 5); // E5
	const statusRange = sheet.getRange(5, 6); // F5
	
	// Update summary cells
	if (Math.abs(totalSurplus) < 0.01) {
		summaryRange.setValue("On Budget").setNumberFormat("@").setFontColor("black");
		statusRange.setValue("");
	} else {
		summaryRange.setValue(Math.abs(totalSurplus)).setNumberFormat('$ #,##0.00'); 
		if (totalSurplus > 0) {
			summaryRange.setFontColor("black");
			// Status updated to reflect Cash Flow SURPLUS
			statusRange.setValue("Under Budget").setFontColor("green"); 
		} else {
			summaryRange.setFontColor("red");
			// Status updated to reflect Cash Flow DEFICIT
			statusRange.setValue("Over Budget").setFontColor("red"); 
		}
	}
	
	// Update last updated timestamp
	sheet.getRange('B3').setValue("Last updated on " + getDateTime());
}

/**
 * Calculates the percentage of actual spending relative to budget.
 * Returns 0 if both are 0, 'N/A' if budget is 0 but actual is not, otherwise actual/budget.
 * @param {number} budget The budgeted amount.
 * @param {number} actual The actual spent amount.
 * @return {number|string} The percentage or 'N/A'.
 */
function getJointItemPercentage(budget, actual) {
	budget = Number(budget);
	actual = Number(actual);
	if (!isFinite(budget) || !isFinite(actual)) return 'N/A';
	
	if (budget === 0) {
		return actual === 0 ? 0 : 'N/A'; // Cannot divide by zero, but zero actual / zero budget is 0% used.
	}
	return actual / budget;
}
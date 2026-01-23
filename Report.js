function regenerateNetWorthReports(){
  // Net Worth Reports
  if (typeof populateNetWorth === 'function') {
    populateNetWorth();
  }
  if (typeof populateJointNetWorth === 'function') {
    populateJointNetWorth();
  }
}

/**
 * Master function that calls all individual report generation functions.
 * You MUST ensure these functions exist in your script project.
 */
function regenerateAllReports() {

  Logger.log("Starting full report regeneration...");

  // Monthly Reports
  if (typeof populateMonthlyBudget === 'function') {
    populateMonthlyBudget();
  }
  if (typeof populateJointMonthlyBudget === 'function') {
    populateJointMonthlyBudget();
  }

  // Yearly Reports
  if (typeof populateYearlyBudget === 'function') {
    populateYearlyBudget();
  }
  // This is the function we corrected earlier
  if (typeof populateJointYearlyBudget === 'function') {
    populateJointYearlyBudget();
  }

  // Net Worth Reports
  if (typeof populateNetWorth === 'function') {
    populateNetWorth();
  }
  if (typeof populateJointNetWorth === 'function') {
    populateJointNetWorth();
  }

  Logger.log("Report regeneration complete.");
}

function confirmPopulateReportToSheets(){

  try{
    populateNetWorth();
    populateJointNetWorth();
    populateMonthlyBudget();
    populateJointMonthlyBudget();
    populateYearlyBudget();
    populateJointYearlyBudget();
    return true;
  }catch(error){
    Logger.log( JSON.stringify("generate report error: "+ error) );
    return false;
  }
}
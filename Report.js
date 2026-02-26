/**
 * Master function that calls all individual report generation functions.
 */
function regenerateAllReports() {
  Logger.log("Starting full report regeneration...");

  if (typeof populateNetWorth === 'function') {
    populateNetWorth();
  }
  if (typeof populateJointNetWorth === 'function') {
    populateJointNetWorth();
  }
  if (typeof populateMonthlyBudget === 'function') {
    populateMonthlyBudget();
  }
  if (typeof populateJointMonthlyBudget === 'function') {
    populateJointMonthlyBudget();
  }
  if (typeof populateYearlyBudget === 'function') {
    populateYearlyBudget();
  }
  if (typeof populateJointYearlyBudget === 'function') {
    populateJointYearlyBudget();
  }

  Logger.log("Report regeneration complete.");
}

function startGenerationOfMonthlyBudget(){
  try{
    populateMonthlyBudget();
  }catch(error){
    SpreadsheetApp.getUi().alert('Something went wrong while generating monthly budget. Please try again.');
    Logger.log("Error in startGenerationOfMonthlyBudget: " + error.toString());
  }
}

function startGenerationOfJointMonthlyBudget(){
  try{
    populateJointMonthlyBudget();
  }catch(error){
    SpreadsheetApp.getUi().alert('Something went wrong while generating joint monthly budget. Please try again.');
    Logger.log("Error in startGenerationOfJointMonthlyBudget: " + error.toString());
  }
}

function startGenerationOfYearlyBudget(){
  try{
    populateYearlyBudget();
  }catch(error){
    SpreadsheetApp.getUi().alert('Something went wrong while generating yearly budget. Please try again.');
    Logger.log("Error in startGenerationOfYearlyBudget: " + error.toString());
  }
}

function startGenerationOfJointYearlyBudget(){
  try{
    populateJointYearlyBudget();
  }catch(error){
    SpreadsheetApp.getUi().alert('Something went wrong while generating joint yearly budget. Please try again.');
    Logger.log("Error in startGenerationOfJointYearlyBudget: " + error.toString());
  }
}

function startReportGenerationTask(){
  try{
    regenerateAllReports();
  }catch(error){
    SpreadsheetApp.getUi().alert('Something went wrong while regenerating reports. Please try again.');
    Logger.log("Error in startReportGenerationTask: " + error.toString());
  }
}

function generateFinancialReport(){
  try{
    SpreadsheetApp.getUi().alert('Financial report generation has started. Do not change anything until the process completes.');
    regenerateAllReports();
    return {
      success: true,
      message: 'Financial report generated successfully.'
    };
  }catch(error){
    SpreadsheetApp.getUi().alert('Please make sure that all the sheets are configured for report generate.');
    Logger.log("Error in generateFinancialReport: " + error.toString());
    return {
      success: false,
      message: error.toString()
    };
  }
}

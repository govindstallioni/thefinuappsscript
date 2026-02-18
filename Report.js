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

  // Net Worth Reports
  if (typeof populateNetWorth === 'function') {
    populateNetWorth();
  }
  if (typeof populateJointNetWorth === 'function') {
    populateJointNetWorth();
  }
  
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


function startGenerationOfMonthlyBudget(){
  try{
    //SpreadsheetApp.getUi().alert('Monthly budget generation has started. Do not change anything until the process completes.');
    populateMonthlyBudget();
  }catch(error){
    SpreadsheetApp.getUi().alert('Something went wrong while generating monthly budget. Please try again.');
    Logger.log("Error in startGenerationOfMonthlyBudget: " + error.toString());
  }
}

function startGenerationOfJointMonthlyBudget(){
  try{
    //SpreadsheetApp.getUi().alert('Joint monthly budget generation has started. Do not change anything until the process completes.');
    populateJointMonthlyBudget();
  }catch(error){
    SpreadsheetApp.getUi().alert('Something went wrong while generating joint monthly budget. Please try again.');
    Logger.log("Error in startGenerationOfJointMonthlyBudget: " + error.toString());
  }
}

function startGenerationOfYearlyBudget(){
  try{
    //SpreadsheetApp.getUi().alert('Yearly budget generation has started. Do not change anything until the process completes.');
    populateYearlyBudget();
  }catch(error){
    SpreadsheetApp.getUi().alert('Something went wrong while generating yearly budget. Please try again.');
    Logger.log("Error in startGenerationOfYearlyBudget: " + error.toString());
  }
}

function startGenerationOfJointYearlyBudget(){
  try{
    //SpreadsheetApp.getUi().alert('Joint yearly budget generation has started. Do not change anything until the process completes.');
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
  }finally{
    SpreadsheetApp.getUi().alert('Reports regenerated successfully.');
  }
}

function generateFinancialReport(){
  try{
    SpreadsheetApp.getUi().alert('Financial report generation has started. Do not change anything until the process completes.');
    regenerateAllReports();
    return {
      status: true,
      message: 'Financial report generated successfully.'
    };
  }catch(error){
    SpreadsheetApp.getUi().alert('Please make sure that all the sheets are configured for report generate.');
    Logger.log("Error in generateFinancialReport: " + error.toString());
    return {
      status: false,
      message: error.toString()
    };
  }
}
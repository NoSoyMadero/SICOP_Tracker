// =================================================================================================
// MAIN ORCHESTRATOR
// This file acts as the entry point for time-driven triggers and main execution flows.
// =================================================================================================

/**
 * Primary execution function for the SICOP processing pipeline.
 * Fetches emails, parses data, and hands off to the AI batch processor.
 */
function main() {
  Logger.log('Starting SICOP Pipeline...');
  
  Logger.log('Fetching and parsing emails from Gmail...');
  processSicopEmails();
  
  Logger.log('Evaluating pending tenders with GenAI...');
  processAiBatches();
  
  Logger.log('Pipeline Execution Complete.');
}

/**
 * Orchestrates the AI evaluation across all time-bound sheets.
 * Uses a shared cache to prevent duplicate GenAI API calls for the same Message ID.
 */
function processAiBatches() {
  // Shared memory map: { "Message_ID": "AI_Summary" }
  const sharedCache = {}; 
  
  batchProcessEmails(CONFIG.SHEETS.MONTHLY, sharedCache);
  batchProcessEmails(CONFIG.SHEETS.WEEKLY, sharedCache);
  batchProcessEmails(CONFIG.SHEETS.TODAY, sharedCache);
}

// =================================================================================================
// DEVELOPER TOOLS
// Utility functions for maintenance, testing, and resetting the environment.
// =================================================================================================

/**
 * HARD RESET: Deletes and rebuilds all core tracking sheets.
 * WARNING: This will permanently wipe all current extracted data (but leaves Tracked_Tenders intact).
 */

function devResetSpreadsheet() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetId = scriptProperties.getProperty('Spreadsheet_ID');
  
  if (!sheetId) {
    Logger.log('Error: Spreadsheet ID not found in script properties.');
    return;
  }
  
  const ss = SpreadsheetApp.openById(sheetId);
  const targetSheets = [CONFIG.SHEETS.TODAY, CONFIG.SHEETS.WEEKLY, CONFIG.SHEETS.MONTHLY];
  
  Logger.log('Initiating Developer Wipe...');

  // 1. Create a temporary sheet to prevent Google Sheets from crashing 
  // (A spreadsheet must always have at least one visible sheet).
  let tempSheet = ss.getSheetByName('Temp_Dev_Wipe');
  if (!tempSheet) {
    tempSheet = ss.insertSheet('Temp_Dev_Wipe');
  }
  
  // 2. Delete the target operational sheets
  targetSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ss.deleteSheet(sheet);
      Logger.log(`Deleted sheet: ${sheetName}`);
    }
  });
  
  // 3. Re-create a pristine baseline sheet 
  const primarySheet = ss.insertSheet(CONFIG.SHEETS.TODAY);
  
  // THE FIX: Standardized header naming to match Spreadsheets.gs ("Technical Specs" instead of "Body")
  primarySheet.appendRow(['Date', 'Tender Number', 'Subject', 'Technical Specs', 'AI Summary', 'Message ID']);
  applyCleanFormatting(primarySheet);
  
  // 4. Clean up the temporary sheet
  ss.deleteSheet(tempSheet);
  
  Logger.log('Dev wipe complete. Spreadsheet reset to a clean state.');
}
// 1. Pass the sheet object directly instead of sheetId/sheetName
function saveToSheet(sheet, data) {
  // Ensure headers exist if the sheet is brand new
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date', 'Tender Number', 'Subject', 'Technical Specs', 'AI Summary', 'Message ID']);
  }
  
  const lastRow = sheet.getLastRow();
  const validIds = new Set(data.map(row => row[5])); 
  
  // 2. Separate the "real" incoming updates from the "EXISTING" placeholders
  const incomingDataMap = new Map();
  data.forEach(row => {
    if (row[0] !== "EXISTING") incomingDataMap.set(row[5], row);
  });
  
  const finalData = [];
  
  // 3. Process existing sheet data IN MEMORY
  if (lastRow > 1) {
    const currentData = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    
    currentData.forEach(row => {
      const msgId = row[5];
      
      // If the ID is still valid for this timeframe (Today/Weekly/Monthly)
      if (validIds.has(msgId)) {
        if (incomingDataMap.has(msgId)) {
          // It's an update: push the NEW row and remove it from the map
          finalData.push(incomingDataMap.get(msgId));
          incomingDataMap.delete(msgId);
        } else {
          // It's unchanged: keep the OLD row
          finalData.push(row);
        }
      }
    });
  }
  
  // 4. Add any brand-new emails that weren't in the sheet at all
  incomingDataMap.forEach(newRow => finalData.push(newRow));
  
  // 5. Sort the array in memory by Date (Column index 0) descending
  finalData.sort((a, b) => new Date(b[0]) - new Date(a[0]));
  
  // 6. Write to sheet in ONE massively fast operation
  if (lastRow > 1) {
    // Clear everything except the header
    sheet.getRange(2, 1, lastRow - 1, 6).clearContent(); 
  }
  
  if (finalData.length > 0) {
    sheet.getRange(2, 1, finalData.length, 6).setValues(finalData);
  }
  
  applyCleanFormatting(sheet);
}

function applyCleanFormatting(sheet) {
  const lastRow = sheet.getLastRow();
  
  // Format Header Row
  sheet.getRange(1, 1, 1, 6)
       .setFontWeight("bold")
       .setBackground("#e0e0e0")
       .setHorizontalAlignment("center")
       .setVerticalAlignment("middle");
       
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 15); 
  
  // Set Column Widths efficiently
  const widths = [90, 180, 250, 150, 400];
  widths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));
  sheet.hideColumns(6); 
  
  // Format Data Rows (Only if data exists)
  if (lastRow > 1) {
    const numDataRows = lastRow - 1;
    const dataRange = sheet.getRange(2, 1, numDataRows, 6);
    
    dataRange.setVerticalAlignment("top").setHorizontalAlignment("left");
    
    // Set Text Wrapping
    sheet.getRange(2, 3, numDataRows, 2).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // Subject & Specs
    sheet.getRange(2, 5, numDataRows, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // Summary
    
    sheet.setRowHeights(2, numDataRows, 45);
  }
}


//--------------------------------------------------------------------------------------------------------------------------------------------------
// Funciones del Tracker CRM
//--------------------------------------------------------------------------------------------------------------------------------------------------

function appendToTracked(tenderData) {
  const sheetId = PropertiesService.getScriptProperties().getProperty('Spreadsheet_ID');
  if (!sheetId) return false;
  
  const ss = SpreadsheetApp.openById(sheetId);
  let sheet = ss.getSheetByName("Tracked_Tenders");
  
  // Create the tracking sheet on the fly if it's the first time
  if (!sheet) {
    sheet = ss.insertSheet("Tracked_Tenders");
    sheet.appendRow(['Date Tracked', 'Tender Date', 'Tender Number', 'Message ID', 'Title']);
    
    // Quick format for the new header
    sheet.getRange(1, 1, 1, 5)
         .setFontWeight("bold")
         .setBackground("#a855f7") // Purple background for tracked
         .setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(5, 400); // Make title column wider
  }
  
  // Append the row
  sheet.appendRow([
    new Date(),           // The exact moment you clicked track
    tenderData.date,      // The original email date
    tenderData.tender,    // 2026LD-...
    tenderData.msgId,     // Gmail ID
    tenderData.title      // Project Title
  ]);
  
  return true;
}



// Fetches a list of all manually tracked Tender IDs
function getTrackedTenderIds() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('Spreadsheet_ID');
  if (!sheetId) return [];
  
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName("Tracked_Tenders");
  
  // If sheet doesn't exist or has no data rows, return empty array
  if (!sheet || sheet.getLastRow() < 2) return []; 
  
  // Tender Numbers are stored in Column 3 (C)
  const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues(); 
  Logger.log(data);
  return data.map(row => String(row[0]).trim()).filter(id => id !== "");
}

// Removes a tender from the tracking sheet
function removeFromTracked(tenderNumber) {
  const sheetId = PropertiesService.getScriptProperties().getProperty('Spreadsheet_ID');
  if (!sheetId) return false;
  
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName("Tracked_Tenders");
  if (!sheet) return false;
  
  // ULTRA-FAST: Let Google's servers find the exact cell in Column C instantly
  const match = sheet.getRange("C:C").createTextFinder(String(tenderNumber).trim()).matchEntireCell(true).findNext();
  
  if (match) {
    sheet.deleteRow(match.getRow());
    return true;
  }
  
  return false;
}
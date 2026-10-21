//--------------------------------------------------------------------------------------------------------------------------------------------------
// Funciones encargadas con la base de datos del SICOP
// Para más información consultar: https://www.observatoriocomprapublica.go.cr/descargas-sicop/
//--------------------------------------------------------------------------------------------------------------------------------------------------

//--------------------------------------------------------------------------------------------------------------------------------------------------
// 1. THE ORCHESTRATOR & SYNC MECHANICS
//--------------------------------------------------------------------------------------------------------------------------------------------------

async function syncSicopDatabase() {
  const date = new Date();
  await downloadSicopMonth(date.getFullYear(), date.getMonth() + 1);
  cleanOldArchives(); 
}

// 🔴 ONE-TIME SETUP: Run this function ONCE from the editor to build the 12-month archive
async function syncHistoricalArchive() {
  const MONTHS_BACK = 12;
  
  for (let i = 0; i <= MONTHS_BACK; i++) {
    const targetDate = new Date();
    targetDate.setMonth(targetDate.getMonth() - i);
    
    const year = targetDate.getFullYear();
    const month = targetDate.getMonth() + 1;
    
    Logger.log(`Downloading historical archive: ${year}-${month}`);
    await downloadSicopMonth(year, month);
    
    // Pause for 5 seconds between months to prevent server rate-limiting
    Utilities.sleep(5000); 
  }
  Logger.log("Historical Archive Complete.");
}

async function downloadSicopMonth(yearNum, monthNum) {
  const rootFolderName = "SICOP_Archive";
  const yearStr = yearNum.toString();
  const monthStr = monthNum.toString().padStart(2, '0');
  
  // 1. Folder Management
  let rootFolders = DriveApp.getFoldersByName(rootFolderName);
  let rootFolder = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder(rootFolderName);
  
  let yearFolders = rootFolder.getFoldersByName(yearStr);
  let yearFolder = yearFolders.hasNext() ? yearFolders.next() : rootFolder.createFolder(yearStr);
  
  let monthFolders = yearFolder.getFoldersByName(monthStr);
  let targetFolder = monthFolders.hasNext() ? monthFolders.next() : yearFolder.createFolder(monthStr);
  
  const targetFolderId = targetFolder.getId();
  
  try {
    // 2. Clear existing CSVs to prevent duplicates before downloading
    const targetFiles = ["DetalleCarteles.csv", "DetalleLineaCartel.csv"];
    targetFiles.forEach(fileName => {
      const existing = targetFolder.getFilesByName(fileName);
      while (existing.hasNext()) existing.next().setTrashed(true);
    });

    Logger.log(`Starting Phase 1: Chunked Download for ${yearStr}/${monthStr}...`);
    const zipFileId = downloadLargeZipChunked(yearStr, monthStr, targetFolderId);
    
    if (zipFileId) {
      Logger.log(`Starting Phase 2: Memory Extraction...`);
      await extractLargeZip(zipFileId, targetFolderId);
      
      // 3. Clean up the massive ZIP file to save Drive space
      DriveApp.getFileById(zipFileId).setTrashed(true);
      Logger.log(`✅ Database sync complete for ${yearStr}/${monthStr}. Temporary ZIP deleted.`);
    }
  } catch (error) {
    Logger.log(`❌ ERROR processing ${yearStr}/${monthStr}: ` + error.message);
  }
}

function cleanOldArchives() {
  const MAX_MONTHS_TO_KEEP = 12;
  const rootFolders = DriveApp.getFoldersByName("SICOP_Archive");
  if (!rootFolders.hasNext()) return;
  
  const rootFolder = rootFolders.next();
  
  const cutoffDate = new Date();
  cutoffDate.setMonth(cutoffDate.getMonth() - MAX_MONTHS_TO_KEEP);
  const cutoffYear = cutoffDate.getFullYear();
  const cutoffMonth = cutoffDate.getMonth() + 1;
  
  const yearFolders = rootFolder.getFolders();
  while (yearFolders.hasNext()) {
    const yearFolder = yearFolders.next();
    const folderYear = Number.parseInt(yearFolder.getName());
    
    if (folderYear < cutoffYear) {
      yearFolder.setTrashed(true);
      Logger.log(`🗑️ Trashed old archive year: ${folderYear}`);
    } else if (folderYear === cutoffYear) {
      const monthFolders = yearFolder.getFolders();
      while (monthFolders.hasNext()) {
        const monthFolder = monthFolders.next();
        const folderMonth = Number.parseInt(monthFolder.getName());
        
        if (folderMonth < cutoffMonth) {
          monthFolder.setTrashed(true);
          Logger.log(`🗑️ Trashed old archive month: ${folderYear}/${monthFolder.getName()}`);
        }
      }
    }
  }
}

//--------------------------------------------------------------------------------------------------------------------------------------------------
// 2. THE CHUNKED DOWNLOADER & MEMORY EXTRACTOR
//--------------------------------------------------------------------------------------------------------------------------------------------------

function downloadLargeZipChunked(yearStr, monthStr, targetFolderId) {
  const zipUrl = `https://dlsaobservatorioprod.blob.core.windows.net/fs-synapse-observatorio-produccion/Zip/${yearStr}${monthStr}.zip`;
  const fileName = `${yearStr}${monthStr}.zip`;
  const chunkSize = 8 * 1024 * 1024; 
  
  let sizeResponse;
  
  // 1. Catch hard network drops (like "Address unavailable")
  try {
    sizeResponse = UrlFetchApp.fetch(zipUrl, { headers: { "Range": "bytes=0-0" }, muteHttpExceptions: true });
  } catch (e) {
    console.error(`⚠️ Network drop reaching SICOP for ${yearStr}/${monthStr}. Skipping.`);
    return null; 
  }
  
  // 2. Catch missing files (HTTP 404 Not Found)
  if (sizeResponse.getResponseCode() === 404) {
    console.log(`⚠️ SICOP has not published data for ${yearStr}/${monthStr} yet (404). Skipping.`);
    return null; 
  }
  
  const headers = sizeResponse.getHeaders();
  const contentRange = headers['Content-Range'] || headers['content-range'];
  
  if (!contentRange) throw new Error("Could not determine file size from server.");
  const totalSize = Number.parseInt(contentRange.split('/')[1]);

  const token = ScriptApp.getOAuthToken();
  const metadata = { name: fileName, parents: [targetFolderId] };
  
  const initResponse = UrlFetchApp.fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable", {
    method: "post",
    headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
    payload: JSON.stringify(metadata),
    muteHttpExceptions: true
  });
  
  const uploadUrl = initResponse.getHeaders()['Location'] || initResponse.getHeaders()['location'];
  
  let start = 0;
  let finalResponse = null;

  while (start < totalSize) {
    let end = Math.min(start + chunkSize - 1, totalSize - 1);
    
    let chunkResponse = UrlFetchApp.fetch(zipUrl, {
      headers: { "Range": `bytes=${start}-${end}` },
      muteHttpExceptions: true
    });
    
    finalResponse = UrlFetchApp.fetch(uploadUrl, {
      method: "put",
      headers: { "Content-Range": `bytes ${start}-${end}/${totalSize}` },
      payload: chunkResponse.getBlob(),
      muteHttpExceptions: true
    });
    
    start = end + 1;
  }
  
  const responseData = JSON.parse(finalResponse.getContentText());
  return responseData.id; 
}

async function extractLargeZip(zipFileId, targetFolderId) {
  const jsZipUrl = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
  
  // Trick JSZip into thinking it's in a browser environment
  globalThis.setImmediate = function(cb) { cb(); };
  
  eval(UrlFetchApp.fetch(jsZipUrl).getContentText());
  const token = ScriptApp.getOAuthToken();
  // ... rest of the code stays exactly the same
  
  const metaResponse = UrlFetchApp.fetch(`https://www.googleapis.com/drive/v3/files/${zipFileId}?fields=size`, {
    headers: { "Authorization": "Bearer " + token }
  });
  const totalSize = Number.parseInt(JSON.parse(metaResponse.getContentText()).size);
  
  const zipData = new Uint8Array(totalSize);
  const chunkSize = 10 * 1024 * 1024;
  let offset = 0;
  
  while (offset < totalSize) {
    let end = Math.min(offset + chunkSize - 1, totalSize - 1);
    let chunkRes = UrlFetchApp.fetch(`https://www.googleapis.com/drive/v3/files/${zipFileId}?alt=media`, {
      headers: { "Authorization": "Bearer " + token, "Range": `bytes=${offset}-${end}` },
      muteHttpExceptions: true
    });
    zipData.set(new Uint8Array(chunkRes.getContent()), offset);
    offset = end + 1;
  }
  
  const zip = new JSZip();
  const loadedZip = await zip.loadAsync(zipData);
  
  const targetFiles = ["DetalleCarteles.csv", "DetalleLineaCartel.csv"];
  for (let i = 0; i < targetFiles.length; i++) {
    const filename = targetFiles[i];
    if (loadedZip.files[filename]) {
      const uncompressedData = await loadedZip.files[filename].async("uint8array");
      uploadMemoryToDrive(uncompressedData, filename, targetFolderId);
    }
  }
}

function uploadMemoryToDrive(uint8Data, fileName, folderId) {
  const token = ScriptApp.getOAuthToken();
  const totalSize = uint8Data.length;
  
  const metadata = { name: fileName, parents: [folderId] };
  const initRes = UrlFetchApp.fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable", {
    method: "post",
    headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
    payload: JSON.stringify(metadata),
    muteHttpExceptions: true
  });
  
  const uploadUrl = initRes.getHeaders()['Location'] || initRes.getHeaders()['location'];
  const chunkSize = 8 * 1024 * 1024; 
  let start = 0;
  
  while (start < totalSize) {
    let end = Math.min(start + chunkSize - 1, totalSize - 1);
    let chunkArray = Array.from(uint8Data.slice(start, end + 1));
    
    UrlFetchApp.fetch(uploadUrl, {
      method: "put",
      headers: { "Content-Range": `bytes ${start}-${end}/${totalSize}` },
      payload: chunkArray,
      muteHttpExceptions: true
    });
    start = end + 1;
  }
}

//--------------------------------------------------------------------------------------------------------------------------------------------------
// 3. THE SEARCH PIPELINE
//--------------------------------------------------------------------------------------------------------------------------------------------------

function getCacheFolder(searchDate) {
  let attempts = 0;
  
  while (attempts < 3) {
    try {
      const rootFolders = DriveApp.getFoldersByName("SICOP_Archive");
      if (!rootFolders.hasNext()) return null;
      
      const yearStr = searchDate.getFullYear().toString();
      const yearFolders = rootFolders.next().getFoldersByName(yearStr);
      if (!yearFolders.hasNext()) return null;
      
      const monthStr = (searchDate.getMonth() + 1).toString().padStart(2, '0');
      const monthFolders = yearFolders.next().getFoldersByName(monthStr);
      
      return monthFolders.hasNext() ? monthFolders.next() : null;
      
    } catch (e) {
      attempts++;
      if (attempts >= 3) {
        console.error(`Drive API crashed accessing folder after 3 attempts: ${e.message}`);
        return null; 
      }
      Utilities.sleep(1000); 
    }
  }
}

function searchTenderInDrive(targetTenderNumber, emailDate) {
  try {
    const MAX_MONTHS_BACK = 12; 

    for (let i = 0; i <= MAX_MONTHS_BACK; i++) {
      Utilities.sleep(500);
      
      const searchDate = new Date(emailDate.getFullYear(), emailDate.getMonth() - i, 1);
      const targetFolder = getCacheFolder(searchDate);
      
      if (!targetFolder) continue; 
      
      const cartelesFiles = targetFolder.getFilesByName("DetalleCarteles.csv");
      if (!cartelesFiles.hasNext()) continue;
      
      const cartelFile = cartelesFiles.next();
      const headerMatches = searchLargeCsvChunked(cartelFile, targetTenderNumber, false);
      
      let internalSicopId = null;
      let tenderHeader = {};
      
      if (headerMatches.length > 0) {
        const row = parseCsvLine(headerMatches[0]);
        if (row.length > 13 && row[3] === targetTenderNumber) {
          internalSicopId = row[0]; 
          tenderHeader = {
            tenderNumber: row[3],
            status: row[6],
            title: row[7],
            budget: row[13]
          };
        }
      }
      
      if (!internalSicopId) {
        continue; 
      }
      
      const lineasFiles = targetFolder.getFilesByName("DetalleLineaCartel.csv");
      if (!lineasFiles.hasNext()) return { header: tenderHeader, items: [] };
      
      const lineaFile = lineasFiles.next();
      const itemMatches = searchLargeCsvChunked(lineaFile, internalSicopId, true);
      
      const lineItems = [];
      
      for (let k = 0; k < itemMatches.length; k++) {
        const row = parseCsvLine(itemMatches[k]);
        if (row.length > 10 && row[0] === internalSicopId) {
          lineItems.push({
            partida: row[2],
            lineNo: row[1],
            qty: parseFloat(row[3]).toString(),
            price: row[4],
            currency: row[5],
            exchangeRate: row[6],
            desc: row[10]
          });
        }
      }
      
      return { header: tenderHeader, items: lineItems };
    }
    
    return null;

  } catch (e) {
    console.error(`Drive API crashed searching for ${targetTenderNumber}: ${e.message}`);
    return null;
  }
}

function searchLargeCsvChunked(file, searchString, findAll = false) {
  const fileId = file.getId();
  const token = ScriptApp.getOAuthToken();
  const url = `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`;
  
  const chunkSize = 10 * 1024 * 1024; 
  let start = 0;
  let end = chunkSize - 1;
  let overlap = "";
  const matches = [];
  
  while (true) {
    const response = UrlFetchApp.fetch(url, {
      headers: {
        "Authorization": "Bearer " + token,
        "Range": `bytes=${start}-${end}`
      },
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    if (code !== 206 && code !== 200) break; 
    
    const chunkText = overlap + response.getContentText("UTF-8");
    const lines = chunkText.split(/\r?\n/);
    
    overlap = lines.pop(); 
    
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].includes(searchString)) {
        matches.push(lines[i]);
        if (!findAll) return matches; 
      }
    }
    
    const headers = response.getHeaders();
    const contentLength = headers['Content-Length'] || headers['content-length'];
    if (!contentLength || Number.parseInt(contentLength) < chunkSize) {
      if (overlap.includes(searchString)) matches.push(overlap);
      break;
    }
    
    start += chunkSize;
    end += chunkSize;
  }
  
  return matches;
}

function parseCsvLine(text) {
  let result = [];
  let startValue = 0;
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    if (text[i] === '"') inQuotes = !inQuotes;
    else if (text[i] === ';' && !inQuotes) { 
      result.push(text.substring(startValue, i).replace(/^"|"$/g, '').trim());
      startValue = i + 1;
    }
  }
  result.push(text.substring(startValue).replace(/^"|"$/g, '').trim());
  return result;
}

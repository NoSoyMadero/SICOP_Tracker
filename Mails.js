//--------------------------------------------------------------------------------------------------------------------------------------------------
// Funciones encargadas de lo relacionado con el correo
//--------------------------------------------------------------------------------------------------------------------------------------------------

// Organize the other functions and send the parameters based on the Config.gs file
function processSicopEmails() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('Spreadsheet_ID'); 
  if (!sheetId) return Logger.log('Error: Spreadsheet ID not found.');
  
  const ss = SpreadsheetApp.openById(sheetId);
  
  const emailConfigs = [
    { 
      sheetName: CONFIG.SHEETS.TODAY, 
      query: `from:${CONFIG.GMAIL.SENDER} ${CONFIG.GMAIL.QUERIES.TODAY}` 
    },
    { 
      sheetName: CONFIG.SHEETS.WEEKLY, 
      query: `from:${CONFIG.GMAIL.SENDER} ${CONFIG.GMAIL.QUERIES.WEEKLY}` 
    },
    { 
      sheetName: CONFIG.SHEETS.MONTHLY, 
      query: `from:${CONFIG.GMAIL.SENDER} ${CONFIG.GMAIL.QUERIES.MONTHLY}` 
    }
  ];

  emailConfigs.forEach(({ sheetName, query }) => {
    // 1. Get the sheet, or create it if it doesn't exist
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    // 2. Now the sheet object is 100% guaranteed to exist
    const existingIds = getExistingMap(sheet); 
    
    const queryEmails = getSicopEmails(query, existingIds);
    saveToSheet(sheet, queryEmails); 
  });
}



function getExistingMap(sheet) {
  const existingMap = new Map(); 
  if (!sheet || sheet.getLastRow() <= 1) return existingMap;
  
  const idData = sheet.getRange(2, 5, sheet.getLastRow() - 1, 2).getValues();
  
  idData.forEach(([status, msgId]) => {
    if (msgId) existingMap.set(msgId, status);
  });
  
  return existingMap; 
}



// Function 1: The Looper (Handles Gmail threads and Cache)
function getSicopEmails(query, existingMap) {
  const threads = GmailApp.search(query);
  const emailList = []; 
  const tenderRegex = /\d{4}[A-Z]{2}-\d{6}-\d{10}/i; 
  
  for (let i = 0; i < threads.length; i++) {
    try {
      // Add a small delay to prevent Google from rate-limiting the script
      Utilities.sleep(250);
      
      const messages = threads[i].getMessages();
      for (let j = 0; j < messages.length; j++) {
        const message = messages[j];
        const messageId = message.getId();
        
        // 1. Circuit Breaker Logic
        if (existingMap?.has(messageId)) {
          const currentStatus = existingMap.get(messageId);
          if (!currentStatus.includes("⏳ PENDIENTE")) {
            // Add an extra "" to account for the new Tracked column
            emailList.push(["EXISTING", "", "", "", "", messageId]);
            continue;
          }
        }
        
        // 2. Extract Base Variables
        const subject = message.getSubject();
        const body = message.getPlainBody();
        
        const subjectMatch = subject.match(tenderRegex);
        const bodyMatch = body.match(tenderRegex);
        let tenderNumber = subjectMatch ? subjectMatch[0] : (bodyMatch ? bodyMatch[0] : "NOT_FOUND");
        
        // 3. Hand off the heavy lifting to our analyzer
        const { dossierText, aiSummary, finalTenderNumber } = analyzeEmailType(message, tenderNumber);
        
        emailList.push([message.getDate(), finalTenderNumber, subject, dossierText, aiSummary, messageId]);
      }
    } catch (e) {
      // If Gmail throws "Operation not allowed", log it and skip to the next thread safely
      console.error(`Gmail error on thread ${i}: ${e.message}`);
      continue;
    }
  }
  return emailList;
}



// Function 2: The Analyzer (Determines what kind of email this is)

function analyzeEmailType(message, tenderNumber) {
  const subjectLower = message.getSubject().toLowerCase();
  const htmlBody = message.getBody(); 
  let dossierText = "N/A";
  let aiSummary = ""; 
  let finalTenderNumber = tenderNumber;
  
  // REGEX: Extract the official title from the SICOP HTML body
  let extractedTitle = "";
  const titleMatch = htmlBody.match(/para\s+<b>(.*?)<\/b>/i);
  if (titleMatch) {
    // Clean up HTML non-breaking spaces (&nbsp; or \u00A0) that SICOP uses
    extractedTitle = titleMatch[1].replace(/&nbsp;/g, ' ').replace(/\u00A0/g, ' ').trim();
  }
  
  if (subjectLower.includes("solicitud") || subjectLower.includes("subsanaci") || subjectLower.includes("Notif") || subjectLower.includes("Notificación") ) {
    dossierText = extractedTitle ? `TÍTULO: ${extractedTitle}\n\nVer correo para detalles del documento a subsanar.` : "Ver correo para detalles del documento a subsanar.";
    aiSummary = "URGENTE: Información Requerida";
    
  } else if (subjectLower.includes("modificaci") || subjectLower.includes("aclaraci")) {
    dossierText = extractedTitle ? `TÍTULO: ${extractedTitle}\n\nVer portal SICOP para revisar los cambios al cartel.` : "Ver portal SICOP para revisar los cambios al cartel.";
    aiSummary = "ALERTA: Modificación/Aclaración al Cartel";
    
  } else if (subjectLower.includes("alerta") || subjectLower.includes("intenci")) {
    aiSummary = "AVISO: Proceso de compra temprana";
    const alertData = parseAlertaHtml(htmlBody);
    dossierText = alertData.text;
    if (finalTenderNumber === "NOT_FOUND" && alertData.tenderId) {
       finalTenderNumber = alertData.tenderId;
    }
    
  } else if (subjectLower.includes("invitaci")) {
    if (finalTenderNumber !== "NOT_FOUND") {
      const rawJson = searchTenderInDrive(finalTenderNumber, message.getDate());
      if (rawJson?.header?.title) {
        dossierText = formatTenderForAI(rawJson);
      } else {
        dossierText = extractedTitle ? `TÍTULO: ${extractedTitle}\n\nDatos aún no publicados por SICOP.` : "Datos no encontrados en el archivo CSV de SICOP.";
        aiSummary = "⏳ PENDIENTE: Datos aún no publicados por SICOP";
      }
    } else {
      dossierText = "No se encontró número de licitación.";
      aiSummary = "ERROR: Sin número de licitación";
    }
    
  } else {
    aiSummary = `INFO: ${message.getSubject()}`; 
  }
  
  return { dossierText, aiSummary, finalTenderNumber };
}




// Function 3: The HTML Parser (Keeps ugly Regex isolated)
function parseAlertaHtml(htmlBody) {
  const tableRowRegex = /<tr>\s*<td>(.*?)<\/td>\s*<td>(.*?)<\/td>\s*<td>(.*?)<\/td>\s*<td>(.*?)<\/td>\s*<\/tr>/ig;
  let match;
  let extractedInfo = [];
  let fallbackTenderId = null;
  
  while ((match = tableRowRegex.exec(htmlBody)) !== null) {
    let idAndType = match[1].replace(/<[^>]+>/g, '').trim();
    let institution = match[2].replace(/<[^>]+>/g, '').trim();
    let title = match[3].replace(/<[^>]+>/g, '').trim();
    let deadline = match[4].replace(/<[^>]+>/g, '').trim();
    
    extractedInfo.push(`Institución: ${institution}\nProyecto: ${title}\nCierre: ${deadline}`);
    
    if (!fallbackTenderId) fallbackTenderId = idAndType.split(' ')[0]; 
  }
  
  return {
    text: extractedInfo.length > 0 ? extractedInfo.join('\n\n') : "Ver correo para detalles.",
    tenderId: fallbackTenderId
  };
}
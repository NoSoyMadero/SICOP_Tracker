//--------------------------------------------------------------------------------------------------------------------------------------------------
// Funciones relacionadas con la IA
//--------------------------------------------------------------------------------------------------------------------------------------------------

// Esta función gestiona las demás funciones de este archivo

function batchProcessEmails(sheetName, cache) {
  const sheetId = PropertiesService.getScriptProperties().getProperty('Spreadsheet_ID');
  if (!sheetId) return Logger.log('Error: Spreadsheet ID not found.');

  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName); 
  if (!sheet) return Logger.log(`Error: Sheet "${sheetName}" not found.`);

  const data = sheet.getDataRange().getValues(); 
  const BATCH_SIZE = 3; 
  let currentBatch = [];

  for (let i = 1; i < data.length; i++) {
    const rowNumber = i + 1;
    const [emailDate, tenderNumber, subject, techSpecs, summary, messageId] = data[i];
    
    if (!messageId) continue; 
    
    const summaryText = summary ? String(summary).toUpperCase() : "";
    const specsText = techSpecs ? String(techSpecs) : "";

    // 1. Populate the cache if this row is already completed
    if (summaryText !== "" && !summaryText.includes("PENDIENTE")) {
      cache[messageId] = summaryText; 
      continue; 
    }

    // NEW BLOCK: Prevent AI evaluation if SICOP data is still missing
    if (specsText.includes("Datos aún no publicados") || specsText.includes("Datos no encontrados")) {
      // Skip this row so it remains "PENDIENTE" in the sheet
      continue; 
    }

    // 2. Read from the cache to prevent duplicate AI calls
    if (cache[messageId]) {
      sheet.getRange(rowNumber, 5).setValue(cache[messageId]);
      continue;
    }

    currentBatch.push({ row: rowNumber, subject: subject, text: techSpecs, key: messageId });

    if (currentBatch.length === BATCH_SIZE) {
      Logger.log(`[${sheetName}] Sending batch to AI: Rows ${currentBatch.map(item => item.row).join(', ')}`);
      processAiBatch(currentBatch, sheet, cache);
      currentBatch = []; 
    }
  }

  if (currentBatch.length > 0) {
    Logger.log(`[${sheetName}] Sending final batch to AI: Rows ${currentBatch.map(item => item.row).join(', ')}`);
    processAiBatch(currentBatch, sheet, cache);
  }
}

// Esta función se encarga de el manejo de la API de IA
// Para más información consultar : [https://docs.google.com/document/d/1EV3-XqPZBxDNe9rMxfNHc0f5ilHAKFAtey2-vOdmH48/edit?tab=t.0]

function callGenAI(promptText) {
  // 1. Fetch API key and fail gracefully if missing
  const apiKey = PropertiesService.getScriptProperties().getProperty('GENAI_API_KEY');
  if (!apiKey) {
    Logger.log('Error: GENAI_API_KEY not found in Script Properties.');
    return null; 
  }

  // 2. Package the request options cleanly
  const options = {
    method: 'post',
    headers: {
      'x-msi-genai-api-key': apiKey, 
      'Content-Type': 'application/json' 
    },
    payload: JSON.stringify({
      userId: CONFIG.GENAI.USER_ID, 
      model: CONFIG.GENAI.MODEL,
      system: CONFIG.GENAI.COMPANY_PROFILE,
      prompt: promptText, 
      modelConfig: { 
        temperature: 0.3, 
        "max tokens": 500 
      }
    }),
    muteHttpExceptions: true
  };

  // 3. Execute and handle the response
  try {
    const response = UrlFetchApp.fetch(CONFIG.GENAI.ENDPOINT, options);
    const json = JSON.parse(response.getContentText());

    if (json.status === true) {
      return json.msg; 
    } 
    
    Logger.log(`API Error: ${json.msg}`);
    return null;
    
  } catch (error) {
    Logger.log(`Request Failed: ${error.toString()}`);
    return null;
  }
}

// Envia los paquetes a la función de IA
function processAiBatch(batch, sheet, cache) {
  if (batch.length === 0) return;

  const promptText = CONFIG.PROMPTS.PROMPT_DOSSIER(batch);
  const aiResponse = callGenAI(promptText);

  if (!aiResponse) {
    Logger.log("Skipping row updates: AI returned no response.");
    return;
  }

  const summaries = aiResponse.split('|||');
  
  batch.forEach((item, index) => {
    const formattedSummary = summaries[index] ? summaries[index].trim() : "Error: Summary not generated.";
    sheet.getRange(item.row, 5).setValue(formattedSummary); // Write to Column 5
    cache[item.key] = formattedSummary; 
  });
}


// Ordena los datos para enviar a la IA

function formatTenderForAI(tenderData) {
  if (!tenderData?.header) {
    return "Información técnica no encontrada en la base de datos.";
  }
  
  const { title, budget, status } = tenderData.header;
  
  // Grab the exchange rate from the first item (it applies to the whole tender)
  const rate = (tenderData.items && tenderData.items[0] && tenderData.items[0].exchangeRate) ? tenderData.items[0].exchangeRate : "N/A";
  
  const headerText = 
  `
  TÍTULO: ${title}
  PRESUPUESTO ESTIMADO: ${budget}
  TIPO DE CAMBIO: ${rate}
  ESTADO: ${status}

  LÍNEAS DE ARTÍCULOS SOLICITADOS:
  `;
  
  const itemsText = (tenderData.items || []).map(item => {
    const shortDesc = item.desc.length > 300 ? `${item.desc.substring(0, 300)}...` : item.desc;
    return `- Partida ${item.partida} (Línea ${item.lineNo}): ${item.qty} unidades [Moneda: ${item.currency}, Precio: ${item.price}]\n  Descripción: ${shortDesc}`;
  }).join('\n');
  
  return headerText + itemsText;
}

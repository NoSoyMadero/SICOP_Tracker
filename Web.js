function doGet() {
  return HtmlService.createTemplateFromFile('Dashboard')
      .evaluate()
      .setTitle('Motorola SICOP Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Helper to fetch the raw email on demand
function fetchEmailBody(messageId) {
  try {
    const message = GmailApp.getMessageById(messageId);
    return message.getBody(); 
  } catch (e) {
    return "<div style='padding:20px; font-family:sans-serif;'>Error: No se pudo cargar el correo original.</div>";
  }
}

// Helper to clean and format raw budget numbers
function formatCurrency(rawBudget, fixedRate) {
  if (!rawBudget) return "";
  
  const numMatch = rawBudget.match(/[\d\.]+/);
  if (!numMatch) return rawBudget;
  
  const num = parseFloat(numMatch[0]);
  if (isNaN(num)) return rawBudget;
  
  const formatted = new Intl.NumberFormat('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(num);
  
  if (rawBudget.toUpperCase().includes("USD") || rawBudget.includes("$")) {
    return `USD $${formatted}`;
  } else if (fixedRate && !isNaN(fixedRate) && fixedRate > 0) {
    const usdApprox = new Intl.NumberFormat('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(num / fixedRate);
    return `₡${formatted} CRC (~$${usdApprox} USD)`;
  } else {
    return `₡${formatted} CRC`;
  }
}

function getDashboardMetrics(timeframe) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheetId = scriptProperties.getProperty('Spreadsheet_ID');
  const ss = SpreadsheetApp.openById(sheetId);
  
  const trackedIdsArray = getTrackedTenderIds();
  const trackedSet = new Set(trackedIdsArray);

  let targetSheetName = "Monthly"; 
  if (timeframe === "TODAY") targetSheetName = "Today";
  if (timeframe === "WEEKLY") targetSheetName = "Weekly";
  
  const sheet = ss.getSheetByName(targetSheetName); 
  if (!sheet) return { urgent: [], alerta: [], relevant: [], aviso: [], pending: [], no_relevante: [], trackedIds: trackedIdsArray };

  const data = sheet.getDataRange().getValues();
  
  // We removed the 'tracked' array here since they go to 'urgent'
  let urgent = [], alerta = [], relevant = [], aviso = [], pending = [], no_relevante = [];

  for (let i = 1; i < data.length; i++) {
    const rawDate = data[i][0];
    const dateStr = rawDate instanceof Date ? rawDate.toLocaleDateString() : rawDate;
    const tenderNum = data[i][1];
    const subject = data[i][2];
    const techSpecs = data[i][3] || "";
    const rawSummary = data[i][4] || ""; 
    const msgId = data[i][5] || ""; 
    const summaryUpper = rawSummary.toUpperCase(); 

    let displayTitle = subject;
    let displayBudget = "";
    let displayInstitution = "";

    if (techSpecs.includes("TÍTULO:")) {
      const titleMatch = techSpecs.match(/TÍTULO:\s*(.*?)(?:\n|$)/i);
      if (titleMatch) displayTitle = titleMatch[1].trim();

      let extractedRate = null;
      const rateMatch = techSpecs.match(/TIPO DE CAMBIO:\s*([\d\.]+)/i);
      if (rateMatch) extractedRate = parseFloat(rateMatch[1].trim());

      const budgetMatch = techSpecs.match(/PRESUPUESTO ESTIMADO:\s*(.*?)(?:\n|$)/i);
      if (budgetMatch) displayBudget = formatCurrency(budgetMatch[1].trim(), extractedRate);
    } 
    else if (summaryUpper.includes("AVISO:")) {
      const instMatch = techSpecs.match(/Institución:\s*(.*?)(?:\n|$)/i);
      if (instMatch) displayInstitution = instMatch[1].trim();

      const projMatch = techSpecs.match(/Proyecto:\s*(.*?)(?:\n|$)/i);
      if (projMatch) displayTitle = projMatch[1].trim();
    }

    if (summaryUpper.includes("RELEVANTE:") && !displayBudget) {
      const budgetMatch = rawSummary.match(/Presupuesto:\s*(.*?)(?:\n|$)/i);
      if (budgetMatch) displayBudget = formatCurrency(budgetMatch[1].trim());
    }

    // NEW: Check the tracked set and save it directly inside the card data
    const isTracked = trackedSet.has(tenderNum);

    const tenderCard = { 
      date: dateStr, 
      tender: tenderNum, 
      subject: subject,
      title: displayTitle,
      budget: displayBudget,
      institution: displayInstitution,
      dossier: techSpecs,
      ai_summary: rawSummary,
      msgId: msgId,
      isTracked: isTracked // <-- THE FLAG
    };

    // ROUTING: If tracked, force it into Urgente. Otherwise, sort normally.
    if (isTracked || summaryUpper.includes("URGENTE")) {
      urgent.push(tenderCard);
    } 
    else if (summaryUpper.includes("NO RELEVANTE")) {
      no_relevante.push(tenderCard);
    }
    else if (summaryUpper.includes("ALERTA")) alerta.push(tenderCard);
    else if (summaryUpper.includes("RELEVANTE")) relevant.push(tenderCard);
    else if (summaryUpper.includes("AVISO")) aviso.push(tenderCard);
    else pending.push(tenderCard); 
  }

  return { urgent, alerta, relevant, aviso, pending, no_relevante, trackedIds: trackedIdsArray };
}
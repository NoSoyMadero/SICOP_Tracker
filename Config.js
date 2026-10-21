//--------------------------------------------------------------------------------------------------------------------------------------------------
// Archivo de configuraciones a lo largo del codigo
//--------------------------------------------------------------------------------------------------------------------------------------------------

const CONFIG = {
  GMAIL: {
    SENDER: 'info@sicop.go.cr',
    MAILBOX: 'sicop1@motorolasolutions.com',
    SUBJECT: 'Invitación',
    QUERIES: {
      TODAY: 'newer_than:2d',
      WEEKLY: 'newer_than:7d',
      MONTHLY: 'newer_than:1m'
    }
  },
  SHEETS: {
    TODAY: 'Today',
    WEEKLY: 'Weekly',
    MONTHLY: 'Monthly'
  },
  GENAI: {
    ENDPOINT: 'https://genai-service.stage.commandcentral.com/app-gateway/api/v2/chat',
    USER_ID: 'QWH348@motorolasolutions.com', 
    MODEL: "Gemini-2_5-Flash", 
    // Broadened profile to include general tech/parts
    COMPANY_PROFILE: "You are a procurement analyst for a telecommunications company (Motorola Solutions Colombia). You extract tenders for radios, software, communication hardware, and general technology. You err on the side of caution: if a tender is vague but could be tech-related, you include it."
  },
  PROMPTS: { // <-- FIXED: Changed '=' to ':'
    PROMPT_DOSSIER: (batch) => {
      let promptText = 
      `You are an expert procurement analyst for Motorola Solutions evaluating government tender dossiers. Evaluate each tender and generate a plain-text summary in Spanish.

      TARGET SCOPE:
      Land Mobile Radio (LMR) systems, P25 technology, trunked networks (troncalizado), Motorola repeater systems, and bulk radio fleet expansions.

      EVALUATION RULES:
      1. RELEVANTE: Radios or LMR infrastructure MUST be the PRIMARY product of the tender, AND it must meet at least one of these criteria: explicit mention of P25/troncalizado, Motorola brand request, large radio fleet quantity, or high-budget telecom infrastructure.
      2. NO RELEVANTE: Mark if the tender meets ANY of these: main product is not radios, minimal quantities (e.g., under 5 units), generic IT equipment (printers, laptops) without P25 specs, or unrelated categories (cleaning, food, civil works).

      OUTPUT FORMAT INSTRUCTIONS:
      - NO introductory text. NO markdown formatting (no asterisks, bold, etc.). Use plain text only.
      - Start tu response immediately with the evaluation tag.
      - Separate evaluations for multiple tenders using exactly three pipes (|||).

      TEMPLATE FOR RELEVANTE:
      RELEVANTE: [Una frase concisa indicando exactamente qué están comprando].
      Presupuesto: [Número formateado con comas, ej. 11,932,598.74]
      Detalle Técnico: [Lista corta separada por comas con los ítems principales y sus cantidades]

      TEMPLATE FOR NO RELEVANTE:
      NO RELEVANTE: [Motivo breve, ej. Equipo no relacionado, cantidad muy baja, o producto principal no es radio].\n\n`;

      batch.forEach((item, index) => {
        promptText += `Tender ${index + 1}:\nSubject: ${item.subject}\nData Dossier:\n${item.text}\n\n`;
      });

      return promptText;
    }
  }
};
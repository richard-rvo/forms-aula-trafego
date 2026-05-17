// ============================================================
// Google Apps Script — S-LAB · IA na Prática · NPS Pós-Aula
//
// Como usar:
//   1. Abra a planilha Google Sheets que vai receber as respostas
//   2. No menu: Extensões > Apps Script
//   3. Cole TODO este código no editor (substitua o conteúdo existente)
//   4. Clique em "Implantar" > "Nova implantação"
//      - Tipo: Aplicativo da Web
//      - Executar como: Eu (sua conta)
//      - Quem tem acesso: Qualquer pessoa
//   5. Copie a URL gerada e cole em index.html na variável SCRIPT_URL
// ============================================================

// ── ⚙️  CONFIGURAÇÃO — ajuste aqui se necessário ──────────────
const SHEET_NAME = 'NPS'; // nome da aba no Sheets
// ─────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const sheet = getOrCreateSheet();
    const data  = e.parameter;

    sheet.appendRow([
      data.timestamp,
      data.sentimento,
      data.parte_valiosa,
      data.mudaria,
      data.testar,
      data.nps,
    ]);

    return response({ success: true });
  } catch (err) {
    return response({ success: false, error: err.message });
  }
}

function doGet() {
  return response({ status: 'Apps Script NPS ativo ✓', projeto: 'S-LAB — IA na Prática' });
}

// ── Helpers ───────────────────────────────────────────────────

function getOrCreateSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);

    sheet.appendRow([
      'Data/Hora',
      'Sentimento pós-aula',
      'Parte mais valiosa',
      'O que mudaria / faltou',
      'O que vai testar primeiro',
      'NPS (0–10)',
    ]);

    // Formata cabeçalho com identidade S-LAB
    const header = sheet.getRange(1, 1, 1, 6);
    header.setBackground('#F08113');   // laranja principal S-LAB
    header.setFontColor('#FFFFFF');
    header.setFontWeight('bold');
    sheet.setFrozenRows(1);

    // Larguras das colunas
    sheet.setColumnWidth(1, 150);  // Data/Hora
    sheet.setColumnWidth(2, 280);  // Sentimento
    sheet.setColumnWidth(3, 200);  // Parte valiosa
    sheet.setColumnWidth(4, 340);  // O que mudaria
    sheet.setColumnWidth(5, 340);  // O que vai testar
    sheet.setColumnWidth(6, 80);   // NPS
  }

  return sheet;
}

function response(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

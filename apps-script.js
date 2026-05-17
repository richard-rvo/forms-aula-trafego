// ============================================================
// Google Apps Script — S-LAB · Tráfego Pago para Captação e Vendas
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
const SHEET_NAME = 'Inscrições'; // nome da aba no Sheets
// ─────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const sheet = getOrCreateSheet();
    const data  = e.parameter;

    sheet.appendRow([
      data.timestamp,
      data.nome,
      data.whatsapp,
      data.email,
      data.idade,
      data.sexo,
      data.profissao,
      data.vende,
      data.produto,
      data.cliente,
      data.preco,
      data.perfil,
      data.seguidores,
      data.frequencia,
      data.anuncios,
      data.gestao,
      data.desafio,
      data.impede,
      data.expectativa,
    ]);

    return response({ success: true });
  } catch (err) {
    return response({ success: false, error: err.message });
  }
}

function doGet() {
  return response({ status: 'Apps Script ativo ✓', projeto: 'S-LAB — Tráfego Pago para Captação e Vendas' });
}

// ── Helpers ───────────────────────────────────────────────────

function getOrCreateSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);

    const headers = [
      'Data/Hora',
      'Nome completo',
      'WhatsApp',
      'E-mail',
      'Idade',
      'Sexo',
      'Profissão / Área de atuação',
      'Situação de vendas',
      'O que vende / pretende vender',
      'Cliente ideal',
      'Faixa de preço',
      'Perfil no Instagram',
      'Seguidores',
      'Frequência de posts',
      'Histórico com anúncios',
      'Gestão do tráfego',
      'Maiores desafios no Instagram',
      'O que impede de crescer',
      'Expectativa com a aula',
    ];

    sheet.appendRow(headers);

    // Formata cabeçalho com identidade S-LAB
    const numCols = headers.length;
    const header  = sheet.getRange(1, 1, 1, numCols);
    header.setBackground('#F08113');
    header.setFontColor('#FFFFFF');
    header.setFontWeight('bold');
    sheet.setFrozenRows(1);

    // Larguras das colunas
    const widths = [150, 200, 130, 220, 100, 100, 220, 240, 340, 340, 130, 200, 130, 180, 260, 200, 260, 340, 340];
    widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }

  return sheet;
}

function response(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

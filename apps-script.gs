// Cole este código no Apps Script da sua planilha:
// Planilha > Extensões > Apps Script > cole e salve > Implantar > Novo implante

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = JSON.parse(e.postData.contents);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Data/Hora', 'Interesse', 'Nome', 'Email', 'WhatsApp', 'Origem']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    }

    const interesse = {
      A: 'Seminário gratuito sobre finanças',
      B: 'Aposentadoria livre de impostos',
      C: 'Planejamento de custos de faculdade',
      D: 'Proteção hipotecária',
      E: 'Benefícios em Vida',
      F: 'Transição de carreira / renda extra',
      G: 'Outro Tema'
    };

    sheet.appendRow([
      new Date().toLocaleString('pt-BR', { timeZone: 'America/New_York' }),
      interesse[data.interesse] || data.interesse,
      data.nome,
      data.email,
      data.whatsapp,
      data.origem
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

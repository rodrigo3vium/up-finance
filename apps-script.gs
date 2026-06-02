// Cole este código no Apps Script da sua planilha:
// Planilha > Extensões > Apps Script > cole e salve
// Depois implante: Implantar > Gerenciar implantações >
//   edite a implantação que o FORMULÁRIO usa (a do AKfycbwkp8...) >
//   Versão: "Nova versão" > Implantar

const EMAIL_CLIENTE = 'francielli@benitesalbuquerque.com.br';

// Etiqueta de versão. Abra a URL /exec do formulário no navegador (GET):
// se aparecer este texto, aquela URL está rodando o código NOVO.
function doGet(e) {
  return ContentService.createTextOutput(
    'Up Finance webhook ATIVO — VERSAO 2 (grava Idioma + envia email)'
  );
}

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = JSON.parse(e.postData.contents);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Data/Hora', 'Interesse', 'Nome', 'Email', 'WhatsApp', 'Origem', 'Idioma']);
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
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

    const interesseTexto = interesse[data.interesse] || data.interesse;
    const dataHora = new Date().toLocaleString('pt-BR', { timeZone: 'America/New_York' });

    // Força coluna WhatsApp (E) como texto para não interpretar o "+" como fórmula
    sheet.getRange('E:E').setNumberFormat('@');

    const newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, 7).setValues([[
      dataHora,
      interesseTexto,
      data.nome,
      data.email,
      data.whatsapp,
      data.origem,
      data.idioma || ''
    ]]);

    // Envia email pro cliente — isolado para não derrubar a gravação do lead se falhar
    try {
      const corpo =
        'Novo lead recebido pelo formulário:\n\n' +
        'Data/Hora: ' + dataHora + '\n' +
        'Interesse: ' + interesseTexto + '\n' +
        'Nome: ' + data.nome + '\n' +
        'Email: ' + data.email + '\n' +
        'WhatsApp: ' + data.whatsapp + '\n' +
        'Origem: ' + data.origem + '\n' +
        'Idioma: ' + (data.idioma || '—');

      MailApp.sendEmail({
        to: EMAIL_CLIENTE,
        subject: '🔔 Novo lead: ' + interesseTexto + ' — ' + data.nome,
        replyTo: data.email,
        body: corpo
      });
    } catch (mailErr) {
      console.error('Falha ao enviar email (lead já gravado): ' + mailErr.toString());
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// === TESTE ===
// Selecione esta função no editor e clique "Executar".
// 1ª vez: o Google vai pedir autorização (inclui permissão de enviar email) — aceite.
// Isso grava um lead de teste na planilha E dispara o email, sem precisar do formulário.
function testarEnvio() {
  doPost({
    postData: {
      contents: JSON.stringify({
        interesse: 'A',
        nome: 'Teste Lead',
        email: 'teste@exemplo.com',
        whatsapp: '+1 (555) 123-4567',
        origem: 'Teste manual',
        idioma: 'Português'
      })
    }
  });
}

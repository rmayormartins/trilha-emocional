// ============================================================
// TRILHA EMOCIONAL — Google Apps Script
// Compatível com o sendResponse() por questão do index.html
//
// ===== SETUP PASSO A PASSO =====
//
// 1. No Google Sheets: Extensões → Apps Script
// 2. Apague TODO o código existente (o myFunction padrão)
// 3. Cole ESTE código inteiro, salve (Ctrl+S)
// 4. Clique em ▶ Executar (com a função "testar" selecionada)
//    → Vai pedir permissão: Avançado → Ir para... → Permitir
//    → Confira na planilha se apareceu uma linha de teste
// 5. Implantar → Nova implantação:
//    - Tipo: Aplicativo da Web
//    - Executar como: EU MESMO
//    - Quem tem acesso: QUALQUER PESSOA  ← IMPORTANTE!
// 6. Copie a URL (começa com https://script.google.com/macros/s/...)
// 7. Cole no SHEETS_URL do index.html
//
// ⚠️  Se mudar o código depois, precisa fazer NOVA implantação
//     (Implantar → Gerenciar implantações → Criar versão)
//     A URL muda! Atualize no index.html.
// ============================================================

const SHEET_NAME = 'Respostas';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getOrCreateSheet();
    var letters = ['A','B','C','D','E','F'];

    sheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.sessionId || '',
      data.participantName || 'Anônimo',
      data.participantAge || '',
      data.fase || '',
      data.tema || '',
      (data.questionIndex != null) ? data.questionIndex + 1 : '',
      data.questionText || '',
      letters[data.selectedIndex] || (data.selectedIndex + 1),
      data.selectedText || '',
      data.isCorrect ? 'SIM' : 'NÃO',
      data.pointsEarned || 0,
      data.responseTimeMs ? Math.round(data.responseTimeMs / 1000) : '',
      data.coinsEarned || 0,
      data.totalCoins || 0,
      data.totalScore || 0
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    // Log do erro para debug (ver em Execuções no Apps Script)
    console.error('doPost erro:', err);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Também aceita GET (útil para testar no navegador)
function doGet(e) {
  return ContentService
    .createTextOutput('✅ Trilha Emocional API está funcionando! Use POST para enviar dados.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'Timestamp',
      'SessionId',
      'Nome',
      'Idade',
      'Fase',
      'Tema',
      'Questão Nº',
      'Pergunta',
      'Opção',
      'Texto da Opção',
      'Acertou?',
      'Pontos',
      'Tempo (s)',
      'Moedas',
      'Moedas Total',
      'Score Total'
    ]);

    // Estilizar cabeçalho
    var headerRange = sheet.getRange(1, 1, 1, 16);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#3D2B5F');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    // Larguras
    sheet.setColumnWidth(1, 170);  // Timestamp
    sheet.setColumnWidth(2, 130);  // SessionId
    sheet.setColumnWidth(3, 120);  // Nome
    sheet.setColumnWidth(4, 60);   // Idade
    sheet.setColumnWidth(5, 180);  // Fase
    sheet.setColumnWidth(6, 160);  // Tema
    sheet.setColumnWidth(7, 80);   // Questão Nº
    sheet.setColumnWidth(8, 280);  // Pergunta
    sheet.setColumnWidth(9, 60);   // Opção
    sheet.setColumnWidth(10, 220); // Texto da Opção
    sheet.setColumnWidth(11, 70);  // Acertou
    sheet.setColumnWidth(12, 60);  // Pontos
    sheet.setColumnWidth(13, 70);  // Tempo
    sheet.setColumnWidth(14, 70);  // Moedas
    sheet.setColumnWidth(15, 90);  // Moedas Total
    sheet.setColumnWidth(16, 85);  // Score Total
  }

  return sheet;
}

// ============================================================
// TESTE — rode esta função para:
// (a) autorizar as permissões
// (b) verificar que os dados chegam na planilha
// ============================================================
function testar() {
  doPost({
    postData: {
      contents: JSON.stringify({
        timestamp: new Date().toISOString(),
        sessionId: 'TESTE-' + Date.now(),
        participantName: 'Aluno Teste',
        participantAge: '10',
        fase: '🎮 Fase 1 – Jogo Online',
        tema: 'Consciência Emocional',
        questionIndex: 0,
        questionText: 'O que João está sentindo?',
        selectedIndex: 2,
        selectedText: '😡 Bravo / com raiva',
        isCorrect: true,
        pointsEarned: 2,
        responseTimeMs: 3200,
        coinsEarned: 100,
        totalCoins: 100,
        totalScore: 2
      })
    }
  });

  // Segundo teste para simular outra questão
  doPost({
    postData: {
      contents: JSON.stringify({
        timestamp: new Date().toISOString(),
        sessionId: 'TESTE-' + Date.now(),
        participantName: 'Aluno Teste',
        participantAge: '10',
        fase: '🎮 Fase 1 – Jogo Online',
        tema: 'Consciência Emocional',
        questionIndex: 1,
        questionText: 'O que você faria?',
        selectedIndex: 1,
        selectedText: 'Pedir para ele se acalmar',
        isCorrect: true,
        pointsEarned: 2,
        responseTimeMs: 4100,
        coinsEarned: 100,
        totalCoins: 200,
        totalScore: 4
      })
    }
  });
}

// ============================================================
// DEBUG — cole a URL no navegador para testar se está online
// Se aparecer "✅ Trilha Emocional API está funcionando!" → OK
// Se der erro 403 → a permissão não está como "Qualquer pessoa"
// Se der erro 404 → a URL está errada
// ============================================================

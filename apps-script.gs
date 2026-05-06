// ============================================================
// TRILHA DAS ATITUDES — Escala de Competência Emocional Digital
// Google Apps Script
// Cole em: Extensões → Apps Script (dentro do Google Sheets)
// ============================================================

const SHEET_NAME = 'Respostas';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet();
    sheet.appendRow([
      data.timestamp,
      data.sessionId,
      data.participantName,
      data.participantAge,
      data.fase,
      data.tema,
      data.questionIndex + 1,
      data.questionText,
      data.selectedIndex + 1,
      data.selectedText,
      data.isCorrect,
      data.pointsEarned,
      data.responseTimeMs,
      data.coinsEarned,
      data.starsEarned,
      data.totalCoins,
      data.totalStars,
      data.totalScore
    ]);
    return ContentService.createTextOutput(JSON.stringify({ok: true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'timestamp', 'sessionId', 'nome', 'idade',
      'fase', 'tema',
      'questao_n', 'enunciado', 'opcao_escolhida_n', 'opcao_texto',
      'acertou', 'pontos_ganhos', 'tempo_resposta_ms',
      'moedas_ganhas', 'estrelas_ganhas',
      'moedas_total', 'estrelas_total', 'score_total'
    ]);
    sheet.getRange(1, 1, 1, 18)
      .setFontWeight('bold')
      .setBackground('#3D2B5F')
      .setFontColor('#fff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 180);  // timestamp
    sheet.setColumnWidth(5, 200);  // fase
    sheet.setColumnWidth(8, 300);  // enunciado
    sheet.setColumnWidth(10, 250); // opcao_texto
  }
  return sheet;
}

// ============================================================
// TESTE MANUAL — rode uma vez para autorizar permissões
// ============================================================
function testar() {
  doPost({
    postData: {
      contents: JSON.stringify({
        timestamp: new Date().toISOString(),
        sessionId: 'TESTE',
        participantName: 'Teste',
        participantAge: '10',
        fase: 'Fase 1 – Jogo Online',
        tema: 'Consciência Emocional',
        questionIndex: 0,
        questionText: 'O que João está sentindo?',
        selectedIndex: 2,
        selectedText: 'Bravo / com raiva',
        isCorrect: true,
        pointsEarned: 2,
        responseTimeMs: 3200,
        coinsEarned: 100,
        starsEarned: 1,
        totalCoins: 100,
        totalStars: 1,
        totalScore: 2
      })
    }
  });
}

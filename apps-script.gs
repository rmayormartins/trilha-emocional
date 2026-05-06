// ============================================================
// TRILHA EMOCIONAL — Google Apps Script
// UMA LINHA POR JOGADOR — recebe tudo no final do jogo
//
// ===== SETUP =====
// 1. No Google Sheets → Extensões → Apps Script
// 2. Apague TODO o código → cole ESTE → salve (Ctrl+S)
// 3. No dropdown ao lado de ▶, selecione "testar" → ▶ Executar
//    → Autorize: Avançado → Ir para... → Permitir
//    → Confira a linha de teste na aba "Respostas"
// 4. Implantar → Nova implantação → Aplicativo da Web
//    → Executar como: EU MESMO
//    → Quem tem acesso: QUALQUER PESSOA  ← NÃO é "com conta Google"!
// 5. Copie a URL → cole no SHEETS_URL do index.html
//
// Mudou o código? → Implantar → Gerenciar implantações
//     → Editar → Versão: Nova versão → Implantar
// ============================================================

var SHEET_NAME = 'Respostas';
var NUM_QUESTIONS = 9;

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getOrCreateSheet();
    var letters = ['A','B','C','D','E','F'];

    // === Montar linha: Info + Q1..Q9 (opção + pontos) + Totais ===
    var row = [
      data.data || '',
      data.hora || '',
      data.nome || 'Anônimo',
      data.idade || '',
      data.sessionId || ''
    ];

    // Preparar arrays para Q1..Q9
    var qOpcoes = [];
    var qPontos = [];
    for (var i = 0; i < NUM_QUESTIONS; i++) {
      qOpcoes.push('');
      qPontos.push('');
    }

    // Preencher com as respostas recebidas
    if (data.respostas && data.respostas.length > 0) {
      for (var j = 0; j < data.respostas.length; j++) {
        var r = data.respostas[j];
        var idx = r.questionIndex;
        if (idx >= 0 && idx < NUM_QUESTIONS) {
          qOpcoes[idx] = letters[r.selectedIndex] || String(r.selectedIndex + 1);
          qPontos[idx] = (r.pointsEarned !== undefined) ? r.pointsEarned : '';
        }
      }
    }

    // Adicionar Q1_opcao, Q1_pontos, Q2_opcao, Q2_pontos, ...
    for (var k = 0; k < NUM_QUESTIONS; k++) {
      row.push(qOpcoes[k]);
      row.push(qPontos[k]);
    }

    // Totais
    var maxScore = NUM_QUESTIONS * 2;
    var score = data.totalScore || 0;
    var pct = maxScore > 0 ? Math.round(score / maxScore * 100) : 0;

    row.push(score);           // Score Total
    row.push(data.totalCoins || 0);  // Moedas
    row.push(data.totalStars || 0);  // Estrelas
    row.push(pct + '%');             // Percentual

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    console.error('doPost erro:', err);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput('✅ Trilha Emocional API funcionando! Use POST para enviar dados.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);

    // Montar cabeçalho
    var header = ['Data', 'Hora', 'Nome', 'Idade', 'SessionId'];
    for (var i = 1; i <= NUM_QUESTIONS; i++) {
      header.push('Q' + i + '_opção');
      header.push('Q' + i + '_pontos');
    }
    header.push('Score Total', 'Moedas', 'Estrelas', 'Percentual');
    sheet.appendRow(header);

    // Estilizar cabeçalho geral
    var lastCol = header.length;
    var headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    headerRange.setFontSize(9);
    sheet.setFrozenRows(1);

    // Cor do cabeçalho: info
    sheet.getRange(1, 1, 1, 5).setBackground('#3D2B5F').setFontColor('#fff');

    // Cor por fase: Q1-Q3 azul (Consciência), Q4-Q6 roxo (Regulação), Q7-Q9 verde (Empatia)
    var phaseColors = ['#5C6BC0', '#5C6BC0', '#5C6BC0', '#AB47BC', '#AB47BC', '#AB47BC', '#26A69A', '#26A69A', '#26A69A'];
    for (var i = 0; i < NUM_QUESTIONS; i++) {
      var col = 6 + i * 2; // coluna da Q(i+1)_opção
      sheet.getRange(1, col, 1, 2).setBackground(phaseColors[i]).setFontColor('#fff');
    }

    // Cor dos totais
    var totStart = 6 + NUM_QUESTIONS * 2;
    sheet.getRange(1, totStart, 1, 4).setBackground('#E91E63').setFontColor('#fff');

    // Larguras
    sheet.setColumnWidth(1, 95);   // Data
    sheet.setColumnWidth(2, 60);   // Hora
    sheet.setColumnWidth(3, 110);  // Nome
    sheet.setColumnWidth(4, 55);   // Idade
    sheet.setColumnWidth(5, 120);  // SessionId
    for (var i = 0; i < NUM_QUESTIONS; i++) {
      sheet.setColumnWidth(6 + i * 2, 55);      // opção
      sheet.setColumnWidth(6 + i * 2 + 1, 55);  // pontos
    }
    var t = 6 + NUM_QUESTIONS * 2;
    sheet.setColumnWidth(t, 75);     // Score
    sheet.setColumnWidth(t + 1, 70); // Moedas
    sheet.setColumnWidth(t + 2, 65); // Estrelas
    sheet.setColumnWidth(t + 3, 75); // Percentual
  }

  return sheet;
}

// ============================================================
// TESTE — selecione "testar" no dropdown e clique ▶ Executar
// ============================================================
function testar() {
  var respostas = [
    {questionIndex:0, selectedIndex:2, selectedText:'Bravo / com raiva', pointsEarned:2, isCorrect:true, responseTimeMs:3200},
    {questionIndex:1, selectedIndex:1, selectedText:'Pedir para ele se acalmar', pointsEarned:2, isCorrect:true, responseTimeMs:4100},
    {questionIndex:2, selectedIndex:1, selectedText:'Ficam irritados', pointsEarned:2, isCorrect:true, responseTimeMs:2800},
    {questionIndex:3, selectedIndex:2, selectedText:'Triste', pointsEarned:2, isCorrect:true, responseTimeMs:3500},
    {questionIndex:4, selectedIndex:2, selectedText:'Fingir que não aconteceu', pointsEarned:2, isCorrect:true, responseTimeMs:4200},
    {questionIndex:5, selectedIndex:0, selectedText:'Conversar para entender', pointsEarned:2, isCorrect:true, responseTimeMs:3000},
    {questionIndex:6, selectedIndex:2, selectedText:'Chateado', pointsEarned:2, isCorrect:true, responseTimeMs:2600},
    {questionIndex:7, selectedIndex:2, selectedText:'Fazer outra coisa que goste', pointsEarned:2, isCorrect:true, responseTimeMs:3800},
    {questionIndex:8, selectedIndex:1, selectedText:'Mandar mensagem apoiando', pointsEarned:2, isCorrect:true, responseTimeMs:2900}
  ];

  doPost({
    postData: {
      contents: JSON.stringify({
        sessionId: 'TESTE-' + Date.now(),
        data: new Date().toLocaleDateString('pt-BR'),
        hora: new Date().toLocaleTimeString('pt-BR', {hour:'2-digit', minute:'2-digit'}),
        nome: 'Maria Silva',
        idade: '10',
        respostas: respostas,
        totalScore: 18,
        totalCoins: 900,
        totalStars: 9
      })
    }
  });

  Logger.log('✅ Linha de teste inserida! Confira a aba "Respostas".');
}

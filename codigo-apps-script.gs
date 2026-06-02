// ============================================================
//  BOLÃO COPA 2026 — Google Apps Script
//  Cole este código em: Extensões > Apps Script da planilha
//  Depois: Implantar > Nova implantação > App da Web
//  Execute como: Você | Acesso: Qualquer pessoa
// ============================================================

var SHEET_ID       = '1vrXtyJVpkHrroZptjpnH-FrZ2OMgZMJTqUJxPcVKJl8';
var WHATSAPP_TOKEN = 'COLE_AQUI_SEU_TOKEN_CALLMEBOT';
var PRECO_UNIT     = 5.00;

var JOGOS = [
  { aba: 'Jogo1 - Brasil x Marrocos', nome: 'Brasil x Marrocos', time1: 'Brasil',  time2: 'Marrocos' },
  { aba: 'Jogo2 - Brasil x Haiti',    nome: 'Brasil x Haiti',    time1: 'Brasil',  time2: 'Haiti'    },
  { aba: 'Jogo3 - Escocia x Brasil',  nome: 'Escocia x Brasil',  time1: 'Escocia', time2: 'Brasil'   }
];

// Cabeçalho base — T1 e T2 são substituídos pelo nome dos times em cada aba
var CABECALHO_BASE = [
  'Codigo', 'Data/Hora', 'Nome', 'WhatsApp', 'Email',
  'GOL_T1', 'GOL_T2', 'Total Palpites', 'Total R$',
  'Comprovante?', 'Pagamento Confirmado?', 'Observacao'
];

function cabecalhoJogo(jogo) {
  return CABECALHO_BASE.map(function(c) {
    if (c === 'GOL_T1') return jogo.time1;
    if (c === 'GOL_T2') return jogo.time2;
    return c;
  });
}

// ── Inicializa abas da planilha ──────────────────────────────
function inicializarPlanilha() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  for (var i = 0; i < JOGOS.length; i++) {
    var jogo = JOGOS[i];
    var cab  = cabecalhoJogo(jogo);
    var aba  = ss.getSheetByName(jogo.aba);
    if (!aba) aba = ss.insertSheet(jogo.aba);
    // Apaga e recria cabeçalho sempre (para atualizar nomes dos times)
    if (aba.getLastRow() > 0) {
      aba.getRange(1, 1, 1, cab.length).setValues([cab]);
    } else {
      var h = aba.getRange(1, 1, 1, cab.length);
      h.setValues([cab]);
      h.setBackground('#002776').setFontColor('#ffffff').setFontWeight('bold');
      aba.setFrozenRows(1);
      aba.setColumnWidths(1, cab.length, 150);
    }
    // Garante formatação do cabeçalho
    aba.getRange(1, 1, 1, cab.length)
       .setBackground('#002776').setFontColor('#ffffff').setFontWeight('bold');
  }
  SpreadsheetApp.getUi().alert('Planilha inicializada com sucesso!');
}

// ── Recebe POST (novo palpite) ───────────────────────────────
function doPost(e) {
  try {
    var dados   = JSON.parse(e.postData.contents);
    var ss      = SpreadsheetApp.openById(SHEET_ID);
    var jogoIdx = parseInt(dados.jogoIndex);
    var jogo    = JOGOS[jogoIdx];
    if (!jogo) throw new Error('Jogo invalido');

    var aba = ss.getSheetByName(jogo.aba);
    if (!aba) {
      aba = ss.insertSheet(jogo.aba);
      aba.getRange(1, 1, 1, CABECALHO.length).setValues([CABECALHO]);
    }

    var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
    var linha = [
      dados.codigo        || '',
      agora,
      dados.nome          || '',
      dados.whatsapp      || '',
      dados.email         || '',
      dados.gol1 != null  ? dados.gol1 : '',
      dados.gol2 != null  ? dados.gol2 : '',
      dados.totalPalpites || 1,
      dados.totalValor    || PRECO_UNIT,
      'Pendente',
      'Nao',
      ''
    ];

    aba.appendRow(linha);
    var ul = aba.getLastRow();
    aba.getRange(ul, 1, 1, CABECALHO.length)
       .setBackground(ul % 2 === 0 ? '#f0f4ff' : '#ffffff');

    return resposta({ mensagem: 'Palpite registrado!' });
  } catch (err) {
    return falha(err.message);
  }
}

// ── Recebe GET (palpite, status ou ping) ─────────────────────
function doGet(e) {
  var action   = (e.parameter.action || '').toLowerCase();
  var callback = e.parameter.callback || '';

  var result;
  if (action === 'palpite') {
    result = salvarPalpite(e.parameter);
  } else if (action === 'status') {
    result = consultarStatus(e.parameter.codigo || '');
  } else {
    result = resposta({ status: 'online', bolao: 'Copa 2026' });
  }

  // Suporte a JSONP (contorna CORS)
  if (callback) {
    var json = result.getContent();
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return result;
}

// ── Salva palpite recebido via GET params ────────────────────
function salvarPalpite(p) {
  try {
    var ss      = SpreadsheetApp.openById(SHEET_ID);
    var jogoIdx = parseInt(p.jogoIndex);
    var jogo    = JOGOS[jogoIdx];
    if (!jogo) throw new Error('Jogo invalido: ' + jogoIdx);

    var cab = cabecalhoJogo(jogo);
    var aba = ss.getSheetByName(jogo.aba);
    if (!aba) {
      aba = ss.insertSheet(jogo.aba);
      var h = aba.getRange(1, 1, 1, cab.length);
      h.setValues([cab]);
      h.setBackground('#002776').setFontColor('#ffffff').setFontWeight('bold');
      aba.setFrozenRows(1);
      aba.setColumnWidths(1, cab.length, 150);
    }

    var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
    var linha = [
      p.codigo                                  || '',
      agora,
      p.nome                                    || '',
      p.whatsapp                                || '',
      p.email                                   || '',
      p.gol1 !== undefined ? parseInt(p.gol1) : '',
      p.gol2 !== undefined ? parseInt(p.gol2) : '',
      p.totalPalpites                           || 1,
      p.totalValor                              || PRECO_UNIT,
      'Pendente',
      'Nao',
      ''
    ];

    aba.appendRow(linha);
    var ul = aba.getLastRow();
    aba.getRange(ul, 1, 1, cab.length)
       .setBackground(ul % 2 === 0 ? '#f0f4ff' : '#ffffff');

    return resposta({ mensagem: 'Palpite registrado!' });
  } catch (err) {
    return falha(err.message);
  }
}

// ── Consulta status por codigo ───────────────────────────────
function consultarStatus(codigo) {
  codigo = codigo.trim().toUpperCase();
  if (!codigo) return falha('Codigo nao informado');

  var ss = SpreadsheetApp.openById(SHEET_ID);

  for (var i = 0; i < JOGOS.length; i++) {
    var jogo = JOGOS[i];
    var aba  = ss.getSheetByName(jogo.aba);
    if (!aba || aba.getLastRow() < 2) continue;

    var dados = aba.getRange(2, 1, aba.getLastRow() - 1, CABECALHO.length).getValues();
    for (var r = 0; r < dados.length; r++) {
      var linha = dados[r];
      if ((linha[0] || '').toString().trim().toUpperCase() === codigo) {
        var confirmado = (linha[10] || '').toString().toLowerCase() === 'sim';
        return resposta({
          encontrado: true,
          confirmado: confirmado,
          codigo:     linha[0],
          jogo:       jogo.nome,
          nome:       linha[2],
          gol1:       linha[5],
          gol2:       linha[6]
        });
      }
    }
  }

  return resposta({ encontrado: false });
}

// ── Trigger: notifica participante ao confirmar pagamento ────
// Configure: Gatilhos (icone relogio) > onEdit > Ao editar
function onEdit(e) {
  try {
    var sheet = e.range.getSheet();
    var col   = e.range.getColumn();
    var row   = e.range.getRow();
    var valor = (e.value || '').toString().toLowerCase();

    // Coluna 11 = Pagamento Confirmado | deve ser "sim"
    if (col !== 11 || row < 2 || valor !== 'sim') return;

    var dados    = sheet.getRange(row, 1, 1, CABECALHO.length).getValues()[0];
    var codigo   = dados[0];
    var nome     = dados[2];
    var whatsapp = dados[3];
    var gol1     = dados[5];
    var gol2     = dados[6];

    var nomJogo  = '';
    for (var i = 0; i < JOGOS.length; i++) {
      if (JOGOS[i].aba === sheet.getName()) { nomJogo = JOGOS[i].nome; break; }
    }

    var msg = encodeURIComponent(
      'BOLAO COPA 2026 - Confirmacao\n\n' +
      'Ola, ' + nome + '!\n\n' +
      'Seu palpite foi CONFIRMADO!\n\n' +
      'Codigo: ' + codigo + '\n' +
      'Jogo: ' + nomJogo + '\n' +
      'Palpite: ' + gol1 + ' x ' + gol2 + '\n\n' +
      'Boa sorte! Vai Brasil!'
    );

    var num = (whatsapp || '').replace(/\D/g, '');
    if (num.length === 11) num = '55' + num;
    if (num.length === 10) num = '55' + num;

    if (num && WHATSAPP_TOKEN !== 'COLE_AQUI_SEU_TOKEN_CALLMEBOT') {
      var url = 'https://api.callmebot.com/whatsapp.php?phone=' + num + '&text=' + msg + '&apikey=' + WHATSAPP_TOKEN;
      UrlFetchApp.fetch(url);
    }
  } catch (err) {
    console.error('Erro no onEdit: ' + err.message);
  }
}

// ── Helpers ──────────────────────────────────────────────────
function resposta(obj) {
  obj.sucesso = true;
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function falha(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ sucesso: false, erro: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

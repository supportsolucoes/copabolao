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

  if (action === 'palpite') {
    var r = salvarPalpite(e.parameter);
    if (callback) return jsonp(callback, r);
    return r;
  }

  if (action === 'status') {
    var codigo = e.parameter.codigo || '';
    var dados  = consultarStatusDados(codigo);
    var html;

    if (!dados.encontrado) {
      html = paginaStatus('❌ Código não encontrado',
        'O código <strong>' + codigo + '</strong> não foi encontrado.<br>Verifique se digitou corretamente.',
        '#fef2f2', '#dc2626');
    } else if (dados.confirmado) {
      html = paginaStatus('✅ Pagamento Confirmado!',
        'Participante: <strong>' + dados.nome + '</strong><br>' +
        'Jogo: <strong>' + dados.jogo + '</strong><br>' +
        'Palpite: <strong>' + dados.gol1 + ' × ' + dados.gol2 + '</strong><br>' +
        'Código: <strong>' + dados.codigo + '</strong>',
        '#f0fdf4', '#15803d');
    } else {
      html = paginaStatus('⏳ Aguardando Confirmação',
        'Participante: <strong>' + dados.nome + '</strong><br>' +
        'Jogo: <strong>' + dados.jogo + '</strong><br>' +
        'Palpite: <strong>' + dados.gol1 + ' × ' + dados.gol2 + '</strong><br>' +
        'Código: <strong>' + dados.codigo + '</strong><br><br>' +
        'Seu comprovante ainda não foi confirmado.<br>Aguarde até 1 hora após o envio.',
        '#fff8e1', '#7a5800');
    }
    return HtmlService.createHtmlOutput(html)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return resposta({ status: 'online', bolao: 'Copa 2026' });
}

function jsonp(callback, result) {
  return ContentService
    .createTextOutput(callback + '(' + result.getContent() + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ── Retorna dados de status como objeto JS ───────────────────
function consultarStatusDados(codigo) {
  codigo = (codigo || '').trim().toUpperCase();
  if (!codigo) return { sucesso: false, encontrado: false };

  var ss = SpreadsheetApp.openById(SHEET_ID);
  for (var i = 0; i < JOGOS.length; i++) {
    var jogo = JOGOS[i];
    var aba  = ss.getSheetByName(jogo.aba);
    if (!aba || aba.getLastRow() < 2) continue;
    var dados = aba.getRange(2, 1, aba.getLastRow() - 1, 12).getValues();
    for (var r = 0; r < dados.length; r++) {
      var linha = dados[r];
      if ((linha[0] || '').toString().trim().toUpperCase() === codigo) {
        var confirmado = (linha[10] || '').toString().toLowerCase() === 'sim';
        return {
          sucesso: true, encontrado: true, confirmado: confirmado,
          codigo: linha[0], jogo: jogo.nome,
          nome: linha[2], gol1: linha[5], gol2: linha[6]
        };
      }
    }
  }
  return { sucesso: true, encontrado: false };
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

// ── Gera página HTML de status ───────────────────────────────
function paginaStatus(titulo, corpo, bgColor, textColor) {
  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Bolão Copa 2026 — Status</title>' +
    '<style>*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:Arial,sans-serif;background:#f5f6fa;display:flex;align-items:center;justify-content:center;min-height:100vh;padding:1rem}' +
    '.card{background:#fff;border-radius:16px;padding:2rem;max-width:420px;width:100%;text-align:center;box-shadow:0 4px 24px rgba(0,0,0,.1)}' +
    '.badge{background:' + bgColor + ';color:' + textColor + ';border-radius:12px;padding:1.2rem;margin-bottom:1.2rem}' +
    '.titulo{font-size:1.4rem;font-weight:700;margin-bottom:.75rem;color:' + textColor + '}' +
    '.corpo{font-size:.95rem;line-height:1.7;color:#374151}' +
    '.btn{display:inline-block;margin-top:1.5rem;background:#002776;color:#fff;border:none;border-radius:10px;padding:.75rem 1.5rem;font-size:1rem;font-weight:600;cursor:pointer;text-decoration:none}' +
    '.logo{font-size:2rem;margin-bottom:.5rem}' +
    '</style></head><body>' +
    '<div class="card">' +
    '<div class="logo">⚽🇧🇷</div>' +
    '<div style="font-family:Arial;font-size:1.1rem;font-weight:700;color:#002776;letter-spacing:1px;margin-bottom:1rem">BOLÃO COPA 2026</div>' +
    '<div class="badge"><div class="titulo">' + titulo + '</div><div class="corpo">' + corpo + '</div></div>' +
    '<a class="btn" href="javascript:window.close()">Fechar</a>' +
    '</div></body></html>';
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

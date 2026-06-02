// ============================================================
//  BOLÃO COPA 2026 — Google Apps Script
//  Cole em: Extensões > Apps Script da planilha
//  Implantar > App da Web | Execute como: Você | Acesso: Qualquer pessoa
// ============================================================

var SHEET_ID   = '1vrXtyJVpkHrroZptjpnH-FrZ2OMgZMJTqUJxPcVKJl8';
var PRECO_UNIT = 5.00;

var JOGOS = [
  { aba: 'Jogo1 - Brasil x Marrocos', nome: 'Brasil x Marrocos', time1: 'Brasil',  time2: 'Marrocos' },
  { aba: 'Jogo2 - Brasil x Haiti',    nome: 'Brasil x Haiti',    time1: 'Brasil',  time2: 'Haiti'    },
  { aba: 'Jogo3 - Escocia x Brasil',  nome: 'Escocia x Brasil',  time1: 'Escocia', time2: 'Brasil'   }
];

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

// ── Inicializa planilha + aba Resultados ─────────────────────
function inicializarPlanilha() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // Abas dos jogos
  for (var i = 0; i < JOGOS.length; i++) {
    var jogo = JOGOS[i];
    var cab  = cabecalhoJogo(jogo);
    var aba  = ss.getSheetByName(jogo.aba);
    if (!aba) aba = ss.insertSheet(jogo.aba);
    var h = aba.getRange(1, 1, 1, cab.length);
    h.setValues([cab]);
    h.setBackground('#002776').setFontColor('#ffffff').setFontWeight('bold');
    aba.setFrozenRows(1);
    aba.setColumnWidths(1, cab.length, 150);
  }

  // Aba Resultados
  var abaRes = ss.getSheetByName('Resultados');
  if (!abaRes) abaRes = ss.insertSheet('Resultados');
  if (abaRes.getLastRow() === 0) {
    var cabRes = [['Jogo', 'Time1', 'Gol_Time1', 'Gol_Time2', 'Time2', 'Premio_Acumulado']];
    abaRes.getRange(1, 1, 1, 6).setValues(cabRes)
      .setBackground('#009c3b').setFontColor('#ffffff').setFontWeight('bold');
    abaRes.setColumnWidths(1, 6, 160);
    // Linhas iniciais para cada jogo
    abaRes.appendRow(['Brasil x Marrocos', 'Brasil', '', '', 'Marrocos', '']);
    abaRes.appendRow(['Brasil x Haiti',    'Brasil', '', '', 'Haiti',    '']);
    abaRes.appendRow(['Escocia x Brasil',  'Escocia','', '', 'Brasil',   '']);
  }

  SpreadsheetApp.getUi().alert('Planilha inicializada! Preencha os resultados na aba "Resultados" após cada jogo.');
}

// ── GET — roteador principal ──────────────────────────────────
function doGet(e) {
  var action = (e.parameter.action || '').toLowerCase();

  if (action === 'palpite')   return salvarPalpite(e.parameter);
  if (action === 'dashboard') return getDashboard();

  return resposta({ status: 'online', bolao: 'Copa 2026' });
}

// ── Salva palpite ─────────────────────────────────────────────
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
    }

    var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
    aba.appendRow([
      p.codigo || '',
      agora,
      p.nome || '',
      p.whatsapp || '',
      p.email || '',
      p.gol1 !== undefined ? parseInt(p.gol1) : '',
      p.gol2 !== undefined ? parseInt(p.gol2) : '',
      p.totalPalpites || 1,
      p.totalValor || PRECO_UNIT,
      'Pendente',
      'Nao',
      ''
    ]);

    var ul = aba.getLastRow();
    aba.getRange(ul, 1, 1, cab.length)
       .setBackground(ul % 2 === 0 ? '#f0f4ff' : '#ffffff');

    return resposta({ mensagem: 'Palpite registrado!' });
  } catch (err) {
    return falha(err.message);
  }
}

// ── Dashboard: prêmio acumulado + ranking ─────────────────────
function getDashboard() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);

    // ── 1. Prêmio acumulado: soma coluna I (índice 8) de todos os jogos ──
    var totalPremio = 0;
    var totalPalpites = 0;
    for (var i = 0; i < JOGOS.length; i++) {
      var aba = ss.getSheetByName(JOGOS[i].aba);
      if (!aba || aba.getLastRow() < 2) continue;
      var vals = aba.getRange(2, 9, aba.getLastRow() - 1, 1).getValues();
      for (var r = 0; r < vals.length; r++) {
        var v = parseFloat(vals[r][0]);
        if (!isNaN(v)) { totalPremio += v; totalPalpites++; }
      }
    }

    // ── 2. Resultados + ranking por jogo ─────────────────────
    var abaRes = ss.getSheetByName('Resultados');
    var ranking = [];

    if (abaRes && abaRes.getLastRow() > 1) {
      var resData = abaRes.getRange(2, 1, abaRes.getLastRow() - 1, 6).getValues();

      for (var ri = 0; ri < resData.length; ri++) {
        var linha   = resData[ri];
        var nomeJogo = (linha[0] || '').toString().trim();
        var g1Real   = linha[2] !== '' ? parseInt(linha[2]) : null;
        var g2Real   = linha[3] !== '' ? parseInt(linha[3]) : null;
        var premio   = linha[5] !== '' ? parseFloat(linha[5]) : null;

        // Só processa jogo com resultado preenchido
        if (g1Real === null || g2Real === null || isNaN(g1Real) || isNaN(g2Real)) continue;

        // Busca palpites do jogo correspondente
        var jogoRef = null;
        for (var ji = 0; ji < JOGOS.length; ji++) {
          if (JOGOS[ji].nome === nomeJogo) { jogoRef = JOGOS[ji]; break; }
        }
        if (!jogoRef) continue;

        var abaJogo = ss.getSheetByName(jogoRef.aba);
        if (!abaJogo || abaJogo.getLastRow() < 2) continue;

        var palpites = abaJogo.getRange(2, 1, abaJogo.getLastRow() - 1, 12).getValues();
        var acertaram = [];

        for (var pi = 0; pi < palpites.length; pi++) {
          var pal = palpites[pi];
          var pgto = (pal[10] || '').toString().toLowerCase();
          if (pgto !== 'sim') continue; // só conta quem pagou
          var p1 = parseInt(pal[5]);
          var p2 = parseInt(pal[6]);
          if (p1 === g1Real && p2 === g2Real) {
            acertaram.push({ nome: pal[2] });
          }
        }

        ranking.push({
          jogo:      nomeJogo,
          resultado: g1Real + ' x ' + g2Real,
          acertaram: acertaram,
          premio:    premio
        });
      }
    }

    return resposta({
      totalPremio:   totalPremio.toFixed(2),
      totalPalpites: totalPalpites,
      ranking:       ranking
    });

  } catch(err) {
    return falha(err.message);
  }
}

// ── POST (legado) ─────────────────────────────────────────────
function doPost(e) {
  try {
    var dados = JSON.parse(e.postData.contents);
    return salvarPalpite(dados);
  } catch(err) {
    return falha(err.message);
  }
}

// ── Helpers ───────────────────────────────────────────────────
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

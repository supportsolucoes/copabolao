// ============================================================
//  BOLÃO COPA 2026 — Google Apps Script
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

function inicializarPlanilha() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
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
  var abaRes = ss.getSheetByName('Resultados');
  if (!abaRes) abaRes = ss.insertSheet('Resultados');
  if (abaRes.getLastRow() === 0) {
    var cabRes = [['Jogo','Time1','Gol_Time1','Gol_Time2','Time2','Premio_Acumulado','Status','Premio_Pago']];
    abaRes.getRange(1, 1, 1, 8).setValues(cabRes)
      .setBackground('#009c3b').setFontColor('#ffffff').setFontWeight('bold');
    abaRes.setColumnWidths(1, 8, 160);
    abaRes.appendRow(['Brasil x Marrocos','Brasil','','','Marrocos','','','']);
    abaRes.appendRow(['Brasil x Haiti',   'Brasil','','','Haiti',   '','','']);
    abaRes.appendRow(['Escocia x Brasil', 'Escocia','','','Brasil',  '','','']);
  }
  SpreadsheetApp.getUi().alert('Planilha inicializada!');
}

function doGet(e) {
  var action = (e.parameter.action || '').toLowerCase();
  if (action === 'palpite')    return salvarPalpite(e.parameter);
  if (action === 'dashboard')  return getDashboard();
  if (action === 'palpites')   return getPalpites();
  if (action === 'ver')        return getPaginaPalpites(e.parameter);
  if (action === 'admin')      return getDadosAdmin();
  if (action === 'pgto')       return confirmarPagamento(e.parameter);
  if (action === 'resultado')  return salvarResultado(e.parameter);
  if (action === 'premiopago') return marcarPremioPago(e.parameter);
  return resposta({ status: 'online', bolao: 'Copa 2026' });
}

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
      p.codigo || '', agora, p.nome || '', p.whatsapp || '', p.email || '',
      p.gol1 !== undefined ? parseInt(p.gol1) : '',
      p.gol2 !== undefined ? parseInt(p.gol2) : '',
      p.totalPalpites || 1, PRECO_UNIT, 'Pendente', 'Nao', ''
    ]);
    var ul = aba.getLastRow();
    aba.getRange(ul, 1, 1, cab.length).setBackground(ul % 2 === 0 ? '#f0f4ff' : '#ffffff');
    return resposta({ mensagem: 'Palpite registrado!' });
  } catch (err) { return falha(err.message); }
}

function lerResultados(ss) {
  var resultados = [];
  var abaRes = ss.getSheetByName('Resultados');
  if (!abaRes || abaRes.getLastRow() < 2) return resultados;
  var numCols = abaRes.getLastColumn();
  var resData = abaRes.getRange(2, 1, abaRes.getLastRow() - 1, Math.max(numCols, 8)).getValues();
  for (var ri = 0; ri < resData.length; ri++) {
    var rg1        = resData[ri][2];
    var rg2        = resData[ri][3];
    var status     = (resData[ri][6] || '').toString().toLowerCase() || 'parcial';
    var premioPago = (resData[ri][7] || '').toString().toLowerCase() === 'sim';
    resultados.push(rg1 !== '' && rg2 !== ''
      ? { g1: parseInt(rg1), g2: parseInt(rg2), status: status, premioPago: premioPago }
      : {});
  }
  return resultados;
}

function calcularPremioAtual(ss) {
  var resultados = lerResultados(ss);
  var ultimoPago = -1;
  for (var i = 0; i < resultados.length; i++) {
    if (resultados[i] && resultados[i].premioPago) ultimoPago = i;
  }
  var qtd = 0;
  for (var i = 0; i < JOGOS.length; i++) {
    if (i <= ultimoPago) continue;
    var aba = ss.getSheetByName(JOGOS[i].aba);
    if (!aba || aba.getLastRow() < 2) continue;
    qtd += aba.getLastRow() - 1;
  }
  return { total: qtd * PRECO_UNIT, qtd: qtd };
}

function getDashboard() {
  try {
    var ss          = SpreadsheetApp.openById(SHEET_ID);
    var premioAtual = calcularPremioAtual(ss);
    var resultados  = lerResultados(ss);
    var ranking     = [];
    for (var ri = 0; ri < resultados.length; ri++) {
      var r = resultados[ri];
      if (!r || r.g1 === undefined) continue;
      var jogoRef = JOGOS[ri];
      var abaJogo = ss.getSheetByName(jogoRef.aba);
      if (!abaJogo || abaJogo.getLastRow() < 2) continue;
      var palpites  = abaJogo.getRange(2, 1, abaJogo.getLastRow() - 1, 12).getValues();
      var acertaram = [];
      for (var pi = 0; pi < palpites.length; pi++) {
        var pal  = palpites[pi];
        var pgto = (pal[10] || '').toString().toLowerCase();
        if (pgto !== 'sim') continue;
        if (parseInt(pal[5]) === r.g1 && parseInt(pal[6]) === r.g2) acertaram.push({ nome: pal[2] });
      }
      ranking.push({ jogo: jogoRef.nome, resultado: r.g1 + ' x ' + r.g2, acertaram: acertaram, premio: premioAtual.total });
    }
    var historico = [];
    for (var i = 0; i < JOGOS.length; i++) {
      var aba = ss.getSheetByName(JOGOS[i].aba);
      var qtdJogo = (aba && aba.getLastRow() > 1) ? aba.getLastRow() - 1 : 0;
      historico.push({ qtd: qtdJogo, premio: qtdJogo * PRECO_UNIT });
    }
    return resposta({
      totalPremio:   premioAtual.total.toFixed(2),
      totalPalpites: premioAtual.qtd,
      ranking:       ranking,
      resultados:    resultados,
      historico:     historico,
      palpites:      getPalpitesDados(ss)
    });
  } catch(err) { return falha(err.message); }
}

function getPalpites() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    return resposta({ palpites: getPalpitesDados(ss) });
  } catch(err) { return falha(err.message); }
}

function getPalpitesDados(ss) {
  var resultado = [];
  for (var i = 0; i < JOGOS.length; i++) {
    var jogo  = JOGOS[i];
    var aba   = ss.getSheetByName(jogo.aba);
    var itens = [];
    if (aba && aba.getLastRow() > 1) {
      var dados = aba.getRange(2, 1, aba.getLastRow() - 1, 7).getValues();
      for (var r = 0; r < dados.length; r++) {
        var nome = (dados[r][2] || '').toString().trim();
        var g1   = dados[r][5];
        var g2   = dados[r][6];
        if (nome && g1 !== '' && g2 !== '') itens.push({ nome: nome, g1: g1, g2: g2 });
      }
    }
    resultado.push({ jogo: jogo.nome, itens: itens });
  }
  return resultado;
}

function getPaginaPalpites(p) {
  var ss      = SpreadsheetApp.openById(SHEET_ID);
  var data    = getPalpitesDados(ss);
  var tabInicial = p && p.tab !== undefined ? parseInt(p.tab) : 0;
  if (isNaN(tabInicial) || tabInicial < 0 || tabInicial >= data.length) tabInicial = 0;
  var abas = data.map(function(jogo, i) {
    return '<button class="tab' + (i===tabInicial?' ativo':'') + '" onclick="trocar(' + i + ')">' + jogo.jogo + '</button>';
  }).join('');
  var conteudos = data.map(function(jogo, i) {
    var lista = jogo.itens.length === 0
      ? '<div class="vazio">Nenhum palpite ainda.</div>'
      : jogo.itens.map(function(p, n) {
          return '<div class="linha"><span class="nome">'+(n+1)+'. '+p.nome+'</span><span class="placar">'+p.g1+' × '+p.g2+'</span></div>';
        }).join('');
    return '<div class="conteudo" id="c'+i+'" style="display:'+(i===tabInicial?'block':'none')+'">'+lista+'</div>';
  }).join('');
  var html = '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1"><title>Palpites do Grupo</title><style>' +
    '*{box-sizing:border-box;margin:0;padding:0}body{font-family:Arial,sans-serif;background:#f5f6fa;padding:1rem}' +
    'h2{font-size:1.2rem;color:#002776;margin-bottom:1rem}.tabs{display:flex;gap:6px;margin-bottom:1rem;flex-wrap:wrap}' +
    '.tab{padding:6px 12px;border:none;border-radius:8px;background:#e0e3ee;color:#6b7280;font-size:.82rem;font-weight:600;cursor:pointer}' +
    '.tab.ativo{background:#002776;color:white}' +
    '.linha{display:flex;justify-content:space-between;align-items:center;background:white;border-radius:8px;padding:.6rem .85rem;margin-bottom:6px;font-size:.88rem}' +
    '.nome{font-weight:600;color:#1a1a2e}.placar{font-family:monospace;font-size:1rem;font-weight:700;color:#002776;background:#f0f4ff;border-radius:6px;padding:2px 10px}' +
    '.vazio{text-align:center;color:#6b7280;padding:2rem;font-size:.85rem}' +
    '</style></head><body><h2>⚽ Palpites do Grupo</h2><div class="tabs">'+abas+'</div>'+conteudos+
    '<script>function trocar(i){document.querySelectorAll(".tab").forEach(function(t,j){t.classList.toggle("ativo",j===i)});document.querySelectorAll(".conteudo").forEach(function(c,j){c.style.display=j===i?"block":"none"});}<\/script></body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDadosAdmin() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var totalPremio = 0, totalPalpites = 0, totalPendentes = 0;
    var palpites = [];
    for (var i = 0; i < JOGOS.length; i++) {
      var jogo = JOGOS[i];
      var aba  = ss.getSheetByName(jogo.aba);
      var itens = [];
      if (aba && aba.getLastRow() > 1) {
        var dados = aba.getRange(2, 1, aba.getLastRow() - 1, 12).getValues();
        for (var r = 0; r < dados.length; r++) {
          var linha = dados[r];
          totalPremio += PRECO_UNIT;
          totalPalpites++;
          var pgto = (linha[10] || '').toString().toLowerCase();
          if (pgto !== 'sim') totalPendentes++;
          itens.push({ codigo: linha[0], nome: linha[2], whats: linha[3], g1: linha[5], g2: linha[6], comp: linha[9], pgto: pgto });
        }
      }
      palpites.push({ jogo: jogo.nome, itens: itens });
    }
    var resultados = lerResultados(ss);
    var ranking = [];
    for (var ji = 0; ji < JOGOS.length; ji++) {
      var res = resultados[ji];
      if (!res || res.g1 === undefined) continue;
      var aba2 = ss.getSheetByName(JOGOS[ji].aba);
      if (!aba2 || aba2.getLastRow() < 2) continue;
      var dados2 = aba2.getRange(2, 1, aba2.getLastRow() - 1, 12).getValues();
      var acertaram = [];
      for (var ri2 = 0; ri2 < dados2.length; ri2++) {
        var l = dados2[ri2];
        if ((l[10] || '').toString().toLowerCase() !== 'sim') continue;
        if (parseInt(l[5]) === res.g1 && parseInt(l[6]) === res.g2) acertaram.push({ nome: l[2], p1: l[5], p2: l[6] });
      }
      ranking.push({ jogo: JOGOS[ji].nome, resultado: res.g1 + ' × ' + res.g2, acertaram: acertaram });
    }
    return resposta({ totalPremio: totalPremio.toFixed(2), totalPalpites: totalPalpites, totalPendentes: totalPendentes, palpites: palpites, resultados: resultados, ranking: ranking });
  } catch(err) { return falha(err.message); }
}

function confirmarPagamento(p) {
  try {
    var ss   = SpreadsheetApp.openById(SHEET_ID);
    var jogo = JOGOS[parseInt(p.jogoIndex)];
    var status = (p.status || 'nao').toLowerCase() === 'sim' ? 'Sim' : 'Nao';
    var aba  = ss.getSheetByName(jogo.aba);
    if (!aba) return falha('Aba não encontrada');
    var dados = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getValues();
    for (var r = 0; r < dados.length; r++) {
      if ((dados[r][0] || '').toString() === p.codigo) {
        aba.getRange(r + 2, 11).setValue(status);
        return resposta({ mensagem: 'Pagamento atualizado!' });
      }
    }
    return falha('Código não encontrado');
  } catch(err) { return falha(err.message); }
}

function salvarResultado(p) {
  try {
    var ss      = SpreadsheetApp.openById(SHEET_ID);
    var jogoIdx = parseInt(p.jogoIndex);
    var abaRes  = ss.getSheetByName('Resultados');
    if (!abaRes) return falha('Aba Resultados não encontrada');
    if (abaRes.getLastColumn() < 7) abaRes.getRange(1, 7).setValue('Status');
    if (abaRes.getLastColumn() < 8) abaRes.getRange(1, 8).setValue('Premio_Pago');
    abaRes.getRange(jogoIdx + 2, 3).setValue(parseInt(p.g1));
    abaRes.getRange(jogoIdx + 2, 4).setValue(parseInt(p.g2));
    abaRes.getRange(jogoIdx + 2, 7).setValue(p.status || 'parcial');
    return resposta({ mensagem: 'Resultado salvo!' });
  } catch(err) { return falha(err.message); }
}

function marcarPremioPago(p) {
  try {
    var ss      = SpreadsheetApp.openById(SHEET_ID);
    var jogoIdx = parseInt(p.jogoIndex);
    var abaRes  = ss.getSheetByName('Resultados');
    if (!abaRes) return falha('Aba Resultados não encontrada');
    if (abaRes.getLastColumn() < 8) abaRes.getRange(1, 8).setValue('Premio_Pago');
    abaRes.getRange(jogoIdx + 2, 8).setValue('Sim');
    return resposta({ mensagem: 'Premio marcado como pago!' });
  } catch(err) { return falha(err.message); }
}

function doPost(e) {
  try {
    var dados = JSON.parse(e.postData.contents);
    return salvarPalpite(dados);
  } catch(err) { return falha(err.message); }
}

function resposta(obj) {
  obj.sucesso = true;
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function falha(msg) {
  return ContentService.createTextOutput(JSON.stringify({ sucesso: false, erro: msg })).setMimeType(ContentService.MimeType.JSON);
}

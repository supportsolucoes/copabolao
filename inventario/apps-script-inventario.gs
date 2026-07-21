// ============================================================
//  INVENTARIO DE MAQUINAS — Google Apps Script
// ============================================================

var SHEET_ID = 'COLE_AQUI_O_ID_DA_SUA_PLANILHA';
var ABA      = 'Inventario';

var CABECALHO = [
  'Data/Hora', 'Hostname', 'Usuario', 'Dominio/Grupo',
  'Fabricante', 'Modelo', 'Numero de Serie', 'Patrimonio',
  'CPU', 'Nucleos', 'RAM (GB)', 'Discos', 'Placa de Video',
  'Sistema Operacional', 'Versao SO', 'IP(s)', 'MAC(s)',
  'Qtd Softwares', 'JSON Completo'
];

function inicializarPlanilha() {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var aba = ss.getSheetByName(ABA);
  if (!aba) aba = ss.insertSheet(ABA);
  var h = aba.getRange(1, 1, 1, CABECALHO.length);
  h.setValues([CABECALHO]);
  h.setBackground('#0f2557').setFontColor('#ffffff').setFontWeight('bold');
  aba.setFrozenRows(1);
  aba.setColumnWidths(1, CABECALHO.length, 140);
  SpreadsheetApp.getUi().alert('Planilha de inventario inicializada!');
}

function getOuCriarAba(ss) {
  var aba = ss.getSheetByName(ABA);
  if (!aba) {
    aba = ss.insertSheet(ABA);
    var h = aba.getRange(1, 1, 1, CABECALHO.length);
    h.setValues([CABECALHO]);
    h.setBackground('#0f2557').setFontColor('#ffffff').setFontWeight('bold');
    aba.setFrozenRows(1);
  }
  return aba;
}

function encontrarLinhaPorHostname(aba, hostname) {
  if (!hostname || aba.getLastRow() < 2) return -1;
  var col = aba.getRange(2, 2, aba.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    if ((col[i][0] || '').toString() === hostname) return i + 2;
  }
  return -1;
}

function doGet(e) {
  var action = (e.parameter.action || '').toLowerCase();
  if (action === 'inventario') return getInventario();
  return resposta({ status: 'online', app: 'Inventario de Maquinas' });
}

function doPost(e) {
  try {
    var dados = JSON.parse(e.postData.contents);
    return salvarInventario(dados);
  } catch (err) { return falha(err.message); }
}

function salvarInventario(d) {
  try {
    var ss  = SpreadsheetApp.openById(SHEET_ID);
    var aba = getOuCriarAba(ss);
    var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');

    var discos = (d.discos || []).map(function (x) {
      return x.modelo + ' (' + x.tamanhoGB + 'GB' + (x.tipo && x.tipo !== 'N/A' ? ', ' + x.tipo : '') + ')';
    }).join(' | ');
    var gpus = (d.gpus || []).map(function (x) { return x.nome; }).join(' | ');
    var ips  = (d.rede || []).map(function (x) { return x.ip; }).filter(Boolean).join(', ');
    var macs = (d.rede || []).map(function (x) { return x.mac; }).filter(Boolean).join(', ');

    var linha = [
      agora, d.hostname || '', d.usuarioLogado || '', d.dominio || '',
      d.fabricante || '', d.modelo || '', d.numeroSerie || '', d.patrimonio || '',
      d.cpu ? d.cpu.nome : '', d.cpu ? d.cpu.nucleos : '', d.ramTotalGB || '',
      discos, gpus,
      d.sistemaOperacional ? d.sistemaOperacional.nome : '',
      d.sistemaOperacional ? d.sistemaOperacional.versao : '',
      ips, macs, (d.softwares || []).length, JSON.stringify(d)
    ];

    var linhaExistente = encontrarLinhaPorHostname(aba, d.hostname);
    if (linhaExistente > 0) {
      aba.getRange(linhaExistente, 1, 1, CABECALHO.length).setValues([linha]);
    } else {
      aba.appendRow(linha);
      var ul = aba.getLastRow();
      aba.getRange(ul, 1, 1, CABECALHO.length).setBackground(ul % 2 === 0 ? '#f0f4ff' : '#ffffff');
    }
    return resposta({ mensagem: 'Inventario registrado com sucesso!' });
  } catch (err) { return falha(err.message); }
}

function getInventario() {
  try {
    var ss  = SpreadsheetApp.openById(SHEET_ID);
    var aba = ss.getSheetByName(ABA);
    var maquinas = [];
    if (aba && aba.getLastRow() > 1) {
      var dados = aba.getRange(2, 1, aba.getLastRow() - 1, CABECALHO.length).getValues();
      for (var i = 0; i < dados.length; i++) {
        var linha = dados[i];
        var completo = {};
        try { completo = JSON.parse(linha[18]); } catch (e) {}
        maquinas.push({
          atualizadoEm: linha[0], hostname: linha[1], usuario: linha[2], dominio: linha[3],
          fabricante: linha[4], modelo: linha[5], numeroSerie: linha[6], patrimonio: linha[7],
          cpu: linha[8], nucleos: linha[9], ramGB: linha[10], discosResumo: linha[11],
          gpuResumo: linha[12], sistemaOperacional: linha[13], versaoSO: linha[14],
          ips: linha[15], macs: linha[16], qtdSoftwares: linha[17], detalhes: completo
        });
      }
    }
    return resposta({ maquinas: maquinas });
  } catch (err) { return falha(err.message); }
}

function resposta(obj) {
  obj.sucesso = true;
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function falha(msg) {
  return ContentService.createTextOutput(JSON.stringify({ sucesso: false, erro: msg })).setMimeType(ContentService.MimeType.JSON);
}

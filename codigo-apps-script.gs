// ============================================================
//  BOLÃO COPA 2026 — Google Apps Script
//  Cole este código em: script.google.com > Novo projeto
//  Depois: Implantar > Novo implantação > App da Web
//  Acesso: Qualquer pessoa
//  Execute como: Você
// ============================================================

const SHEET_ID = 'COLE_AQUI_O_ID_DA_SUA_PLANILHA'; // Ex: '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms'
const WHATSAPP_ADMIN = '5511999999999'; // Seu número com DDI+DDD, sem espaços

// Cabeçalhos de cada aba (jogo)
const JOGOS = [
  { aba: 'Jogo1 - Brasil x Marrocos',  data: '13/06/2026', hora: '19:00' },
  { aba: 'Jogo2 - Brasil x Haiti',     data: '19/06/2026', hora: '21:30' },
  { aba: 'Jogo3 - Escócia x Brasil',   data: '24/06/2026', hora: '19:00' },
];

const CABECALHO = [
  'Código',
  'Data/Hora Inscrição',
  'Nome',
  'WhatsApp',
  'E-mail',
  'Palpite (Time 1)',
  'Palpite (Time 2)',
  'Comprovante Enviado?',
  'Pagamento Confirmado?',
  'Observação',
];

// ─── Inicializa planilha criando abas se não existirem ───────
function inicializarPlanilha() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  JOGOS.forEach(jogo => {
    let aba = ss.getSheetByName(jogo.aba);
    if (!aba) {
      aba = ss.insertSheet(jogo.aba);
    }
    // Só escreve cabeçalho se a aba estiver vazia
    if (aba.getLastRow() === 0) {
      const headerRow = aba.getRange(1, 1, 1, CABECALHO.length);
      headerRow.setValues([CABECALHO]);
      headerRow.setBackground('#002776');
      headerRow.setFontColor('#ffffff');
      headerRow.setFontWeight('bold');
      aba.setFrozenRows(1);
      aba.setColumnWidth(1, 160);
      aba.setColumnWidth(2, 180);
      aba.setColumnWidth(3, 140);
      aba.setColumnWidth(4, 200);
      aba.setColumnWidth(5, 120);
      aba.setColumnWidth(6, 120);
      aba.setColumnWidth(7, 160);
      aba.setColumnWidth(8, 160);
      aba.setColumnWidth(9, 200);
    }
  });

  return ContentService.createTextOutput('Planilha inicializada com sucesso!');
}

// ─── Recebe POST do formulário ───────────────────────────────
function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);

    const jogoIdx = parseInt(dados.jogoIndex);
    const jogo = JOGOS[jogoIdx];
    if (!jogo) throw new Error('Jogo inválido: ' + jogoIdx);

    let aba = ss.getSheetByName(jogo.aba);
    if (!aba) {
      aba = ss.insertSheet(jogo.aba);
      const headerRow = aba.getRange(1, 1, 1, CABECALHO.length);
      headerRow.setValues([CABECALHO]);
      headerRow.setBackground('#002776');
      headerRow.setFontColor('#ffffff');
      headerRow.setFontWeight('bold');
      aba.setFrozenRows(1);
    }

    const agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');

    const linha = [
      dados.codigo   || '',
      agora,
      dados.nome     || '',
      dados.whatsapp || '',
      dados.email    || '',
      dados.gol1 !== undefined ? dados.gol1 : '',
      dados.gol2 !== undefined ? dados.gol2 : '',
      'Pendente',
      'Não',
      '',
    ];

    aba.appendRow(linha);

    // Formata a linha recém adicionada
    const ultimaLinha = aba.getLastRow();
    const bgColor = ultimaLinha % 2 === 0 ? '#f0f4ff' : '#ffffff';
    aba.getRange(ultimaLinha, 1, 1, CABECALHO.length).setBackground(bgColor);

    return ContentService
      .createTextOutput(JSON.stringify({ sucesso: true, mensagem: 'Palpite registrado!' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ sucesso: false, erro: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── Responde GET (teste de conectividade) ───────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'online', bolao: 'Copa 2026' }))
    .setMimeType(ContentService.MimeType.JSON);
}

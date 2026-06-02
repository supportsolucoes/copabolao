// ============================================================
//  BOLÃO COPA 2026 — Google Apps Script
//  Cole este código em: script.google.com > Novo projeto
//  Depois: Implantar > Novo implantação > App da Web
//  Acesso: Qualquer pessoa | Execute como: Você
// ============================================================

const SHEET_ID        = 'COLE_AQUI_O_ID_DA_SUA_PLANILHA';
const WHATSAPP_TOKEN  = 'COLE_AQUI_SEU_TOKEN_CALLMEBOT'; // ver instruções abaixo
const PRECO_UNITARIO  = 5.00;

const JOGOS = [
  { aba: 'Jogo1 - Brasil x Marrocos', nome: 'Brasil × Marrocos' },
  { aba: 'Jogo2 - Brasil x Haiti',    nome: 'Brasil × Haiti'    },
  { aba: 'Jogo3 - Escócia x Brasil',  nome: 'Escócia × Brasil'  },
];

const CABECALHO = [
  'Código', 'Data/Hora', 'Nome', 'WhatsApp', 'E-mail',
  'Palpite T1', 'Palpite T2', 'Total Palpites', 'Total R$',
  'Comprovante?', 'Pagamento Confirmado?', 'Observação',
];

// ── Inicializa abas da planilha ──────────────────────────────
function inicializarPlanilha() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  JOGOS.forEach(jogo => {
    let aba = ss.getSheetByName(jogo.aba);
    if (!aba) aba = ss.insertSheet(jogo.aba);
    if (aba.getLastRow() === 0) {
      const h = aba.getRange(1, 1, 1, CABECALHO.length);
      h.setValues([CABECALHO]);
      h.setBackground('#002776').setFontColor('#ffffff').setFontWeight('bold');
      aba.setFrozenRows(1);
      aba.setColumnWidths(1, CABECALHO.length, 140);
    }
  });
  return ContentService.createTextOutput('✅ Planilha inicializada!');
}

// ── Recebe POST (novo palpite) ───────────────────────────────
function doPost(e) {
  try {
    const dados   = JSON.parse(e.postData.contents);
    const ss      = SpreadsheetApp.openById(SHEET_ID);
    const jogoIdx = parseInt(dados.jogoIndex);
    const jogo    = JOGOS[jogoIdx];
    if (!jogo) throw new Error('Jogo inválido');

    let aba = ss.getSheetByName(jogo.aba);
    if (!aba) {
      aba = ss.insertSheet(jogo.aba);
      aba.getRange(1, 1, 1, CABECALHO.length).setValues([CABECALHO]);
    }

    const agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
    const linha = [
      dados.codigo        || '',
      agora,
      dados.nome          || '',
      dados.whatsapp      || '',
      dados.email         || '',
      dados.gol1 !== undefined ? dados.gol1 : '',
      dados.gol2 !== undefined ? dados.gol2 : '',
      dados.totalPalpites || 1,
      dados.totalValor    || PRECO_UNITARIO,
      'Pendente',
      'Não',
      '',
    ];

    aba.appendRow(linha);
    const ul = aba.getLastRow();
    aba.getRange(ul, 1, 1, CABECALHO.length)
       .setBackground(ul % 2 === 0 ? '#f0f4ff' : '#ffffff');

    return ok({ mensagem: 'Palpite registrado com sucesso!' });
  } catch (err) {
    return erro(err.message);
  }
}

// ── Recebe GET (consulta de status ou confirmação manual) ────
function doGet(e) {
  const action = (e.parameter.action || '').toLowerCase();

  if (action === 'status') {
    return consultarStatus(e.parameter.codigo || '');
  }

  // Teste de conectividade
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'online', bolao: 'Copa 2026' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Consulta status por código ────────────────────────────────
function consultarStatus(codigo) {
  codigo = codigo.trim().toUpperCase();
  if (!codigo) return erro('Código não informado');

  const ss = SpreadsheetApp.openById(SHEET_ID);

  for (const jogo of JOGOS) {
    const aba = ss.getSheetByName(jogo.aba);
    if (!aba || aba.getLastRow() < 2) continue;

    const dados = aba.getRange(2, 1, aba.getLastRow() - 1, CABECALHO.length).getValues();
    for (const linha of dados) {
      if ((linha[0] || '').toString().trim().toUpperCase() === codigo) {
        const confirmado = (linha[10] || '').toString().toLowerCase() === 'sim';
        return ok({
          encontrado:  true,
          confirmado,
          codigo:      linha[0],
          jogo:        jogo.nome,
          nome:        linha[2],
          gol1:        linha[5],
          gol2:        linha[6],
          data:        linha[1],
        });
      }
    }
  }

  return ok({ encontrado: false });
}

// ── Trigger: dispara quando você edita a planilha ────────────
// Configure em: Gatilhos (relógio) > onEdit > Ao editar
function onEdit(e) {
  try {
    const sheet  = e.range.getSheet();
    const col    = e.range.getColumn();
    const row    = e.range.getRow();
    const valor  = (e.value || '').toString().toLowerCase();

    // Coluna 11 = "Pagamento Confirmado?" | linha > 1 (não é cabeçalho) | valor = "sim"
    if (col !== 11 || row < 2 || valor !== 'sim') return;

    const dados = sheet.getRange(row, 1, 1, CABECALHO.length).getValues()[0];
    const codigo    = dados[0];
    const nome      = dados[2];
    const whatsapp  = dados[3]; // ex: (11) 99999-9999
    const gol1      = dados[5];
    const gol2      = dados[6];

    // Descobre qual jogo pela aba
    const nomJogo = JOGOS.find(j => j.aba === sheet.getName())?.nome || 'Jogo';

    // Monta mensagem
    const msg = encodeURIComponent(
      `🇧🇷 *BOLÃO COPA 2026 — Confirmação*\n\n` +
      `Olá, *${nome}*! 🎉\n\n` +
      `✅ Seu palpite foi *confirmado* com sucesso!\n\n` +
      `🎫 Código: *${codigo}*\n` +
      `⚽ Jogo: ${nomJogo}\n` +
      `🎯 Palpite: *${gol1} × ${gol2}*\n\n` +
      `Boa sorte! Que o Brasil ganhe! 🇧🇷🏆`
    );

    // Formata número (remove tudo que não é dígito, adiciona 55 se não tiver)
    let num = (whatsapp || '').replace(/\D/g, '');
    if (num.length === 11) num = '55' + num; // DDD + número sem DDI
    if (num.length === 10) num = '55' + num; // DDD + número curto

    // Envia via CallMeBot (WhatsApp gratuito para notificações)
    if (num && WHATSAPP_TOKEN !== 'COLE_AQUI_SEU_TOKEN_CALLMEBOT') {
      const url = `https://api.callmebot.com/whatsapp.php?phone=${num}&text=${msg}&apikey=${WHATSAPP_TOKEN}`;
      UrlFetchApp.fetch(url);
    }

  } catch (err) {
    console.error('Erro no onEdit:', err.message);
  }
}

// ── Helpers ───────────────────────────────────────────────────
function ok(obj) {
  return ContentService
    .createTextOutput(JSON.stringify({ sucesso: true, ...obj }))
    .setMimeType(ContentService.MimeType.JSON);
}
function erro(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ sucesso: false, erro: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  COMO CONFIGURAR O WHATSAPP AUTOMÁTICO (CallMeBot - grátis)
// ============================================================
//  1. Adicione o número +34 644 59 37 23 nos seus contatos
//  2. Envie a mensagem: "I allow callmebot to send me messages"
//  3. Você receberá um token (ex: 1234567)
//  4. Cole o token na constante WHATSAPP_TOKEN acima
//
//  Para disparar quando você confirma:
//  1. No Apps Script, clique no ícone de relógio (Gatilhos)
//  2. Adicionar gatilho > função: onEdit > evento: Da planilha > Ao editar
// ============================================================

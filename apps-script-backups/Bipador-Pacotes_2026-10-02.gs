// Backup do codigo-fonte do Apps Script "Bipador Pacotes" (doGet/doPost que o
// index.html usa via URL_SCRIPT), salvo em 02/10/2026, ANTES de adicionar a
// rota ?action=pendentes (pacotes sem prateleira).
//
// Projeto Apps Script: "Bipador Pacotes"
// fileId: 1yoGTz0BRMrSDGITLGVlJoAR0TZ6zXft1nHdDagbDnP_mBpqqLvIFDiBK
// Planilha vinculada (SHEET_ID): 1XZgj3q49fk8fq6MS2dsBKo-9i81dP03uMW_m40Nd6D8
//   ("beta mercado livre coleta backup", aba "Historico")
//
// Se algo der errado depois de colar a versao nova no editor, e so colar
// este arquivo de volta no lugar do codigo atual e dar "Nova versao" no
// deploy de novo, pra voltar exatamente a este estado.

const NOME_ABA = 'Histórico';
const SHEET_ID = '1XZgj3q49fk8fq6MS2dsBKo-9i81dP03uMW_m40Nd6D8';
const COL_PACOTE = 2;
const COL_PRATELEIRA = 3;
const COL_DATA = 4;
const COL_OPERADOR = 7;
const COL_MOVIMENTACOES = 8;
const NOME_ABA_LOG = 'log';

const PRATELEIRAS_VALIDAS = [
  '(L2) Prat 1','(L2) Prat 2','(L2) Prat 3','(L2) Prat 4','(L2) Prat 5',
  '(L2) Prat 6','(L2) Prat 7','(L2) Prat 8','(L2) Prat 9','(L2) Prat 10',
  '(L2) Prat 11','(L2) Prat 12','(L2) Prat 13','(L2) Prat 14','(L2) Prat 15',
  '(L2) Prat 16','(L2) Prat 17','(L2) Prat 18','(L2) Prat 19','(L2) Prat 20',
  '(L2) Prat 21','(L2) Prat 22','(L2) Prat 23','(L2) Prat 24','(L2) Prat 25',
  '(L2) Prat 26','(L2) Prat 27','(L2) Prat 28','(L2) Prat 29','(L2) Prat 30',
  '(L2) Prat 31','(L2) Prat 32','(L2) Prat 33','(L2) Prat 34','(L2) Prat 35',
  '(L2) Prat 36','(L2) Prat 37','(L2) Prat 38','(L2) Prat 39','(L2) Prat 40',
  '(L2) Prat 41','(L2) Prat 42',
  '(L1) Prat 60','(L1) Prat 61','(L1) Prat 62','(L1) Prat 63','(L1) Prat 64',
  '(L2) Gond 1','(L2) Gond 3','(L2) Gond 2','(L2) Gond 4','(L2) Gond 5',
  '(L2) Gond 6','(L2) Gond 7',
  '(L2) Chao','(L1) Chao',
  '(L2) Parede','(L2) Pre Triagem'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    let p = (e && e.parameter) || {};
    if ((!p.pacote || !p.prateleira) && e && e.postData && e.postData.contents) {
      try { const j = JSON.parse(e.postData.contents); p = Object.assign({}, p, j); } catch(_) {}
    }
    const pacote = String(p.pacote || '').trim();
    const prateleira = String(p.prateleira || '').trim();
    const operador = String(p.operador || '').trim();

    if (!pacote) return _resp({ ok: false, erro: 'Pacote vazio' });
    if (!prateleira) return _resp({ ok: false, erro: 'Prateleira vazia' });
    if (PRATELEIRAS_VALIDAS.indexOf(prateleira) === -1) {
      return _resp({ ok: false, erro: 'Prateleira invalida: ' + prateleira });
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const aba = ss.getSheetByName(NOME_ABA);
    if (!aba) return _resp({ ok: false, erro: 'Aba nao encontrada: ' + NOME_ABA });
    const url = ss.getUrl();
    const tz = ss.getSpreadsheetTimeZone();

    const lastRow = aba.getLastRow();
    let linhaExistente = -1;
    if (lastRow >= 2) {
      const valores = aba.getRange(2, COL_PACOTE, lastRow - 1, 1).getValues();
      for (let i = 0; i < valores.length; i++) {
        if (String(valores[i][0]).trim() === pacote) { linhaExistente = i; break; }
      }
    }

    let atualizado = false;
    let moved = false;
    let linhaPlanilha;
    let prateleiraAnterior = '';
    const agora = new Date();

    if (linhaExistente >= 0) {
      atualizado = true;
      linhaPlanilha = linhaExistente + 2;

      prateleiraAnterior = String(aba.getRange(linhaPlanilha, COL_PRATELEIRA).getValue() || '').trim();
      const movAnterior = String(aba.getRange(linhaPlanilha, COL_MOVIMENTACOES).getValue() || '');

      if (prateleiraAnterior && prateleiraAnterior !== prateleira) {
        moved = true;
        const agoraStr = Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm');
        const novaEntrada = prateleira + ' — ' + agoraStr + ' (' + (operador || 's/op') + ')';
        const movAtualizada = movAnterior ? (movAnterior + '\n' + novaEntrada) : novaEntrada;
        aba.getRange(linhaPlanilha, COL_MOVIMENTACOES).setValue(movAtualizada);
      }

      aba.getRange(linhaPlanilha, COL_PRATELEIRA).setValue(prateleira);
      aba.getRange(linhaPlanilha, COL_DATA).setValue(agora);
      if (operador) aba.getRange(linhaPlanilha, COL_OPERADOR).setValue(operador);
    } else {
      linhaPlanilha = Math.max(lastRow, 1) + 1;
      aba.getRange(linhaPlanilha, 1, 1, 4).setValues([['', pacote, prateleira, agora]]);
      if (operador) aba.getRange(linhaPlanilha, COL_OPERADOR).setValue(operador);
      // Registra cadastro inicial na coluna H (Movimentacoes)
      const agoraStr = Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm');
      const entradaInicial = prateleira + ' — ' + agoraStr + ' (' + (operador || 's/op') + ')';
      aba.getRange(linhaPlanilha, COL_MOVIMENTACOES).setValue(entradaInicial);
    }

    const verificacao = String(aba.getRange(linhaPlanilha, COL_PACOTE).getValue() || '').trim();
    if (verificacao !== pacote) {
      return _resp({ ok: false, erro: 'Readback falhou: esperava "' + pacote + '" mas leu "' + verificacao + '"' });
    }

    _gravarLog(pacote, prateleira);

    return _resp({ ok: true, atualizado: atualizado, moved: moved, prateleiraAnterior: prateleiraAnterior, prateleira: prateleira, pacote: pacote, linha: linhaPlanilha, url: url });
  } catch (err) {
    return _resp({ ok: false, erro: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

function doGet() {
  return _resp({ ok: true, total: PRATELEIRAS_VALIDAS.length });
}

function _resp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function _gravarLog(pacote, prateleira) {
  try {
    const log = SpreadsheetApp.openById(SHEET_ID).getSheetByName(NOME_ABA_LOG);
    if (!log) return;
    log.appendRow(['', '', pacote, prateleira, new Date()]);
  } catch(e) {}
}

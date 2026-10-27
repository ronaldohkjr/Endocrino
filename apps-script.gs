/**
 * Backend da página "Busca Endocrinologista".
 * Cole este código no Apps Script de uma Planilha Google (Extensões > Apps Script).
 *
 * 1) Troque o PIN abaixo por um código que só vocês dois saibam.
 * 2) Implantar > Nova implantação > Tipo: App da Web
 *      Executar como: Eu
 *      Quem pode acessar: Qualquer pessoa
 * 3) Copie a URL que termina em /exec e mande para o Claude.
 */
const PIN = '0000'; // <<< TROQUE AQUI

const ABA = 'Contatos';
const COLUNAS = ['id', 'medica', 'data', 'canal', 'resultado', 'nota', 'autor', 'criadoEm'];

function aba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.appendRow(COLUNAS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function listar_() {
  const sh = aba_();
  const vals = sh.getDataRange().getDisplayValues();
  const head = vals.shift();
  return vals.filter(r => r[0]).map(r => {
    const o = {};
    head.forEach((h, i) => (o[h] = r[i]));
    return o;
  });
}

function doGet(e) {
  if (String(e.parameter.pin || '') !== PIN) return json_({ ok: false, erro: 'pin' });
  return json_({ ok: true, contatos: listar_() });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, erro: 'formato' }); }
  if (String(body.pin || '') !== PIN) return json_({ ok: false, erro: 'pin' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = aba_();
    if (body.action === 'add') {
      const c = body.contato || {};
      const linha = [
        Utilities.getUuid(),
        String(c.medica || '').slice(0, 40),
        String(c.data || '').slice(0, 10),
        String(c.canal || '').slice(0, 30),
        String(c.resultado || '').slice(0, 30),
        String(c.nota || '').slice(0, 600),
        String(c.autor || '').slice(0, 40),
        new Date().toISOString()
      ];
      // Força texto na coluna de data para o Sheets não converter
      linha[2] = "'" + linha[2];
      sh.appendRow(linha);
    } else if (body.action === 'delete') {
      const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => r[0]);
      const i = ids.indexOf(String(body.id));
      if (i >= 0) sh.deleteRow(i + 2);
    } else {
      return json_({ ok: false, erro: 'acao' });
    }
    return json_({ ok: true, contatos: listar_() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Saber Não Basta: lista de convidados
 * Colar em: Google Sheet > Extensões > Apps Script (substituir tudo)
 * Depois: Implementar > Nova implementação > App da Web
 *   Executar como: Eu · Quem tem acesso: Qualquer pessoa
 * Copiar o URL que termina em /exec e colar em DADOS.planilha (convidado/index.html).
 */
const CABECALHO = ['Data', 'Nome completo', 'E-mail', 'Telemóvel', 'Quem convidou'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = (e && e.parameter) || {};
    if (p.site) return resposta({ ok: true }); // robô (campo isca preenchido)

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheets()[0]; // primeira aba da planilha
    if (sh.getLastRow() === 0) {
      sh.appendRow(CABECALHO);
      sh.getRange(1, 1, 1, CABECALHO.length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), limpa(p.nome), limpa(p.email), limpa(p.telefone), limpa(p.indicou)]);
    return resposta({ ok: true });
  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* Corta o tamanho e impede que um valor seja lido como fórmula (=, +, -, @).
   O apóstrofo também mantém o "+351" do telemóvel como texto. */
function limpa(v) {
  const s = String(v || '').trim().slice(0, 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

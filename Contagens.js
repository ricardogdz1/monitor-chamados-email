/**
 * Contagens: períodos de apuração com início/fim definidos pelo gestor (não seguem o mês).
 *
 * A aba "Contagens" guarda um corte por linha. Cada contagem vai do seu início até o
 * início da próxima; a última linha é a contagem atual (em aberto).
 * Para corrigir um encerramento feito por engano, basta editar/apagar a linha na planilha.
 */

const ABA_CONTAGENS = 'Contagens';

function abaContagens_() {
  const ss = planilha_();
  let aba = ss.getSheetByName(ABA_CONTAGENS);
  if (!aba) {
    aba = ss.insertSheet(ABA_CONTAGENS);
    aba.getRange(1, 1, 1, 2).setValues([['inicio', 'registradoEm']]).setFontWeight('bold');
    aba.setFrozenRows(1);
  }
  return aba;
}

/** Inícios das contagens (ms), em ordem crescente. */
function lerContagens_() {
  const aba = abaContagens_();
  const n = aba.getLastRow();
  if (n < 2) return [];
  return aba.getRange(2, 1, n - 1, 1).getValues()
    .map(r => ms_(r[0]))
    .filter(v => v)
    .sort((a, b) => a - b);
}

/**
 * Encerra a contagem atual agora e inicia a próxima no mesmo instante.
 * Na primeira vez não existe contagem registrada: inicioMs informa quando a atual começou.
 */
function encerrarContagem(inicioMs) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const aba = abaContagens_();
    const agora = new Date();
    const linhas = [];
    if (!lerContagens_().length) {
      if (!inicioMs) throw new Error('Informe quando a contagem atual começou.');
      if (inicioMs >= agora.getTime()) throw new Error('O início precisa ser anterior a agora.');
      linhas.push([new Date(inicioMs), agora]);
    }
    linhas.push([agora, agora]);
    aba.getRange(aba.getLastRow() + 1, 1, linhas.length, 2).setValues(linhas).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    return lerContagens_();
  } finally {
    lock.releaseLock();
  }
}

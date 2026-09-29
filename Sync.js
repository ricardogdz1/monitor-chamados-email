/**
 * Sincronização Gmail -> Planilha.
 *
 * Funções para rodar pelo editor:
 *   instalar()        -> cria a planilha, o gatilho automático e faz a 1ª importação
 *   sincronizar()     -> importa emails novos (roda sozinho a cada 10 min)
 *   diagnosticar()    -> grava exemplos de emails na aba "Diagnóstico" para ajuste dos rótulos
 *   reprocessarTudo() -> apaga os dados e relê todos os emails (use após ajustar o Config)
 */

function instalar() {
  planilha_();
  garantirGatilho_();
  sincronizar();
  Logger.log('Instalado. Planilha: ' + planilha_().getUrl());
}

/** Gatilho de sincronização do usuário atual (cada colaborador tem o seu, lendo o próprio Gmail). */
function garantirGatilho_() {
  const existe = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'sincronizar');
  if (!existe) ScriptApp.newTrigger('sincronizar').timeBased().everyMinutes(CONFIG.INTERVALO_SYNC_MIN).create();
}

const PAGINA_THREADS = 50;
const PAUSA_COTA_MS = 60 * 60 * 1000;
const ehErroDeCota_ = e => /muitas vezes|too many times|limit exceeded|cota|quota/i.test(String(e && e.message || e));

function sincronizar() {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(1000)) return 'Sincronização já em andamento.';
  try {
    planilha_(); // garante a planilha (e a migração do estado antigo) antes de ler as propriedades
    const inicio = Date.now();
    const props = props_();

    // Cota diária do Gmail esgotada: não insiste (cada tentativa gasta mais cota)
    const pausaAte = Number(props.getProperty('PAUSA_COTA_ATE') || 0);
    if (pausaAte > inicio) {
      return 'Limite diário de leituras do Gmail atingido. Nova tentativa após ' +
        Utilities.formatDate(new Date(pausaAte), CONFIG.FUSO, 'dd/MM HH:mm') + '.';
    }

    const ultimaSync = Number(props.getProperty('ULTIMA_SYNC') || 0);
    let offset = Number(props.getProperty('OFFSET_BUSCA') || 0);

    // Só os emails desde a última sincronização (com 1 h de margem), em segundos desde 1970
    let query = CONFIG.GMAIL_QUERY;
    if (ultimaSync) query += ' after:' + Math.floor((ultimaSync - 3600 * 1000) / 1000);

    const chamados = lerChamados_();
    const processados = lerProcessados_();
    const novos = [];
    let terminou = false;
    let erroCota = null;
    const tempoEsgotado = () => Date.now() - inicio > CONFIG.TEMPO_MAX_EXECUCAO_MS;

    try {
      pagina:
      while (!tempoEsgotado()) {
        const threads = GmailApp.search(query, offset, PAGINA_THREADS);
        if (!threads.length) { terminou = true; break; }
        const mensagens = GmailApp.getMessagesForThreads(threads);
        for (const lista of mensagens) {
          for (const msg of lista) {
            // Para no meio da página: a próxima execução relê a página, pulando o que já foi processado
            if (tempoEsgotado()) break pagina;
            const id = msg.getId();
            if (processados[id]) continue;
            const d = lerMensagem_(msg);
            processados[id] = true;
            novos.push([id, d ? d.numero : '', d ? d.tipo : 'ignorado', new Date()]);
            if (d) mesclar_(chamados, d);
          }
        }
        offset += threads.length;
      }
    } catch (e) {
      if (!ehErroDeCota_(e)) throw e;
      erroCota = e;
      props.setProperty('PAUSA_COTA_ATE', String(Date.now() + PAUSA_COTA_MS));
    } finally {
      if (novos.length) {
        salvarChamados_(chamados);
        registrarProcessados_(novos);
      }
    }

    if (terminou) {
      props.setProperty('ULTIMA_SYNC', String(inicio));
      props.deleteProperty('OFFSET_BUSCA');
      props.deleteProperty('PAUSA_COTA_ATE');
    } else {
      // Importação grande (ou limite do Gmail): continua de onde parou na próxima execução
      props.setProperty('OFFSET_BUSCA', String(offset));
    }
    const resumo = novos.length + ' email(s) novo(s) processado(s)' +
      (erroCota ? ' — limite diário do Gmail atingido, continua mais tarde.' : terminou ? '.' : ' — continua na próxima execução.');
    Logger.log(resumo);
    return resumo;
  } finally {
    lock.releaseLock();
  }
}

function reprocessarTudo() {
  const ss = planilha_();
  [ABA_CHAMADOS, ABA_PROCESSADOS].forEach(nome => {
    const aba = ss.getSheetByName(nome);
    if (aba.getLastRow() > 1) aba.deleteRows(2, aba.getLastRow() - 1);
  });
  const props = props_();
  props.deleteProperty('ULTIMA_SYNC');
  props.deleteProperty('OFFSET_BUSCA');
  props.deleteProperty('PAUSA_COTA_ATE');
  return sincronizar();
}

/**
 * Agrupa os emails pelo padrão do assunto (número trocado por ####) e grava
 * 1 exemplo de cada padrão, com o que foi extraído, na aba "Diagnóstico".
 */
function diagnosticar() {
  const ss = planilha_();
  const aba = ss.getSheetByName(ABA_DIAGNOSTICO) || ss.insertSheet(ABA_DIAGNOSTICO);
  aba.clear();
  const grupos = {};
  const threads = GmailApp.search(CONFIG.GMAIL_QUERY, 0, 300);
  GmailApp.getMessagesForThreads(threads).forEach(lista => lista.forEach(msg => {
    const padrao = msg.getSubject().replace(/\d{4}-\d{6}/g, '####').replace(/\d+/g, '9');
    const chave = msg.getFrom().replace(/.*</, '').replace(/>.*/, '') + ' | ' + padrao;
    if (grupos[chave]) { grupos[chave].qtd++; return; }
    grupos[chave] = { qtd: 1, msg: msg };
  }));
  const linhas = Object.keys(grupos).sort((a, b) => grupos[b].qtd - grupos[a].qtd).map(chave => {
    const g = grupos[chave], msg = g.msg, d = lerMensagem_(msg);
    const extraido = d ? JSON.stringify(Object.assign({}, d, { resumo: undefined, resumoResolucao: undefined }), null, 1) : '';
    return [g.qtd, d ? d.tipo : 'sem número no assunto', msg.getFrom(), msg.getSubject(), extraido,
            msg.getPlainBody().slice(0, 20000)];
  });
  aba.getRange(1, 1, 1, 6).setValues([['qtd', 'tipo detectado', 'remetente', 'assunto (exemplo)', 'dados extraídos', 'corpo (texto)']]).setFontWeight('bold');
  if (linhas.length) aba.getRange(2, 1, linhas.length, 6).setValues(linhas);
  aba.setColumnWidth(1, 50).setColumnWidths(2, 4, 220).setColumnWidth(6, 600);
  Logger.log('Veja a aba Diagnóstico: ' + ss.getUrl());
  linhas.forEach(l => Logger.log(l[0] + 'x | ' + l[1] + ' | ' + l[2] + ' | ' + l[3]));
}

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
    // Início da primeira execução desta busca. Uma conversa que recebe email novo durante uma busca
    // longa sobe para o topo e é pulada pelo offset; a próxima busca, a partir deste início, a relê.
    const inicioBusca = offset ? Number(props.getProperty('INICIO_BUSCA') || inicio) : inicio;

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
      props.setProperty('ULTIMA_SYNC', String(inicioBusca));
      props.deleteProperty('OFFSET_BUSCA');
      props.deleteProperty('INICIO_BUSCA');
      props.deleteProperty('PAUSA_COTA_ATE');
    } else {
      // Importação grande (ou limite do Gmail): continua de onde parou na próxima execução
      props.setProperty('OFFSET_BUSCA', String(offset));
      props.setProperty('INICIO_BUSCA', String(inicioBusca));
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
  props.deleteProperty('INICIO_BUSCA');
  props.deleteProperty('PAUSA_COTA_ATE');
  return sincronizar();
}

/**
 * Relê os emails dos últimos dias sem apagar nada (os já processados são pulados).
 * Use quando uma sincronização interrompida deixou emails para trás.
 */
function ressincronizarUltimosDias(dias) {
  dias = dias || 2;
  const lock = LockService.getUserLock();
  lock.waitLock(30000); // não deixa uma sincronização em andamento sobrescrever o novo ponto de partida
  try {
    planilha_();
    const props = props_();
    props.setProperty('ULTIMA_SYNC', String(Date.now() - dias * 24 * 3600 * 1000));
    props.deleteProperty('OFFSET_BUSCA');
    props.deleteProperty('INICIO_BUSCA');
    props.deleteProperty('PAUSA_COTA_ATE');
  } finally {
    lock.releaseLock();
  }
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

const ABA_DIAGNOSTICO_DIA = 'Diagnóstico por dia';

/**
 * Compara, dia a dia, o que está no Gmail com o que está na planilha e no painel.
 * Só lê os dados (não altera chamados nem processados); o resultado vai para a aba
 * "Diagnóstico por dia". Sem parâmetros, analisa do dia 1 do mês atual até hoje.
 *
 * Colunas por dia (dia = primeiro email do chamado no Gmail):
 *   gmail          -> chamados com email nesse dia
 *   naPlanilha     -> desses, quantos existem na aba Chamados
 *   noPainel       -> desses, quantos passam pelo filtro de operador (aparecem no painel)
 *   naoProcessados -> chamados com algum email que não está na aba Processados
 * E, pela data de abertura gravada na planilha:
 *   planilhaAbertura / painelAbertura -> chamados abertos nesse dia (na planilha / no painel)
 */
function diagnosticarPorDia(inicio, fim) {
  const fuso = CONFIG.FUSO;
  const hoje = new Date();
  inicio = inicio ? new Date(inicio) : new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  fim = fim ? new Date(fim) : new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() + 1);
  const dia = d => Utilities.formatDate(d, fuso, 'yyyy-MM-dd');

  const mapa = lerChamados_();
  const processados = lerProcessados_();
  const eu = operadorDoUsuario_(mapa).nome;

  // Gmail: primeiro email de cada chamado no período e se algum email ficou sem processar.
  // A busca por data do Gmail não usa o nosso fuso: busca 1 dia a mais de cada lado e filtra abaixo.
  const umDia = 24 * 3600 * 1000;
  const query = CONFIG.GMAIL_QUERY +
    ' after:' + Utilities.formatDate(new Date(inicio.getTime() - umDia), fuso, 'yyyy/MM/dd') +
    ' before:' + Utilities.formatDate(new Date(fim.getTime() + umDia), fuso, 'yyyy/MM/dd');
  const doGmail = {};
  const t0 = Date.now();
  let offset = 0, incompleto = false;
  while (true) {
    if (Date.now() - t0 > CONFIG.TEMPO_MAX_EXECUCAO_MS) { incompleto = true; break; }
    const threads = GmailApp.search(query, offset, 100);
    if (!threads.length) break;
    GmailApp.getMessagesForThreads(threads).forEach(lista => lista.forEach(msg => {
      const numero = extrairNumero_(msg.getSubject());
      const data = msg.getDate();
      if (!numero || data < inicio || data >= fim) return;
      const g = doGmail[numero] || (doGmail[numero] = { primeiro: data, naoProcessados: 0 });
      if (data < g.primeiro) g.primeiro = data;
      if (!processados[msg.getId()]) g.naoProcessados++;
    }));
    offset += threads.length;
  }

  const dias = {};
  const linhaDia = d => dias[d] || (dias[d] = { gmail: 0, naPlanilha: 0, noPainel: 0, naoProcessados: 0, planilhaAbertura: 0, painelAbertura: 0 });
  for (let d = new Date(inicio); d < fim; d.setDate(d.getDate() + 1)) linhaDia(dia(d));

  const fora = []; // chamados do Gmail que não aparecem no painel
  Object.keys(doGmail).forEach(numero => {
    const g = doGmail[numero], c = mapa[numero], r = linhaDia(dia(g.primeiro));
    r.gmail++;
    if (g.naoProcessados) r.naoProcessados++;
    if (c) r.naPlanilha++;
    if (c && ehMeu_(c, eu)) { r.noPainel++; return; }
    fora.push([numero, dia(g.primeiro), c ? 'na planilha, outro operador' : 'fora da planilha',
               c ? txt_(c.operador) : '', c && c.dataAbertura instanceof Date ? dia(c.dataAbertura) : '', g.naoProcessados]);
  });
  Object.keys(mapa).forEach(numero => {
    const c = mapa[numero];
    if (!(c.dataAbertura instanceof Date) || c.dataAbertura < inicio || c.dataAbertura >= fim) return;
    const r = linhaDia(dia(c.dataAbertura));
    r.planilhaAbertura++;
    if (ehMeu_(c, eu)) r.painelAbertura++;
  });

  const ss = planilha_();
  const aba = ss.getSheetByName(ABA_DIAGNOSTICO_DIA) || ss.insertSheet(ABA_DIAGNOSTICO_DIA);
  aba.clear();
  const cab = ['dia', 'gmail', 'naPlanilha', 'noPainel', 'naoProcessados', 'planilhaAbertura', 'painelAbertura'];
  const linhas = Object.keys(dias).sort().map(d => [d].concat(cab.slice(1).map(k => dias[d][k])));
  aba.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
  aba.getRange(2, 1, linhas.length, cab.length).setValues(linhas);

  const col = cab.length + 2;
  aba.getRange(1, col).setValue('Operador do painel: ' + (eu || '(não identificado)') +
    (incompleto ? ' — ATENÇÃO: tempo esgotado, Gmail lido só em parte' : '')).setFontWeight('bold');
  const cabFora = ['chamado', 'diaGmail', 'situação', 'operador na planilha', 'abertura na planilha', 'emails não processados'];
  aba.getRange(3, col, 1, cabFora.length).setValues([cabFora]).setFontWeight('bold');
  fora.sort((a, b) => a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0);
  if (fora.length) aba.getRange(4, col, fora.length, cabFora.length).setValues(fora);

  Logger.log('Operador do painel: ' + eu + (incompleto ? ' (Gmail lido só em parte)' : ''));
  Logger.log(cab.join(' | '));
  linhas.forEach(l => Logger.log(l.join(' | ')));
  Logger.log(fora.length + ' chamado(s) do Gmail fora do painel. Detalhes na aba "' + ABA_DIAGNOSTICO_DIA + '": ' + ss.getUrl());
}

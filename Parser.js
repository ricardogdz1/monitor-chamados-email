/**
 * Leitura (parse) dos emails do DeskManager (Suporte - Agro1 <helpdesk@agro1.inf.br>).
 *
 * Tipos reconhecidos:
 *   aberto     -> "ABERTO Nº 0926-005066"
 *   andamento  -> "Transferência Nº ...", "Aguardando Cliente Nº ..." (qualquer "<status> Nº <número>")
 *   resolvido  -> "Resolvido Nº ...", "Finalizado ...", "Finalização Automatica ..."
 *   atribuido  -> "Você recebeu o chamado #..."
 *   avaliacao  -> corpo com "Questão: 1 ... atribua uma nota" (assunto "CLIENTE - SOLICITANTE #número")
 */

function extrairNumero_(assunto) {
  const m = String(assunto || '').match(CONFIG.REGEX_NUMERO);
  return m ? m[1] : null;
}

function identificarTipo_(assunto, texto) {
  if (/atribua uma nota|Quest[aã]o:\s*1/i.test(texto)) return { tipo: 'avaliacao' };
  if (/voc[eê] recebeu o chamado/i.test(assunto)) return { tipo: 'atribuido' };
  const m = assunto.match(/^\s*(.+?)\s*(?:N[º°o]\.?|#)\s*:?\s*#?\s*\d{4}-\d{6}/i);
  const rotulo = m ? m[1].replace(/^(Agro1\s*-\s*)/i, '').replace(/[|:\-]+$/, '').trim() : '';
  if (/^(re|res|enc|fw|fwd)\s*:/i.test(assunto)) return { tipo: 'ignorar' };
  if (/resolvid|finaliz|fechad|conclu|encerrad/i.test(rotulo || assunto)) return { tipo: 'resolvido', rotulo: 'Resolvido' };
  if (/^abert/i.test(rotulo)) return { tipo: 'aberto', rotulo: 'Aberto' };
  if (rotulo && rotulo.length <= 40) return { tipo: 'andamento', rotulo: capitalizar_(rotulo) };
  return { tipo: 'ignorar' };
}

function capitalizar_(s) {
  return s.toLowerCase().replace(/(^|\s)\S/g, c => c.toUpperCase());
}

/** Converte o HTML do email em texto com 1 linha por linha de tabela (td separados por espaço). */
function htmlParaTexto_(html) {
  return String(html || '')
    .replace(/<(script|style|head)[\s\S]*?<\/\1>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(tr|p|div|h\d|li|table)>/gi, '\n')
    .replace(/<\/t[dh]>/gi, '\t')
    .replace(/<img[^>]*\balt="([^"]*)"[^>]*>/gi, ' $1 ')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"').replace(/&#39;/g, "'")
    .replace(/&#(\d+);/g, (m, n) => String.fromCharCode(Number(n)))
    .split('\n').map(l => l.replace(/[ \t\u00a0]+/g, ' ').trim()).join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function escaparRegex_(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Procura "Rótulo: valor" (ou rótulo numa linha e valor na próxima). */
function lerCampo_(texto, rotulos) {
  const grupo = rotulos.map(escaparRegex_).join('|');
  let m = texto.match(new RegExp('^\\s*(?:' + grupo + ')\\s*:\\s*(\\S.*)$', 'im'));
  if (m) return limparValor_(m[1]);
  m = texto.match(new RegExp('^\\s*(?:' + grupo + ')\\s*:\\s*\\n\\s*(\\S.*)$', 'im'));
  return m ? limparValor_(m[1]) : '';
}

function limparValor_(v) {
  return String(v).replace(/\s+/g, ' ').trim().slice(0, 250);
}

/** "23/09/2026 - 08:28:19" ou "25/09/26 - 10:09:00" -> Date */
function lerDataHora_(s) {
  const m = String(s || '').match(/(\d{2})\/(\d{2})\/(\d{2,4})(?:\s*-?\s*(\d{2}):(\d{2})(?::(\d{2}))?)?/);
  if (!m) return null;
  let ano = Number(m[3]);
  if (ano < 100) ano += 2000;
  return new Date(ano, Number(m[2]) - 1, Number(m[1]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0));
}

/** Lê as questões da avaliação. A opção escolhida vem marcada com "*****". */
function lerAvaliacao_(texto) {
  const r = { notaAtendimento: '', notaSistema: '', resolucaoCliente: '', comentario: '' };
  const blocos = texto.split(/Quest[aã]o:\s*\d+/i).slice(1);
  blocos.forEach(b => {
    const pergunta = (b.match(/^\s*(.+)/) || ['', ''])[1];
    const marcada = b.match(/^\s*\*{3,}\s*(.+)$/m);
    const valor = marcada ? marcada[1].trim() : '';
    if (!valor) return;
    if (CONFIG.NOTA_ATENDIMENTO.test(pergunta)) r.notaAtendimento = Number(valor) || valor;
    else if (CONFIG.NOTA_SISTEMA.test(pergunta)) r.notaSistema = Number(valor) || valor;
    else if (/considera|resolvid/i.test(pergunta)) r.resolucaoCliente = valor;
  });
  const motivo = texto.match(/Motivo da escolha:[^\n]*\n([\s\S]*?)(?:\n\s*Para visualizar|\n\s*\* Este e-mail|$)/i);
  if (motivo) r.comentario = motivo[1].replace(/\s+/g, ' ').trim().slice(0, 2000);
  if (!r.comentario) r.comentario = lerCampo_(texto, CONFIG.ROTULOS.comentario);
  return r;
}

/** Conversa (histórico) do chamado, para o resumo. */
function lerHistorico_(texto) {
  const i = texto.search(/Hist[oó]rico do chamado/i);
  let t = i >= 0 ? texto.slice(i) : texto;
  const fim = t.search(/\n\s*Para visualizar o conte[uú]do/i);
  if (fim > 0) t = t.slice(0, fim);
  return t.trim();
}

/** Assunto da avaliação: "21441 - ARLINDO ... BELLINCANTA - Danubia #0926-005788" */
function lerClienteDoAssunto_(assunto) {
  const antes = assunto.split('#')[0].trim();
  const i = antes.lastIndexOf(' - ');
  if (i < 0) return { cliente: antes, solicitante: '' };
  return { cliente: antes.slice(0, i).trim(), solicitante: antes.slice(i + 3).trim() };
}

/** Transforma um GmailMessage em um objeto com os dados do chamado. */
function lerMensagem_(msg) {
  const assunto = msg.getSubject();
  const numero = extrairNumero_(assunto);
  if (!numero) return null;

  const texto = htmlParaTexto_(msg.getBody()) || msg.getPlainBody();
  const t = identificarTipo_(assunto, texto);
  const R = CONFIG.ROTULOS;

  const operadorBruto = lerCampo_(texto, R.operador);
  const partesOp = operadorBruto.split('|').map(s => s.trim());

  const d = {
    numero: numero,
    tipo: t.tipo,
    rotulo: t.rotulo || '',
    data: msg.getDate(),
    msgId: msg.getId(),
    cliente: lerCampo_(texto, R.cliente),
    solicitante: lerCampo_(texto, R.solicitante),
    operador: partesOp[0] || '',
    grupo: partesOp[1] || '',
    assunto: lerCampo_(texto, R.assunto),
    dataAbertura: lerDataHora_(lerCampo_(texto, R.dataAbertura))
  };

  if (t.tipo === 'resolvido' || t.tipo === 'andamento' || t.tipo === 'aberto') {
    d.resumo = lerHistorico_(texto).slice(0, CONFIG.TAMANHO_MAX_RESUMO);
    d.descricaoAcao = lerCampo_(texto, R.descricaoAcao);
  }
  if (t.tipo === 'avaliacao') {
    Object.assign(d, lerAvaliacao_(texto));
    const doAssunto = lerClienteDoAssunto_(assunto);
    d.cliente = doAssunto.cliente || d.cliente;
    d.solicitante = doAssunto.solicitante;
    d.dataFinalizacao = lerDataHora_(lerCampo_(texto, ['Data de Finalização', 'Data de Finalizacao']));
  }
  return d;
}

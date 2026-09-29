/**
 * Aplicativo web (painel) e funções chamadas pela página.
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Painel de Chamados')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function include(nome) {
  return HtmlService.createHtmlOutputFromFile(nome).getContent();
}

const ms_ = v => (v instanceof Date ? v.getTime() : null);
const nota_ = v => (v === '' || v === null || v === undefined || isNaN(Number(v)) ? null : Number(v));
const txt_ = v => String(v === null || v === undefined ? '' : v);
// "13536 - UNIDADE X" e "UNIDADE X" são o mesmo cliente: remove o código do início
const cliente_ = v => txt_(v).replace(/^\s*\d+\s*-\s*/, '').trim();

// Compara nomes ignorando maiúsculas, acentos e espaços extras
const normNome_ = v => txt_(v).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();

/**
 * Nome do usuário como operador no DeskManager. O escolhido no painel tem prioridade;
 * senão, o operador mais frequente nos chamados "Você recebeu o chamado" do próprio Gmail;
 * senão, o operador cujo nome contém todas as partes do email (ricardo.ferronato -> Ricardo Ferronato).
 */
function operadorDoUsuario_(mapa) {
  const salvo = props_().getProperty('OPERADOR');
  if (salvo) return { nome: salvo, confirmado: true };
  const contagem = {};
  Object.keys(mapa).forEach(n => {
    const c = mapa[n];
    if (c.atribuidoAMim === 'Sim' && c.operador) contagem[c.operador] = (contagem[c.operador] || 0) + 1;
  });
  let nome = Object.keys(contagem).sort((a, b) => contagem[b] - contagem[a])[0] || '';
  if (!nome) {
    const partes = normNome_(Session.getEffectiveUser().getEmail().split('@')[0].replace(/[._-]+/g, ' ')).split(' ').filter(Boolean);
    const candidatos = {};
    Object.keys(mapa).forEach(n => {
      const op = mapa[n].operador;
      if (op && partes.length && partes.every(p => normNome_(op).split(' ').indexOf(p) >= 0)) candidatos[op] = (candidatos[op] || 0) + 1;
    });
    nome = Object.keys(candidatos).sort((a, b) => candidatos[b] - candidatos[a])[0] || '';
  }
  return { nome: nome, confirmado: false };
}

const ehMeu_ = (c, operador) => !!operador && normNome_(c.operador) === normNome_(operador);

/** Lista resumida dos chamados do usuário (sem a conversa, para ficar leve). */
function getDados() {
  garantirGatilho_();
  const mapa = lerChamados_();
  const eu = operadorDoUsuario_(mapa);
  const operadores = {};
  const lista = [];
  Object.keys(mapa).forEach(n => {
    const c = mapa[n];
    if (c.operador) operadores[c.operador] = true;
    if (!ehMeu_(c, eu.nome)) return;
    lista.push({
      n: txt_(c.numero), st: txt_(c.status), cl: cliente_(c.cliente), so: txt_(c.solicitante),
      op: txt_(c.operador), gr: txt_(c.grupo), as: txt_(c.assunto),
      da: ms_(c.dataAbertura), dr: ms_(c.dataResolucao), dv: ms_(c.dataAvaliacao),
      na: nota_(c.notaAtendimento), ns: nota_(c.notaSistema),
      rc: txt_(c.resolucaoCliente), co: txt_(c.comentario)
    });
  });
  const props = props_();
  const pausaCota = Number(props.getProperty('PAUSA_COTA_ATE') || 0);
  return {
    chamados: lista,
    operador: eu.nome,
    operadorConfirmado: eu.confirmado,
    operadores: Object.keys(operadores).sort((a, b) => a.localeCompare(b)),
    contagens: lerContagens_(),
    ultimaSync: Number(props.getProperty('ULTIMA_SYNC') || 0),
    importando: !!props.getProperty('OFFSET_BUSCA'),
    pausaCota: pausaCota > Date.now() ? pausaCota : 0,
    planilhaUrl: planilha_().getUrl()
  };
}

/** Grava o nome de operador escolhido pelo usuário no painel. */
function definirOperador(nome) {
  nome = txt_(nome).trim();
  if (!nome) throw new Error('Escolha o seu nome de operador.');
  props_().setProperty('OPERADOR', nome);
  return getDados();
}

/** Detalhe de um chamado: dados + emails originais (HTML, com as imagens da conversa). */
function getDetalhe(numero) {
  const mapa = lerChamados_();
  const c = mapa[numero];
  if (!c || !ehMeu_(c, operadorDoUsuario_(mapa).nome)) throw new Error('Chamado não encontrado: ' + numero);
  const vistos = {};
  const emails = [];
  [['Resolução', c.msgResolucao], ['Última atualização', c.msgUltima], ['Abertura', c.msgAbertura], ['Avaliação', c.msgAvaliacao]]
    .forEach(([rotulo, id]) => {
      if (!id || vistos[id]) return;
      vistos[id] = true;
      try {
        const msg = GmailApp.getMessageById(id);
        emails.push({
          rotulo: rotulo, assunto: msg.getSubject(), data: msg.getDate().getTime(),
          html: msg.getBody(),
          link: 'https://mail.google.com/mail/u/0/#all/' + msg.getThread().getId()
        });
      } catch (e) {
        emails.push({ rotulo: rotulo, erro: 'Não foi possível abrir o email (' + e.message + ')' });
      }
    });
  return {
    numero: txt_(c.numero), status: txt_(c.status), cliente: cliente_(c.cliente), solicitante: txt_(c.solicitante),
    operador: txt_(c.operador), grupo: txt_(c.grupo), assunto: txt_(c.assunto), descricaoAcao: txt_(c.descricaoAcao),
    dataAbertura: ms_(c.dataAbertura), dataResolucao: ms_(c.dataResolucao), dataAvaliacao: ms_(c.dataAvaliacao),
    notaAtendimento: nota_(c.notaAtendimento), notaSistema: nota_(c.notaSistema),
    resolucaoCliente: txt_(c.resolucaoCliente), comentario: txt_(c.comentario),
    resumo: txt_(c.resumo), emails: emails
  };
}

function sincronizarAgora() {
  return sincronizar();
}

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

/** Lista resumida de todos os chamados (sem a conversa, para ficar leve). */
function getDados() {
  const mapa = lerChamados_();
  const lista = Object.keys(mapa).map(n => {
    const c = mapa[n];
    return {
      n: txt_(c.numero), st: txt_(c.status), cl: cliente_(c.cliente), so: txt_(c.solicitante),
      op: txt_(c.operador), gr: txt_(c.grupo), as: txt_(c.assunto),
      da: ms_(c.dataAbertura), dr: ms_(c.dataResolucao), dv: ms_(c.dataAvaliacao),
      na: nota_(c.notaAtendimento), ns: nota_(c.notaSistema),
      rc: txt_(c.resolucaoCliente), co: txt_(c.comentario), eu: c.atribuidoAMim === 'Sim'
    };
  });
  const props = PropertiesService.getScriptProperties();
  return {
    chamados: lista,
    contagens: lerContagens_(),
    ultimaSync: Number(props.getProperty('ULTIMA_SYNC') || 0),
    importando: !!props.getProperty('OFFSET_BUSCA'),
    planilhaUrl: planilha_().getUrl()
  };
}

/** Detalhe de um chamado: dados + emails originais (HTML, com as imagens da conversa). */
function getDetalhe(numero) {
  const c = lerChamados_()[numero];
  if (!c) throw new Error('Chamado não encontrado: ' + numero);
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

/**
 * Armazenamento em Google Planilhas (criada automaticamente na pasta do Drive).
 */

const ABA_CHAMADOS = 'Chamados';
const ABA_PROCESSADOS = 'Processados';
const ABA_DIAGNOSTICO = 'Diagnóstico';

const COLUNAS = [
  'numero', 'status', 'cliente', 'solicitante', 'operador', 'grupo', 'assunto', 'descricaoAcao',
  'dataAbertura', 'dataResolucao', 'dataAvaliacao', 'dataAtualizacao',
  'notaAtendimento', 'notaSistema', 'resolucaoCliente', 'comentario', 'atribuidoAMim',
  'msgAbertura', 'msgResolucao', 'msgAvaliacao', 'msgUltima', 'resumo'
];

function planilha_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PLANILHA_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* recria abaixo */ }
  }
  const ss = SpreadsheetApp.create(CONFIG.NOME_PLANILHA);
  DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(CONFIG.PASTA_DRIVE_ID));

  const aba = ss.getSheets()[0].setName(ABA_CHAMADOS);
  aba.setFrozenRows(1);
  const proc = ss.insertSheet(ABA_PROCESSADOS);
  proc.getRange(1, 1, 1, 4).setValues([['msgId', 'numero', 'tipo', 'processadoEm']]).setFontWeight('bold');

  props.setProperty('PLANILHA_ID', ss.getId());
  return ss;
}

function lerChamados_() {
  const aba = planilha_().getSheetByName(ABA_CHAMADOS);
  const valores = aba.getDataRange().getValues();
  if (valores.length < 2) return {};
  const cab = valores[0];
  const mapa = {};
  for (let i = 1; i < valores.length; i++) {
    const obj = {};
    cab.forEach((c, j) => { if (c) obj[c] = valores[i][j]; });
    if (obj.numero) mapa[obj.numero] = obj;
  }
  return mapa;
}

function salvarChamados_(mapa) {
  const aba = planilha_().getSheetByName(ABA_CHAMADOS);
  if (aba.getLastRow() > 0) aba.getRange(1, 1, aba.getLastRow(), aba.getLastColumn()).clearContent();
  aba.getRange(1, 1, 1, COLUNAS.length).setValues([COLUNAS]).setFontWeight('bold');
  const linhas = Object.keys(mapa).sort().map(n => COLUNAS.map(c => {
    const v = mapa[n][c];
    return v === undefined || v === null ? '' : v;
  }));
  if (linhas.length) {
    aba.getRange(2, 1, linhas.length, 1).setNumberFormat('@'); // número do chamado como texto
    aba.getRange(2, 1, linhas.length, COLUNAS.length).setValues(linhas);
  }
}

function lerProcessados_() {
  const aba = planilha_().getSheetByName(ABA_PROCESSADOS);
  const n = aba.getLastRow();
  const set = {};
  if (n > 1) aba.getRange(2, 1, n - 1, 1).getValues().forEach(r => set[r[0]] = true);
  return set;
}

function registrarProcessados_(linhas) {
  if (!linhas.length) return;
  const aba = planilha_().getSheetByName(ABA_PROCESSADOS);
  aba.getRange(aba.getLastRow() + 1, 1, linhas.length, 4).setValues(linhas);
}

/** Aplica os dados de um email ao registro do chamado (a ordem de leitura dos emails não importa). */
function mesclar_(mapa, d) {
  if (d.tipo === 'ignorar') return;
  const c = mapa[d.numero] || (mapa[d.numero] = { numero: d.numero, status: 'Aberto' });
  const temValor = v => v !== '' && v !== undefined && v !== null;
  const definir = (campo, valor) => { if (temValor(valor)) c[campo] = valor; };
  const seVazio = (campo, valor) => { if (!temValor(c[campo])) definir(campo, valor); };
  const oficial = ['aberto', 'andamento', 'resolvido'].indexOf(d.tipo) >= 0;
  const maisRecente = !c.dataAtualizacao || d.data >= c.dataAtualizacao;

  // Cliente/solicitante: emails de abertura/andamento/resolução têm prioridade sobre a avaliação
  (oficial ? definir : seVazio)('cliente', d.cliente);
  (oficial ? definir : seVazio)('solicitante', d.solicitante);
  if (d.dataAbertura && (!c.dataAbertura || d.dataAbertura < c.dataAbertura)) c.dataAbertura = d.dataAbertura;

  // Operador/assunto: vale o do email de resolução; senão, o do email mais recente
  if (d.tipo === 'resolvido' || (oficial && !c.msgResolucao && maisRecente)) {
    definir('operador', d.operador);
    definir('grupo', d.grupo);
    definir('assunto', d.assunto);
  } else {
    seVazio('operador', d.operador);
    seVazio('grupo', d.grupo);
    seVazio('assunto', d.assunto);
  }

  if (d.tipo === 'atribuido') c.atribuidoAMim = 'Sim';

  if (d.tipo === 'aberto') {
    if (!c.msgAbertura) c.msgAbertura = d.msgId;
    seVazio('resumo', d.resumo);
  }

  if (oficial && maisRecente) {
    c.dataAtualizacao = d.data;
    c.msgUltima = d.msgId;
    if (c.status !== 'Resolvido') c.status = d.rotulo || c.status;
    if (!c.msgResolucao && d.resumo) c.resumo = d.resumo; // conversa mais recente = mais completa
  }

  if (d.tipo === 'resolvido') {
    c.status = 'Resolvido';
    if (!c.dataResolucao || !c.msgResolucao || d.data > c.dataResolucao) {
      c.dataResolucao = d.data;
      c.msgResolucao = d.msgId;
      definir('descricaoAcao', d.descricaoAcao);
      definir('resumo', d.resumo);
    }
  }

  if (d.tipo === 'avaliacao') {
    c.dataAvaliacao = d.data;
    c.msgAvaliacao = d.msgId;
    definir('notaAtendimento', d.notaAtendimento);
    definir('notaSistema', d.notaSistema);
    definir('resolucaoCliente', d.resolucaoCliente);
    definir('comentario', d.comentario);
    if (!c.dataResolucao && d.dataFinalizacao) c.dataResolucao = d.dataFinalizacao;
    c.status = 'Resolvido';
  }
}

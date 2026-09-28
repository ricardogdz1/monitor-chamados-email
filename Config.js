/**
 * CONFIGURAÇÕES DO PAINEL DE CHAMADOS
 * Ajuste aqui caso o DeskManager mude o formato dos emails.
 * Depois de alterar, rode reprocessarTudo().
 */
const CONFIG = {
  // Pasta do Google Drive onde a planilha de dados é criada
  PASTA_DRIVE_ID: '14i97cnr5rTbys2vsH4ofdSLLDvmC3RIT',
  NOME_PLANILHA: 'Painel de Chamados - Dados',

  // Busca do Gmail que encontra os emails do DeskManager
  GMAIL_QUERY: '(from:helpdesk@agro1.inf.br OR from:suporteagro1@agro1.inf.br) -"o que achou"',

  // Número do chamado no assunto. Ex.: 0926-005381
  REGEX_NUMERO: /\b(\d{4}-\d{6})\b/,

  // Rótulos procurados no corpo do email (formato "Rótulo: valor")
  ROTULOS: {
    cliente:       ['Cliente'],
    solicitante:   ['Solicitante'],
    operador:      ['Operador/Grupo', 'Operador'],
    assunto:       ['Assunto'],
    dataAbertura:  ['Data - Hora', 'Data e Hora'],
    descricaoAcao: ['Descrição da Ação', 'Descricao da Acao'],
    comentario:    ['Comentário', 'Comentario', 'Observação', 'Sugestão']
  },

  // Textos das perguntas da avaliação
  NOTA_ATENDIMENTO: /profissional|atendente|atendimento prestado/i,
  NOTA_SISTEMA: /ferramenta|sistema|software/i,

  // Limites
  TAMANHO_MAX_RESUMO: 40000,     // caracteres guardados da conversa na planilha (limite da célula: 50.000)
  TEMPO_MAX_EXECUCAO_MS: 4.5 * 60 * 1000,
  INTERVALO_SYNC_MIN: 10,
  FUSO: 'America/Sao_Paulo'
};

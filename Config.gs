/**
 * ----------------------------------------------------------------------------
 * ARQUIVO DE CONFIGURAÇÃO - CONTROLE DE INDUSTRIALIZAÇÃO (ROTEAMENTO)
 * ----------------------------------------------------------------------------
 * * Versão: 1.1.0 (Com roteamento de e-mail por Destinatário)
 */

var CONFIG = {
  // --- E-MAILS DE DESTINO ---
  
  // 1. E-mail para pendências da "INNOVA PETIKO"
  EMAIL_ESPECIFICO: "higorpiccirilli@gmail.com", 
  
  // 2. E-mail para TODO O RESTO (Padrão)
  EMAIL_PADRAO: "higorpiccirilli@gmail.com",
  
  // --- REGRA DE ROTEAMENTO ---
  // Nome exato (ou parte principal) para identificar o cliente específico
  NOME_ALVO_ESPECIFICO: "INNOVA PETIKO INDUSTRIA DE ARTIGOS PARA ANIMAIS LTDA",

  // --- GERAL ---
  NOME_ABA: "Painel pendencias de retorno", // Nome da aba de resumo

  // --- ESTRUTURA DA PLANILHA (ABA PENDENCIAS) ---
  LINHA_INICIAL: 4,
  
  // Índices das Colunas (Javascript começa no 0)
  // A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9
  
  INDEX_COLUNA_NF: 6,          // Coluna G (NF Remessa)
  INDEX_COLUNA_DIAS: 9,        // Coluna J (Dias Restantes)
  
  // ATENÇÃO: Na aba 'Pendencias', a Razão Social (que vem da coluna R original)
  // geralmente cai na Coluna F (Índice 5) dependendo da sua QUERY.
  // Se estiver na coluna R mesmo na aba Pendencias, mude para 17.
  INDEX_COLUNA_DESTINATARIO: 5, // Coluna F (Razão Social / Destinatário)

  // --- STATUS E CORES (HTML) ---
  CATEGORIAS: {
    VENCIDO: {
      texto: "⚫ VENCIDO",
      estilo: "color: white; background-color: black; padding: 2px 6px; border-radius: 4px; font-weight: bold;"
    },
    URGENTE: {
      texto: "🔴 SUPER URGENTE (0-19 dias)",
      estilo: "color: #8B0000; background-color: #ffcccc; padding: 2px 6px; border-radius: 4px; font-weight: bold;"
    },
    ATENCAO: {
      texto: "🟠 ATENÇÃO (20-39 dias)",
      estilo: "color: #b45f06; background-color: #ffe599; padding: 2px 6px; border-radius: 4px;"
    },
    MORNO: {
      texto: "🟡 MORNO (40-69 dias)",
      estilo: "color: #7f6000; background-color: #ffffcc; padding: 2px 6px; border-radius: 4px;"
    },
    PRAZO: {
      texto: "🟢 NO PRAZO (+70 dias)",
      estilo: "color: green; font-weight: bold;"
    }
  }
};

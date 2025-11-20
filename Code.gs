/**
 * ----------------------------------------------------------------------------
 * 🏭 CONTROLE DE PRAZOS DE INDUSTRIALIZAÇÃO - SCRIPT PRINCIPAL
 * ----------------------------------------------------------------------------
 * * 📄 DESCRIÇÃO DO SISTEMA:
 * Este script é o motor de automação para gestão fiscal de remessas. 
 * Ele monitora a aba de "Pendencias", identifica Notas Fiscais que ainda não 
 * retornaram e calcula o prazo restante para evitar passivos fiscais.
 * * 👥 CRIADORES:
 * - Lógica de Negócio e Estrutura: [Seu Nome]
 * - Desenvolvimento e Refatoração: Gemini AI
 * * 🔖 VERSÃO:
 * 1.1.0 - Build Atual (Inclui Roteamento de E-mail por Destinatário)
 * * ⚙️ O QUE ESTE SCRIPT FAZ:
 * 1. Leitura: Acessa a aba 'Pendencias' e lê os dados a partir da linha configurada.
 * 2. Deduplicação: Se uma NF tem 10 itens, o script agrupa em apenas 1 aviso para não lotar o e-mail.
 * 3. Roteamento: Separa NFs de clientes específicos (VIPs) para gestores dedicados.
 * 4. Formatação: Gera um e-mail HTML com cores de semáforo (Preto/Vermelho/Amarelo/Verde).
 * * 🚀 GUIA RÁPIDO DE USO:
 * - Para configurar E-mails ou Nomes de Abas: Edite APENAS o arquivo 'Config.gs'.
 * - Para testar: Selecione a função 'enviarRelatorioIndustrializacao' acima e clique em 'Executar'.
 * - Para automatizar: Crie um Acionador (Trigger) no menu "Acionadores" (ícone de relógio) à esquerda.
 * * ----------------------------------------------------------------------------
 */

function enviarRelatorioIndustrializacao() {
  Logger.log(">>> INICIANDO O SCRIPT COM ROTEAMENTO...");

  // 1. Acesso à Planilha
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NOME_ABA);
  
  if (!sheet) {
    Logger.log("ERRO CRÍTICO: Aba '" + CONFIG.NOME_ABA + "' não encontrada.");
    return;
  }

  var lastRow = sheet.getLastRow();
  
  if (lastRow < CONFIG.LINHA_INICIAL) {
    Logger.log("AVISO: A planilha parece vazia.");
    return;
  }

  // Pega dados da Linha 4 até o final, Colunas A até J (10 colunas)
  var numLinhasDados = lastRow - (CONFIG.LINHA_INICIAL - 1);
  var dados = sheet.getRange(CONFIG.LINHA_INICIAL, 1, numLinhasDados, 10).getValues();
  
  Logger.log("Lendo " + dados.length + " linhas de dados...");
  
  // 2. Listas Separadas (Roteamento)
  var relatorioPadrao = {};      // Para todos os outros
  var relatorioEspecifico = {};  // Para Innova Petiko
  var nfsProcessadas = {};       // Para evitar duplicidade global

  var contadorPadrao = 0;
  var contadorEspecifico = 0;

  for (var i = 0; i < dados.length; i++) {
    var linhaReal = i + CONFIG.LINHA_INICIAL;
    var nf = dados[i][CONFIG.INDEX_COLUNA_NF];
    var diasRestantes = dados[i][CONFIG.INDEX_COLUNA_DIAS];
    var destinatario = dados[i][CONFIG.INDEX_COLUNA_DESTINATARIO]; // Agora lê o Destinatário
    
    // Verifica se pegou os dados corretamente nas primeiras linhas
    if (i < 3) {
      Logger.log("Linha " + linhaReal + " | NF: " + nf + " | Dest: " + destinatario);
    }

    if (nf === "" || diasRestantes === "") continue;
    if (nfsProcessadas[nf]) continue; // Já processou essa NF nesta execução

    var infoStatus = obterStatusCategoria(diasRestantes);
    var itemRelatorio = {
      destinatario: destinatario, // Guarda o nome para exibir na tabela
      dias: diasRestantes,
      statusTexto: infoStatus.texto,
      statusEstilo: infoStatus.estilo
    };

    // --- ROTEAMENTO ---
    // Verifica se é o cliente específico (Innova)
    // .trim() remove espaços extras, .toUpperCase() ignora maiúsculas/minúsculas
    if (destinatario && destinatario.toString().trim().toUpperCase() === CONFIG.NOME_ALVO_ESPECIFICO.toUpperCase()) {
      relatorioEspecifico[nf] = itemRelatorio;
      contadorEspecifico++;
    } else {
      relatorioPadrao[nf] = itemRelatorio;
      contadorPadrao++;
    }

    nfsProcessadas[nf] = true; // Marca como processada
  }

  // 3. Envio dos E-mails (Separados)
  
  // Envia para o E-mail Específico (Se houver dados)
  if (contadorEspecifico > 0) {
    Logger.log("Enviando relatório ESPECÍFICO (" + contadorEspecifico + " pendências)...");
    enviarEmailFormatado(relatorioEspecifico, CONFIG.EMAIL_ESPECIFICO, "Relatório: " + CONFIG.NOME_ALVO_ESPECIFICO);
  }

  // Envia para o E-mail Padrão (Se houver dados)
  if (contadorPadrao > 0) {
    Logger.log("Enviando relatório PADRÃO (" + contadorPadrao + " pendências)...");
    enviarEmailFormatado(relatorioPadrao, CONFIG.EMAIL_PADRAO, "Relatório Geral de Pendências");
  }

  if (contadorEspecifico === 0 && contadorPadrao === 0) {
    Logger.log("Nenhuma pendência encontrada para enviar.");
  }
}

/**
 * Função auxiliar para determinar a categoria
 */
function obterStatusCategoria(dias) {
  var d = parseInt(dias);
  if (isNaN(d)) return CONFIG.CATEGORIAS.PRAZO;

  if (d < 0) return CONFIG.CATEGORIAS.VENCIDO;
  if (d <= 19) return CONFIG.CATEGORIAS.URGENTE;
  if (d <= 39) return CONFIG.CATEGORIAS.ATENCAO;
  if (d <= 69) return CONFIG.CATEGORIAS.MORNO;
  return CONFIG.CATEGORIAS.PRAZO;
}

/**
 * Função auxiliar de envio (Agora aceita o e-mail de destino e título como parametro)
 */
function enviarEmailFormatado(dadosRelatorio, emailDestino, tituloEmail) {
  try {
    var listaOrdenada = Object.keys(dadosRelatorio).map(function(nf) {
      return { nf: nf, dados: dadosRelatorio[nf] };
    });

    listaOrdenada.sort(function(a, b) {
      return a.dados.dias - b.dados.dias;
    });

    var html = "<div style='font-family: Arial, sans-serif; color: #333;'>";
    html += "<h2 style='color: #202124;'>" + tituloEmail + "</h2>";
    html += "<p>Seguem as NFs pendentes de retorno:</p>";
    
    html += "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse; width: 100%; border: 1px solid #ddd;'>";
    html += "<tr style='background-color: #f1f3f4;'><th>NF Remessa</th><th>Destinatário</th><th>Dias Restantes</th><th>Status</th></tr>";

    for (var k = 0; k < listaOrdenada.length; k++) {
      var item = listaOrdenada[k];
      html += "<tr>";
      html += "<td><strong>" + item.nf + "</strong></td>";
      html += "<td>" + item.dados.destinatario + "</td>";
      html += "<td style='text-align: center;'>" + item.dados.dias + "</td>";
      html += "<td><span style='" + item.dados.statusEstilo + "'>" + item.dados.statusTexto + "</span></td>";
      html += "</tr>";
    }
    html += "</table></div>";

    MailApp.sendEmail({
      to: emailDestino,
      subject: "[Controle Industrialização] " + tituloEmail,
      htmlBody: html
    });
    
    Logger.log("SUCESSO: E-mail enviado para " + emailDestino);

  } catch (e) {
    Logger.log("ERRO AO ENVIAR E-MAIL PARA " + emailDestino + ": " + e.toString());
  }
}

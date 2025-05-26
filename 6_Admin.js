/**
 * Adiciona um menu personalizado à planilha quando ela é aberta.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Administração')
    .addItem('Limpar Todos os Dados de Teste', 'confirmClearAllData')
    .addSeparator() // Adiciona uma linha divisória no menu
    .addItem('Atualizar Dados de Cliente', 'showUpdateClientSidebar')
    .addToUi();
}

/**
 * Exibe a barra lateral para atualização de dados do cliente.
 */
function showUpdateClientSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('UpdateClient.html')
    .setTitle('Atualizador de Clientes')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Procura os dados de um cliente nas planilhas com base no termo de busca.
 * Esta função é chamada pela barra lateral (sidebar).
 * @param {string} searchTerm O Nome ou CPF do cliente.
 * @returns {Object|null} Um objeto com os dados do cliente ou nulo se não encontrado.
 */
function findClientData(searchTerm) {
  if (!searchTerm) throw new Error("Termo de busca não pode ser vazio.");
  return SheetService.findClientRowData(searchTerm.trim());
}

/**
 * Atualiza os dados de um cliente em todas as abas relevantes.
 * Esta função é chamada pela barra lateral (sidebar).
 * @param {Object} updateObject O objeto com o ID e os novos valores a serem atualizados.
 * @returns {string} Uma mensagem de sucesso.
 */
function updateClientData(updateObject) {
  if (!updateObject || !updateObject.id) throw new Error("Dados de atualização inválidos.");
  SheetService.updateClientRowData(updateObject);
  return "Cliente atualizado com sucesso!";
}

/**
 * Exibe um diálogo de confirmação MUITO IMPORTANTE antes de apagar os dados.
 * É uma medida de segurança para evitar a exclusão acidental de dados.
 */
function confirmClearAllData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ ATENÇÃO: AÇÃO IRREVERSÍVEL ⚠️',
    'Você tem certeza que deseja apagar TODAS as respostas do formulário e TODOS os dados das abas "Cliente", "Agenciamento" e "Comissao"?\n\nEsta ação não pode ser desfeita.',
    ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    try {
      clearAllTestData();
      ui.alert('Dados de teste foram limpos com sucesso!');
    } catch (e) {
      Logger.log(e);
      ui.alert(`Ocorreu um erro durante a limpeza: ${e.message}`);
    }
  } else {
    ui.alert('A operação de limpeza foi cancelada.');
  }
}


/**
 * Função principal que executa a limpeza dos dados.
 * 1. Limpa as abas de dados processados, mantendo os cabeçalhos.
 * 2. Limpa TODAS as respostas do formulário vinculado (o que também limpa a aba de respostas).
 */
function clearAllTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Iniciando processo de limpeza de dados...");

  // Parte 1: Limpar as abas de dados processados
  const sheetsToClear = [
    CONFIG.SHEET_NAMES.CLIENTS,
    CONFIG.SHEET_NAMES.AGENCY_FEES,
    CONFIG.SHEET_NAMES.COMMISSIONS
  ];

  sheetsToClear.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      // Se houver mais do que apenas a linha de cabeçalho, limpa o conteúdo
      if (lastRow > 1) {
        // Limpa da segunda linha até a última, todas as colunas que têm conteúdo
        const lastColumn = sheet.getLastColumn();
        if (lastColumn > 0) { // Verifica se há colunas para limpar
          sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
        }
        Logger.log(`Dados da aba "${sheetName}" limpos.`);
      }
    } else {
      Logger.log(`Aba "${sheetName}" não encontrada para limpeza.`);
    }
  });

  // Parte 2: Limpar as respostas do formulário vinculado
  const formUrl = ss.getFormUrl();
  if (formUrl) {
    try {
      const form = FormApp.openByUrl(formUrl);
      form.deleteAllResponses();
      Logger.log("Todas as respostas do formulário vinculado foram excluídas.");
    } catch (e) {
      Logger.log(`Erro ao tentar limpar respostas do formulário: ${e.message}. Verifique se o formulário ainda está vinculado e as permissões estão corretas.`);
    }
  } else {
    Logger.log("Nenhum formulário vinculado encontrado para limpar.");
  }

  Logger.log("Processo de limpeza de dados concluído.");
}
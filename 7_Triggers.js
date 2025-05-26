/**
 * @OnlyCurrentDoc
 * Arquivo principal para os gatilhos (Triggers).
 */

/**
 * Função acionada pelo envio do formulário. Orquestra a leitura, cálculo e escrita dos dados de uma nova apólice.
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e O objeto de evento do formulário.
 */
function onFormSubmitTrigger(e) {
  const functionName = "onFormSubmitTrigger";
  Logger.log(`Iniciando ${functionName}`);

  try {
    if (!e || !e.values) throw new Error("Evento de formulário inválido ou sem dados.");

    // 1. Mapeia os dados brutos do formulário para um objeto organizado
    const formData = Utils.mapFormToObject(e.values);

    // 2. Prepara os dados para cada aba, realizando os cálculos iniciais
    const newPolicyData = CalculationService.prepareNewPolicyData(formData);

    // 3. Escreve os dados processados em cada uma das abas da planilha
    SheetService.appendRow(CONFIG.SHEET_NAMES.CLIENTS, newPolicyData.clientRow);
    SheetService.appendRow(CONFIG.SHEET_NAMES.AGENCY_FEES, newPolicyData.agencyRow);
    SheetService.appendRow(CONFIG.SHEET_NAMES.COMMISSIONS, newPolicyData.commissionRow);

    Logger.log(`${functionName} concluído com sucesso para ID ${newPolicyData.id}.`);

  } catch (error) {
    Logger.log(`ERRO GRAVE em ${functionName}: ${error.message}\nStack: ${error.stack}`);
    Utils.sendErrorEmail(
      `Erro Crítico no Processamento de Formulário`,
      `Ocorreu um erro grave durante a execução de ${functionName}:<br><br><b>Erro:</b> ${error.message}<br><b>Dados do Formulário:</b><br>${JSON.stringify(e.namedValues)}`
    );
  }
}

/**
 * Função acionada mensalmente por um gatilho de tempo. Orquestra a leitura, cálculo e envio do relatório de comissões.
 */
function sendMonthlyReportTrigger() {
  const functionName = "sendMonthlyReportTrigger";
  Logger.log(`Iniciando ${functionName}`);

  try {
    // 1. Lê e consolida todos os dados relevantes das planilhas
    const allData = SheetService.getConsolidatedData();

    // 2. Realiza todos os cálculos de pagamentos, descontos e estornos
    const paymentData = CalculationService.calculateMonthlyPayments(allData);

    // 3. Gera e envia o relatório por e-mail
    ReportService.sendBrokerReport(paymentData);

    Logger.log(`${functionName} concluído com sucesso.`);

  } catch (error) {
    Logger.log(`ERRO GRAVE em ${functionName}: ${error.message}\nStack: ${error.stack}`);
    Utils.sendErrorEmail(
      `Erro Crítico na Geração do Relatório Mensal`,
      `Ocorreu um erro grave durante a execução de ${functionName}:<br><br><b>Erro:</b> ${error.message}`
    );
  }
}
/**
 * @OnlyCurrentDoc
 */

// --- Configurações ---
const NOME_ABA_CLIENTES = "ClientesStatus";
const NOME_ABA_AGENCIAMENTOS = "Agenciamentos";
const NOME_ABA_COMISSOES = "Comissoes";
const EMAIL_ADMIN_ERRO = "vitoroliveiraguedeshugo@gmail.com";

// Mapeamento Colunas Formulário -> Chave no Objeto 'dadosForm'
// Manter a mesma ordem das perguntas do formulário
const MAP_FORM_TO_DATA_KEY = {
  0: 'dtCadastro',           // dtCadastro (geralmente a primeira coluna)
  1: 'clienteNome',         // Nome do Cliente
  2: 'clienteCPF',          // CPF do Cliente
  3: 'corretor',            // Corretor Responsável
  4: 'contrato',             // contrato/Seguro
  5: 'premioMensal',        // Valor Prêmio Mensal (R$)
  6: 'formaPagto',          // Forma Pagto Cliente (Se você tiver/precisar)
  7: 'taxaAgencP1Perc',     // Taxa Agenciamento P1 (%)
  8: 'taxaAgencP2Perc',     // Taxa Agenciamento P2 (%)
  9: 'pagarP2AposMeses',    // Pagar Agenciamento P2 após (Meses Cliente)
  10: 'taxaComissaoPerc',    // Taxa Comissão (%)
  11: 'inicioComissaoApos',  // Início Comissão (Após Meses)
  12: 'estornoAgenc',        // Estorno Agenc. (Meses)
  13: 'taxaEstornoPerc'      // Taxa Estorno Agenc (%)
};


// Função Principal: Esta função é chamada automaticamente quando o formulário é enviado
// Esta função executa as funções auxiliares para preencher as abas correspondentes 
function onFormSubmit(e) {
  const functionName = "onFormSubmit";
  Logger.log("Iniciando " + functionName);
  Logger.log("Evento recebido (e): " + JSON.stringify(e));
  try {
    if (!e || !e.values) throw new Error("Evento de formulário inválido ou sem dados.");

    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetAgenciamentos = sheet.getSheetByName(NOME_ABA_AGENCIAMENTOS);
    const sheetClientes = sheet.getSheetByName(NOME_ABA_CLIENTES);
    const sheetComissoes = sheet.getSheetByName(NOME_ABA_COMISSOES);

    // Validação das Abas
    if (!sheetClientes) throw new Error("Aba '" + NOME_ABA_CLIENTES + "' não encontrada.");
    if (!sheetAgenciamentos) throw new Error("Aba '" + NOME_ABA_AGENCIAMENTOS + "' não encontrada.");
    if (!sheetComissoes) throw new Error("Aba '" + NOME_ABA_COMISSOES + "' não encontrada.");

    // --- Mapeia dados brutos para um objeto estruturado ---
    const dadosFormRaw = e.values;
    const dadosForm = {};
    for (const index in MAP_FORM_TO_DATA_KEY) {
      if (dadosFormRaw[index] !== undefined) {
        dadosForm[MAP_FORM_TO_DATA_KEY[index]] = dadosFormRaw[index];
      } else {
        Logger.log("Aviso: Dado esperado no índice " + index + " (" + MAP_FORM_TO_DATA_KEY[index] + ") não encontrado.");
        dadosForm[MAP_FORM_TO_DATA_KEY[index]] = null;
      }
    }
    // Adiciona dtCadastro se não veio do form
    if (!dadosForm.dtCadastro) dadosForm.dtCadastro = new Date();
    Logger.log("Dados do formulário mapeados: " + JSON.stringify(dadosForm));


    // --- Gerar ID Único ---
    const idAgenciamento = "AG-" + new Date().getTime() + "-" + Math.floor(Math.random() * 100);
    Logger.log("ID Agenciamento gerado: " + idAgenciamento);

    // Preencher as abas com os dados do formulário
    let writingSuccess = true;
    try {
      setClientStatus(dadosForm, idAgenciamento, sheetClientes);
      Logger.log("Sucesso ao popular " + NOME_ABA_CLIENTES);
    } catch (errCliente) {
      writingSuccess = false;
      Logger.log("ERRO em setClientStatus: " + errCliente + "\nStack: " + errCliente.stack);
    }

    try {
      setAgencyStatus(dadosForm, idAgenciamento, sheetAgenciamentos);
      Logger.log("Sucesso ao popular " + NOME_ABA_AGENCIAMENTOS);
    } catch (errAgenc) {
      writingSuccess = false;
      Logger.log("ERRO em setAgencyStatus: " + errAgenc + "\nStack: " + errAgenc.stack);
    }

    try {
      setCommissionStatus(dadosForm, idAgenciamento, sheetComissoes);
      Logger.log("Sucesso ao popular " + NOME_ABA_COMISSOES);
    } catch (errComiss) {
      writingSuccess = false;
      Logger.log("ERRO em setCommissionStatus: " + errComiss + "\nStack: " + errComiss.stack);
    }

    // --- Log Final e Notificação de Erro Grave ---
    if (writingSuccess) {
      Logger.log(functionName + " concluído com sucesso para ID " + idAgenciamento);
    } else {
      Logger.log("ERRO: " + functionName + " concluído com falhas para ID " + idAgenciamento + ". Verifique os logs acima.");
      MailApp.sendEmail(EMAIL_ADMIN_ERRO, "Erro Urgente no Processamento de Formulário (Parcial)",
        "Script: " + functionName + "\nID Agenciamento: " + idAgenciamento + "\nOcorreram erros ao preencher uma ou mais abas. Verifique os Logs na planilha.");
    }

  } catch (error) {
    // Erro na função principal (ex: aba não encontrada, evento inválido)
    Logger.log("ERRO GRAVE em " + functionName + ": " + error + "\nStack: " + error.stack);
    MailApp.sendEmail(EMAIL_ADMIN_ERRO, "Erro Crítico no Processamento de Formulário",
      "Script: " + functionName + "\nErro: " + error + "\nO processamento do formulário falhou completamente. Verifique os Logs.");
  }
}

/**
 * Serviço para todas as interações com a Planilha Google.
 * ESTRUTURA FINAL E CORRIGIDA: Usa o "Revealing Module Pattern" (IIFE)
 */
const SheetService = (function () {

  // --- FUNÇÕES INTERNAS / PRIVADAS ---

  function _getSheet(name) {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); // Correto, usa a planilha ativa
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(`Aba "${name}" não foi encontrada.`);
    return sheet;
  }

  function _mapRowToObject(row, headers) {
    const obj = {};
    headers.forEach((header, i) => {
      if (header) obj[header] = row[i];
    });
    return obj;
  }

  function _mapObjectToRow(obj, headers) {
    return headers.map(header => (obj[header] !== undefined ? obj[header] : ""));
  }

  // Função interna para obter todos os dados consolidados, incluindo o número da linha na planilha
  function _getConsolidatedDataInternal() { // Renomeada para clareza
    const consolidatedData = {};
    const C = CONFIG.COLUMN_NAMES;

    const clientSheet = _getSheet(CONFIG.SHEET_NAMES.CLIENTS);
    const clientValues = clientSheet.getDataRange().getValues();
    const clientHeaders = clientValues.shift(); // Remove e retorna o cabeçalho

    clientValues.forEach((row, i) => {
      const clientObj = _mapRowToObject(row, clientHeaders);
      const id = clientObj[C.ID];
      if (id) {
        clientObj.linhaPlanilha = i + 2; // Número da linha real na planilha (dados começam na linha 2)
        consolidatedData[id] = { client: clientObj, agency: null, commission: null };
      }
    });

    const agencySheet = _getSheet(CONFIG.SHEET_NAMES.AGENCY_FEES);
    const agencyValues = agencySheet.getDataRange().getValues();
    const agencyHeaders = agencyValues.shift();

    agencyValues.forEach((row, i) => {
      const agencyObj = _mapRowToObject(row, agencyHeaders);
      const id = agencyObj[C.ID];
      if (consolidatedData[id]) {
        agencyObj.linhaPlanilha = i + 2; // Número da linha real
        consolidatedData[id].agency = agencyObj;
      }
    });

    const commissionSheet = _getSheet(CONFIG.SHEET_NAMES.COMMISSIONS);
    const commissionValues = commissionSheet.getDataRange().getValues();
    const commissionHeaders = commissionValues.shift();

    commissionValues.forEach((row, i) => {
      const commissionObj = _mapRowToObject(row, commissionHeaders);
      const id = commissionObj[C.ID];
      if (consolidatedData[id]) {
        commissionObj.linhaPlanilha = i + 2; // Número da linha real
        consolidatedData[id].commission = commissionObj;
      }
    });

    // Logger.log("Dados consolidados: " + JSON.stringify(consolidatedData)); // Para depuração, se necessário
    return consolidatedData;
  }

  // --- OBJETO RETORNADO (A API PÚBLICA DO SERVIÇO) ---
  return {
    appendRow: function (sheetName, rowObject) {
      const sheet = _getSheet(sheetName);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const rowAsArray = _mapObjectToRow(rowObject, headers);
      sheet.appendRow(rowAsArray);
    },

    getConsolidatedData: function () { // Função pública que chama a interna
      return _getConsolidatedDataInternal();
    },

    findClientRowData: function (searchTerm) {
      const C = CONFIG.COLUMN_NAMES;
      const allData = _getConsolidatedDataInternal(); // Pega todos os dados já com as linhas
      let foundPolicyData = null;

      // Itera sobre os dados consolidados para encontrar o cliente
      for (const id in allData) {
        const policy = allData[id];
        if (policy.client) {
          const clientName = String(policy.client[C.CLIENTE_NOME] || "").trim().toLowerCase();
          const clientCPF = String(policy.client[C.CLIENTS.CPF] || "").trim();
          const searchLower = searchTerm.toLowerCase();

          if (clientName === searchLower || clientCPF === searchTerm) {
            foundPolicyData = policy;
            break;
          }
        }
      }

      if (!foundPolicyData || !foundPolicyData.client || !foundPolicyData.agency || !foundPolicyData.commission) {
        Logger.log("Cliente não encontrado ou dados incompletos para o termo: " + searchTerm);
        return null;
      }

      const { client, agency, commission } = foundPolicyData;

      return {
        id: client[C.ID], // O ID vem do objeto client
        clienteNome: client[C.CLIENTE_NOME],
        statusCliente: client[C.CLIENTS.STATUS],
        dataCancelamento: client[C.CLIENTS.DATA_CANCELAMENTO],
        inadimplenteMes: client[C.CLIENTS.INADIMPLENTE],
        p1Pago: agency[C.AGENCY.P1_PAGO],
        p2Pago: agency[C.AGENCY.P2_PAGO],
        estornoAplicado: agency[C.AGENCY.ESTORNO_APLICADO],
        estornoDeduzido: agency[C.AGENCY.ESTORNO_DEDUZIDO],
        comissoesPagasQtd: commission[C.COMMISSION.QTD_PAGA] || 0
      };
    },

    updateClientRowData: function (updateObject) {
      const C = CONFIG.COLUMN_NAMES;
      const allData = _getConsolidatedDataInternal();
      const policyData = allData[updateObject.id];

      if (!policyData || !policyData.client || !policyData.agency || !policyData.commission) {
        throw new Error("Não foi possível encontrar os dados completos na planilha para o ID: " + updateObject.id + ". Verifique os logs e se o cliente existe em todas as 3 abas principais.");
      }

      const { client, agency, commission } = policyData; // Contêm .linhaPlanilha

      const sheetClientes = _getSheet(CONFIG.SHEET_NAMES.CLIENTS);
      const sheetAgenc = _getSheet(CONFIG.SHEET_NAMES.AGENCY_FEES);
      const sheetComiss = _getSheet(CONFIG.SHEET_NAMES.COMMISSIONS);

      const headersClientes = sheetClientes.getRange(1, 1, 1, sheetClientes.getLastColumn()).getValues()[0];
      const headersAgenc = sheetAgenc.getRange(1, 1, 1, sheetAgenc.getLastColumn()).getValues()[0];
      const headersComiss = sheetComiss.getRange(1, 1, 1, sheetComiss.getLastColumn()).getValues()[0];

      const updateCell = (sheet, row, headers, columnName, value) => {
        const colIndex = headers.indexOf(columnName);
        if (colIndex !== -1 && row) {
          sheet.getRange(row, colIndex + 1).setValue(value);
        } else if (!row) {
          Logger.log(`AVISO: Número da linha inválido para a aba ${sheet.getName()}. A coluna "${columnName}" (valor: ${value}) não foi atualizada.`);
        } else {
          Logger.log(`AVISO: A coluna "${columnName}" não foi encontrada na aba ${sheet.getName()} e não foi atualizada (valor: ${value}).`);
        }
      };

      updateCell(sheetClientes, client.linhaPlanilha, headersClientes, C.CLIENTS.STATUS, updateObject.statusCliente);
      updateCell(sheetClientes, client.linhaPlanilha, headersClientes, C.CLIENTS.DATA_CANCELAMENTO, updateObject.dataCancelamento || "");
      updateCell(sheetClientes, client.linhaPlanilha, headersClientes, C.CLIENTS.INADIMPLENTE, updateObject.inadimplenteMes);

      updateCell(sheetAgenc, agency.linhaPlanilha, headersAgenc, C.AGENCY.P1_PAGO, updateObject.p1Pago);
      updateCell(sheetAgenc, agency.linhaPlanilha, headersAgenc, C.AGENCY.P2_PAGO, updateObject.p2Pago);
      updateCell(sheetAgenc, agency.linhaPlanilha, headersAgenc, C.AGENCY.ESTORNO_APLICADO, updateObject.estornoAplicado);
      updateCell(sheetAgenc, agency.linhaPlanilha, headersAgenc, C.AGENCY.ESTORNO_DEDUZIDO, updateObject.estornoDeduzido);

      updateCell(sheetComiss, commission.linhaPlanilha, headersComiss, C.COMMISSION.QTD_PAGA, updateObject.comissoesPagasQtd);
    }
  };
})();
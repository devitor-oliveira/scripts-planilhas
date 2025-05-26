/**
 * Serviço para todos os cálculos de negócio.
 * Funções aqui devem ser "puras": recebem dados, retornam resultados, sem efeitos colaterais.
 */
const CalculationService = {
  /**
   * A partir dos dados do formulário, prepara os objetos para cada aba da planilha.
   * @param {Object} formData Objeto com os dados vindos do formulário.
   * @returns {Object} Um objeto contendo o ID gerado e os objetos de linha para cada aba.
   */
  prepareNewPolicyData(formData) {
    const C = CONFIG.COLUMN_NAMES;

    const id = `AG-${new Date().getTime()}-${Math.floor(Math.random() * 100)}`;
    const premio = Utils.parseBrazilianNumber(formData.premioMensal);
    if (premio <= 0) throw new Error("Prêmio mensal inválido.");

    const taxaP1 = Utils.parseBrazilianNumber(formData.taxaAgencP1Perc);
    const taxaP2 = Utils.parseBrazilianNumber(formData.taxaAgencP2Perc);
    const taxaEstorno = Utils.parseBrazilianNumber(formData.taxaEstornoPerc);
    const taxaComissao = Utils.parseBrazilianNumber(formData.taxaComissaoPerc);

    // const agencP1Pago = formData.agencP1PagoNaProposta === 'Sim' ? 'S' : 'N'; // Mantido comentado pois FORM_MAP não tem a chave

    const valorP1 = premio * (taxaP1 / 100.0);
    const valorP2 = premio * (taxaP2 / 100.0);
    const valorEstornoPotencial = (valorP1 + valorP2) * (taxaEstorno / 100.0);
    const valorComissaoParcela = premio * (taxaComissao / 100.0);

    return {
      id: id,
      clientRow: {
        [C.ID]: id,
        [C.CLIENTS.DT_CADASTRO]: formData.dtCadastro,
        [C.CLIENTE_NOME]: formData.clienteNome,
        [C.CLIENTS.CPF]: formData.clienteCPF,
        [C.CORRETOR]: formData.corretor,
        [C.CLIENTS.CONTRATO]: formData.contrato,
        [C.CLIENTS.PREMIO]: premio,
        [C.CLIENTS.ESTORNO_MESES]: parseInt(formData.estornoAgencMeses) || 0,
        [C.CLIENTS.STATUS]: "Ativo",
        [C.CLIENTS.MESES_PAGOS]: 0,
        [C.CLIENTS.INADIMPLENTE]: "N",
        [C.CLIENTS.DATA_STATUS]: new Date(),
        [C.CLIENTS.DATA_CANCELAMENTO]: ""
      },
      agencyRow: {
        [C.ID]: id,
        [C.CORRETOR]: formData.corretor,
        [C.CLIENTE_NOME]: formData.clienteNome,
        [C.AGENCY.TAXA_P1]: taxaP1,
        [C.AGENCY.TAXA_P2]: taxaP2,
        [C.AGENCY.PAGAR_P2_APOS]: parseInt(formData.pagarP2AposMeses) || 1,
        [C.AGENCY.VALOR_P1]: valorP1,
        [C.AGENCY.VALOR_P2]: valorP2,
        [C.AGENCY.TAXA_ESTORNO]: taxaEstorno,
        [C.AGENCY.VALOR_ESTORNO_POTENCIAL]: valorEstornoPotencial,
        [C.AGENCY.P1_PAGO]: "N",
        [C.AGENCY.P2_PAGO]: "N",
        [C.AGENCY.ESTORNO_APLICADO]: "N",
        [C.AGENCY.ESTORNO_DEDUZIDO]: "N",
      },
      commissionRow: {
        [C.ID]: id,
        [C.CORRETOR]: formData.corretor,
        [C.CLIENTE_NOME]: formData.clienteNome,
        [C.COMMISSION.TAXA]: taxaComissao,
        [C.COMMISSION.INICIO_APOS]: parseInt(formData.inicioComissaoApos) || 0,
        [C.COMMISSION.VALOR_PARCELA]: valorComissaoParcela,
        [C.COMMISSION.QTD_PAGA]: 0,
      }
    };
  },

  calculateMonthlyPayments(allData) {
    const C = CONFIG.COLUMN_NAMES;
    const ADMIN_FEE = CONFIG.BUSINESS_RULES.ADMIN_FEE_PERCENTAGE;
    const paymentsByBroker = {};

    for (const id in allData) {
      const { client, agency, commission } = allData[id];

      if (!client || !agency || !commission) {
        Logger.log(`Aviso: Dados ausentes para ID ${id}. Pulando.`);
        continue;
      }

      const corretor = client[C.CORRETOR];
      if (!corretor) {
        Logger.log(`Aviso: Corretor não definido para ID ${id}. Pulando.`);
        continue;
      }

      let statusCliente = client[C.CLIENTS.STATUS];
      if (client[C.CLIENTS.DATA_CANCELAMENTO] && statusCliente === "Ativo") {
        statusCliente = "Cancelado";
      }

      let paymentP1Bruto = 0, paymentP2Bruto = 0, paymentCommissionBruto = 0, clawback = 0;
      const details = [];

      const mesesPagos = parseInt(client[C.CLIENTS.MESES_PAGOS]) || 0;
      const pagarP2Apos = parseInt(agency[C.AGENCY.PAGAR_P2_APOS]) || 1;
      const inicioComissaoApos = parseInt(commission[C.COMMISSION.INICIO_APOS]) || 0;
      const comissoesPagasQtd = parseInt(commission[C.COMMISSION.QTD_PAGA]) || 0;
      const inadimplente = String(client[C.CLIENTS.INADIMPLENTE]).toUpperCase() === 'S';
      const premioMensalCliente = parseFloat(client[C.CLIENTS.PREMIO]) || 0; // <-- ADICIONADO

      if (String(agency[C.AGENCY.P1_PAGO]).toUpperCase() === 'S' && mesesPagos === 0) {
        details.push(`Ag. P1 de ${Utils.toCurrency(parseFloat(agency[C.AGENCY.VALOR_P1]))} pago na proposta (informativo)`);
      }

      if (String(agency[C.AGENCY.P1_PAGO]).toUpperCase() !== 'S') {
        paymentP1Bruto = parseFloat(agency[C.AGENCY.VALOR_P1]) || 0;
        details.push(`Pagamento Ag.P1: ${Utils.toCurrency(paymentP1Bruto)}`);
      }

      if (parseFloat(agency[C.AGENCY.VALOR_P2]) > 0 && String(agency[C.AGENCY.P2_PAGO]).toUpperCase() !== 'S' && mesesPagos >= pagarP2Apos) {
        paymentP2Bruto = parseFloat(agency[C.AGENCY.VALOR_P2]) || 0;
        details.push(`Pagamento Ag.P2: ${Utils.toCurrency(paymentP2Bruto)}`);
      }

      if (inadimplente && statusCliente === 'Ativo' && mesesPagos > inicioComissaoApos) {
        details.push(`AVISO: Comissão suspensa devido à inadimplência neste mês.`);
      }

      if (statusCliente === 'Ativo' && !inadimplente && mesesPagos > inicioComissaoApos && mesesPagos > comissoesPagasQtd) {
        const comissoesACalcular = mesesPagos - comissoesPagasQtd;
        if (comissoesACalcular > 0) {
          paymentCommissionBruto = comissoesACalcular * (parseFloat(commission[C.COMMISSION.VALOR_PARCELA]) || 0);
          details.push(`${comissoesACalcular}x Comissão: ${Utils.toCurrency(paymentCommissionBruto)}`);
        }
      }

      const estornoAgencMeses = parseInt(client[C.CLIENTS.ESTORNO_MESES]) || 0;
      if (statusCliente === "Cancelado" && mesesPagos < estornoAgencMeses) {
        if (String(agency[C.AGENCY.ESTORNO_APLICADO]).toUpperCase() === 'S' && String(agency[C.AGENCY.ESTORNO_DEDUZIDO]).toUpperCase() !== 'S') {
          clawback = parseFloat(agency[C.AGENCY.VALOR_ESTORNO_POTENCIAL]) || 0;
          details.push(`Estorno: ${Utils.toCurrency(-clawback)}`);
        } else {
          details.push(`AVISO: Estorno Potencial não aplicado`);
        }
      }

      let infoAgenciamentoP2 = "";
      if (parseFloat(agency[C.AGENCY.VALOR_P2]) > 0) {
        if (String(agency[C.AGENCY.P2_PAGO]).toUpperCase() === 'S') {
          infoAgenciamentoP2 = "Agenciamento P2 já foi pago.";
        } else if (mesesPagos < pagarP2Apos) {
          infoAgenciamentoP2 = `Agenciamento P2 de ${Utils.toCurrency(parseFloat(agency[C.AGENCY.VALOR_P2]))} será pago após o ${pagarP2Apos}º mês.`;
        }
      }

      let infoComissao = "";
      const valorParcelaComissao = parseFloat(commission[C.COMMISSION.VALOR_PARCELA]) || 0;
      if (valorParcelaComissao > 0) {
        if (statusCliente === 'Ativo' && !inadimplente && mesesPagos <= inicioComissaoApos) {
          infoComissao = `Corretagem de ${Utils.toCurrency(valorParcelaComissao)}/mês iniciará após o ${inicioComissaoApos}º mês.`;
        } else if (statusCliente === 'Ativo' && !inadimplente) {
          infoComissao = `Corretagem de ${Utils.toCurrency(valorParcelaComissao)}/mês em andamento.`;
        }
      }

      const totalBrutoItem = paymentP1Bruto + paymentP2Bruto + paymentCommissionBruto;
      const totalLiquidoItem = (totalBrutoItem - clawback) * (1 - ADMIN_FEE);
      const detalhesString = details.length > 0 ? details.join(' | ') : "Nenhuma movimentação financeira este mês.";

      if (totalLiquidoItem !== 0 || detalhesString.includes("AVISO") || infoAgenciamentoP2 || infoComissao) {
        if (!paymentsByBroker[corretor]) {
          paymentsByBroker[corretor] = { totalBruto: 0, totalLiquido: 0, items: [] };
        }
        paymentsByBroker[corretor].totalBruto += totalBrutoItem - clawback;
        paymentsByBroker[corretor].totalLiquido += totalLiquidoItem;
        paymentsByBroker[corretor].items.push({
          id: id,
          cliente: client[C.CLIENTE_NOME],
          premioMensalCliente: premioMensalCliente, // <-- ADICIONADO
          valorBrutoItem: totalBrutoItem - clawback,
          valorLiquidoItem: totalLiquidoItem,
          detalhes: detalhesString,
          infoAgenciamentoP2: infoAgenciamentoP2,
          infoComissao: infoComissao,
          isAviso: detalhesString.includes("AVISO"),
          taxaP1: agency[C.AGENCY.TAXA_P1],
          taxaP2: agency[C.AGENCY.TAXA_P2],
          taxaComissao: commission[C.COMMISSION.TAXA],
        });
      }
    }
    return paymentsByBroker; // Ajuste: Não estamos mais retornando delinquentClients daqui, isso será feito em ReportService
  }
};
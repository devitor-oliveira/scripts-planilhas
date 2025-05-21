// Preencher a aba ClientesStatus
function setClientStatus(dadosForm, idAgenciamento, sheetClientes) {
    // Prepara a linha na ordem correta das colunas da aba ClientesStatus
    const setClientRows = [
        idAgenciamento,
        dadosForm.dtCadastro,
        dadosForm.clienteNome,
        dadosForm.clienteCPF,
        dadosForm.corretor,
        dadosForm.contrato,
        parseFloat(dadosForm.premioMensal) || 0,
        parseInt(dadosForm.estornoAgenc) || 0,
        // Status Manuais
        "Ativo", // Status Cliente
        0,       // Cliente Meses Pagos (Qtd)
        "N",     // Inadimplente Mês Atual? (S/N)
        new Date() // Data Status
    ];
    sheetClientes.appendRow(setClientRows);
}

// Preencher a aba Agenciamentos
function setAgencyStatus(dadosForm, idAgenciamento, sheetAgenciamentos) {
    // Extrai e calcula valores específicos para esta aba
    const premioMensal = parseFloat(dadosForm.premioMensal) || 0;
    const taxaAgencP1Perc = parseFloat(dadosForm.taxaAgencP1Perc) || 0;
    const taxaAgencP2Perc = parseFloat(dadosForm.taxaAgencP2Perc) || 0;
    const pagarP2AposMeses = parseInt(dadosForm.pagarP2AposMeses) || 1;
    const taxaEstornoPerc = parseFloat(dadosForm.taxaEstornoPerc) || 0;

    if (premioMensal <= 0) throw new Error("Prêmio mensal inválido para cálculo de Agenciamento.");

    const valorAgenciamentoP1 = premioMensal * (taxaAgencP1Perc / 100.0);
    const valorAgenciamentoP2 = premioMensal * (taxaAgencP2Perc / 100.0);
    const valorEstornoPotencial = (valorAgenciamentoP1 + valorAgenciamentoP2) * (taxaEstornoPerc / 100.0);

    // Prepara a linha na ordem correta das colunas da aba Agenciamentos
    const setAgencyRows = [
        idAgenciamento,
        dadosForm.corretor,
        dadosForm.clienteNome,
        taxaAgencP1Perc,
        taxaAgencP2Perc,
        pagarP2AposMeses,
        valorAgenciamentoP1,
        valorAgenciamentoP2,
        taxaEstornoPerc,
        valorEstornoPotencial,
        // Status Manuais
        "N", // Agenc Parte 1 Pago? (S/N)
        "N", // Agenc Parte 2 Pago? (S/N)
        "N", // Estorno Aplicado? (S/N)
        "N"  // Estorno Já Deduzido? (S/N)
    ];
    sheetAgenciamentos.appendRow(setAgencyRows);
}

// Preencher a aba Comissões
function setCommissionStatus(dadosForm, idAgenciamento, sheetComissoes) {
    // Extrai e calcula valores específicos para esta aba
    const premioMensal = parseFloat(dadosForm.premioMensal) || 0;
    const taxaComissaoPerc = parseFloat(dadosForm.taxaComissaoPerc) || 0;
    const inicioComissaoApos = parseInt(dadosForm.inicioComissaoApos) || 0;

    if (premioMensal <= 0) throw new Error("Prêmio mensal inválido para cálculo de Comissão.");

    const valorComissaoParcela = premioMensal * (taxaComissaoPerc / 100.0);

    // Prepara a linha na ordem correta das colunas da aba Comissoes
    const sheetComissionsRows = [
        idAgenciamento,
        dadosForm.corretor,
        dadosForm.clienteNome,
        taxaComissaoPerc,
        inicioComissaoApos,
        valorComissaoParcela,
        // Status Manuais
        0 // Comissões Pagas ao Corretor (Qtd/meses)
    ];
    sheetComissoes.appendRow(sheetComissionsRows);
}
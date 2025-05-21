/**
 * @OnlyCurrentDoc
 */

// --- Configurações ---
const ID_PLANILHA_MESTRE_MENSAL = SpreadsheetApp.getActiveSpreadsheet().getId();
const NOME_ABA_CLIENTES_MENSAL = "ClientesStatus";
const NOME_ABA_AGENCIAMENTOS_MENSAL = "Agenciamentos";
const NOME_ABA_COMISSOES_MENSAL = "Comissoes";
const EMAIL_RESPONSAVEL_MENSAL = "vitoroliveiraguedeshugo@gmail.com";
const ASSUNTO_EMAIL_MENSAL = "Relatório Mensal de Pagamentos a Corretores";
const TAXA_ADMINISTRATIVA = 0.20;


// Função auxiliar para formatar moeda
const toCurrencyFormat = value => {
    if (typeof value !== 'number' || isNaN(value)) {
        return Number(0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
    }
    return Number(value).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}


/**
 * Script Mensal: Lê as 3 abas, calcula pagamentos/descontos
 * incluindo taxa administrativa, e envia um email com o relatório geral.
 */
function sendEmailMonthly() {
    const functionName = "sendEmailMonthly";
    const hoje = new Date();

    const mesAno = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "MMMM 'de' yyyy");
    const assuntoEmail = ASSUNTO_EMAIL_MENSAL + " - " + mesAno;
    Logger.log("Iniciando " + functionName + " para " + mesAno);

    try {
        const ss = SpreadsheetApp.openById(ID_PLANILHA_MESTRE_MENSAL);
        const sheetClientes = ss.getSheetByName(NOME_ABA_CLIENTES_MENSAL);
        const sheetAgenciamentos = ss.getSheetByName(NOME_ABA_AGENCIAMENTOS_MENSAL);
        const sheetComissoes = ss.getSheetByName(NOME_ABA_COMISSOES_MENSAL);

        if (!sheetClientes) throw new Error("Aba '" + NOME_ABA_CLIENTES_MENSAL + "' não encontrada.");
        if (!sheetAgenciamentos) throw new Error("Aba '" + NOME_ABA_AGENCIAMENTOS_MENSAL + "' não encontrada.");
        if (!sheetComissoes) throw new Error("Aba '" + NOME_ABA_COMISSOES_MENSAL + "' não encontrada.");

        // Ler dados das abas
        const dataClientes = sheetClientes.getDataRange().getValues();
        const dataAgenciamentos = sheetAgenciamentos.getDataRange().getValues();
        const dataComissoes = sheetComissoes.getDataRange().getValues();

        // --- Mapeamento: Cabeçalho -> Índice ---
        // Cria um mapa de cabeçalhos para cada aba. Ignora cabeçalhos vazios.
        const headerMapClientes = {};
        dataClientes[0].forEach((header, i) => { if (header) headerMapClientes[String(header).trim()] = i; });

        const headerMapAgenc = {};
        dataAgenciamentos[0].forEach((header, i) => { if (header) headerMapAgenc[String(header).trim()] = i; });

        const headerMapComiss = {};
        dataComissoes[0].forEach((header, i) => { if (header) headerMapComiss[String(header).trim()] = i; });

        // Função interna para buscar dados pelo nome do cabeçalho
        // Inicializa objeto para rastrear cabeçalhos ausentes dentro desta execução
        const missingHeaders = {};
        function getVal(rowData, map, headerName) {
            const index = map[headerName];
            if (index === undefined) {
                if (!missingHeaders[headerName]) { // Loga apenas uma vez por execução
                    Logger.log("Aviso: Cabeçalho '" + headerName + "' não encontrado na planilha mestre correspondente.");
                    missingHeaders[headerName] = true;
                }
                return null;
            }
            return (rowData[index] !== undefined && rowData[index] !== '') ? rowData[index] : null;
        }

        // --- Processamento e Consolidação em Memória ---
        const clienteInfo = {};
        Logger.log("Processando dados da aba " + NOME_ABA_CLIENTES_MENSAL);

        for (let i = 1; i < dataClientes.length; i++) {
            const id = getVal(dataClientes[i], headerMapClientes, 'ID Agenciamento');
            if (id) {
                // Garante que o ID não é vazio antes de adicionar
                clienteInfo[id] = {
                    id: id,
                    linhaCliente: i + 1,
                    corretor: getVal(dataClientes[i], headerMapClientes, 'Corretor Responsável'),
                    clienteNome: getVal(dataClientes[i], headerMapClientes, 'Cliente Nome'),
                    status: getVal(dataClientes[i], headerMapClientes, 'Status Cliente'),
                    mesesPagos: parseInt(getVal(dataClientes[i], headerMapClientes, 'Cliente Meses Pagos (Qtd)')) || 0,
                    inadimplenteMes: getVal(dataClientes[i], headerMapClientes, 'Inadimplente Mês Atual? (S/N)'),
                    estornoAgenc: parseInt(getVal(dataClientes[i], headerMapClientes, 'Estorno Agenc. (Meses)')) || 0,
                    agenciamento: null,
                    comissao: null
                };
            } else if (dataClientes[i].some(cell => cell !== '')) {
                Logger.log("Aviso: ID Agenciamento vazio ou inválido na linha " + (i + 1) + " da aba " + NOME_ABA_CLIENTES_MENSAL + ". Linha ignorada.");
            }
        }

        Logger.log("Processando dados da aba " + NOME_ABA_AGENCIAMENTOS_MENSAL);
        for (let i = 1; i < dataAgenciamentos.length; i++) {
            const id = getVal(dataAgenciamentos[i], headerMapAgenc, 'ID Agenciamento');

            if (id && clienteInfo[id]) {
                clienteInfo[id].agenciamento = {
                    linhaAgenc: i + 1,
                    valorP1Base: parseFloat(getVal(dataAgenciamentos[i], headerMapAgenc, 'Valor Agenc Parte 1 (R$)')) || 0,
                    valorP2Base: parseFloat(getVal(dataAgenciamentos[i], headerMapAgenc, 'Valor Agenc Parte 2 (R$)')) || 0,
                    pagarP2AposMeses: parseInt(getVal(dataAgenciamentos[i], headerMapAgenc, 'Pagar P2 Após Meses')) || 1,
                    p1Pago: getVal(dataAgenciamentos[i], headerMapAgenc, 'Agenc Parte 1 Pago? (S/N)'),
                    p2Pago: getVal(dataAgenciamentos[i], headerMapAgenc, 'Agenc Parte 2 Pago? (S/N)'),
                    valorEstornoPotencialBase: parseFloat(getVal(dataAgenciamentos[i], headerMapAgenc, 'Valor Estorno Potencial (R$)')) || 0,
                    estornoAplicado: getVal(dataAgenciamentos[i], headerMapAgenc, 'Estorno Aplicado? (S/N)'),
                    estornoDeduzido: getVal(dataAgenciamentos[i], headerMapAgenc, 'Estorno Já Deduzido? (S/N)')
                };
            } else if (id && dataAgenciamentos[i].some(cell => cell !== '')) {
                Logger.log("Aviso: ID Agenciamento " + id + " (Linha " + (i + 1) + " Agenciamentos) não encontrado na aba ClientesStatus.");
            }
        }

        Logger.log("Processando dados da aba " + NOME_ABA_COMISSOES_MENSAL);
        for (let i = 1; i < dataComissoes.length; i++) {
            const id = getVal(dataComissoes[i], headerMapComiss, 'ID Agenciamento');

            if (id && clienteInfo[id]) {
                clienteInfo[id].comissao = {
                    linhaComiss: i + 1,
                    inicioApos: parseInt(getVal(dataComissoes[i], headerMapComiss, 'Início Comissão (Após Meses)')) || 0,
                    valorParcelaBase: parseFloat(getVal(dataComissoes[i], headerMapComiss, 'Valor Comissão Parcela (R$)')) || 0,
                    comissoesPagasQtd: parseInt(getVal(dataComissoes[i], headerMapComiss, 'Comissões Pagas Corretor (Qtd)')) || 0
                };
            } else if (id && dataComissoes[i].some(cell => cell !== '')) {
                Logger.log("Aviso: ID Agenciamento " + id + " (Linha " + (i + 1) + " Comissoes) não encontrado na aba ClientesStatus.");
            }
        }
        Logger.log("Dados das abas lidos e consolidados em memória.");


        const pagamentosPorCorretor = {};
        // --- Loop Principal pelos Dados ---
        for (let id in clienteInfo) {
            let info = clienteInfo[id];

            // Verifica se existe os dados de agenciamento e comissão e se o corretor está definido
            if (!info.agenciamento) {
                Logger.log("Aviso: Dados de Agenciamento ausentes para ID " + id + ". Pulando cálculo para este ID.");
                continue;
            }
            if (!info.comissao) {
                Logger.log("Aviso: Dados de Comissão ausentes para ID " + id + ". Pulando cálculo para este ID.");
                continue;
            }
            if (!info.corretor) {
                Logger.log("Aviso: Corretor não definido para ID " + id + ". Pulando cálculo para este ID.");
                continue;
            }

            let pagarAgencP1Final = 0;
            let pagarAgencP2Final = 0;
            let pagarComissaoFinal = 0;
            let comissoesCalculadasNesteCiclo = 0;
            let descontarEstornoFinal = 0;
            let detalhesLinha = [];

            // 1. Calcular Agenciamento P1 a pagar (COM TAXA)
            if (String(info.agenciamento.p1Pago).toUpperCase() !== 'S') {
                pagarAgencP1Final = info.agenciamento.valorP1Base * (1 - TAXA_ADMINISTRATIVA);
                detalhesLinha.push("Ag.P1:" + formatarMoedaMensal(info.agenciamento.valorP1Base) + "+Tx=" + formatarMoedaMensal(pagarAgencP1Final));
            }

            // 2. Calcular Agenciamento P2 a pagar (COM TAXA)
            if (info.agenciamento.valorP2Base > 0 &&
                String(info.agenciamento.p2Pago).toUpperCase() !== 'S' &&
                info.mesesPagos >= info.agenciamento.pagarP2AposMeses) {
                pagarAgencP2Final = info.agenciamento.valorP2Base * (1 - TAXA_ADMINISTRATIVA);
                detalhesLinha.push("Ag.P2:" + formatarMoedaMensal(info.agenciamento.valorP2Base) + "+Tx=" + formatarMoedaMensal(pagarAgencP2Final));
            }

            // 3. Calcular Comissão a pagar (COM TAXA)
            if (info.status === "Ativo" &&
                String(info.inadimplenteMes).toUpperCase() !== 'S' &&
                info.mesesPagos > info.comissao.inicioApos &&
                info.mesesPagos > info.comissao.comissoesPagasQtd) {

                comissoesCalculadasNesteCiclo = info.mesesPagos - info.comissao.comissoesPagasQtd;
                if (comissoesCalculadasNesteCiclo > 0) {
                    let comissaoBaseTotal = comissoesCalculadasNesteCiclo * info.comissao.valorParcelaBase;
                    pagarComissaoFinal = comissaoBaseTotal * (1 - TAXA_ADMINISTRATIVA);

                    let detalheComissao = (comissoesCalculadasNesteCiclo === 1) ?
                        "Comissão(" + formatarMoedaMensal(info.comissao.valorParcelaBase) + "+Tx=" + formatarMoedaMensal(pagarComissaoFinal) + ")" :
                        comissoesCalculadasNesteCiclo + "xComissões(" + formatarMoedaMensal(comissaoBaseTotal) + "+Tx=" + formatarMoedaMensal(pagarComissaoFinal) + ")";
                    detalhesLinha.push(detalheComissao);
                }
            }

            // 4. Calcular Estorno a descontar (SEM TAXA - sobre valor base)
            if (info.status === "Cancelado/Inadimplente" && info.mesesPagos < info.estornoAgenc) {
                if (String(info.agenciamento.estornoAplicado).toUpperCase() === 'S') {
                    if (String(info.agenciamento.estornoDeduzido).toUpperCase() !== 'S') {
                        descontarEstornoFinal = info.agenciamento.valorEstornoPotencialBase;
                        detalhesLinha.push("Estorno: " + formatarMoedaMensal(-descontarEstornoFinal));
                    }
                } else {
                    detalhesLinha.push("Estorno Potencial (Verificar Aba Agenciamentos)");
                }
            }

            // --- Calcular total líquido e agrupar ---
            let totalLiquidoLinha = pagarAgencP1Final + pagarAgencP2Final + pagarComissaoFinal - descontarEstornoFinal;
            let detalhesString = detalhesLinha.join(' | ');

            // Adiciona ao corretor se houver valor a pagar/descontar ou aviso de estorno
            if (totalLiquidoLinha !== 0 || detalhesString.includes("Estorno Potencial")) {
                let corretor = info.corretor;
                if (!pagamentosPorCorretor[corretor]) {
                    pagamentosPorCorretor[corretor] = { total: 0, itens: [] };
                }

                let valorRealItem = totalLiquidoLinha;
                if (detalhesString === "Estorno Potencial (Verificar Aba Agenciamentos)") {
                    valorRealItem = 0;
                }

                pagamentosPorCorretor[corretor].total += valorRealItem;
                pagamentosPorCorretor[corretor].itens.push({
                    cliente: info.clienteNome,
                    valorItem: valorRealItem,
                    detalhes: detalhesString,
                    pagouAgP1: pagarAgencP1Final > 0,
                    pagouAgP2: pagarAgencP2Final > 0,
                    pagouNComissoes: comissoesCalculadasNesteCiclo,
                    descontouEstorno: descontarEstornoFinal > 0,
                    idAgenciamento: id
                });
            }
        }

        // --- Envio de Email ---
        let totalGeralPago = 0;
        let temPagamentos = false;
        let corretorToTemplate = [];
        let corretoresOrdenados = Object.keys(pagamentosPorCorretor).sort();

        corretoresOrdenados.forEach(corretor => {
            const dadosCorretor = pagamentosPorCorretor[corretor];
            const dadosFormatados = dadosCorretor.itens.map(item => {
                return {
                    cliente: item.cliente,
                    valor: item.valorItem,
                    detalhes: item.detalhes,
                    isWarning: item.detalhes.includes("Estorno Potencial"),
                };
            });
            corretorToTemplate.push({
                corretor: corretor,
                totalLiquido: dadosCorretor.total,
                itens: dadosFormatados
            });
            totalGeralPago += dadosCorretor.total;
        });

        const templateEmailData = {
            corretores: corretorToTemplate,
            totalGeralPago: totalGeralPago,
            temPagamentos: corretorToTemplate.length > 0
        };

        const emailTemplate = HtmlService.createTemplateFromFile('templateEmail.html');
        emailTemplate.data = templateEmailData;
        emailBody = emailTemplate.evaluate().getContent();

        // Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss")";

        MailApp.sendEmail({
            to: EMAIL_RESPONSAVEL_MENSAL,
            subject: assuntoEmail,
            htmlBody: emailBody
        });

        Logger.log(functionName + " concluído. Email enviado para " + EMAIL_RESPONSAVEL_MENSAL);

    } catch (error) {
        Logger.log("ERRO GRAVE em " + functionName + ": " + error + "\nStack: " + error.stack);
        MailApp.sendEmail(EMAIL_RESPONSAVEL_MENSAL, "ERRO GRAVE no Relatório Mensal de Pagamentos",
            "Script: " + functionName + "\nErro: " + error + "\nO relatório mensal não pôde ser gerado. Verifique os Logs na planilha mestre.");
    }
}


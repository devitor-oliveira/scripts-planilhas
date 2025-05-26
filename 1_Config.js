/**
 * Arquivo de Configuração Central.
 * Altere os valores aqui para configurar o comportamento de todo o script.
 */
const CONFIG = {
  // ID da Planilha Ativa (não precisa mudar)
  SPREADSHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),

  // Nomes exatos das abas da planilha
  SHEET_NAMES: {
    CLIENTS: "Cliente",
    AGENCY_FEES: "Agenciamento",
    COMMISSIONS: "Comissao",
  },

  // Configurações de E-mail
  EMAIL_SETTINGS: {
    ADMIN: "@gmail.com", // <-- MUDE AQUI
    MONTHLY_REPORT_RECIPIENT: "@gmail.com", // <-- MUDE AQUI
    MONTHLY_REPORT_SUBJECT: "Relatório Mensal de Pagamentos a Corretores",
  },

  // Regras de Negócio
  BUSINESS_RULES: {
    ADMIN_FEE_PERCENTAGE: 0.20, // Taxa administrativa de 20%
  },

  // Mapeamento das colunas do formulário para chaves de dados (a ordem importa!)
  FORM_MAP: [
    'dtCadastro',         // Coluna 1 (Timestamp)
    'clienteNome',        // Coluna 2
    'clienteCPF',         // Coluna 3
    'corretor',           // Coluna 4
    'contrato',           // Coluna 5
    'premioMensal',       // Coluna 6
    'formaPagto',         // Coluna 7
    'taxaAgencP1Perc',    // Coluna 8
    'taxaAgencP2Perc',    // Coluna 9
    'pagarP2AposMeses',   // Coluna 10
    'taxaComissaoPerc',   // Coluna 11
    'inicioComissaoApos', // Coluna 12
    'estornoAgencMeses',  // Coluna 13
    'taxaEstornoPerc'     // Coluna 14
  ],

  // Nomes exatos das colunas em cada aba (para leitura e escrita)
  COLUMN_NAMES: {
    // Chave Universal
    ID: "ID Agenciamento",
    CORRETOR: "Corretor Responsável",
    CLIENTE_NOME: "Cliente Nome",

    // Aba ClientesStatus
    CLIENTS: {
      CPF: "Cliente CPF",
      CONTRATO: "Contrato/Seguro",
      PREMIO: "Valor Prêmio Mensal (R$)",
      ESTORNO_MESES: "Estorno Agenc. (Meses)",
      STATUS: "Status Cliente",
      MESES_PAGOS: "Cliente Meses Pagos (Qtd)",
      INADIMPLENTE: "Inadimplente Mês Atual? (S/N)",
      DATA_STATUS: "Data Status",
      DT_CADASTRO: "Data Cadastro"
    },
    // Aba Agenciamentos
    AGENCY: {
      TAXA_P1: "Taxa Agenc Parte 1 (%)",
      TAXA_P2: "Taxa Agenc Parte 2 (%)",
      PAGAR_P2_APOS: "Pagar P2 Após Meses",
      VALOR_P1: "Valor Agenc Parte 1 (R$)",
      VALOR_P2: "Valor Agenc Parte 2 (R$)",
      TAXA_ESTORNO: "Taxa Estorno Agenc (%)",
      VALOR_ESTORNO_POTENCIAL: "Valor Estorno Potencial (R$)",
      P1_PAGO: "Agenc Parte 1 Pago? (S/N)",
      P2_PAGO: "Agenc Parte 2 Pago? (S/N)",
      ESTORNO_APLICADO: "Estorno Aplicado? (S/N)",
      ESTORNO_DEDUZIDO: "Estorno Já Deduzido? (S/N)"
    },
    // Aba Comissões
    COMMISSION: {
      TAXA: "Taxa Comissão (%)",
      INICIO_APOS: "Início Comissão (Após Meses)",
      VALOR_PARCELA: "Valor Comissão Parcela (R$)",
      QTD_PAGA: "Comissões Pagas Corretor (Qtd)"
    }
  }
};
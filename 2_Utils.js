/**
 * Funções utilitárias e auxiliares genéricas.
 */
const Utils = {
  /**
   * Converte uma string de número em formato brasileiro (ex: "1.500,75") para um float.
   * Esta versão é mais robusta para diferenciar separadores de milhar e decimais.
   * @param {string | number} numberString A string ou número a ser convertido.
   * @returns {number} O número convertido.
   */
  parseBrazilianNumber(numberString) {
    if (typeof numberString === 'number') {
      return numberString;
    }
    if (typeof numberString !== 'string' || !numberString) {
      return 0;
    }

    let cleaned = numberString.trim();

    const hasComma = cleaned.includes(',');
    const hasDot = cleaned.includes('.');

    // Caso 1: Formato com vírgula decimal (ex: "1.500,75" ou "199,50")
    // A vírgula é o indicador mais forte de formato brasileiro.
    if (hasComma) {
      // Remove os pontos de milhar e troca a vírgula decimal por ponto.
      cleaned = cleaned.replace(/\./g, '').replace(',', '.');
    }
    // Caso 2: Formato apenas com pontos (ambíguo: "5.000" vs "364.62")
    else if (hasDot) {
      const lastDotIndex = cleaned.lastIndexOf('.');
      const digitsAfterLastDot = cleaned.substring(lastDotIndex + 1).length;

      // Se a parte após o último ponto tem 3 dígitos, consideramos todos os pontos
      // como separadores de milhar e os removemos.
      if (digitsAfterLastDot === 3) {
        cleaned = cleaned.replace(/\./g, '');
      }
      // Se a parte após o último ponto tem 1 ou 2 dígitos (ex: "364.62"),
      // consideramos o número já no formato correto para parseFloat e não fazemos nada.
    }

    return parseFloat(cleaned) || 0;
  },

  /**
   * Converte um valor numérico para o formato de moeda BRL.
   * @param {number} value O valor a ser formatado.
   * @returns {string} O valor formatado como moeda.
   */
  toCurrency(value) {
    if (typeof value !== 'number' || isNaN(value)) {
      value = 0;
    }
    return value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  },

  /**
   * Mapeia os valores de um array (vindo do formulário) para um objeto com chaves nomeadas.
   * @param {Array<any>} formValues O array `e.values` do evento do formulário.
   * @returns {Object} Um objeto com os dados mapeados.
   */
  mapFormToObject(formValues) {
    const data = {};
    CONFIG.FORM_MAP.forEach((key, index) => {
      data[key] = (formValues[index] !== undefined) ? formValues[index] : null;
    });
    if (!data.dtCadastro) data.dtCadastro = new Date();
    return data;
  },

  /**
   * Envia um e-mail de notificação de erro para o administrador.
   * @param {string} subject O assunto do e-mail.
   * @param {string} body O corpo do e-mail (pode ser HTML).
   */
  sendErrorEmail(subject, body) {
    MailApp.sendEmail({
      to: CONFIG.EMAIL_SETTINGS.ADMIN,
      subject: `ALERTA SCRIPT PLANILHAS: ${subject}`,
      htmlBody: body
    });
  }
};
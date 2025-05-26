/**
 * Serviço para geração e envio de relatórios.
 */
const ReportService = {
  /**
   * Prepara os dados calculados e envia o e-mail do relatório mensal para os corretores.
   * @param {Object} paymentData Os dados de pagamento calculados, agrupados por corretor.
   */
  sendBrokerReport(paymentData) {
    const hoje = new Date();
    // A linha abaixo foi corrigida de 'Federer' para 'yyyy'
    const mesAno = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "MMMM 'de' yyyy");
    const assunto = `${CONFIG.EMAIL_SETTINGS.MONTHLY_REPORT_SUBJECT} - ${mesAno}`;

    let totalGeralBruto = 0;
    let totalGeralLiquido = 0;
    const corretoresOrdenados = Object.keys(paymentData).sort();
    
    const corretoresParaTemplate = corretoresOrdenados.map(corretor => {
      totalGeralBruto += paymentData[corretor].totalBruto;
      totalGeralLiquido += paymentData[corretor].totalLiquido;
      return {
        corretor: corretor,
        totalBrutoCorretor: paymentData[corretor].totalBruto,
        totalLiquidoCorretor: paymentData[corretor].totalLiquido,
        itens: paymentData[corretor].items,
      };
    });

    const despesaAdminTotal = totalGeralBruto - totalGeralLiquido;

    const templateData = {
      corretores: corretoresParaTemplate,
      totalGeralBruto: totalGeralBruto,
      totalGeralLiquido: totalGeralLiquido,
      despesaAdminTotal: despesaAdminTotal,
      temPagamentos: corretoresParaTemplate.length > 0,
      TAXA_ADMINISTRATIVA: CONFIG.BUSINESS_RULES.ADMIN_FEE_PERCENTAGE,
      toCurrencyFormat: Utils.toCurrency
    };

    const template = HtmlService.createTemplateFromFile('email.html');
    template.data = templateData;
    const emailBody = template.evaluate().getContent();
    const nomeArquivo = `Relatorio_Pagamentos_${mesAno}.pdf`;

    const pdfAnexo = Utilities.newBlob(emailBody, MimeType.HTML).getAs(MimeType.PDF).setName(nomeArquivo);

    MailApp.sendEmail({
      to: CONFIG.EMAIL_SETTINGS.MONTHLY_REPORT_RECIPIENT,
      subject: assunto,
      htmlBody: emailBody,
      attachments: [pdfAnexo]
    });
    
    Logger.log(`Relatório enviado para ${CONFIG.EMAIL_SETTINGS.MONTHLY_REPORT_RECIPIENT} com PDF anexado.`);
  }
};
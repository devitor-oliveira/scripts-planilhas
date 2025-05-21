# Scripts Planilhas Google
 
Lembre-se de atualizar as variáveis de configuração no código!

## Descrição dos Scripts e Funções

### Script email.js
- Mapear dados e realizar os cálculos finais (definir formatos)
- Preparar os dados do template HTML e enviá-lo por e-mail
- Executado Mensalmente

### Script Principal
- Acionado quando há resposta no formulário
- Executar as funções auxiliares
- Mapeamento das perguntas para transformar em objeto
- Mapeamento dos dados da planilha

### Script writeSheets
- Funções auxiliares para transcrever os dados da planilha principal para as abas
- Informação sobre as colunas que serão preenchidas manualmente
- Funções executadas pelo script principal
- Sem Acionadores

## Como configurar os acionadores (Gatilhos/Triggers)

### Acionador do Evento de envio do Formulário (onFormSubmit)

 1. Na planilha, vá em Extensões > Apps Script.
 2. No editor, vá em Acionadores (relógio) > + Adicionar Acionador.
 3. Configure:
    - Função a ser executada: `onFormSubmit`
    - Escolha qual implantação deve ser executada: `Principal`
    - Selecione a origem do evento: `Da planilha`
    - Selecione o tipo de evento: `Ao enviar formulário`
    - Configurações de notificação de falha: `Notificar-me imediatamente`
 4. Clique em "Salvar" e autorize as permissões necessárias.


### Acionador Mensal do Evento para enviar o e-mail com relatório

1. Na planilha, vá em Extensões > Apps Script.
 2. No editor, vá em Acionadores (relógio) > + Adicionar Acionador.
 3. Configure:
    - Função a ser executada: `sendEmailMonthly`
    - Escolha qual implantação deve ser executada: `Principal`
    - Selecione o tipo de evento: `Baseado no tempo`
    - Selecione o tipo de acionador baseado em tempo: `Contador de Meses`
    - Selecione o dia desejado: `Exemplo: Dia 1°`
    - Selecione a hora do dia: `Exemplo: 8:00 às 9:00`
    - Configurações de notificação de falha: `Notificar-me imediatamente`
 4. Clique em "Salvar" e autorize as permissões necessárias.

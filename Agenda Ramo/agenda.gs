// Cria uma "trigger" que analisa a planilha a cada 1 minuto.
function comeceaqui() {
    try {
        // Deleta qualquer "trigger" existente.
        var triggers = ScriptApp.getProjectTriggers();
        for (var i in triggers)
            ScriptApp.deleteTrigger(triggers[i]);
        // Cria "trigger" temporal.
        ScriptApp.newTrigger('evento')
            .timeBased()
            .everyMinutes(1)
            .create();
    } catch (error) {
        // Aviso de erro.
        throw new Error("Esse código deve ser adicionado na planilha de respostas de um formulário!");
    }
}
// Para criar um evento no calendário.
function evento() {
    // Carrega planilha ativa.
    var ss = SpreadsheetApp.getActive();
    // Pega o valor valores da pasta ativa.
    var dados = ss.getDataRange().getValues();
    // Seleciona a última linha escrita da tabela.
    var ultimaLinha = ss.getLastRow() - 1; 
    // Opções do evento: Descrição e Localização.
    var options = {
        description: dados[ultimaLinha][2],
        location: dados[ultimaLinha][5],
    };

    for (var i = 1; i < ultimaLinha + 1; i++) {
        var validacao = dados[i][6];
        var escrita = dados[i][7];
        if (validacao === "Yes" && escrita == "") {
            // SpreadsheetApp.getActiveSheet() - Seleciona a pasta ativa.
            // getRange(linha, coluna) seleciona a célula indicada. 
            // setValue("valor") atribui o valor a célula indicada.
            SpreadsheetApp.getActiveSheet().getRange(i + 1, 8).setValue('Ok');
            // "event" seleciona o Calendário Padrão e Cria um evento (título, data de início, data de final, opções{descrição e localização})
            var event = CalendarApp.getDefaultCalendar().createEvent(
                dados[i][1],
                new Date(dados[i][3]),
                new Date(dados[i][4]),
                options);
        }
    }
}
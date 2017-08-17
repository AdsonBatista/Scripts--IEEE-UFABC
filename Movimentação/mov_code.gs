// ---------------------------------------------------------------------------
// Função para arquivos "javascript" e "Stylesheet" - Boas práticas de programação.
// https://developers.google.com/apps-script/guides/html/best-practices#separate_html_css_and_javascript 
// ---------------------------------------------------------------------------
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
}
// ---------------------------------------------------------------------------
// Funções para pegar as X últimas letras de uma "String".
// var texto = exemplo
// texto.right(3);
// saida = plo
// ---------------------------------------------------------------------------
String.prototype.right = function() {
        return this.substr(this.length - (arguments[0] == undefined ? 1 : parseInt(arguments[0])), this.length);
}
// ---------------------------------------------------------------------------
// Funções para pegar as X primeiras letras de uma "String".
// var texto = exemplo
// texto.left(3);
// saida = exe
// ---------------------------------------------------------------------------
String.prototype.left = function() {
    return this.substr(0, arguments[0] == undefined ? 1 : parseInt(arguments[0]));
}
// ---------------------------------------------------------------------------
// Funções para para escrever um número sempre utilizando um número X de casas.
// pad(6,4) 
// saida 0006
// ---------------------------------------------------------------------------
function pad(number, length) {
    var str = '' + number;
    while (str.length < length) {
        str = '0' + str;
    }
    return str;
}
// ---------------------------------------------------------------------------
// ID`s - http://alicekeeler.com/2013/08/03/google-docs-unique-id/
// ---------------------------------------------------------------------------
// ID da Planilha de submissão de dados.
var submissionSSKey = '1qXFlxbzs2IKLAIl1awYByP8QQGm5F_CGWg97XVIryCQ';
// ---------------------------------------------------------------------------
// ID da Pasta que será salvo os arquivos.
var folderId = "0B8CcpExpMKFlSFZQMmtaYWpNbWs";
// ---------------------------------------------------------------------------
//
// ---------------------------------------------------------------------------
// doGet - inicia o processo do "WebApp".
// doPost - envia dados para serem processados quando formulário é enviado.
// createTemplateFromFile - cria o template do formulário com base no modelo mov_form.html.
// evaluate(); retorna valores do html.
// ---------------------------------------------------------------------------
function doGet(e) {
    var template = HtmlService.createTemplateFromFile('mov_form.html');
    template.action = ScriptApp.getService().getUrl();
    return template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
//
// ---------------------------------------------------------------------------
// Ações com os dados do formulário.
// ---------------------------------------------------------------------------
//
function processForm(theForm) {
    // Define a seguda pasta da planilha ativa como pasta de trabalho (planilha adminstrativa).
    // Retira os números dos recibos dela.
    var sheet = SpreadsheetApp.openById(submissionSSKey).getSheets()[2];
    //
    // ---------------------------------------------------------------------------
    // Extrair variáveis do formulário.
    // ---------------------------------------------------------------------------
    //
    // Busca nome do capítulo nas respostas do formulário.
    var unidade = theForm.unidade;
    // Busca data da movimentação, ela é extraída no formato de string DD/MM/YYYY.
    var data = theForm.data;
    // Busca valor da movimentação.
    var valor = theForm.valor;
    // Fixa o valor em duas casas decimais e troca o ponto pela vírgula.
    valor = parseFloat(valor).toFixed(2).toString().replace(/\./g, ',');
    // Descrição do gasto.
    var descricaogasto = theForm.descricaogasto;
    // E-mail.
    var email = theForm.email;
    // Transação de entrada ou saída.
    var tipotransacao = theForm.tipotransacao;
    // ---------------------------------------------------------------------------
    // Tratamentos para montagem do nome do arquivo .
    // Código da unidade + Mês + Tipo de Movimentação + Número da Movimentação + Ano.
    // ---------------------------------------------------------------------------  
    // Define a variável range como valor da busca linha e coluna na sheets ativa.
    var range;
    // Define um ID para cada unidade.
    var idunidade;
    // Busca na planilha uma célula específica (neste caso o número do recibo).
    // range = sheet.getRange(linha,coluna);
    if (unidade == "Ramo") {
        range = sheet.getRange(2, 2);
        idunidade = "00";
    } else if (unidade == "AESS") {
        range = sheet.getRange(3, 2);
        idunidade = "01";
    } else if (unidade == "CS") {
        range = sheet.getRange(4, 2);
        idunidade = "02";
    } else if (unidade == "CPMT") {
        range = sheet.getRange(5, 2);
        idunidade = "03";
    } else if (unidade == "EMBS") {
        range = sheet.getRange(6, 2);
        idunidade = "04";
    } else if (unidade == "PES") {
        range = sheet.getRange(7, 2);
        idunidade = "05";
    } else if (unidade == "RAS") {
        range = sheet.getRange(8, 2);
        idunidade = "06";
    } else if (unidade == "TEMS") {
        range = sheet.getRange(9, 2);
        idunidade = "07";
    }
    // Defino o número do recibo com 4 algarismos.
    var numrecibo = pad(range.getValue(), 4);
    // Chama a função "right" para pegar os 2 últimos digitos da data preenchido no formulário.
    var ano = data.right(2);
    // Para pegar o mês eu pego os 7 caracteres da direita e depois os 2 da esquerda e define obrigatoriamente como duas casas decimais.
    var mes = pad(data.right(7).left(2), 2);
    // Variável de data- DD/MM/YYYY
    // right(2) - YY
    // right(7) - MM//YYYY
    // left(2) - MM
    var idtrasacao;
    if (tipotransacao == "Saída") {
        idtrasacao = "S";
    } else {
        (tipotransacao == "Entrada")
        idtrasacao = "E";
    }
    //
    // ---------------------------------------------------------------------------
    // Arquivo Anexo.
    // ---------------------------------------------------------------------------  
    //
    // Pego o arquivo do formulário.
    var arquivo = theForm.myFile;
    // Extrai o nome do arquivo.
    var nomearquivo = arquivo.getName();
    // Regex para buscar a extensão do arquivo e depois retiro a extensão do arquivo.
    // https://stackoverflow.com/a/680982/1677912
    var extensionfinder = /(?:\.([^.]+))?$/;
    var ext = extensionfinder(nomearquivo)[1];
    // Define um novo nome para o arquivo.
    var novoName = idunidade + mes + idtrasacao + numrecibo + ano;
    // Cria a variável com novo nome e extensão do arquivo.
    var novonomearquivo = novoName + '.' + ext;
    // Nomeia o arquivo mantendo a extensão original.
    arquivo.setName(novonomearquivo);
    // Define pasta que vou salvar.
    var folder = DriveApp.getFolderById(folderId);
    // Cria o arquivo na pasta definida.
    var doc = folder.createFile(arquivo);
    // Coloca a "URL" com direito de visualização. Pega "URL" do arquivo.
    var urlarquivo = doc.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT).getUrl();
    //
    // ---------------------------------------------------------------------------
    // Página de agradecimento.
    // ---------------------------------------------------------------------------
    //
    // Cria o "template" .
    var template = HtmlService.createTemplateFromFile('mov_Thanks.html');
    // Escreve o Valor no "template".
    template.valor = valor;
    // Escreve a unidade no "template".
    template.unidade = unidade;
    // Escreve a data no "template".
    template.data = data;
    // Escreve o tipo de transação no "template".
    template.tipotransacao = tipotransacao;
    // Escreve a descrição do gasto no "template".
    template.descricaogasto = descricaogasto;
    // Escreve o e-mail no "template".
    template.email = email;
    // Escreve o Nome do arquivo no "template".
    template.novonomearquivo = novonomearquivo;
    // Escreve o "URL" do arquivo no "template".
    template.urlarquivo = urlarquivo;
    //
    // ---------------------------------------------------------------------------
    //  Gravar dados na Planilha de dados
    // ---------------------------------------------------------------------------
    //
    // SpreadsheetApp.openById(submissionSSKey).getSheets()[0] Abre a planilha ativa utilizando um ID e abre a primeira pasta de trabalho [0].
    var sheet = SpreadsheetApp.openById(submissionSSKey).getSheets()[1];
    // Cria o elemento da primeira coluna da planilha "timestamp".
    var row = [new Date()];
    // Pega a última linha escrita.
    var lastRow = sheet.getLastRow();
    // Gravar dados do formulário na planilha.
    // Seleciona e Preenche a primeira linha em branco "[lastRow+1]" com os valores na ordem dada pelo vetor em "setValues".
    // getRange(linha, coluna, numlinhas, numColunas) reserva um vetor para preenchimento dos dados.
    var targetRange = sheet.getRange(lastRow + 1, 1, 1, 9).setValues([
        [row, valor, unidade, data, tipotransacao, descricaogasto, email, novonomearquivo, urlarquivo]
    ]);
    //
    // ---------------------------------------------------------------------------
    //  E-MAIL
    // ---------------------------------------------------------------------------
    //
    // Assunto do e-mail.
    var subject = "Aviso de movimentação";
    // Dados do Rementente.
    var remetente = "IEEE UFABC<contato@ieeeufabc.org>";
    //
    // ---------------------------------------------------------------------------
    //  Corpo do E-MAIL
    // ---------------------------------------------------------------------------
    //
    var html = HtmlService.createTemplateFromFile('mail_template');
    // Escreve as variáveis no template do e-mail.
    html.valor = valor;
    html.unidade = unidade;
    html.data = data;
    html.tipotransacao = tipotransacao
    html.descricaogasto = descricaogasto;
    html.email = email
    html.novonomearquivo = novonomearquivo;
    html.urlarquivo = urlarquivo;
    // Fecha o "template" e pega o modelo e salva na variável htmlBody.
    var htmlBody = html.evaluate().getContent();
  
    // ---------------------------------------------------------------------------
    // Enviar E-MAIL
    // ---------------------------------------------------------------------------
    // E-mail de teste.
    MailApp.sendEmail("adson.batista@live.com", subject, html, {
        name: remetente,
        htmlBody: htmlBody
    });
    // Envia o e-mail com recibo para o solicitante.
    MailApp.sendEmail(email, subject, html, {
        name: remetente,
        htmlBody: htmlBody
    });

    //Retorna o texto HTML na página.
    return template.evaluate().getContent();
}
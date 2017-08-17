var dropbox = "Atividades";

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
}


// Faz o formulário a partir de um modelo HTML
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('forms.html');
  template.action = ScriptApp.getService().getUrl();
  return template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  //return HtmlService.createHtmlOutputFromFile('forms.html').setSandboxMode(HtmlService.SandboxMode.IFRAME).setTitle("Sandbox Upload Form");  
}

  
// cria uma planilha na pasta que o script está
function initSpreadSheet(){
var folder = DriveApp.getFoldersByName(dropbox).next();;//gets first folder with the given foldername
var new_ss = SpreadsheetApp.create('Report Atividade')
var copyFile=DriveApp.getFileById(new_ss.getId());
folder.addFile(copyFile);
PropertiesService.getScriptProperties().setProperty('sheet_id', copyFile.getId());
var ss = SpreadsheetApp.openById(copyFile.getId()); // ID of the spread_sheet
var sheet = ss.getActiveSheet();
sheet.appendRow(['Timestamp','Nome do Evento','Unidade','Tipo de evento','Data','Nome do responsável','Parceiros','Email do responsável','Descrição do evento','Observações','Link Folder','Time']);
DriveApp.getRootFolder().removeFile(copyFile);
}

function uploadFilesToDrive(base64Data, fileName, user_dir_name, init_user) {

  try
  {
    // Pegar diretório de root
    // Nome da pasta Root, Esta pasta já deve existir!
 
    // Busca a pasta de root e procura a proxima
    var folder = DriveApp.getFoldersByName(dropbox).next();
    // Busca pasta criada no formulário para o projeto
    var user_dirs = folder.getFoldersByName(user_dir_name);
    // if procura a pasta a user_dirs e init_user (não entendi ela...) se encontrar não encontrar ele cria ela se encontrar ele entra nela
    if (!user_dirs.hasNext() && init_user) {
       var user_dir = folder.createFolder(user_dir_name);
    } else {
       var user_dir = folder.getFoldersByName(user_dir_name).next();
    }

    // Leitura e processamento dos arquivos de upload
    var splitBase = base64Data.split(','),
        type = splitBase[0].split(';')[0].replace('data:','');

    var byteCharacters = Utilities.base64Decode(splitBase[1]);
    var ss = Utilities.newBlob(byteCharacters, type);
    ss.setName(fileName);

    // Salva o arquivo na pasta definida
    var file = user_dir.createFile(ss);
    return "OK";
  }
  // se erro entra na condicão init_user e retorna um texto caso contrario chama a função uploadFilesToDrive
  catch(e){
    if(init_user)
     return e.toString();
    else
    {
    uploadFilesToDrive(base64Data, fileName, user_dir_name, init_user);
    }
  }
}
function get_shared_folders(user_dir_name){
  try{
    var dropbox = "Atividades"; // Root folder Name, musts be already initialized!
    var folder = DriveApp.getFoldersByName(dropbox).next();
    var user_dir = folder.getFoldersByName(user_dir_name).next().setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
    return user_dir.getUrl();
  }catch(e){
     return get_shared_folders(user_dir_name);

  }
}

function logSheet(name, unidade,tipoatividade, data,resp_name,parceiros,email,descricao,observacao,imagem,user_time){
 var row = new Date();
 var unidades = unidade.toString();
 var sheet_id = PropertiesService.getScriptProperties().getProperty('sheet_id');
 var ss = SpreadsheetApp.openById(sheet_id); // ID of the spread_sheet
 var sheet = ss.getActiveSheet();
  if (imagem =="Sim"){
    var url = get_shared_folders(unidade + ' - ' +data + ' - ' + name);
  }else{var url = "Não existiu Upload Imagem"}
   if (parceiros == "")
    {
parceiros = "-";  
    }
     if (observacao == "")
    {
 observacao = "-";     
    }
 sheet.appendRow([row,name, unidades,tipoatividade, data,resp_name,parceiros,email,descricao,observacao, url, user_time,imagem]);

  var template = HtmlService.createTemplateFromFile('obrigado.html');
    // Escreve o Valor no template
    template.name = name;
    template.valor = url;
    // Escreve a unidade no template
    template.unidade = unidades;
    // Escreve a data no template
    template.data = data;
    // Escreve o tipo de transação no template
    template.tipotransacao = tipoatividade;
    // Escreve a descrição do gasto no template
    template.descricaogasto = resp_name;
    // Escreve o email no template
    template.email = email;
    // Escreve o Nome do arquivo no template
    template.novonomearquivo = descricao;
    // Escreve o URL do arquivo no template
    template.urlarquivo = url;
  
     return template.evaluate().getContent();
}

function Extenso(n,moeda,moedas,centavo,centavos){
  var j,x,m,r,ri,rd,d,i,casas,erro;
  var v1=0,v2=0,v3=0,v4=0,v5=0,v6=0;
  r="";
  rd="";
  ri="";
  i=parseInt(n);
  d=n-i;
  d=d.toFixed(2);
  d=d*100;
  d=d.toFixed(0);
  casas=i.toString().length;

  if(n=="?"){return "Função Extenso() Marcelo Camargo - marcelocamargo@gmail.com";}
  if(n<0){return "Erro: número negativo";}
  if(moeda!=null){if(moedas==null || centavo==null || centavos==null || moeda=="" || moedas=="" || centavo=="" || centavos==""){return "Erro: parâmetros de moeda";}}

  if(d==100){
    d=0;
    i=i+1;
  }

  if(casas>12){
    v5=(parseInt(i/1000000000000)*1000000000000-parseInt(i/1000000000000000)*1000000000000000)/1000000000000;
    if(v5>0){
      j="";
      x=CentenaExtenso(v5);
      if(v5>1){ri=ri+j+x+" trilhões";}else{ri=ri+j+x+" trilhão";}
    }
  }
  if(casas>9){
    v4=(parseInt(i/1000000000)*1000000000-parseInt(i/1000000000000)*1000000000000)/1000000000;
    if(v4>0){
      if(v5){j=", ";}else{j="";}
      x=CentenaExtenso(v4);
      if(v4>1){ri=ri+j+x+" bilhões";}else{ri=ri+j+x+" bilhão";}
    }
  }
  if(casas>6){
    v3=(parseInt(i/1000000)*1000000-parseInt(i/1000000000)*1000000000)/1000000;
    if(v3>0){
      if(v4+v5){j=", ";}else{j="";}
      x=CentenaExtenso(v3);
      if(v3>1){ri=ri+j+x+" milhões";}else{ri=ri+j+x+" milhão";}
    }
  }
  if(casas>3){
    v2=(parseInt(i/1000)*1000-parseInt(i/1000000)*1000000)/1000;
    if(v2>0){
      if(v3+v4+v5){j=", ";}else{j="";}
      x=CentenaExtenso(v2);
      if(v2==1){
        ri=ri+j+"mil";
      } else {
        ri=ri+j+x+" mil";
      }
    }
  }
  if(casas>0){
    v1=(parseInt(i).toFixed(0))-(parseInt(i/1000).toFixed(0)*1000);
    if(v1>0){
      if(v2+v3+v4+v5){if(v1<=100){j=" e ";}else{j=", ";}}else{j="";}
      x=CentenaExtenso(v1);
      ri=ri+j+x;
    }
  }

  if(moeda==null){
    moedas="reais";
    moeda="real";
    centavos="centavos";
    centavo="centavo";
  }
  if((d!=0 && moeda=="inteiro") || moeda!="inteiro"){
    if(i>0 && !v1){ri=ri+" de "+moedas;}
    else if(i>1 && v1==1){ri=ri+" "+moedas;}
    else if(v1==1){ri=ri+" "+moeda;}
    else if(v1>1){ri=ri+" "+moedas;}
    else if(i==1){ri=ri+" "+moeda;}
  }
  
  if(d==1){
    rd="um "+centavo;
  } else if(d>1 && d<100){
    rd=CentenaExtenso(d)+" "+centavos;
  }
  if(i<1 && d>0 && moeda!="inteiro"){
    rd=rd+" de "+moeda;
  }else if(i==0 && d==0){
    rd="zero "+moeda;
  }  

  if(d>0 && i>0){
    rd=" e "+rd;
  }
    
  r=ri+rd;
  return r;
}

function CentenaExtenso(n){
  var u,d,c,casas;
  var r="";
  var t1=["um","dois","três","quatro","cinco","seis","sete","oito","nove"];
  var t2=["dez","onze","doze","treze","quatorze","quinze","dezesseis","dezessete","dezoito","dezenove"];
  var t3=["vinte","trinta","quarenta","cinquenta","sessenta","setenta","oitenta","noventa"];
  var t4=["cento","duzentos","trezentos","quatrocentos","quinhentos","seiscentos","setecentos","oitocentos","novecentos"];
  casas=n.toString().length;
  u=0;d=0;c=0;
  if(n>0) {u=parseInt(n.toString().substr(casas-1,1));}
  if(n>9) {d=parseInt(n.toString().substr(casas-2,1));}
  if(n>99){c=parseInt(n.toString().substr(casas-3,1));}
  if(n==100){return "cem";}
  else {
    if(c>0){
      r=r+t4[c-1];
      if(d>0 || u>0){r=r+" e ";}
    }
    if(d>1){
      r=r+t3[d-2];
      if(u>0){r=r+" e ";}
    } else if(d==1 && u>=0){
      r=r+t2[d+u-1];
    }
    if(u>0 && d!=1){
      r=r+t1[u-1];
    }
  }
  return r;
}


function send_Rec_Email() {
  // ID do modelo recibo no Google Docs
  var recibotemplateId = "14Zlj5zwYyWHhAUnBG9dMFYTzeClIa_xvayxJsB8k_Os"
  var recibotempDoc = "Recibo";
  
  // Carrega planilha ativa
  var ss = SpreadsheetApp.getActive()
  //var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Recibos")

  var dados = ss.getDataRange().getValues();
  var ultimaLinha = ss.getLastRow() - 1; //Pega a última linha da tabela

  // informações da planilha
  // var nome_completo = dados[ultimaLinha]['coloque aqui o nº da coluna onde ficara o Nome completo da pessoa, lembrando que a coluna A vale 0(zero)'];
  var nome_completo = dados[ultimaLinha][6];
  
  // var evento = dados[ultimaLinha]['coloque aqui o nº da coluna onde ficara o valor do curso, lembrando que a coluna A vale 0(zero)'];
  var evento = dados[ultimaLinha][4];
  
  // var unidade = dados[ultimaLinha]['coloque aqui o nº da coluna onde ficara o unidade responsavel, lembrando que a coluna A vale 0(zero)'];
  var unidade = dados[ultimaLinha][5];
  
  // var valor = dados[ultimaLinha]['coloque aqui o nº da coluna onde ficara o valor do curso, lembrando que a coluna A vale 0(zero)'];
  var valor = dados[ultimaLinha][11];
  
    // var valor = dados[ultimaLinha]['coloque aqui o nº da coluna onde ficara o valor do curso, lembrando que a coluna A vale 0(zero)'];
  var valorextenso = Extenso(valor);
  
  // var CPF = dados[ultimaLinha]['coloque aqui o nº da coluna onde ficara o CPF da pessoa, lembrando que a coluna A vale 0(zero)'];
  
  var CPF2 = String(dados[ultimaLinha][7]);
  if (CPF2.lenght != 11){
    CPF2= '0'.concat(CPF2);
  }
  var CPF = CPF2.replace(/^(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4");
  // var id_recibo = dados[ultimaLinha]['coloque aqui o nº da coluna onde ficara o id_recibo da transaçao, lembrando que a coluna A vale 0(zero)'];
  var id_recibo = dados[ultimaLinha][0];
  
  // var datarecibo = dados[ultimaLinha]['coloque aqui o nº da coluna onde ficara o id_recibo da transaçao, lembrando que a coluna A vale 0(zero)'];
  var datarecibo = dados[ultimaLinha][2];

  //var destinatariorecibo = dados[ultimaLinha]['coloque aqui o nº da coluna onde ficara o e-mail, lembrando que a coluna A vale 0(zero)'];
  var destinatariorecibo = dados[ultimaLinha][9];
  
  
  //var destinatariorecibo = dados[ultimaLinha]['coloque aqui o nº da coluna onde ficara o e-mail, lembrando que a coluna A vale 0(zero)'];
  var idrecibo = dados[ultimaLinha][0];
  
  
  // defino destinatario do Canhoto - Fixo 
  var destinatariocanhoto = "adson.batista@live.com";


  //EMAIL
  // Assunto do email
  var subject = "Recibo IEEE UFABC";
  // Mensagem do Corpo 
  //var body = "Segue anexo o recibo " + id_recibo;
    // Mensagem do Corpo 
  var html  =  
    '<body>' + 
    '<h2><b>Olá ' + nome_completo +'!'+ '</h2></b>' +
    'Você está recebendo este e-mail pois no dia ' +'<b>' + datarecibo +'</b>'+ 
     ' você efetuou um pagamento no valor de <b> ' + valor + ' ('+ valorextenso + ') ' +  '</b>' + 'referente ao ' +'<b>'+ evento +'</b>'+'<br>'+
    'Seu recibo foi anexado neste email e pode ser identificado pelo ID' +'<b>'+ id_recibo +'</b>' +'.'+
    '</body>'
  
  // Dados do Rementente
  var remetente = "IEEE UFABC<contato@ieeeufabc.org>";

  // Cria um recibo temporário, recupera o ID e o abre
  var idCopia = DriveApp.getFileById(recibotemplateId).makeCopy(recibotempDoc +'_' + id_recibo).getId();
  // var idCopia = DriveApp.getFileById(recibotemplateId).makeCopy(recibotempDoc +'_' + id_recibo + '_' + nome_completo).getId();
  var docCopia = DocumentApp.openById(idCopia);

  // recupera o corpo do recibo
  var bodyCopia = docCopia.getActiveSection();

  // faz o replace das variáveis do template, salva e fecha o documento temporario
  bodyCopia.replaceText("NOME", nome_completo);
  bodyCopia.replaceText("NUMEROCPF", CPF);
  bodyCopia.replaceText("VALOR", valor);
  bodyCopia.replaceText("VALEXTENSO", valorextenso);
  bodyCopia.replaceText("CURSO", evento);
  bodyCopia.replaceText("DATARECIBO", datarecibo);
  bodyCopia.replaceText("IDRECIBO", idrecibo);
  docCopia.saveAndClose();

  // abre o documento temporario como PDF utilizando o seu ID
  var recibo_pdf = DriveApp.getFileById(idCopia).getAs("application/pdf");

  //Pastas Drive para Salvar recibos
  var folderramoID = "0B8CcpExpMKFldjB3QWMxODdOdkk"; 
  var folderAESSID = "0B8CcpExpMKFlSnd6RVprakJYN3M"; 
  var folderCSID = "0B8CcpExpMKFldjB3QWMxODdOdkk"
  var folderCPMTID = "0B8CcpExpMKFldjB3QWMxODdOdkk"
  var folderEMBSID = "0B8CcpExpMKFldjB3QWMxODdOdkk"
  var folderPESID = "0B8CcpExpMKFldjB3QWMxODdOdkk"
  var folderRASID = "0B8CcpExpMKFldjB3QWMxODdOdkk"
  var folderTEMSID = "0B8CcpExpMKFldjB3QWMxODdOdkk"
  var folder_recibo_CS = DriveApp.getFolderById(folderCSID);
  var folder_recibo_CPMT = DriveApp.getFolderById(folderCPMTID);
  var folder_recibo_EMBS = DriveApp.getFolderById(folderEMBSID);
  var folder_recibo_PES = DriveApp.getFolderById(folderPESID);
  var folder_recibo_RAS = DriveApp.getFolderById(folderRASID);
  var folder_recibo_TEMS = DriveApp.getFolderById(folderTEMSID);
  var folder_recibo_ramo = DriveApp.getFolderById(folderramoID);
  var folder_recibo_AESS = DriveApp.getFolderById(folderAESSID);
  
  //salva pdf na pasta do ID
  if (unidade == "Ramo") {folder_recibo_ramo.createFile(recibo_pdf)
    }
  else if (unidade == "AESS") {folder_recibo_AESS.createFile(recibo_pdf)
    }
  else if (unidade == "CS") {folder_recibo_CS.createFile(recibo_pdf)
    }
  else if (unidade == "CPMT") {folder_recibo_CPMT.createFile(recibo_pdf)
    }
  else if (unidade == "EMBS") {folder_recibo_EMBS.createFile(recibo_pdf)
    }
  else if (unidade == "PES") {folder_recibo_PES.createFile(recibo_pdf)
    }
  else if (unidade == "RAS") {folder_recibo_RAS.createFile(recibo_pdf)
    }
  else if (unidade == "TEMS") {folder_recibo_TEMS.createFile(recibo_pdf)
    }
  
  // envia o email com recibo para destinatario
  // MailApp.sendEmail(destinatariorecibo, subject, body, {name: remetente, attachments: recibo_pdf});
  MailApp.sendEmail(destinatariorecibo, subject, html, {name: remetente, htmlBody: html, attachments: recibo_pdf});
  // envia o email recibo para email do ramo  
  MailApp.sendEmail(destinatariocanhoto, subject, html, {name: remetente, htmlBody: html, attachments: recibo_pdf});
  
  // apaga o documento temporário
  DriveApp.getFileById(idCopia).setTrashed(true);  
}


function doPost(e) { /*exportar() {/**/

  // Os dados enviados no corpo da requisição POST estarão em e.postData.contents
  var dadosRecebidos = JSON.parse(e.postData.contents);

  // Acessa os parâmetros
  const colunas = dadosRecebidos.colunas;
  const colunas_data = dadosRecebidos.colunas_data
  const tickets_enviados = dadosRecebidos.tickets;
  const coluna_ticket = dadosRecebidos.coluna_ticket;
  const nome_aba_fonte = dadosRecebidos.nome_aba_fonte;
  const fonte = dadosRecebidos.fonte;

  // Abre a planilha pelo ID
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  
  var aba = planilha.getSheetByName(nome_aba_fonte);

  //Obtém todos os tickets da coluna A
  const n_linhas = aba.getMaxRows()
  var tickets_aba = aba.getRange(1,coluna_ticket,n_linhas).getValues().map(linha => linha[0]);

  // Obtém todos os dados da aba
  const n_colunas = aba.getMaxColumns()
  var dados = []
  tickets_enviados.forEach(ticket => {
    const linha = tickets_aba.indexOf(ticket) + 1
    if(linha > 0){
      const dados_linha = aba.getRange(linha,1,1,n_colunas).getValues()[0]
      dados.push(dados_linha)
    }
  })

  //protocolos
  let colunas_export_abc = colunas
  
  // Converte os dados para o formato CSV
  var csv = dados.map(function(linha) {
    if(linha[colunas_export_abc[0]] === "") return undefined
    var dados_linha = colunas_export_abc.map((coluna,i) => {
      var conteudoCelula = linha[coluna].toString().replace(/"/g, '""');
      if(conteudoCelula !== "" && conteudoCelula !== undefined && conteudoCelula !== null) {
        if (colunas_data.indexOf(coluna) >= 0) {
          var data = new Date(conteudoCelula);
          // Se a data for válida (não um "Invalid Date"), formata a data
          if (!isNaN(data.getTime())) {
            conteudoCelula = Utilities.formatDate(data, Session.getScriptTimeZone(), "dd/MM/yyyy");
          }
          return conteudoCelula;
        }
        if(i === colunas_export_abc.length-1){
          return fonte
        }
        return '"' + conteudoCelula + '"';
      }
      
    })
    //console.log(dados_linha.join(','))
    return dados_linha.join(',');
  })
  
  csv = csv.filter(linha => linha !== undefined)
  csv = csv.join('\n');
  
  // Retorna o CSV com o tipo de conteúdo adequado
  return ContentService
    .createTextOutput(csv)
    .setMimeType(ContentService.MimeType.CSV);
}

function doGet(e) { /*exportar() {/**/

  // Abre a planilha pelo ID
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  
  var aba = planilha.getSheetByName("CDC");

  //Obtém todos os tickets da coluna A
  console.log("iniciando getvalues")
  var dados = aba.getDataRange().getValues()
  
  console.log("iniciando leitura dos dados")
  // Converte os dados para o formato CSV
  var csv = dados.map((linha,id_linha) => {
    var dados_linha = linha.map((dado_coluna) => {
      var conteudoCelula = dado_coluna.toString().replace(/"/g, '""');
      if(conteudoCelula !== "" && conteudoCelula !== undefined && conteudoCelula !== null) {
        return '"' + conteudoCelula + '"';
      }
    })
    //console.log(dados_linha.join(','))
    return dados_linha.join(',');
  })

  console.log("iniciando pulo de linhas csv")
  
  csv = csv.filter(linha => linha !== undefined)
  csv = csv.join('\n');
  
  // Retorna o CSV com o tipo de conteúdo adequado
  return ContentService
    .createTextOutput(csv)
    .setMimeType(ContentService.MimeType.CSV);
}

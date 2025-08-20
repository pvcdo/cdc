function atualizacaoAzul() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const aba_atu_azul = ss.getSheetByName('Paulo - azul')

  const aba_cdc = ss.getSheetByName('CDC')
  
  var lin_atu = 2
  var continuar = true

  while (continuar) {
    const dados = aba_atu_azul.getRange(lin_atu,1,1,4).getValues()[0]
    
    if(dados[0] == ""){
      continuar = false
      console.log('chegamos no fim')
      break
    }

    const atualizar = parseInt(aba_atu_azul.getRange(lin_atu,9).getValue())
    const inserido = aba_atu_azul.getRange(lin_atu,10).getValue()

    if (atualizar == 1 || inserido == "inserido") {
      lin_atu++
      console.log(dados[2] + ' já estava na planilha ou já inserido anteriormente por aqui')
      continue
    }
  
    const linha_cdc = aba_atu_azul.getRange(lin_atu,12).getValue()

    aba_cdc.getRange(linha_cdc,38).setValue(dados[1]) //nome
    aba_cdc.getRange(linha_cdc,39).setValue(dados[2]) // matricula
    aba_cdc.getRange(linha_cdc,44).setValue(dados[3]) // inicio

    aba_atu_azul.getRange(lin_atu,10).setValue('inserido')

    console.log(dados[0]+' inserido')
    
    lin_atu++

  }
  
}

function atualizacaoAzul_Otimizada() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba_atu_azul = ss.getSheetByName('Paulo - azul');
  const aba_cdc = ss.getSheetByName('CDC');

  // PASSO 1: LER TODOS OS DADOS DE UMA VEZ
  // Pega todos os dados da planilha de origem, da linha 2 até a última linha com conteúdo.
  // Isso transforma uma potencial centena de leituras em apenas UMA.
  const intervaloDados = aba_atu_azul.getRange(2, 1, aba_atu_azul.getLastRow() - 1, 12);
  const todosOsDados = intervaloDados.getValues();

  // Prepara um array para armazenar os status "inserido" que serão escritos de volta.
  const statusParaAtualizar = [];

  // PASSO 2: PROCESSAR OS DADOS EM MEMÓRIA (MUITO RÁPIDO)
  // Agora, iteramos sobre o array 'todosOsDados', o que não envolve chamadas à planilha.
  for (let i = 0; i < todosOsDados.length; i++) {
    const linhaAtual = todosOsDados[i];
    
    // Os índices do array são baseados em 0 (coluna A = 0, B = 1, etc.)
    const id = linhaAtual[0]; // Coluna A
    const nome = linhaAtual[1]; // Coluna B
    const matricula = linhaAtual[2]; // Coluna C
    const inicio = linhaAtual[3]; // Coluna D
    const atualizar = parseInt(linhaAtual[8]); // Coluna I
    const inserido = linhaAtual[9]; // Coluna J
    const linha_cdc = linhaAtual[11]; // Coluna L
    
    // Se a primeira célula da linha estiver vazia, paramos.
    if (id === "") {
      console.log('Chegamos no fim dos dados');
      break; // Sai do loop
    }

    // Mantém a mesma lógica de verificação
    if (atualizar === 1 || inserido === "inserido") {
      console.log(matricula + ' já estava na planilha ou já inserido anteriormente por aqui');
      statusParaAtualizar.push(['inserido']); // Mantém o status para reescrever e garantir consistência
      continue; // Pula para a próxima iteração
    }

    // As escritas na aba 'CDC' são mais difíceis de agrupar porque são em linhas
    // aleatórias. No entanto, como eliminamos todas as leituras do loop,
    // estas poucas escritas serão muito mais rápidas.
    aba_cdc.getRange(linha_cdc, 38).setValue(nome);
    aba_cdc.getRange(linha_cdc, 39).setValue(matricula);
    aba_cdc.getRange(linha_cdc, 44).setValue(inicio);
    
    // Em vez de escrever 'inserido' na planilha agora, adicionamos ao nosso array de status.
    statusParaAtualizar.push(['inserido']);
    console.log(id + ' inserido');
  }

  // PASSO 3: ESCREVER OS RESULTADOS DE VOLTA (EM LOTE)
  // Após o loop terminar, se tivermos status para atualizar,
  // escrevemos todos de uma vez só na coluna J ('inserido').
  if (statusParaAtualizar.length > 0) {
    // Escreve o array de status na coluna 10 (J), começando da linha 2.
    aba_atu_azul.getRange(2, 10, statusParaAtualizar.length, 1).setValues(statusParaAtualizar);
    console.log('Status de ' + statusParaAtualizar.length + ' linhas foram atualizados na planilha de origem.');
  }
}

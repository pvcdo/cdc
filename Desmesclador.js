/**
 * @OnlyCurrentDoc
 *
 * Cria um menu personalizado na planilha quando ela é aberta.
 
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Ferramentas Especiais')
      .addItem('Apenas Desmesclar Células', 'unmergeSelectedCells')
      .addToUi();
}

/**
 * Desfaz a mesclagem de todas as células dentro do intervalo 
 * que o usuário selecionou.

function unmergeSelectedCells() {
  // Pega o intervalo de células que o usuário selecionou ativamente.
  const range = SpreadsheetApp.getActiveRange();

  // A única ação necessária é esta linha, que desfaz todas as mesclagens no intervalo.
  range.breakApart();

  // Opcional: Mostra uma mensagem confirmando que a operação foi concluída.
  SpreadsheetApp.getUi().alert('As células na sua seleção foram desmescladas com sucesso!');
}
**/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("📌 Despachos padronizados");
  menu.addItem("📜 Despacho","despacho")
        .addToUi();
}

/**
 * Helper function to normalize string values for comparison.
 * @param {any} value The value to normalize.
 * @returns {string} The normalized string (trimmed, uppercase).
 */
function normalizeString(value) {
  return String(value || '').trim().toUpperCase();
}

/**
 * Helper function to show a standardized error alert.
 * @param {GoogleAppsScript.Base.Ui} ui The UI object.
 * @param {string} message The message to display.
 */
function showErrorAlert(ui, message) {
  ui.alert("Erro", message, ui.ButtonSet.OK);
}

/**
 * Validates the active sheet and selected row.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh The active sheet.
 * @param {GoogleAppsScript.Base.Ui} ui The UI object.
 * @param {string} expectedSheetName The name of the expected sheet.
 * @returns {number | null} The active row index (1-based) if valid, otherwise null.
 */
function validateActiveRow(sh, ui, expectedSheetName) {
  if (sh.getName() !== expectedSheetName) {
    showErrorAlert(ui, `Por favor, navegue para a aba '${expectedSheetName}' e selecione a linha desejada antes de executar esta função.`);
    return null;
  }

  const activeRange = sh.getActiveRange();
  // Assume que a linha 1 é o cabeçalho
  if (!activeRange || activeRange.getHeight() > 1 || activeRange.getRow() <= 1) {
    showErrorAlert(ui, `Por favor, selecione uma única célula ou a linha de dados que deseja processar na aba '${expectedSheetName}'.`);
    return null;
  }
  return activeRange.getRow();
}

/**
 * Extracts initial data from a selected row.
 * @param {Array<any>} selectedRowData The array of data from the selected row.
 * @returns {Object} An object containing the extracted initial data.
*/
function extractInitialData(selectedRowData) {
  return {
    prof_saiu: selectedRowData[18],        // Coluna S - Profissional a ser substituído (índice 18)
    mot_cont_achado: selectedRowData[20],  // Coluna U - Motivo da Contratação (índice 20)
    cargo_cdc_achado: selectedRowData[5],  // Coluna F - Cargo (índice 5)
    especialidade_cdc_achado: selectedRowData[6], // Coluna G - Especialidade (índice 6)
    unidade_achado: selectedRowData[11],   // Coluna L - Unidade/Lotação (índice 11)
    area_at_cdc_achado: selectedRowData[14], // Coluna O - Área de Atuação (índice 14)
    clas_cdc_achado: selectedRowData[13],  // Coluna N - Classificação (índice 13)
    cargaH_cdc_achado: selectedRowData[8], // Coluna I - Carga Horária (índice 8)
    equipe_cdc_achado: selectedRowData[10] // Coluna K - Equipe (índice 10)
  };
}

/**
 * Validates the extracted initial data.
 * @param {GoogleAppsScript.Base.Ui} ui The UI object.
 * @param {Object} initialData The object containing initial data.
 * @returns {boolean} True if all required fields are present, false otherwise.
 */
function validateInitialData(ui, initialData) {
  const fieldsToValidate = {
    prof_saiu: 'Profissional a ser substituído (Coluna S)',
    mot_cont_achado: 'Motivo da Contratação (Coluna U)',
    cargo_cdc_achado: 'Cargo (Coluna F)',
    especialidade_cdc_achado: 'Especialidade (Coluna G)',
    unidade_achado: 'Lotação/Unidade (Coluna L)',
    area_at_cdc_achado: 'Área de Atuação (Coluna O)',
    clas_cdc_achado: 'Classificação (Coluna N)',
    cargaH_cdc_achado: 'Carga Horária (Coluna I)',
    equipe_cdc_achado: 'Equipe (Coluna K)'
  };

  for (const key in fieldsToValidate) {
    if (normalizeString(initialData[key]) === "") {
      showErrorAlert(ui, `Erro na aba 'Despachos': O campo '${fieldsToValidate[key]}' está vazio para a linha selecionada.`);
      return false;
    }
  }
  return true;
}

/**
 * Formats a date string (YYYY-MM-DD) to dd/MM/yyyy.
 * @param {string} dateInput The date string in YYYY-MM-DD format.
 * @returns {string} The formatted date string.
 * @throws {Error} If the date format is invalid or formatting fails.
 */
function formatInputDate(dateInput) {
  try {
    const dateObj = new Date(dateInput + 'T00:00:00'); // Adds T00:00:00 to avoid timezone issues
    if (isNaN(dateObj.getTime())) { // Use .getTime() for robust NaN check
      throw new Error("Formato de data inválido. Use AAAA-MM-DD.");
    }
    return Utilities.formatDate(dateObj, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
  } catch (e) {
    throw new Error("Erro ao formatar a data: " + e.message);
  }
}

/**
 * Determines the 'inciso_atual' and 'repo_dataf_texto' based on the inciso option.
 * @param {number} inciso_sim_value The numeric value of the inciso option.
 * @param {string} repo_dataf_raw The raw replacement date from the form (if applicable).
 * @returns {Object} An object with inciso_atual and repo_dataf_texto.
 * @throws {Error} If inciso option is invalid or required data is missing.
 */
function getIncisoDetails(inciso_sim_value, repo_dataf_raw) {
  let inciso_atual;
  let repo_dataf_texto;

  switch (inciso_sim_value) {
    case 1: // TEMPORÁRIO 2 ANOS - Inciso IV
      inciso_atual = "Lei 11175/2019, Art. 2º, Inciso IV - carência de pessoal em decorrência de afastamentos ou licença de servidores.";
      if (!repo_dataf_raw) {
        throw new Error("Data fim da reposição não informada para inciso temporário.");
      }
      repo_dataf_texto = formatInputDate(repo_dataf_raw);
      break;
    case 2: // DEFINITIVO 2 ANOS - Inciso V
      inciso_atual = "Lei 11175/2019, Art. 2º, Inciso V - número de servidores efetivos insuficiente para a continuidade dos serviços públicos essenciais.";
      repo_dataf_texto = 'a reposição por efetivo';
      break;
    case 3: // TEMPORÁRIO 4 ANOS - Inciso VI
      inciso_atual = "Lei 11175/2019, Art. 2º, inciso VI: carência de pessoal para o desempenho de atividades sazonais, projetos temporários ou emergenciais que não justifiquem a criação de cargo efetivo.";
      if (!repo_dataf_raw) {
        throw new Error("Data fim da reposição não informada para inciso temporário.");
      }
      repo_dataf_texto = formatInputDate(repo_dataf_raw);
      break;
    default:
      throw new Error("Opção de inciso inválida. Por favor, selecione uma opção válida para o inciso.");
  }
  return { inciso_atual, repo_dataf_texto };
}

/**
 * Determines the normalized plantão string based on the numeric input.
 * @param {number} plantao_sim_value The numeric value of the plantão option.
 * @returns {string} The normalized plantão string.
 * @throws {Error} If the plantão option is invalid.
 */
function getPlantaoInput(plantao_sim_value) {
  switch (plantao_sim_value) {
    case 1:
      return "NÃO SE APLICA";
    case 2:
      return "Plantões Dias de Semana";
    case 3:
      return "Plantões Fds";
    case 4:
      return "Plantões Dias de Semana + 1Fds";
    default:
      throw new Error("Opção de plantão inválida. Operação cancelada.");
  }
}

/**
 * Searches the 'Simulador cadm 01/12/2024' sheet for matching data.
 * @param {Object} criteria The search criteria (cargo, area, equipe, plantao, classificacao, cargaHoraria).
 * @returns {Object} An object containing the found simulated data (vencimentof, abonof, etc.).
 * @throws {Error} If the simulator sheet is not found or no matching data is found.
 */
function getSimuladorData(criteria) {
  const ssSimulador = SpreadsheetApp.getActiveSpreadsheet();
  const abaSimulador = ssSimulador.getSheetByName("Simulador cadm 01/12/2024");

  if (!abaSimulador) {
    throw new Error("Erro: Aba 'Simulador cadm 01/12/2024' não encontrada na planilha. Verifique o nome da aba.");
  }

  const dadosSimulador = abaSimulador.getDataRange().getValues();
  let dadosEncontrados = {};
  let dadosSimuladorEncontrados = false;

  const normalizedCriteria = {
    cargo: normalizeString(criteria.cargo),
    area: normalizeString(criteria.area),
    equipe: normalizeString(criteria.equipe),
    plantao: normalizeString(criteria.plantao),
    classificacao: normalizeString(criteria.classificacao),
    cargaHoraria: normalizeString(criteria.cargaHoraria)
  };

  for (let i = 0; i < dadosSimulador.length; i++) {
    const simuladorProfissional = normalizeString(dadosSimulador[i][0]);
    const simuladorArea = normalizeString(dadosSimulador[i][1]);
    const simuladorEquipe = normalizeString(dadosSimulador[i][2]);
    const simuladorPlantao = normalizeString(dadosSimulador[i][3]);
    const simuladorClassificacao = normalizeString(dadosSimulador[i][4]);
    const simuladorCargaHoraria = normalizeString(dadosSimulador[i][5]);

    if (simuladorProfissional === normalizedCriteria.cargo &&
        simuladorArea === normalizedCriteria.area &&
        simuladorEquipe === normalizedCriteria.equipe &&
        simuladorPlantao === normalizedCriteria.plantao &&
        simuladorClassificacao === normalizedCriteria.classificacao &&
        simuladorCargaHoraria === normalizedCriteria.cargaHoraria) {

      dadosEncontrados = {
        vencimentof: dadosSimulador[i][6],
        abonof: dadosSimulador[i][7],
        plantdds_f: dadosSimulador[i][8],
        plantddsfds_f: dadosSimulador[i][9],
        plantfds_f: dadosSimulador[i][10],
        redecomp_f: dadosSimulador[i][11],
        insalub_f: dadosSimulador[i][12],
        premio_f: dadosSimulador[i][13],
        compenfer_f: dadosSimulador[i][14],
        remun_b: dadosSimulador[i][15],
        custom: dadosSimulador[i][21],
        custoa: dadosSimulador[i][22]
      };
      dadosSimuladorEncontrados = true;
      break;
    }
  }

  if (!dadosSimuladorEncontrados) {
    throw new Error("Erro: Não existe essa combinação de critérios (Cargo, Área, Equipe, Plantão, Classificação, Carga Horária) na aba 'Simulador cadm 01/12/2024'. Verifique os dados das colunas A a F.");
  }
  return dadosEncontrados;
}


// Refactored `despacho()`
function despacho(){
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const activeRowIndex = sh.getActiveRange().getRow();
  const motivo = sh.getRange(activeRowIndex, 21).getValue();
  const motivoNormalizado = normalizeString(motivo);

  const processFunctions = {
    "RESCISÃO CONTRATUAL": processarResci_Termi,
    "TÉRMINO CONTRATUAL": processarResci_Termi, // Both map to the same function
    "FALECIMENTO": processar_falecimento,
    "APOSENTADORIA": processar_aposent, // Uncomment if function is provided
    "EXONERAÇÃO": processar_exo,       // Uncomment if function is provided
    "MOVIMENTAÇÃO": processar_movi      // Uncomment if function is provided
  };

  const funcToCall = processFunctions[motivoNormalizado];

  if (funcToCall) {
    funcToCall();
  } else {
    showErrorAlert(ui, `O motivo "${motivo}" na Coluna U da linha selecionada não corresponde a um tipo de despacho padronizado.`);
  }
}

// Refactored `processarResci_Termi()`
function processarResci_Termi() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();

  const activeRowIndex = validateActiveRow(sh, ui, "CDC");
  if (activeRowIndex === null) return;

  const selectedRowData = sh.getRange(activeRowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
  const initialData = extractInitialData(selectedRowData);

  if (!validateInitialData(ui, initialData)) return;

  const htmlTemplate = HtmlService.createTemplateFromFile('SidebarRescisao');
  htmlTemplate.initialData = JSON.stringify(initialData);
  const htmlOutput = htmlTemplate.evaluate()
    .setTitle('Despacho de Rescisão/Término Contratual')
    .setWidth(300);
  ui.showSidebar(htmlOutput);
}

// Refactored `processarDadosRescisaoSidebar()`
function processarDadosRescisaoSidebar(formData, initialDataStr) {
  const ui = SpreadsheetApp.getUi();
  const initialData = JSON.parse(initialDataStr);

  const impacto = parseInt(formData.impacto);
  const impacto_ccg = formData.impacto_ccg;
  const mat = formData.mat;
  const plantao_sim_value = parseInt(formData.plantao_sim);
  const inciso_sim_value = parseInt(formData.inciso_sim);

  let impacto_novo;
  if (impacto === 1) {
    impacto_novo = "sem impacto";
  } else if (impacto === 2) {
    if (!impacto_ccg) {
      throw new Error("CCG não informada. Operação cancelada.");
    }
    impacto_novo = "com impacto, CCG " + impacto_ccg;
  } else {
    throw new Error("Opção de impacto inválida. Por favor, digite 1 ou 2.");
  }

  const { inciso_atual, repo_dataf_texto } = getIncisoDetails(inciso_sim_value, formData.repo_dataf);
  const data_termino_formatada = formatInputDate(formData.data);
  const plantao_input = getPlantaoInput(plantao_sim_value);

  // Alteração para que o cargo seja "TÉCNICO EM ENFERMAGEM" se a especialidade for "TÉCNICO EM ENFERMAGEM"
  let cargo_for_simulador = initialData.cargo_cdc_achado;
  if (normalizeString(initialData.especialidade_cdc_achado) === "TÉCNICO EM ENFERMAGEM") {
      cargo_for_simulador = "TÉCNICO EM ENFERMAGEM";
  }

  const simuladorCriteria = {
    cargo: cargo_for_simulador, // Usando o cargo potencialmente modificado
    area: initialData.area_at_cdc_achado,
    equipe: initialData.equipe_cdc_achado,
    plantao: plantao_input,
    classificacao: initialData.clas_cdc_achado,
    cargaHoraria: initialData.cargaH_cdc_achado
  };

  const { vencimentof, abonof, plantdds_f, plantddsfds_f, plantfds_f, redecomp_f,
          insalub_f, premio_f, compenfer_f, remun_b, custom, custoa } = getSimuladorData(simuladorCriteria);

  const despacho =
    `Fica autorizado até ${repo_dataf_texto}, ${impacto_novo}, em razão de ${initialData.mot_cont_achado} em ${data_termino_formatada} do profissional ${initialData.prof_saiu}, MAT ${mat}, ${initialData.cargo_cdc_achado}, ${initialData.especialidade_cdc_achado}, ${initialData.equipe_cdc_achado}, ${initialData.cargaH_cdc_achado}, ${initialData.unidade_achado}.\n` +
    `Motivo da contratação: ${inciso_atual}\n\n` +
    `Vencimento: ${vencimentof}\n` +
    `Abono de fixação: ${abonof}\n` +
    `Rede complementar: ${redecomp_f}\n` +
    `Insalubridade: ${insalub_f}\n` +
    `Prêmio Pró-Família: ${premio_f}\n` +
    `Plantão dia de semana: ${plantdds_f}\n` +
    `Plantão dia de semana + 1 Fds: ${plantddsfds_f}\n` +
    `Plantão fim de semana: ${plantfds_f}\n` +
    `Complementação enfermagem: ${compenfer_f}\n` +
    `Remuneração (bruto): ${remun_b}\n` +
    `Custo mensal: ${custom}\n` +
    `Custo anual: ${custoa}\n\n` +
    `;${impacto_ccg};${initialData.cargo_cdc_achado};${initialData.especialidade_cdc_achado};${initialData.equipe_cdc_achado};` +
    `${initialData.cargaH_cdc_achado};${initialData.unidade_achado};`;

  return despacho;
}

// Refactored `processar_falecimento()`
function processar_falecimento() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();

  const activeRowIndex = validateActiveRow(sh, ui, "CDC");
  if (activeRowIndex === null) return;

  const selectedRowData = sh.getRange(activeRowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
  const initialData = extractInitialData(selectedRowData);

  if (!validateInitialData(ui, initialData)) return;

  const htmlTemplate = HtmlService.createTemplateFromFile('SidebarFalecimento');
  htmlTemplate.initialData = JSON.stringify(initialData);
  const htmlOutput = htmlTemplate.evaluate()
      .setTitle('Despacho de Falecimento')
      .setWidth(300);
  ui.showSidebar(htmlOutput);
}

// Refactored `processarDadosFalecimentoSidebar()`
function processarDadosFalecimentoSidebar(formData, initialDataStr) {
  const ui = SpreadsheetApp.getUi();
  const initialData = JSON.parse(initialDataStr);

  const impacto = parseInt(formData.impacto);
  const impacto_ccg = formData.impacto_ccg;
  const mat = formData.mat;
  const plantao_sim_value = parseInt(formData.plantao_sim);
  const inciso_sim_value = parseInt(formData.inciso_sim);

  let impacto_novo;
  if (impacto === 1) {
    impacto_novo = "sem impacto";
  } else if (impacto === 2) {
    if (!impacto_ccg) {
      throw new Error("CCG não informada. Operação cancelada.");
    }
    impacto_novo = "com impacto, CCG " + impacto_ccg;
  } else {
    throw new Error("Opção de impacto inválida. Por favor, digite 1 ou 2.");
  }

  const dataFormatted = formatInputDate(formData.data);
  const { inciso_atual, repo_dataf_texto } = getIncisoDetails(inciso_sim_value, formData.repo_dataf);
  const plantao_input = getPlantaoInput(plantao_sim_value);

  // Alteração para que o cargo seja "TÉCNICO EM ENFERMAGEM" se a especialidade for "TÉCNICO EM ENFERMAGEM"
  let cargo_for_simulador = initialData.cargo_cdc_achado;
  if (normalizeString(initialData.especialidade_cdc_achado) === "TÉCNICO EM ENFERMAGEM") {
      cargo_for_simulador = "TÉCNICO EM ENFERMAGEM";
  }

  const simuladorCriteria = {
    cargo: cargo_for_simulador, // Usando o cargo potencialmente modificado
    area: initialData.area_at_cdc_achado,
    equipe: initialData.equipe_cdc_achado,
    plantao: plantao_input,
    classificacao: initialData.clas_cdc_achado,
    cargaHoraria: initialData.cargaH_cdc_achado
  };

  const { vencimentof, abonof, plantdds_f, plantddsfds_f, plantfds_f, redecomp_f,
          insalub_f, premio_f, compenfer_f, remun_b, custom, custoa } = getSimuladorData(simuladorCriteria);

  const despacho =
    `Fica autorizado até ${repo_dataf_texto}, ${impacto_novo}, em razão de falecimento em ${dataFormatted} do profissional ${initialData.prof_saiu}, MAT ${mat}, ${initialData.cargo_cdc_achado}, ${initialData.especialidade_cdc_achado}, ${initialData.equipe_cdc_achado}, ${initialData.cargaH_cdc_achado}, ${initialData.unidade_achado}.\n` +
    `Motivo da contratação: ${inciso_atual}\n\n` +
    `Vencimento: ${vencimentof}\n` +
    `Abono de fixação: ${abonof}\n` +
    `Rede complementar: ${redecomp_f}\n` +
    `Insalubridade: ${insalub_f}\n` +
    `Prêmio Pró-Família: ${premio_f}\n` +
    `Plantão dia de semana: ${plantdds_f}\n` +
    `Plantão dia de semana + 1 Fds: ${plantddsfds_f}\n` +
    `Plantão fim de semana: ${plantfds_f}\n` +
    `Complementação enfermagem: ${compenfer_f}\n` +
    `Remuneração (bruto): ${remun_b}\n` +
    `Custo mensal: ${custom}\n` +
    `Custo anual: ${custoa}\n\n` +
    `;${impacto_ccg};${initialData.cargo_cdc_achado};${initialData.especialidade_cdc_achado};${initialData.equipe_cdc_achado};` +
    `${initialData.cargaH_cdc_achado};${initialData.unidade_achado};`;

  return despacho;
}

// Placeholder functions as they were mentioned but not fully provided in the original text
function processar_aposent() {
  const ui = SpreadsheetApp.getUi();
  showErrorAlert(ui, "Função para aposentadoria ainda não implementada.");
}

function processar_exo() {
  const ui = SpreadsheetApp.getUi();
  showErrorAlert(ui, "Função para exoneração ainda não implementada.");
}

function processar_movi() {
  const ui = SpreadsheetApp.getUi();
  showErrorAlert(ui, "Função para movimentação ainda não implementada.");
}
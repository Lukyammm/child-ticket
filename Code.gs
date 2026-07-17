// Google Apps Script - Casa da Gestante
// SPREADSHEET_ID: Insira o ID real da sua planilha aqui.
const SPREADSHEET_ID = 'SUBSTITUA_PELO_ID_DA_PLANILHA';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Casa da Gestante - Controle de Dietas')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Retorna as abas (Dias) disponíveis na planilha
function getDatasDisponiveis() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = ss.getSheets();
    return sheets.map(s => s.getName());
  } catch(e) { 
    // Fallback Mock para testes sem ID da planilha
    return ['17/07/2026', '18/07/2026']; 
  }
}

// Retorna os pacientes da aba especificada
function getPacientesPorData(dataStr) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(dataStr);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const pacientes = [];
    
    // Assumindo que a linha 0 é o cabeçalho
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0] && !row[1]) continue; // Pula linhas vazias
      
      pacientes.push({
        id: i, // Usado como referência para edição (número da linha)
        prontuario_nome: row[0] || "",
        diagnostico: row[1] || "",
        dieta: row[2] || "",
        desjejum: row[3] || "",
        colacao: row[4] || "",
        almoco: row[5] || "",
        lanche: row[6] || "",
        jantar: row[7] || "",
        ceia: row[8] || "",
        observacao: row[9] || ""
      });
    }
    return pacientes;
  } catch(e) {
    // Retorna mock
    return [
      { id: 1, prontuario_nome: "21768 - Maria Evilene Menezes Oliveira", diagnostico: "Puerpera", dieta: "Geral", desjejum: "CAFE COMPLETO + 1 OVO", colacao: "SUCO + BISCOITO", almoco: "GERAL COM FRANGO", lanche: "VITAMINA + BISCOITO", jantar: "BRANDA DE FRANGO", ceia: "MINGAU", observacao: "" },
      { id: 2, prontuario_nome: "32083 - Antonia de Araujo Silva", diagnostico: "Puérpera", dieta: "Geral HAS DM", desjejum: "CAFÉ COMPLETO DM", colacao: "SUCO DM", almoco: "GERAL HAS DM", lanche: "VITAMINA DM", jantar: "BRANDA DM HAS", ceia: "MINGAU DM + 1 FRUTA", observacao: "" }
    ];
  }
}

// Salva ou edita um paciente na aba da data especificada
function salvarPaciente(dataStr, paciente) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(dataStr);
    
    // Se a aba para esse dia não existir, cria uma nova com cabeçalho
    if (!sheet) {
      sheet = ss.insertSheet(dataStr);
      sheet.appendRow([
        "Prontuário / Paciente", "Diagnóstico", "Dieta", 
        "Desjejum - 6h", "Colação - 9h", "Almoço - 12h", 
        "Lanche - 15h", "Jantar - 18h", "Ceia - 21h", "Observação"
      ]);
    }
    
    const rowData = [
      paciente.prontuario_nome,
      paciente.diagnostico,
      paciente.dieta,
      paciente.desjejum,
      paciente.colacao,
      paciente.almoco,
      paciente.lanche,
      paciente.jantar,
      paciente.ceia,
      paciente.observacao
    ];

    if (paciente.id) {
      // Edita linha existente (id é o índice do array, logo linha do sheet é id + 1)
      sheet.getRange(paciente.id + 1, 1, 1, 10).setValues([rowData]);
    } else {
      // Adiciona nova linha
      sheet.appendRow(rowData);
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

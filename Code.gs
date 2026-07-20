// Google Apps Script - Casa da Gestante / Oncológico

// Senha da área de administração (cadastro de itens de dieta).
// ALTERE esta senha antes de usar em produção.
const ADMIN_PASSWORD = 'admin123';

// Prefixo usado nas abas do perfil Oncológico, para não colidir com as abas
// já existentes do perfil Casa da Gestante (que continuam usando o nome puro da data).
const PREFIXO_ONCOLOGICO = 'ONCO - ';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Casa da Gestante - Controle de Dietas')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Planilha de dados: criada automaticamente no primeiro uso e o ID fica salvo
// nas Propriedades do Script. Não é necessário configurar nenhum ID manualmente.
function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');

  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {
      // ID salvo não é mais válido (planilha excluída/movida) - recria abaixo
    }
  }

  const ss = SpreadsheetApp.create('Casa da Gestante - Dados');
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

// Resolve o nome real da aba na planilha para o perfil e data informados
function resolveSheetName_(perfil, dataStr) {
  return perfil === 'ONCOLOGICO' ? (PREFIXO_ONCOLOGICO + dataStr) : dataStr;
}

// Retorna as abas (Dias) disponíveis na planilha para o perfil informado
function getDatasDisponiveis(perfil) {
  try {
    const ss = getSpreadsheet_();
    const sheets = ss.getSheets();
    const nomes = sheets.map(s => s.getName());

    if (perfil === 'ONCOLOGICO') {
      return nomes
        .filter(n => n.indexOf(PREFIXO_ONCOLOGICO) === 0)
        .map(n => n.substring(PREFIXO_ONCOLOGICO.length));
    }
    // Gestante: abas que NÃO usam o prefixo Oncológico (mantém compatibilidade com abas antigas)
    return nomes.filter(n => n.indexOf(PREFIXO_ONCOLOGICO) !== 0 && n !== 'ITENS_DIETA');
  } catch(e) {
    // Fallback mock em caso de erro inesperado
    return ['17/07/2026', '18/07/2026'];
  }
}

// Retorna os pacientes da aba especificada, para o perfil informado
function getPacientesPorData(dataStr, perfil) {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(resolveSheetName_(perfil, dataStr));
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const pacientes = [];

    // Assumindo que a linha 0 é o cabeçalho
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0] && !row[1]) continue; // Pula linhas vazias

      pacientes.push({
        rowNumber: i + 1, // Número real da linha na planilha (1-indexed)
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
    if (perfil === 'ONCOLOGICO') {
      return [
        { rowNumber: 2, prontuario_nome: "40521 - Carlos Eduardo Menezes", diagnostico: "Oncológico", dieta: "Branda Hipercalórica", desjejum: "MINGAU DE AVEIA + FRUTA", colacao: "SUCO ADOÇADO", almoco: "BRANDA COM FRANGO DESFIADO", lanche: "VITAMINA HIPERCALÓRICA", jantar: "SOPA CREMOSA", ceia: "MINGAU + COMPLEMENTO", observacao: "" }
      ];
    }
    return [
      { rowNumber: 2, prontuario_nome: "21768 - Maria Evilene Menezes Oliveira", diagnostico: "Puerpera", dieta: "Geral", desjejum: "CAFE COMPLETO + 1 OVO", colacao: "SUCO + BISCOITO", almoco: "GERAL COM FRANGO", lanche: "VITAMINA + BISCOITO", jantar: "BRANDA DE FRANGO", ceia: "MINGAU", observacao: "" },
      { rowNumber: 3, prontuario_nome: "32083 - Antonia de Araujo Silva", diagnostico: "Puérpera", dieta: "Geral HAS DM", desjejum: "CAFÉ COMPLETO DM", colacao: "SUCO DM", almoco: "GERAL HAS DM", lanche: "VITAMINA DM", jantar: "BRANDA DM HAS", ceia: "MINGAU DM + 1 FRUTA", observacao: "" }
    ];
  }
}

// Salva ou edita um paciente na aba da data especificada, para o perfil informado
function salvarPaciente(dataStr, paciente, perfil) {
  try {
    const ss = getSpreadsheet_();
    const sheetName = resolveSheetName_(perfil, dataStr);
    let sheet = ss.getSheetByName(sheetName);

    // Se a aba para esse dia não existir, cria uma nova com cabeçalho
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
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

    if (paciente.rowNumber) {
      // Edita linha existente usando rowNumber
      sheet.getRange(paciente.rowNumber, 1, 1, 10).setValues([rowData]);
      return { success: true, rowNumber: paciente.rowNumber };
    } else {
      // Adiciona nova linha
      const lastRow = sheet.getLastRow();
      const newRowNumber = lastRow + 1;
      sheet.appendRow(rowData);
      return { success: true, rowNumber: newRowNumber };
    }
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// === Administração ===

// Verifica a senha de acesso à área de administração
function verificarSenhaAdmin(senha) {
  return { success: senha === ADMIN_PASSWORD };
}

// Garante que a aba do catálogo de itens de dieta existe.
// O catálogo é uma única lista de itens, compartilhada entre todas as
// refeições (Desjejum, Colação, Almoço, Lanche, Jantar, Ceia) e entre os
// dois perfis — não há divisão por horário ou categoria.
function getItensSheet_(ss) {
  let sheet = ss.getSheetByName('ITENS_DIETA');
  if (!sheet) {
    sheet = ss.insertSheet('ITENS_DIETA');
    sheet.appendRow(['Item']);
  }
  return sheet;
}

// Retorna o catálogo de itens de dieta (lista única, sem categorias)
function getItensDieta() {
  try {
    const ss = getSpreadsheet_();
    const sheet = getItensSheet_(ss);
    const data = sheet.getDataRange().getValues();

    const itens = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) itens.push(data[i][0]);
    }
    return itens;
  } catch(e) {
    // Fallback mock em caso de erro inesperado
    return ['CAFÉ COMPLETO', 'CAFÉ COMPLETO DM', 'MINGAU DE AVEIA', 'SUCO', 'SUCO DM', 'VITAMINA', 'VITAMINA DM', 'GERAL', 'GERAL HAS DM', 'BRANDA', 'BISCOITO', 'MINGAU'];
  }
}

// Adiciona um novo item ao catálogo
function adicionarItemDieta(item) {
  try {
    const ss = getSpreadsheet_();
    const sheet = getItensSheet_(ss);
    sheet.appendRow([item]);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// Remove um item do catálogo
function removerItemDieta(item) {
  try {
    const ss = getSpreadsheet_();
    const sheet = getItensSheet_(ss);
    const data = sheet.getDataRange().getValues();

    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === item) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

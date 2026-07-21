// Google Apps Script - Casa da Gestante / Oncológico

// Senha INICIAL da área de administração. Depois do primeiro uso, a senha pode
// ser alterada pelo próprio painel admin (fica salva em Script Properties e
// esta constante passa a ser ignorada). ALTERE antes de usar em produção.
const ADMIN_PASSWORD = 'admin123';

// Senha vigente: a salva em Script Properties (se o admin já trocou) ou a constante.
function getAdminPassword_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || ADMIN_PASSWORD;
}

// Confere a senha admin em chamadas sensíveis vindas do frontend.
function checarSenhaAdmin_(senha) {
  return String(senha || '') === getAdminPassword_();
}

// Prefixo usado nas abas do perfil Oncológico, para não colidir com as abas
// já existentes do perfil Casa da Gestante (que continuam usando o nome puro da data).
const PREFIXO_ONCOLOGICO = 'ONCO - ';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Casa da Gestante - Controle de Dietas')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Guarda o ID da planilha sempre que o script tem uma planilha ativa disponível
// (menus, onOpen, execução manual pelo editor). Serve de apoio para o caso raro
// de getActiveSpreadsheet() não estar disponível no contexto atual do Web App.
function onOpen() {
  getSpreadsheet_();
}

// Retorna a planilha onde o script está vinculado (a mesma em que foi colado)
function getSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) {
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
    return ss;
  }

  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) {
    return SpreadsheetApp.openById(id);
  }

  throw new Error('Não foi possível identificar a planilha. Verifique se este script foi adicionado pela própria planilha em Extensões > Apps Script.');
}

// Cabeçalho padrão das abas de dia (mesma ordem usada em getPacientesPorData)
const HEADER_ = [
  "Prontuário / Paciente", "Diagnóstico", "Dieta",
  "Desjejum - 6h", "Colação - 9h", "Almoço - 12h",
  "Lanche - 15h", "Jantar - 18h", "Ceia - 21h", "Observação"
];

// Resolve o nome real da aba na planilha para o perfil e data informados
function resolveSheetName_(perfil, dataStr) {
  return perfil === 'ONCOLOGICO' ? (PREFIXO_ONCOLOGICO + dataStr) : dataStr;
}

// Converte "dd/mm/yyyy" em Date (meia-noite local). Retorna null se não bater o formato.
function parseDataBR_(str) {
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(str || '').trim());
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}

// Entre as abas-data do perfil, encontra a data mais recente ESTRITAMENTE anterior
// à data pedida. Retorna a string "dd/mm/yyyy" da aba anterior, ou null.
function encontrarDiaAnterior_(ss, perfil, dataStr) {
  const alvo = parseDataBR_(dataStr);
  if (!alvo) return null;

  const datasDisponiveis = getDatasDisponiveis(perfil); // já filtra por perfil e remove ITENS_DIETA
  let melhor = null, melhorData = null;
  datasDisponiveis.forEach(d => {
    const dt = parseDataBR_(d);
    if (dt && dt.getTime() < alvo.getTime() && (!melhorData || dt.getTime() > melhorData.getTime())) {
      melhor = d;
      melhorData = dt;
    }
  });
  return melhor;
}

// Cria uma aba nova para `sheetName`, com cabeçalho formatado e as linhas de pacientes fornecidas.
function criarAbaComPacientes_(ss, sheetName, pacientes) {
  const sheet = ss.insertSheet(sheetName);
  sheet.appendRow(HEADER_);
  const headerRange = sheet.getRange(1, 1, 1, HEADER_.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#e3f0ea");
  sheet.setFrozenRows(1);
  if (pacientes && pacientes.length) {
    const linhas = pacientes.map(p => [
      p.prontuario_nome, p.diagnostico, p.dieta,
      p.desjejum, p.colacao, p.almoco,
      p.lanche, p.jantar, p.ceia, p.observacao
    ]);
    sheet.getRange(2, 1, linhas.length, HEADER_.length).setValues(linhas);
  }
  return sheet;
}

// Lê os pacientes de uma aba já existente (sem herança). Uso interno.
function lerPacientesDaAba_(sheet) {
  const data = sheet.getDataRange().getValues();
  const pacientes = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue; // Pula linhas vazias
    pacientes.push({
      rowNumber: i + 1,
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
        .map(n => n.substring(PREFIXO_ONCOLOGICO.length))
        .filter(n => parseDataBR_(n) !== null);
    }
    // Gestante: abas sem o prefixo Oncológico cujo nome é uma data válida
    // (exclui automaticamente ITENS_DIETA, SISTEMA_LOGS, ALTAS e afins).
    return nomes.filter(n => n.indexOf(PREFIXO_ONCOLOGICO) !== 0 && parseDataBR_(n) !== null);
  } catch(e) {
    return [];
  }
}

// Retorna os pacientes da aba da data/perfil informados.
//
// Retorna { pacientes, herdadoDe }:
//  - Se a aba do dia existe com dados, `pacientes` são os dados reais e herdadoDe = null.
//  - Se a aba ainda não existe (ou só tem cabeçalho) e a data é hoje ou futura,
//    devolve uma PRÉVIA com os pacientes do último dia registrado anterior
//    (herdadoDe = "dd/mm/yyyy"), SEM gravar nada na planilha. A aba só é criada
//    de fato quando alguém edita/salva/dá alta naquele dia (ver materializarDia).
//    Isso evita o bug de "congelar" uma cópia do dia anterior só por visualizar
//    um dia futuro — paciente que recebe alta hoje não reaparece amanhã.
// Datas passadas nunca mostram herança (evita reescrever histórico).
function getPacientesPorData(dataStr, perfil) {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(resolveSheetName_(perfil, dataStr));

    if (sheet && sheet.getLastRow() >= 2) {
      return { pacientes: lerPacientesDaAba_(sheet), herdadoDe: null };
    }

    if (dataEhHojeOuFutura_(dataStr)) {
      const diaAnterior = encontrarDiaAnterior_(ss, perfil, dataStr);
      if (diaAnterior) {
        const anterior = ss.getSheetByName(resolveSheetName_(perfil, diaAnterior));
        const pacientes = anterior ? lerPacientesDaAba_(anterior) : [];
        // Prévia: sem rowNumber, pois essas linhas ainda não existem nesta aba.
        pacientes.forEach(p => { p.rowNumber = null; });
        if (pacientes.length) return { pacientes: pacientes, herdadoDe: diaAnterior };
      }
    }

    return { pacientes: [], herdadoDe: null };
  } catch(e) {
    return { pacientes: [], herdadoDe: null };
  }
}

// Cria de fato a aba do dia (se ainda não existir com dados), copiando os
// pacientes do último dia anterior NAQUELE MOMENTO — já sem quem recebeu alta.
// Chamado pelo frontend antes da primeira edição/salvamento em um dia que está
// sendo exibido como prévia herdada. Retorna os pacientes reais (com rowNumber).
function materializarDia(dataStr, perfil) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const ss = getSpreadsheet_();
      const sheetName = resolveSheetName_(perfil, dataStr);
      let sheet = ss.getSheetByName(sheetName);

      if (!sheet || sheet.getLastRow() < 2) {
        let herdados = [];
        let diaAnterior = null;
        if (dataEhHojeOuFutura_(dataStr)) {
          diaAnterior = encontrarDiaAnterior_(ss, perfil, dataStr);
          const anterior = diaAnterior ? ss.getSheetByName(resolveSheetName_(perfil, diaAnterior)) : null;
          herdados = anterior ? lerPacientesDaAba_(anterior) : [];
        }
        if (!sheet) {
          sheet = criarAbaComPacientes_(ss, sheetName, herdados);
        } else if (herdados.length) {
          const linhas = herdados.map(p => [
            p.prontuario_nome, p.diagnostico, p.dieta,
            p.desjejum, p.colacao, p.almoco,
            p.lanche, p.jantar, p.ceia, p.observacao
          ]);
          sheet.getRange(2, 1, linhas.length, HEADER_.length).setValues(linhas);
        }
        logEvent('DIA_MATERIALIZADO', perfil, { dataStr: dataStr, copiadoDe: diaAnterior, qtdPacientes: herdados.length });
      }

      return { success: true, pacientes: lerPacientesDaAba_(sheet) };
    } finally {
      lock.releaseLock();
    }
  } catch(e) {
    return { success: false, error: e.toString(), pacientes: [] };
  }
}

// True se dataStr ("dd/mm/yyyy") é hoje ou uma data futura (comparação por dia).
function dataEhHojeOuFutura_(dataStr) {
  const d = parseDataBR_(dataStr);
  if (!d) return false;
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  return d.getTime() >= hoje.getTime();
}

// Salva ou edita um paciente na aba da data especificada, para o perfil informado
function salvarPaciente(dataStr, paciente, perfil) {
  try {
    const ss = getSpreadsheet_();
    const sheetName = resolveSheetName_(perfil, dataStr);
    let sheet = ss.getSheetByName(sheetName);

    // Se a aba para esse dia não existir, cria uma nova com cabeçalho.
    // Usa sempre sheetName (com prefixo do perfil quando Oncológico) — antes
    // era criada com o nome puro da data, misturando os dois perfis.
    if (!sheet) {
      sheet = criarAbaComPacientes_(ss, sheetName, []);
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
      logEvent('PACIENTE_EDITADO', perfil, { dataStr, rowNumber: paciente.rowNumber, paciente });
      return { success: true, rowNumber: paciente.rowNumber };
    } else {
      // Adiciona nova linha
      const lastRow = sheet.getLastRow();
      const newRowNumber = lastRow + 1;
      sheet.appendRow(rowData);
      logEvent('PACIENTE_ADICIONADO', perfil, { dataStr, rowNumber: newRowNumber, paciente });
      return { success: true, rowNumber: newRowNumber };
    }
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// Garante a aba de registro de ALTAS. Fica VISÍVEL na planilha de propósito:
// é o registro permanente e auditável de quem recebeu alta, quando e por quem.
function getAltasSheet_(ss) {
  let sheet = ss.getSheetByName('ALTAS');
  if (!sheet) {
    sheet = ss.insertSheet('ALTAS');
    sheet.appendRow(['Data / Hora da Alta', 'Dia da Lista', 'Perfil', 'Prontuário / Paciente', 'Diagnóstico', 'Dieta', 'Registrado por']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e3f0ea');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Remove (dá "alta") um paciente da aba da data/perfil informados, pelo rowNumber.
// Além de remover da lista do dia (para deixar de ser herdado nos próximos), o
// paciente é registrado permanentemente na aba ALTAS — a alta nunca some sem rastro.
function removerPaciente(dataStr, rowNumber, perfil) {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(resolveSheetName_(perfil, dataStr));
    if (!sheet) return { success: false, error: 'Aba não encontrada.' };

    const linha = Number(rowNumber);
    if (!linha || linha < 2 || linha > sheet.getLastRow()) {
      return { success: false, error: 'Linha inválida.' };
    }
    const valoresRemovidos = sheet.getRange(linha, 1, 1, 10).getValues()[0];

    let usuario = 'Desconhecido';
    try { usuario = Session.getActiveUser().getEmail() || 'Desconhecido'; } catch(e) {}

    // Registra a alta ANTES de apagar a linha: se algo falhar, o pior caso é o
    // paciente continuar na lista — nunca uma alta sem registro.
    getAltasSheet_(ss).appendRow([
      new Date(), dataStr,
      perfil === 'ONCOLOGICO' ? 'Oncológico' : 'Casa da Gestante',
      valoresRemovidos[0], valoresRemovidos[1], valoresRemovidos[2], usuario
    ]);

    sheet.deleteRow(linha);
    logEvent('PACIENTE_REMOVIDO', perfil, { dataStr, rowNumber, pacienteNome: valoresRemovidos[0] });
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// Retorna as últimas `limite` altas registradas (mais recentes primeiro).
function getAltas(limite) {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName('ALTAS');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, altas: [] };

    const max = Math.min(Number(limite) || 100, 300);
    const total = sheet.getLastRow() - 1;
    const qtd = Math.min(total, max);
    const primeira = sheet.getLastRow() - qtd + 1;
    const valores = sheet.getRange(primeira, 1, qtd, 7).getValues();

    const tz = Session.getScriptTimeZone();
    const altas = valores.map(row => ({
      dataHora: row[0] instanceof Date ? Utilities.formatDate(row[0], tz, 'dd/MM/yyyy HH:mm') : String(row[0]),
      diaLista: String(row[1] || ''),
      perfil: String(row[2] || ''),
      paciente: String(row[3] || ''),
      diagnostico: String(row[4] || ''),
      dieta: String(row[5] || ''),
      usuario: String(row[6] || '')
    })).reverse();
    return { success: true, altas: altas };
  } catch(e) {
    return { success: false, error: e.toString(), altas: [] };
  }
}

// === Administração ===

// Verifica a senha de acesso à área de administração
function verificarSenhaAdmin(senha) {
  const sucesso = checarSenhaAdmin_(senha);
  logEvent(sucesso ? 'LOGIN_ADMIN_SUCESSO' : 'LOGIN_ADMIN_FALHA', 'ADMIN', {});
  return { success: sucesso };
}

// Altera a senha de administração (salva em Script Properties).
function alterarSenhaAdmin(senhaAtual, novaSenha) {
  try {
    if (!checarSenhaAdmin_(senhaAtual)) return { success: false, error: 'Senha atual incorreta.' };
    const nova = String(novaSenha || '').trim();
    if (nova.length < 4) return { success: false, error: 'A nova senha precisa ter pelo menos 4 caracteres.' };
    PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', nova);
    logEvent('SENHA_ADMIN_ALTERADA', 'ADMIN', {});
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// Visão geral de todos os dias registrados, nos dois perfis, com contagem de
// pacientes. Usado no painel admin ("ver tudo").
function getResumoDias(senha) {
  try {
    if (!checarSenhaAdmin_(senha)) return { success: false, error: 'Acesso negado.', dias: [] };
    const ss = getSpreadsheet_();
    const dias = [];
    ss.getSheets().forEach(sheet => {
      const nome = sheet.getName();
      if (nome === 'ITENS_DIETA' || nome === 'SISTEMA_LOGS') return;
      const ehOnco = nome.indexOf(PREFIXO_ONCOLOGICO) === 0;
      const dataStr = ehOnco ? nome.substring(PREFIXO_ONCOLOGICO.length) : nome;
      if (!parseDataBR_(dataStr)) return; // ignora abas que não são de dia
      dias.push({
        data: dataStr,
        perfil: ehOnco ? 'ONCOLOGICO' : 'GESTANTE',
        qtdPacientes: Math.max(0, sheet.getLastRow() - 1)
      });
    });
    return { success: true, dias: dias };
  } catch(e) {
    return { success: false, error: e.toString(), dias: [] };
  }
}

// Exclui a aba de um dia inteiro (ex: dia criado por engano). Só via admin.
function removerDia(dataStr, perfil, senha) {
  try {
    if (!checarSenhaAdmin_(senha)) return { success: false, error: 'Acesso negado.' };
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(resolveSheetName_(perfil, dataStr));
    if (!sheet) return { success: false, error: 'Aba não encontrada.' };
    const qtd = Math.max(0, sheet.getLastRow() - 1);
    ss.deleteSheet(sheet);
    logEvent('DIA_REMOVIDO', perfil, { dataStr: dataStr, qtdPacientes: qtd });
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// Retorna os últimos `limite` eventos do log do sistema (mais recentes primeiro).
// Usado no painel admin para auditoria: quem mexeu, quando e no quê.
function getLogs(senha, limite) {
  try {
    if (!checarSenhaAdmin_(senha)) return { success: false, error: 'Acesso negado.', logs: [] };
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName('SISTEMA_LOGS');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, logs: [] };

    const max = Math.min(Number(limite) || 100, 300);
    const total = sheet.getLastRow() - 1;
    const qtd = Math.min(total, max);
    const primeira = sheet.getLastRow() - qtd + 1;
    const valores = sheet.getRange(primeira, 1, qtd, 5).getValues();

    const tz = Session.getScriptTimeZone();
    const logs = valores.map(row => ({
      dataHora: row[0] instanceof Date ? Utilities.formatDate(row[0], tz, 'dd/MM/yyyy HH:mm') : String(row[0]),
      usuario: String(row[1] || ''),
      acao: String(row[2] || ''),
      perfil: String(row[3] || ''),
      detalhes: String(row[4] || '')
    })).reverse();
    return { success: true, logs: logs };
  } catch(e) {
    return { success: false, error: e.toString(), logs: [] };
  }
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
    return [];
  }
}

// Adiciona um novo item ao catálogo
function adicionarItemDieta(item) {
  try {
    const ss = getSpreadsheet_();
    const sheet = getItensSheet_(ss);
    sheet.appendRow([item]);
    logEvent('ITEM_DIETA_ADICIONADO', 'ADMIN', { item });
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
        logEvent('ITEM_DIETA_REMOVIDO', 'ADMIN', { item });
        break;
      }
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// === Rastreabilidade (Logs) ===
function logEvent(acao, perfil, detalhesObj) {
  try {
    const ss = getSpreadsheet_();
    let sheet = ss.getSheetByName('SISTEMA_LOGS');
    if (!sheet) {
      sheet = ss.insertSheet('SISTEMA_LOGS');
      sheet.appendRow(["Data / Hora", "Usuário", "Ação", "Perfil", "Detalhes"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#333333").setFontColor("#ffffff");
      sheet.setFrozenRows(1);
      sheet.hideSheet(); // Oculta a aba para evitar que enfermeiros apaguem sem querer
    }
    
    let usuario = "Anônimo";
    try {
      usuario = Session.getActiveUser().getEmail() || "Desconhecido";
    } catch(e) {}
    
    const detalhes = detalhesObj ? JSON.stringify(detalhesObj) : "";
    
    sheet.appendRow([
      new Date(),
      usuario,
      acao,
      perfil || "-",
      detalhes
    ]);
  } catch(e) {
    // Falha silenciosa nos logs para não quebrar o fluxo principal
  }
}

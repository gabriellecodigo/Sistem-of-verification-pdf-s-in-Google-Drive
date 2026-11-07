/**
 * Verificador genérico de arquivos do Google Drive para Google Sheets.
 * Usa DriveApp (serviço integrado do Apps Script), sem Drive API avançada.
 * Os dados de origem nunca são alterados; cada execução cria uma aba de relatório.
 */
var DV_CONFIG = {
  MENU_NAME: "Verificador Drive",
  RESULT_PREFIX: "DV_RESULTADO_",
  HEADER_ROW: 1,
  DEFAULT_FIRST_DATA_ROW: 2,
  MAX_LOCK_WAIT_MS: 30000,
  COLORS: {
    header: "#1F4E78",
    found: "#D9EAD3",
    missing: "#F4CCCC",
    duplicate: "#FCE5CD",
    skipped: "#E7E6E6"
  },
  HEADERS: [
    "Status",
    "Planilha de origem",
    "Linha de origem",
    "Coluna de origem",
    "Valor da célula",
    "Chave comparada",
    "Arquivo encontrado",
    "Link do arquivo",
    "Pasta encontrada",
    "Link da pasta",
    "Quantidade de arquivos",
    "Tratamento da duplicidade",
    "_DV_GROUP_ID",
    "_DV_FILE_ID"
  ]
};

/** Adiciona o menu do verificador ao abrir a planilha. */
function onOpen() {
  DV_addMenu_();
}

function DV_addMenu_() {
  SpreadsheetApp.getUi()
    .createMenu(DV_CONFIG.MENU_NAME)
    .addItem("Abrir painel", "DV_showSidebar")
    .addToUi();
}

function DV_showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("DriveVerifierSidebar")
    .setTitle("Verificador de arquivos");
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Retorna as abas disponíveis e a aba ativa para montar os seletores do painel. */
function DV_getInitialData() {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error("Abra uma planilha Google Sheets para usar o verificador.");
  return {
    activeSheetId: book.getActiveSheet().getSheetId(),
    sheets: book.getSheets().map(function (sheet) {
      return { id: sheet.getSheetId(), name: sheet.getName() };
    })
  };
}

/** Retorna cabeçalhos e índices de coluna da aba escolhida. */
function DV_getSheetColumns(sheetId) {
  var sheet = DV_getSourceSheet_(sheetId);
  var lastColumn = sheet.getLastColumn();
  if (!lastColumn) return { sheetName: sheet.getName(), columns: [] };
  var headers = sheet.getRange(DV_CONFIG.HEADER_ROW, 1, 1, lastColumn).getDisplayValues()[0];
  return {
    sheetName: sheet.getName(),
    columns: headers.map(function (header, index) {
      return {
        index: index + 1,
        letter: DV_columnLetter_(index + 1),
        header: String(header || "").trim() || "Coluna " + DV_columnLetter_(index + 1)
      };
    })
  };
}

/** Faz a leitura completa do Drive, compara as colunas escolhidas e cria um relatório separado. */
function DV_verify(config) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(DV_CONFIG.MAX_LOCK_WAIT_MS)) {
    throw new Error("A planilha está sendo atualizada por outro processo. Tente novamente em instantes.");
  }

  var resultSheet = null;
  try {
    var book = SpreadsheetApp.getActiveSpreadsheet();
    if (!book) throw new Error("Abra uma planilha Google Sheets para usar o verificador.");

    var options = DV_validateConfig_(config);
    var sourceSheet = DV_getSourceSheet_(options.sheetId);
    var sourceLastColumn = sourceSheet.getLastColumn();
    var requestedColumns = options.sourceColumns.slice();
    if (options.conditionColumn) requestedColumns.push(options.conditionColumn);
    if (requestedColumns.some(function (column) { return column > sourceLastColumn; })) {
      throw new Error("Uma das colunas escolhidas não existe na aba selecionada.");
    }
    var source = DV_readSource_(sourceSheet, options);
    if (!source.records.length) {
      throw new Error("Não encontrei valores preenchidos nas colunas selecionadas a partir da linha " + options.firstDataRow + ".");
    }

    var folder = DV_getFolder_(options.folderReference);
    var matchesByGroup = DV_findDriveMatches_(folder, source.groups, options);
    var prepared = DV_buildReportRows_(source.records, source.groups, matchesByGroup, sourceSheet);
    var resultName = DV_makeResultSheetName_();

    try {
      resultSheet = book.insertSheet(resultName);
      DV_writeReport_(resultSheet, prepared.rows);
    } catch (writeError) {
      if (resultSheet) book.deleteSheet(resultSheet);
      resultSheet = null;
      throw writeError;
    }

    return {
      resultSheetId: resultSheet.getSheetId(),
      resultSheetName: resultSheet.getName(),
      summary: prepared.summary,
      duplicateGroups: prepared.duplicateGroups
    };
  } catch (error) {
    throw new Error(error && error.message ? error.message : String(error));
  } finally {
    lock.releaseLock();
  }
}

/** Lê somente as colunas selecionadas e a coluna condicional, em blocos contíguos. */
function DV_readSource_(sheet, options) {
  var lastRow = sheet.getLastRow();
  var columnsToRead = options.sourceColumns.slice();
  if (options.conditionColumn) columnsToRead.push(options.conditionColumn);
  if (lastRow < options.firstDataRow) return { records: [], groups: {} };

  var requested = columnsToRead.filter(function (column, index, all) { return all.indexOf(column) === index; })
    .sort(function (left, right) { return left - right; });
  var valuesByColumn = {};
  var headersByColumn = {};
  var rangeStart = 0;
  while (rangeStart < requested.length) {
    var rangeEnd = rangeStart + 1;
    while (rangeEnd < requested.length && requested[rangeEnd] === requested[rangeEnd - 1] + 1) rangeEnd++;
    var firstColumn = requested[rangeStart];
    var width = rangeEnd - rangeStart;
    var block = sheet.getRange(DV_CONFIG.HEADER_ROW, firstColumn, lastRow - DV_CONFIG.HEADER_ROW + 1, width).getDisplayValues();
    for (var offset = 0; offset < width; offset++) {
      var column = firstColumn + offset;
      headersByColumn[column] = String(block[0][offset] || "").trim() || "Coluna " + DV_columnLetter_(column);
      valuesByColumn[column] = [];
      for (var dataIndex = options.firstDataRow - DV_CONFIG.HEADER_ROW; dataIndex < block.length; dataIndex++) {
        valuesByColumn[column].push(block[dataIndex][offset]);
      }
    }
    rangeStart = rangeEnd;
  }

  var groups = {};
  var records = [];

  options.sourceColumns.forEach(function (column) {
    var header = headersByColumn[column];
    for (var i = 0; i < valuesByColumn[column].length; i++) {
      var value = String(valuesByColumn[column][i] || "").trim();
      if (!value) continue;
      var sourceRow = options.firstDataRow + i;
      var conditionValue = options.conditionColumn
        ? String(valuesByColumn[options.conditionColumn][i] || "").trim()
        : "";
      var skipped = options.skipValues.length &&
        options.skipValues.indexOf(DV_normalize_(conditionValue, options.caseSensitive)) !== -1;
      var record = {
        row: sourceRow,
        column: column,
        columnHeader: header,
        value: value,
        conditionValue: conditionValue,
        skipped: Boolean(skipped),
        groupId: ""
      };

      if (!record.skipped) {
        var normalized = DV_normalize_(value, options.caseSensitive);
        var groupId = JSON.stringify([sheet.getSheetId(), column, normalized]);
        record.groupId = groupId;
        if (!groups[groupId]) {
          groups[groupId] = {
            id: groupId,
            column: column,
            columnHeader: header,
            value: value,
            normalizedValue: normalized,
            filesById: {}
          };
        }
      }
      records.push(record);
    }
  });
  return { records: records, groups: groups };
}

/** Percorre a pasta e subpastas e guarda metadados apenas dos arquivos correspondentes. */
function DV_findDriveMatches_(rootFolder, groups, options) {
  var matches = {};
  var groupList = Object.keys(groups).map(function (id) { return groups[id]; });
  var fileNameRegex = options.fileNameRegex
    ? new RegExp(options.fileNameRegex, options.caseSensitive ? "" : "i")
    : null;
  var exactGroups = {};
  if (options.matchMode === "exact") {
    groupList.forEach(function (group) {
      if (!exactGroups[group.normalizedValue]) exactGroups[group.normalizedValue] = [];
      exactGroups[group.normalizedValue].push(group);
    });
  }

  function makeFolderNode_(folder, parentPath) {
    var id = folder.getId();
    var name = folder.getName();
    return {
      folder: folder,
      id: id,
      path: parentPath ? parentPath + " / " + name : name,
      url: "https://drive.google.com/drive/folders/" + id
    };
  }

  var stack = [makeFolderNode_(rootFolder, "")];
  var visited = {};
  while (stack.length) {
    var node = stack.pop();
    if (visited[node.id]) continue;
    visited[node.id] = true;

    var files = node.folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      if (!DV_extensionAllowed_(name, options.extensions)) continue;
      var fileKey = DV_extractFileKey_(name, fileNameRegex);
      if (!fileKey) continue;
      var normalizedFileKey = DV_normalize_(fileKey, options.caseSensitive);
      var candidateGroups = options.matchMode === "exact"
        ? (exactGroups[normalizedFileKey] || [])
        : groupList.filter(function (group) {
          return options.matchMode === "startsWith"
            ? normalizedFileKey.indexOf(group.normalizedValue) === 0
            : normalizedFileKey.indexOf(group.normalizedValue) !== -1;
        });
      if (!candidateGroups.length) continue;

      var id = file.getId();
      var entry = {
        id: id,
        name: name,
        key: fileKey,
        url: file.getUrl(),
        path: node.path,
        folderUrl: node.url
      };
      candidateGroups.forEach(function (group) {
        if (!group.filesById[id]) group.filesById[id] = entry;
      });
    }

    var folders = node.folder.getFolders();
    while (folders.hasNext()) stack.push(makeFolderNode_(folders.next(), node.path));
  }

  Object.keys(groups).forEach(function (groupId) {
    matches[groupId] = Object.keys(groups[groupId].filesById).map(function (fileId) {
      return groups[groupId].filesById[fileId];
    }).sort(function (left, right) {
      return left.name.localeCompare(right.name) || left.id.localeCompare(right.id);
    });
  });
  return matches;
}

/** Monta linhas detalhadas e estatísticas sem escrever na aba de origem. */
function DV_buildReportRows_(records, groups, matchesByGroup, sourceSheet) {
  var rows = [];
  var duplicateGroups = [];
  var stats = { checked: 0, found: 0, missing: 0, duplicated: 0, skipped: 0, filesMatched: 0 };

  Object.keys(groups).forEach(function (groupId) {
    var group = groups[groupId];
    var files = matchesByGroup[groupId] || [];
    group.files = files;
    if (files.length > 1) {
      duplicateGroups.push({
        id: groupId,
        columnHeader: group.columnHeader,
        value: group.value,
        count: files.length
      });
    }
  });

  records.forEach(function (record) {
    if (record.skipped) {
      stats.skipped++;
      rows.push(DV_makeReportRow_("Ignorado pela condição", sourceSheet.getName(), record, "", null, 0, "", ""));
      return;
    }

    stats.checked++;
    var group = groups[record.groupId];
    var files = group.files || [];
    if (!files.length) {
      stats.missing++;
      rows.push(DV_makeReportRow_("Não encontrado", sourceSheet.getName(), record,
        record.value, null, 0, "", ""));
      return;
    }

    if (files.length > 1) stats.duplicated++;
    else stats.found++;
    stats.filesMatched += files.length;
    files.forEach(function (file) {
      rows.push(DV_makeReportRow_(files.length > 1 ? "Duplicado" : "Encontrado", sourceSheet.getName(), record,
        file.key, file, files.length, files.length > 1 ? "Pendente" : "", files.length > 1 ? group.id : ""));
    });
  });

  duplicateGroups.sort(function (left, right) {
    return left.columnHeader.localeCompare(right.columnHeader) || left.value.localeCompare(right.value);
  });
  return {
    rows: rows,
    duplicateGroups: duplicateGroups,
    summary: {
      sheet: sourceSheet.getName(),
      valuesChecked: stats.checked,
      valuesFound: stats.found,
      valuesMissing: stats.missing,
      duplicateValues: stats.duplicated,
      filesMatched: stats.filesMatched,
      skippedByCondition: stats.skipped,
      duplicateGroups: duplicateGroups.length,
      reportRows: rows.length
    }
  };
}

function DV_makeReportRow_(status, sourceSheetName, record, comparisonKey, file, fileCount, decision, groupId) {
  return {
    values: [
      status,
      sourceSheetName,
      record.row,
      record.columnHeader + " (" + DV_columnLetter_(record.column) + ")",
      record.value,
      comparisonKey,
      file ? file.name : "",
      file ? "Abrir arquivo" : "",
      file ? file.path : "",
      file ? "Abrir pasta" : "",
      fileCount,
      decision,
      groupId,
      file ? file.id : ""
    ],
    fileUrl: file ? file.url : "",
    folderUrl: file ? file.folderUrl : "",
    status: status
  };
}

/** Escreve a aba de relatório em blocos e mantém as colunas internas ocultas. */
function DV_writeReport_(sheet, reportRows) {
  sheet.getRange(1, 1, 1, DV_CONFIG.HEADERS.length)
    .setValues([DV_CONFIG.HEADERS])
    .setBackground(DV_CONFIG.COLORS.header)
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1).setNote(
    "Relatório criado pelo Verificador Drive. Os dados da aba de origem não foram alterados."
  );
  if (!reportRows.length) return;

  var values = reportRows.map(function (row) { return row.values; });
  sheet.getRange(2, 1, values.length, DV_CONFIG.HEADERS.length).setValues(values);
  sheet.getRange(2, 8, reportRows.length, 1).setRichTextValues(reportRows.map(function (row) {
    return [DV_makeLinkText_(row.values[7], row.fileUrl)];
  }));
  sheet.getRange(2, 10, reportRows.length, 1).setRichTextValues(reportRows.map(function (row) {
    return [DV_makeLinkText_(row.values[9], row.folderUrl)];
  }));

  var statusBackgrounds = reportRows.map(function (row) {
    var color = row.status === "Encontrado" ? DV_CONFIG.COLORS.found :
      row.status === "Não encontrado" ? DV_CONFIG.COLORS.missing :
      row.status === "Duplicado" ? DV_CONFIG.COLORS.duplicate : DV_CONFIG.COLORS.skipped;
    return [color];
  });
  sheet.getRange(2, 1, reportRows.length, 1).setBackgrounds(statusBackgrounds);
  sheet.getRange(1, 1, reportRows.length + 1, 12).createFilter();
  sheet.hideColumns(13, 2);
  sheet.autoResizeColumns(1, 12);
  sheet.setColumnWidth(5, 220);
  sheet.setColumnWidth(7, 240);
  sheet.setColumnWidth(9, 260);
  sheet.setColumnWidth(12, 190);
}

function DV_makeLinkText_(text, url) {
  var builder = SpreadsheetApp.newRichTextValue().setText(String(text || ""));
  if (url && text) builder.setLinkUrl(String(url));
  return builder.build();
}

/** Busca detalhes de uma duplicidade selecionada no painel lateral. */
function DV_getDuplicateDetails(resultSheetId, groupId) {
  var sheet = DV_getResultSheet_(resultSheetId);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("O relatório não contém linhas para revisar.");
  var values = sheet.getRange(2, 7, lastRow - 1, 8).getDisplayValues(); // G:N
  var files = [];
  var seenFileIds = {};
  values.forEach(function (row) {
    if (row[6] !== groupId) return; // M
    if (seenFileIds[row[7]]) return;
    seenFileIds[row[7]] = true;
    files.push({ id: row[7], name: row[0], path: row[2] }); // N, G, I
  });
  if (files.length < 2) throw new Error("A duplicidade mudou ou não está mais no relatório.");
  return { id: groupId, files: files };
}

/** Aplica as decisões de duplicidade somente na coluna de decisão do relatório. */
function DV_applyDuplicateDecisions(resultSheetId, decisions) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(DV_CONFIG.MAX_LOCK_WAIT_MS)) {
    throw new Error("A planilha está sendo atualizada por outro processo. Tente novamente em instantes.");
  }
  try {
    var sheet = DV_getResultSheet_(resultSheetId);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { updatedGroups: 0 };

    var range = sheet.getRange(2, 12, lastRow - 1, 3); // L:N
    var data = range.getDisplayValues();
    var choices = {};
    (decisions || []).forEach(function (decision) {
      if (!decision || !decision.groupId || choices[decision.groupId]) {
        throw new Error("Há uma decisão de duplicidade ausente ou repetida.");
      }
      if (decision.action !== "KEEP_ALL" && decision.action !== "USE_FILE") {
        throw new Error("Uma decisão de duplicidade não é válida.");
      }
      choices[decision.groupId] = decision;
    });

    var filesByGroup = {};
    data.forEach(function (row) {
      var groupId = row[1];
      if (!groupId) return;
      if (!filesByGroup[groupId]) filesByGroup[groupId] = {};
      filesByGroup[groupId][row[2]] = true;
    });
    Object.keys(filesByGroup).forEach(function (groupId) {
      var decision = choices[groupId];
      if (!decision) throw new Error("Falta decidir como tratar uma duplicidade do relatório.");
      if (decision.action === "USE_FILE" && !filesByGroup[groupId][String(decision.fileId || "")]) {
        throw new Error("O arquivo escolhido não pertence à duplicidade selecionada.");
      }
    });
    Object.keys(choices).forEach(function (groupId) {
      if (!filesByGroup[groupId]) throw new Error("Uma duplicidade mudou desde a leitura do relatório.");
    });

    data.forEach(function (row) {
      var decision = choices[row[1]];
      if (!decision) return;
      row[0] = decision.action === "KEEP_ALL"
        ? "Manter todos"
        : (row[2] === String(decision.fileId) ? "Arquivo selecionado" : "Alternativa não selecionada");
    });
    range.setValues(data);
    return { updatedGroups: Object.keys(choices).length };
  } finally {
    lock.releaseLock();
  }
}

function DV_getResultSheet_(sheetId) {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = book.getSheetById(Number(sheetId));
  if (!sheet || sheet.getName().indexOf(DV_CONFIG.RESULT_PREFIX) !== 0) {
    throw new Error("A aba de relatório não foi encontrada ou não pertence ao Verificador Drive.");
  }
  return sheet;
}

function DV_activateResultSheet(resultSheetId) {
  var sheet = DV_getResultSheet_(resultSheetId);
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  return { name: sheet.getName() };
}

function DV_validateConfig_(config) {
  if (!config || typeof config !== "object") throw new Error("As opções de verificação estão incompletas.");
  var folderReference = String(config.folderReference || "").trim();
  if (!folderReference) throw new Error("Informe o link ou ID da pasta do Drive.");
  var sheetId = Number(config.sheetId);
  if (!isFinite(sheetId)) throw new Error("Escolha uma aba de origem válida.");
  var sourceColumns = (config.sourceColumns || []).map(Number).filter(function (column, index, all) {
    return Number.isInteger(column) && column > 0 && all.indexOf(column) === index;
  });
  if (!sourceColumns.length) throw new Error("Selecione pelo menos uma coluna para comparar.");
  var firstDataRow = Number(config.firstDataRow || DV_CONFIG.DEFAULT_FIRST_DATA_ROW);
  if (!Number.isInteger(firstDataRow) || firstDataRow < 2) throw new Error("A primeira linha de dados deve ser 2 ou maior.");
  var matchMode = String(config.matchMode || "exact");
  if (["exact", "startsWith", "contains"].indexOf(matchMode) === -1) {
    throw new Error("Escolha um modo de comparação válido.");
  }
  var fileNameRegex = String(config.fileNameRegex || "").trim();
  if (fileNameRegex.length > 200) throw new Error("O padrão de nome deve ter no máximo 200 caracteres.");
  if (fileNameRegex) {
    try {
      new RegExp(fileNameRegex, config.caseSensitive ? "" : "i");
    } catch (error) {
      throw new Error("O padrão de nome não é uma expressão regular válida: " + error.message);
    }
    if (!DV_regexHasCaptureGroup_(fileNameRegex)) {
      throw new Error("O padrão precisa ter um grupo de captura entre parênteses, por exemplo ^(.+?)_manifesto.*$.");
    }
  }
  var conditionColumn = config.conditionColumn ? Number(config.conditionColumn) : 0;
  if (conditionColumn && (!Number.isInteger(conditionColumn) || conditionColumn < 1)) {
    throw new Error("Escolha uma coluna condicional válida.");
  }
  var skipValues = String(config.skipValues || "").split(/[\n,;]+/).map(function (value) {
    return DV_normalize_(value, Boolean(config.caseSensitive));
  }).filter(Boolean);
  if (conditionColumn && !skipValues.length) {
    throw new Error("Informe pelo menos um valor para ignorar na coluna condicional.");
  }
  var extensions = String(config.extensions || "").split(/[\n,;]+/).map(function (value) {
    return value.trim().toLowerCase().replace(/^\*?\./, "");
  }).filter(Boolean);
  return {
    folderReference: folderReference,
    sheetId: sheetId,
    sourceColumns: sourceColumns,
    firstDataRow: firstDataRow,
    matchMode: matchMode,
    fileNameRegex: fileNameRegex,
    conditionColumn: conditionColumn,
    skipValues: skipValues,
    extensions: extensions,
    caseSensitive: Boolean(config.caseSensitive)
  };
}

function DV_getSourceSheet_(sheetId) {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = book.getSheetById(Number(sheetId));
  if (!sheet) throw new Error("A aba de origem selecionada não existe.");
  return sheet;
}

function DV_getFolder_(reference) {
  var match = reference.match(/\/folders\/([A-Za-z0-9_-]+)/) ||
    reference.match(/[?&]id=([A-Za-z0-9_-]+)/) ||
    reference.match(/^([A-Za-z0-9_-]+)$/);
  if (!match) throw new Error("O link da pasta não parece válido. Cole o link do Drive ou somente o ID.");
  try {
    var resourceKeyMatch = reference.match(/[?&]resourcekey=([^&#]+)/i);
    if (resourceKeyMatch && DriveApp.getFolderByIdAndResourceKey) {
      return DriveApp.getFolderByIdAndResourceKey(match[1], decodeURIComponent(resourceKeyMatch[1]));
    }
    return DriveApp.getFolderById(match[1]);
  } catch (error) {
    throw new Error("Não foi possível abrir essa pasta. Confirme o acesso da conta e autorize o serviço Drive.");
  }
}

function DV_extractFileKey_(fileName, expression) {
  if (expression) {
    var match = String(fileName || "").match(expression);
    return match && match[1] !== undefined ? String(match[1]).trim() : "";
  }
  var name = String(fileName || "").trim();
  var dot = name.lastIndexOf(".");
  return dot > 0 ? name.slice(0, dot).trim() : name;
}

function DV_regexHasCaptureGroup_(pattern) {
  var escaped = false;
  var inCharacterClass = false;
  for (var i = 0; i < pattern.length; i++) {
    var character = pattern.charAt(i);
    if (escaped) {
      escaped = false;
      continue;
    }
    if (character === "\\") {
      escaped = true;
      continue;
    }
    if (character === "[") {
      inCharacterClass = true;
      continue;
    }
    if (character === "]" && inCharacterClass) {
      inCharacterClass = false;
      continue;
    }
    if (character !== "(" || inCharacterClass) continue;
    if (pattern.charAt(i + 1) !== "?") return true;
    if (pattern.charAt(i + 2) === "<" && ["=", "!"].indexOf(pattern.charAt(i + 3)) === -1) return true;
  }
  return false;
}

function DV_extensionAllowed_(fileName, extensions) {
  if (!extensions.length) return true;
  var match = String(fileName || "").match(/\.([^.]*)$/);
  if (!match) return false;
  return extensions.indexOf(match[1].toLowerCase()) !== -1;
}

function DV_normalize_(value, caseSensitive) {
  var normalized = String(value || "").trim().replace(/\s+/g, " ");
  return caseSensitive ? normalized : normalized.toLowerCase();
}

function DV_columnLetter_(column) {
  var value = Number(column);
  var letter = "";
  while (value > 0) {
    var remainder = (value - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    value = Math.floor((value - 1) / 26);
  }
  return letter;
}

function DV_makeResultSheetName_() {
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  return DV_CONFIG.RESULT_PREFIX + timestamp + "_" + Utilities.getUuid().slice(0, 6);
}

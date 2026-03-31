/**
 * Exporta mensalmente a aba "inspecoes" em .xlsx e .csv,
 * salva os arquivos na pasta do Drive indicada em colaboradores!J2
 * e limpa os dados da aba (mantendo o cabeçalho) após sucesso.
 */
function exportMonthlyInspectionsBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inspectionsSheet = getRequiredSheet_(SHEETS.INSPECOES);
  var folder = getBackupFolderFromCollaborators_();
  var now = new Date();
  var month = Utilities.formatDate(now, 'Etc/GMT-3', 'MM');
  var year = Utilities.formatDate(now, 'Etc/GMT-3', 'yyyy');
  var baseName = 'inspecoes-' + month + '.' + year;

  var tempSpreadsheet = SpreadsheetApp.create(baseName + '-tmp');

  try {
    var tempSheet = tempSpreadsheet.getSheets()[0];
    tempSheet.setName(SHEETS.INSPECOES);

    var lastRow = inspectionsSheet.getLastRow();
    var lastCol = inspectionsSheet.getLastColumn();
    if (lastCol < 1) {
      throw new Error('A aba "inspecoes" não possui colunas para exportação.');
    }

    var values = inspectionsSheet.getRange(1, 1, Math.max(lastRow, 1), lastCol).getValues();
    tempSheet.getRange(1, 1, values.length, values[0].length).setValues(values);

    var tempId = tempSpreadsheet.getId();
    var xlsxBlob = exportSpreadsheetBlob_(tempId, 'xlsx', baseName + '.xlsx');
    var csvBlob = exportSpreadsheetBlob_(tempId, 'csv', baseName + '.csv');

    folder.createFile(xlsxBlob);
    folder.createFile(csvBlob);

    clearInspectionsDataRows_();

    return {
      ok: true,
      files: [baseName + '.xlsx', baseName + '.csv'],
      folderId: folder.getId()
    };
  } finally {
    DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
  }
}

/**
 * Cria (ou recria) o gatilho mensal para executar no dia 1 às 03:00 (GMT+3).
 */
function createMonthlyInspectionsBackupTrigger() {
  var functionName = 'exportMonthlyInspectionsBackup';
  var timezone = 'Etc/GMT-3';

  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .onMonthDay(1)
    .atHour(3)
    .inTimezone(timezone)
    .create();

  return {
    ok: true,
    functionName: functionName,
    schedule: 'Mensal, dia 1 às 03:00 (GMT+3)',
    timezone: timezone
  };
}

function getBackupFolderFromCollaborators_() {
  var collaboratorsSheet = getRequiredSheet_(SHEETS.COLABORADORES);
  var folderLink = String(collaboratorsSheet.getRange('J2').getValue() || '').trim();

  if (!folderLink) {
    throw new Error('Link da pasta de backup não encontrado em colaboradores!J2.');
  }

  var folderIdMatch = folderLink.match(/[-\w]{25,}/);
  if (!folderIdMatch) {
    throw new Error('Não foi possível extrair o ID da pasta a partir de colaboradores!J2.');
  }

  return DriveApp.getFolderById(folderIdMatch[0]);
}

function exportSpreadsheetBlob_(spreadsheetId, format, filename) {
  var exportUrl = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=' + encodeURIComponent(format);
  var response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: false
  });

  return response.getBlob().setName(filename);
}

function clearInspectionsDataRows_() {
  var inspectionsSheet = getRequiredSheet_(SHEETS.INSPECOES);
  var lastRow = inspectionsSheet.getLastRow();

  if (lastRow <= 1) {
    return;
  }

  inspectionsSheet.getRange(2, 1, lastRow - 1, inspectionsSheet.getLastColumn()).clearContent();
}

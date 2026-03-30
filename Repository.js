var SHEETS = {
  INSPECOES: 'inspecoes',
  COLABORADORES: 'colaboradores',
  CAD_DEFEITOS: 'cad_defeitos',
  CAD_ORIGENS: 'cad_origens',
  VIEW_RELACOES_ATIVAS: 'view_relacoes_ativas',
  VIEW_POSICOES_ATIVAS: 'view_posicoes_ativas',
  VIEW_DEFEITOS_ATIVOS: 'view_defeitos_ativos'
};

/**
 * Lê catálogo com filtro de ativos e retorno textual.
 */
function getActiveCatalogValues_(sheetName, valueColumnIndex, activeColumnIndex) {
  var sheet = getRequiredSheet_(sheetName);
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  var values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  return values
    .filter(function (row) {
      return isActiveFlag_(row[activeColumnIndex - 1]);
    })
    .map(function (row) {
      return String(row[valueColumnIndex - 1]).trim();
    })
    .filter(function (value) {
      return value.length > 0;
    });
}

/**
 * Retorna colaboradores ativos para seleção de equipe.
 */
function getActiveCollaborators_() {
  var sheet = getRequiredSheet_(SHEETS.COLABORADORES);
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

  return values
    .filter(function (row) {
      return isActiveFlag_(row[2]);
    })
    .map(function (row) {
      return {
        id: String(row[0]).trim(),
        name: String(row[1]).trim()
      };
    })
    .filter(function (c) {
      return c.id && c.name;
    });
}

function getCollaboratorsForControlList_() {
  var sheet = getRequiredSheet_(SHEETS.COLABORADORES);
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

  return values
    .map(function (row) {
      return {
        id: String(row[0] || '').trim(),
        name: String(row[1] || '').trim(),
        active: isActiveFlag_(row[2]),
        isEditing: false,
        isNew: false
      };
    })
    .filter(function (row) {
      return row.id || row.name;
    });
}

function saveCollaboratorsControlList_(rows) {
  var sheet = getRequiredSheet_(SHEETS.COLABORADORES);
  var lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  }

  if (!rows.length) {
    return;
  }

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
}

/**
 * Lê matriz de defeitos x posições da aba cad_defeitos.
 * Formato esperado:
 * - Linha 1 (a partir da coluna B): posições
 * - Coluna A (a partir da linha 2): defeitos
 * - Interseção: "x" para relacionamento ativo
 */
function getDefectsByPositionCatalog_() {
  var viewCatalog = getDefectsByPositionFromViews_();
  if (viewCatalog) {
    return viewCatalog;
  }

  var sheet = getRequiredSheet_(SHEETS.CAD_DEFEITOS);
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 2) {
    return {
      posicoes: [],
      defeitos: [],
      defeitosPorPosicao: {},
      paresAtivos: {}
    };
  }

  var values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  var header = values[0];

  var posicoes = [];
  var defectsByPosition = {};
  var activePairs = {};

  for (var columnIndex = 1; columnIndex < header.length; columnIndex += 1) {
    var posicao = String(header[columnIndex] || '').trim();
    if (!posicao) {
      continue;
    }
    posicoes.push(posicao);
    defectsByPosition[posicao] = [];
  }

  var defectSet = {};

  for (var rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    var row = values[rowIndex];
    var defeito = String(row[0] || '').trim();
    if (!defeito) {
      continue;
    }

    defectSet[defeito] = true;

    for (var matrixColumn = 1; matrixColumn < header.length; matrixColumn += 1) {
      var headerPosicao = String(header[matrixColumn] || '').trim();
      if (!headerPosicao || !defectsByPosition[headerPosicao]) {
        continue;
      }

      if (isActiveMatrixFlag_(row[matrixColumn])) {
        defectsByPosition[headerPosicao].push(defeito);
        activePairs[getPositionDefectKey_(headerPosicao, defeito)] = true;
      }
    }
  }

  return {
    posicoes: posicoes,
    defeitos: Object.keys(defectSet),
    defeitosPorPosicao: defectsByPosition,
    paresAtivos: activePairs
  };
}

function getDefectsByPositionFromViews_() {
  var relationsSheet = getOptionalSheet_(SHEETS.VIEW_RELACOES_ATIVAS);
  if (!relationsSheet || relationsSheet.getLastRow() < 2) {
    return null;
  }

  var values = relationsSheet.getRange(2, 1, relationsSheet.getLastRow() - 1, Math.max(relationsSheet.getLastColumn(), 2)).getValues();
  var defectsByPosition = {};
  var activePairs = {};
  var defectSet = {};
  var hasAtLeastOnePair = false;

  values.forEach(function (row) {
    var posicao = String(row[0] || '').trim();
    var defeito = String(row[1] || '').trim();
    var status = row.length > 2 ? row[2] : 'x';

    if (!posicao || !defeito || !isActiveMatrixFlag_(status)) {
      return;
    }

    if (!defectsByPosition[posicao]) {
      defectsByPosition[posicao] = [];
    }
    defectsByPosition[posicao].push(defeito);
    activePairs[getPositionDefectKey_(posicao, defeito)] = true;
    defectSet[defeito] = true;
    hasAtLeastOnePair = true;
  });

  if (!hasAtLeastOnePair) {
    return null;
  }

  var posicoesFromView = getSimpleListFromView_(SHEETS.VIEW_POSICOES_ATIVAS, 1);
  var defeitosFromView = getSimpleListFromView_(SHEETS.VIEW_DEFEITOS_ATIVOS, 1);

  return {
    posicoes: posicoesFromView.length ? posicoesFromView : Object.keys(defectsByPosition),
    defeitos: defeitosFromView.length ? defeitosFromView : Object.keys(defectSet),
    defeitosPorPosicao: defectsByPosition,
    paresAtivos: activePairs
  };
}

/**
 * Grava linhas em lote no final da aba.
 */
function appendRowsBatch_(sheetName, rows) {
  if (!rows || !rows.length) {
    return;
  }

  var sheet = getRequiredSheet_(sheetName);
  var start = sheet.getLastRow() + 1;
  var width = rows[0].length;
  sheet.getRange(start, 1, rows.length, width).setValues(rows);
}

function getRequiredSheet_(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Aba obrigatória não encontrada: ' + sheetName + '. Rode setupSchema().');
  }
  return sheet;
}

function getOptionalSheet_(sheetName) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

function getSimpleListFromView_(sheetName, valueColumnIndex) {
  var sheet = getOptionalSheet_(sheetName);
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  var values = sheet.getRange(2, valueColumnIndex, sheet.getLastRow() - 1, 1).getValues();
  var result = [];
  var seen = {};

  values.forEach(function (row) {
    var value = String(row[0] || '').trim();
    if (!value || seen[value]) {
      return;
    }
    seen[value] = true;
    result.push(value);
  });

  return result;
}

function getCollaboratorsEditPassword_() {
  var sheet = getRequiredSheet_(SHEETS.COLABORADORES);
  return String(sheet.getRange('E2').getValue() || '').trim();
}

function isOriginRequired_() {
  var sheet = getRequiredSheet_(SHEETS.COLABORADORES);
  var flagValue = sheet.getRange('G2').getValue();
  return flagValue === true || String(flagValue || '').trim().toUpperCase() === 'TRUE';
}

function isActiveFlag_(value) {
  if (value === true || value === 1) {
    return true;
  }

  var normalized = String(value || '').trim().toLowerCase();
  return normalized === 'true' || normalized === '1' || normalized === 'sim' || normalized === 'ativo';
}

function isActiveMatrixFlag_(value) {
  var normalized = String(value || '').trim().toLowerCase();
  return normalized === 'x' || isActiveFlag_(value);
}

function getPositionDefectKey_(posicao, defeito) {
  return String(posicao).trim() + '||' + String(defeito).trim();
}

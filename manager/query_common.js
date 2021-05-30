var activeSpreedsheet = {
  "id": SpreadsheetApp.getActiveSpreadsheet().getId(),
  "range": "Журнал!A2:K"
}

function insertFormulaDataBase(query) {
  var ssPatternSpreadsheet = SpreadsheetApp.openById(ID_DATABASE_MANAGER)
  var sheetSQL = ssPatternSpreadsheet.getSheetByName("SQL")
  sheetSQL.getRange(1, 1).setFormula(query)
  SpreadsheetApp.flush()
  return sheetSQL.getDataRange().getValues()
}

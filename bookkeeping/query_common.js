var spreedsheet_bookkeeper = {
  "id": SpreadsheetApp.getActiveSpreadsheet().getId(),
  "range": "Журнал!A2:K"
}

function insertFormulaDataBase(query) {
  var ssPatternSpreadsheet = SpreadsheetApp.openById(ID_DATABASE)
  var sheetSQL = ssPatternSpreadsheet.getSheetByName("SQL")
  sheetSQL.getRange(1, 1).setFormula(query)
  return sheetSQL.getDataRange().getValues()
}

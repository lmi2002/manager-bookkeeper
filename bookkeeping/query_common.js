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

function insertTemporaryDataBaseActOfReconciliation(query, nameFile) {
  // Создает новый файл
  let createdSS = SpreadsheetApp.create(nameFile)

  // Открывает файл в новой вкладке. 
  openNewSpreadsheet(createdSS)
  createdSS.getSheets()[0].getRange(1, 1).setFormula(query)
  return createdSS.getDataRange().getValues()
}

function getInvoiceForHTMLSQL() { 
  let SQL = "select Col4, Col5, Col6, Col10, Col11 where Col4 like upper ('%" + globalNumInvoiceFromPromt + "%') order by Col4, Col5 asc"
  let Query = '=QUERY(IMPORTRANGE("' + activeSpreedsheet.id + '"; ' + '"' + activeSpreedsheet.range + '"); \"' +SQL+ '\")'
  return getDataAndConvertDate(insertFormulaDataBase(Query))
}

function getDataFromSheetJournalSQL() {
  activeSpreedsheet.range = "Журнал!A1:K"
  let SQL = "select Col4, Col5, Col6, Col10, Col11"
  let Query = '=QUERY(IMPORTRANGE("' + activeSpreedsheet.id + '"; ' + '"' + activeSpreedsheet.range + '"); \"' +SQL+ '\")'
  return insertFormulaDataBase(Query)
}

function getInvoiceAggregatorSummSQL(numInvoice) { 
  let SQL = "select sum(Col11) where Col4 = '" + numInvoice + "' and Col11 is not null"
  let Query = '=QUERY(IMPORTRANGE("' + activeSpreedsheet.id + '"; ' + '"' + activeSpreedsheet.range + '"); \"' +SQL+ '\")'
  let list = insertFormulaDataBase(Query)
  if (list.length > 1) {
    return list[1][0]
  }
  return 0
}

function getNumInvoiceFilterFromSheetJournalSQL(numInvoice) {
  activeSpreedsheet.range = "Журнал!A2:L"
  let SQL = "select * where Col4 = " + "'" + numInvoice + "'"
  let Query = '=QUERY(IMPORTRANGE("' + activeSpreedsheet.id + '"; ' + '"' + activeSpreedsheet.range + '"); \"' +SQL+ '\")'
  return insertFormulaDataBase(Query)
}

function mostSQL(){
  let startday = "'2021-05-01'"
  let finishday = "'2021-05-05'"
  let contragent_code = "432423423"

  let targetRange = 'Журнал!A:K';
  let SQL = "select E, D, F, J, K where B = " + contragent_code + " and (E >= DATE " +  startday + " and E <= DATE " +  finishday + ") and K is not null"
  let Query = '=QUERY('+targetRange+';\"'+SQL+'\")'
  return insertFormulaDataBase(Query) 
}

function getInvoiceForHTMLSQL() { 
  let SQL = "select Col4, Col5, Col6, Col10, Col11 where Col4 like upper ('%" + globalNumInvoiceFromPromt + "%') order by Col4, Col5 asc"
  let Query = '=QUERY(IMPORTRANGE("' + spreedsheet_bookkeeper.id + '"; ' + '"' + spreedsheet_bookkeeper.range + '"); \"' +SQL+ '\")'
  return getDataAndConvertDate(insertFormulaDataBase(Query))
}

function getDataFromSheetJournalSQL() {
  spreedsheet_bookkeeper.range = "Журнал!A1:K"
  let SQL = "select Col4, Col5, Col6, Col10, Col11"
  let Query = '=QUERY(IMPORTRANGE("' + spreedsheet_bookkeeper.id + '"; ' + '"' + spreedsheet_bookkeeper.range + '"); \"' +SQL+ '\")'
  return insertFormulaDataBase(Query)
}

function getInvoiceAggregatorSummSQL(numInvoice) { 
  let SQL = "select sum(Col11) where Col4 = '" + numInvoice + "' and Col11 is not null"
  let Query = '=QUERY(IMPORTRANGE("' + spreedsheet_bookkeeper.id + '"; ' + '"' + spreedsheet_bookkeeper.range + '"); \"' +SQL+ '\")'
  let list = insertFormulaDataBase(Query)
  if (list.length > 1) {
    return list[1][0]
  }
  return 0
}

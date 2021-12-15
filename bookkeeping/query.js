function getInvoiceFromSheetJournalSQL(code, idContragent){

  spreedsheet_bookkeeper.range = 'Журнал!A1:P'
  let SQL
  
  if (code) {
    SQL = "select * where Col2 = " + code + " or Col16 = " + "'" + idContragent + "' order by Col5 asc"
  }
  else {
    SQL = "select * where Col16 = " + "'" + idContragent + "' order by Col5 asc"
  }

  let Query = 'QUERY(IMPORTRANGE("' + spreedsheet_bookkeeper.id + '"; ' + '"' + spreedsheet_bookkeeper.range + '");\"'+SQL+'\")'
  return insertFormulaDataBase(Query)
}

function getInvoiceToSheetActOfReconciliationForPeriodSQL(startDate, finishDate){

  spreedsheet_bookkeeper.range = 'Акт сверки!A1:P'
  let SQL = "select * where Col5 >= DATE " +  "'" + startDate + "'" + " and Col5 <= DATE " +  "'" + finishDate + "' order by Col5 asc,Col4 asc,Col10 asc"
  let Query = 'QUERY(IMPORTRANGE("' + spreedsheet_bookkeeper.id + '"; ' + '"' + spreedsheet_bookkeeper.range + '");\"'+SQL+'\")'
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

function getInvoicesPrevPeriodSummSQL(date){
  spreedsheet_bookkeeper.range = 'Акт сверки!A1:P'
  let SQL = "select count(Col4),sum(Col6) where Col5 < DATE " +  "'" + date + "'" + "group by Col4"
  let SQL1 = "select Col2/Col1"
  let SQL2 = "select sum(Col1)"
  let Query = '=QUERY(QUERY(QUERY(IMPORTRANGE("' + spreedsheet_bookkeeper.id + '"; ' + '"' + spreedsheet_bookkeeper.range + '");\"'+SQL+'\");\"'+SQL1+'\"); \"'+SQL2+'\")'
  let list = insertFormulaDataBase(Query)
  if (list.length > 1) {
    return list[1][0]
  }
  return 0
}

function getPaidInvoicesPrevPeriodSummSQL(date){
  spreedsheet_bookkeeper.range = 'Акт сверки!A1:P'
  let SQL = "select sum(Col11) where Col5 < DATE " +  "'" + date + "'"
  let Query = 'QUERY(IMPORTRANGE("' + spreedsheet_bookkeeper.id + '"; ' + '"' + spreedsheet_bookkeeper.range + '");\"'+SQL+'\")'
  let list = insertFormulaDataBase(Query)
  if (list.length > 1) {
    return list[1][0]
  }
  return 0
}

function getInvoicesPeriodSummSQL(startDate, finishDate){

  spreedsheet_bookkeeper.range = 'Акт сверки!A1:P'
  let SQL = "select count(Col4),sum(Col6) where Col5 >= DATE " +  "'" + startDate + "'" + " and Col5 <= DATE " +  "'" + finishDate + "'" + "group by Col4"
  let SQL1 = "select Col2/Col1"
  let SQL2 = "select sum(Col1)"
  let Query = '=QUERY(QUERY(QUERY(IMPORTRANGE("' + spreedsheet_bookkeeper.id + '"; ' + '"' + spreedsheet_bookkeeper.range + '");\"'+SQL+'\");\"'+SQL1+'\"); \"'+SQL2+'\")'
  let list = insertFormulaDataBase(Query)
  if (list.length > 1) {
    return list[1][0]
  }
  return 0
}

function getPaidInvoicesPeriodSummSQL(startDate, finishDate){

  spreedsheet_bookkeeper.range = 'Акт сверки!A1:P'
  let SQL = "select sum(Col11) where Col5 >= DATE " +  "'" + startDate + "'" + " and Col5 <= DATE " +  "'" + finishDate + "'"
  let Query = 'QUERY(IMPORTRANGE("' + spreedsheet_bookkeeper.id + '"; ' + '"' + spreedsheet_bookkeeper.range + '");\"'+SQL+'\")'
  let list = insertFormulaDataBase(Query)
  if (list.length > 1) {
    return list[1][0]
  }
  return 0
}


function getQtyContragentAndCodeSQL() {

  spreedsheet_bookkeeper.range = 'Акт сверки!A1:P'
  let SQL = "select Col1, Col2, count(Col1) where Col1 is not null group by Col1, Col2 Label count(Col1) 'Кол-во строк'"
  let Query = '=QUERY(IMPORTRANGE("' + spreedsheet_bookkeeper.id + '"; ' + '"' + spreedsheet_bookkeeper.range + '"); \"' +SQL+ '\")'
  return insertFormulaDataBase(Query)
}

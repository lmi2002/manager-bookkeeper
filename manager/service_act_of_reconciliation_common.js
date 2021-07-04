function differenceAmountInvoice() {
  let obj = getObjSpreadsheetApp()
  let numInvoice = obj.values_list[0][3]
  let sum = getInvoiceAggregatorSummSQL(numInvoice)
  return obj.values_list[0][5] - sum
}

function addPaymentInvoiceToBookkeeperJournal(list) {
  let ss = SpreadsheetApp.openById(ID_BOOKKEEPER)
  let sheetJournal = ss.getSheetByName("Журнал")
  for (var i = 0; i < list.length; i++) {
    var lastRow = sheetJournal.getLastRow()
    for (var y = 0; y < list[i].length; y++) {
      sheetJournal.getRange(lastRow + 1, y + 1).setValue(list[i][y]).setBackground("#FFF2CC")
    }  
  }
}

function addPaymentInvoiceToBookkeeperSheetPayment(list) {
  let ss = SpreadsheetApp.openById(ID_BOOKKEEPER)
  let sheetPayment = ss.getSheetByName("Оплата")
  let firstElem = list[0]
  let lastRow = sheetPayment.getLastRow() + 1
  let lastColumn = sheetPayment.getLastColumn()
  let = nal = Math.floor(Number(firstElem[5]) * 0.9091)

  let listDate = list.map(function(el) {
    return el[9]
  })

  let maxDate = listDate.reduce(function(prevElem, currentElem) {
    if (prevElem > currentElem) {
      return prevElem
    }
    else {
      return currentElem
    }
  })

  sheetPayment.getRange(lastRow, 1).setValue(firstElem[2])
  sheetPayment.getRange(lastRow, 2).setValue(nal)
  sheetPayment.getRange(lastRow, 3).setValue(firstElem[5])
  sheetPayment.getRange(lastRow, 4).setValue(maxDate)
  sheetPayment.getRange(lastRow, 5).setValue(firstElem[0])
  sheetPayment.getRange(lastRow, 11).setValue(firstElem[3])

  sheetPayment.getRange(lastRow, 1, 1, lastColumn).setBackground("#FFF2CC")
}

// Получить номер договора на листе Свободные (функционал не используется)
function getNumContractFromSheetAvailable(listDataJournal) {

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui = SpreadsheetApp.getUi()
  let sheetAvailable = ss.getSheetByName("СВОБОДНЫЕ");
  let listData = readRange(speadsheetManager.id, speadsheetManager.rangeSheetAvailable)
  for (let i = 0; i < listData.length; i++ ) {
    if (listData[i][0] == listDataJournal[0][2]) {
      var numRow = i + 2
      break
    }
  }  

  if (numRow) {
    return sheetAvailable.getRange(numRow,13).getValue()
  }
  else {     
    ui.alert("Не найден код бытовки " + listDataJournal[0][2] + " на листе Свободные!")
  }
              
}

// При полной оплате
function makePaymentToSheetAvailable(listDataJournal) {

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui = SpreadsheetApp.getUi()
  let sh_data = ss.getSheetByName("СВОБОДНЫЕ");
  let listData = readRange(speadsheetManager.id, speadsheetManager.rangeSheetAvailable)
  for (let i = 0; i < listData.length; i++ ) {
    if (listData[i][0] == listDataJournal[0][2]) {
      var numRow = i + 2
      break
    }
  }  

  if (numRow) {
    sh_data.getRange(numRow,14).setValue("Cчет оплачен полностью")
  }
  else {     
    ui.alert("Не найден код бытовки " + listDataJournal[0][2] + " на листе Свободные!")
  }
              
}

// При частичной оплате
function makePartialPaymentToSheetAvailable(listDataJournal) {

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui = SpreadsheetApp.getUi()
  let sh_data = ss.getSheetByName("СВОБОДНЫЕ");
  let listData = readRange(speadsheetManager.id, speadsheetManager.rangeSheetAvailable)
  for (let i = 0; i < listData.length; i++ ) {
    if (listData[i][0] == listDataJournal[0][2]) {
      var numRow = i + 2
      break
    }
  }

  if (numRow) {
    sh_data.getRange(numRow,14).setValue("Cчет оплачен частично")
  }
  else {     
    ui.alert("Не найден код бытовки " + listDataJournal[0][2] + " на листе Свободные!")
  }             
}

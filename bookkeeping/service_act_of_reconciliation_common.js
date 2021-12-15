function differenceAmountInvoice() {
  let obj = getObjSpreadsheetApp()
  let numInvoice = obj.values_list[0][3]
  let sum = getInvoiceAggregatorSummSQL(numInvoice)
  return obj.values_list[0][5] - sum
}

function addListInvoiceToSheetActOfReconciliation() {

  if (SpreadsheetApp.getActiveSheet().getSheetName() == 'Журнал') {

    let obj = getObjSpreadsheetApp()
    let list = obj.values_list
    let code = list[0][1]
    let contragent = list[0][0]
    let idContragent = list[0][15]
    let ui = SpreadsheetApp.getUi()
    let a = 0
    let response = ui.alert("Список для акта сверки", "Код: " + code  + "\n Контрагент: " +  contragent, ui.ButtonSet.YES_NO); 
            
    // Добавляет данные в лист Акт сверки при нажатии кнопки ОК
    if (response == ui.Button.YES) {     
        obj.act_ss.getSheetByName("Акт сверки").clear()
        SpreadsheetApp.flush()
        if (!idContragent) {
          idContragent = deleteSymForIDContragent(contragent)
        }
        let result = getInvoiceFromSheetJournalSQL(code, idContragent)
        for (let el of result) {
          if (a != 0 ) {
            el[4] = getStrDay_1(el[4])
            el[6] = getStrDay_1(el[6])
            if (el[9] != "") {
              el[9] = getStrDay_1(el[9])
            }  
          }
          a++
        }
        appendToMultipleRanges(result, speadsheetBookkeeper.id, speadsheetBookkeeper.rangeSheetActOfReconciliation)   
    }
  }  
  else {
    SpreadsheetApp.getUi().alert('Перейдите на вкладку Журнал!')
  }  
}

// Не изменять!!!
function deleteSymForIDContragent(str) {
  return str.replace(/[\s\"\'\-,.:;+<>«»\\\”\“]/g, "").toLowerCase();
}

function clearsheetActOfReconciliation(sheet) {
  sheet.clear()
}

function copyActOfReconciliationFile(objAOR) {

  let contragent = objAOR.contragent
  let formatStartDate = objAOR.formatStartDate
  let formatFinishDate = objAOR.formatFinishDate
  let invoicesForPeriod = objAOR.invoicesForPeriod
  let balancePrev = objAOR.balancePrev
  let balanceCurrent = objAOR.balanceCurrent
  let sumInvoicesPeriod = objAOR.sumInvoicesPeriod
  let sumPaidInvoicesPeriod = objAOR.sumPaidInvoicesPeriod
  let balanceFinish = balanceCurrent + balancePrev
  

  let period = formatStartDate + " - " + formatFinishDate 
  let nameFile = 'Акт сверки ' + contragent + " (" + period + ")"
  let ssTemplateActOfReconciliationFile = SpreadsheetApp.openById(ID_TEMPLATE_ACT_RECONCILIATION)
  let ssCopyTemplateActOfReconciliationFile = ssTemplateActOfReconciliationFile.copy(nameFile)
  let sheetCopyTemplateActOfReconciliationFile = ssCopyTemplateActOfReconciliationFile.getSheets()[0]
  let idSSCopyTemplateActOfReconciliationFile = ssCopyTemplateActOfReconciliationFile.getId()
  let list = Array()
 
  let textActOfReconciliation_1 = getTextActOfReconciliation_1(period, contragent)
  let textActOfReconciliation_2 = getTextActOfReconciliation_2(contragent)

  // Шапка тоблицы
  sheetCopyTemplateActOfReconciliationFile.getRange('A3').setValue(textActOfReconciliation_1)
  sheetCopyTemplateActOfReconciliationFile.getRange('A5').setValue(textActOfReconciliation_2)
  sheetCopyTemplateActOfReconciliationFile.getRange('E7').setValue("За даними " + contragent + ", грн")
  sheetCopyTemplateActOfReconciliationFile.getRange('A9').setValue("Сальдо на "+ formatStartDate)
  
  if (balancePrev > 0) {
    sheetCopyTemplateActOfReconciliationFile.getRange('C9').setValue(balancePrev)
  }
  else if (balancePrev < 0) {
    sheetCopyTemplateActOfReconciliationFile.getRange('D9').setValue(Math.abs(balancePrev))
  }

  sheetCopyTemplateActOfReconciliationFile.getRange('E9').setValue("Сальдо на "+ formatStartDate )


  // Добавление таблицы счетов
  for (let i = 0; i < invoicesForPeriod.length; i++) {
    if (i != 0 ) {
      let invoice = "Рах.№ " + invoicesForPeriod[i][3]
      // let sum = invoicesForPeriod[i][5].toFixed(2)
      if (invoicesForPeriod[i][10] == "" ) {
        list.push([getStrDay_1(invoicesForPeriod[i][4]), invoice, invoicesForPeriod[i][5]])
      }
      else if (invoicesForPeriod[i-1][3] == invoice) {
        list.push([getStrDay_1(invoicesForPeriod[i][9]), "Сплачено", "", invoicesForPeriod[i][10]])
      }
      else {
        list.push([getStrDay_1(invoicesForPeriod[i][4]), invoice, invoicesForPeriod[i][5]])
        list.push([getStrDay_1(invoicesForPeriod[i][9]), "Сплачено", "", invoicesForPeriod[i][10]])
      }
    }
  }

  appendToMultipleRanges(list, idSSCopyTemplateActOfReconciliationFile, "Акт сверки!A10:D")


  // Итог по таблице счетов
  let lastRow = sheetCopyTemplateActOfReconciliationFile.getLastRow()

  sheetCopyTemplateActOfReconciliationFile.getRange('A'+ Number(lastRow + 1) + ":B" + Number(lastRow + 1)).merge().setValue("Обороти за період").setFontWeight("bold")

  if (sumInvoicesPeriod > 0) {
    sheetCopyTemplateActOfReconciliationFile.getRange('C'+ Number(lastRow + 1)).setValue(sumInvoicesPeriod).setFontWeight("bold")
  }

  if (sumPaidInvoicesPeriod > 0) {
    sheetCopyTemplateActOfReconciliationFile.getRange('D'+ Number(lastRow + 1)).setValue(sumPaidInvoicesPeriod).setFontWeight("bold")
  }  

  sheetCopyTemplateActOfReconciliationFile.getRange('E'+ Number(lastRow + 1) + ":F" + Number(lastRow + 1)).merge().setValue("Обороти за період").setFontWeight("bold")

  sheetCopyTemplateActOfReconciliationFile.getRange('A'+ Number(lastRow + 2) + ":B" + Number(lastRow + 2)).merge().setValue("Сальдо кінцеве").setFontWeight("bold")

  if (balanceFinish > 0) {
    sheetCopyTemplateActOfReconciliationFile.getRange('C'+ Number(lastRow + 2)).setValue(balanceFinish).setFontWeight("bold")
  }  
  else if (balanceFinish < 0){ 
    sheetCopyTemplateActOfReconciliationFile.getRange('D'+ Number(lastRow + 2)).setValue(Math.abs(balanceFinish)).setFontWeight("bold")
  }



  sheetCopyTemplateActOfReconciliationFile.getRange('E'+ Number(lastRow + 2) + ":F" + Number(lastRow + 2)).merge().setValue("Сальдо кінцеве").setFontWeight("bold")

  // В таблицу добавляю границу
  sheetCopyTemplateActOfReconciliationFile.getRange('A10:H'+ Number(lastRow + 2)).setBorder(true, null, true, true, true, true)

  // В таблицу устанавливаю формат #0,00 для следующих столбцов C,D,G,H
  sheetCopyTemplateActOfReconciliationFile.getRange('C10:C'+ Number(lastRow + 2)).setNumberFormat("0.00")
  sheetCopyTemplateActOfReconciliationFile.getRange('D10:D'+ Number(lastRow + 2)).setNumberFormat("0.00")
  sheetCopyTemplateActOfReconciliationFile.getRange('G10:G'+ Number(lastRow + 2)).setNumberFormat("0.00")
  sheetCopyTemplateActOfReconciliationFile.getRange('H10:H'+ Number(lastRow + 2)).setNumberFormat("0.00")

  // Футер Акта
  sheetCopyTemplateActOfReconciliationFile.getRange('A'+ Number(lastRow + 4)).setValue("За даними ФОП "+  DIRECTOR_NAME_SHORT)
  sheetCopyTemplateActOfReconciliationFile.getRange('A'+ Number(lastRow + 5) + ":D" + Number(lastRow + 5)).merge().setValue(getTextActOfReconciliation_3(formatFinishDate, balanceFinish, contragent)).setWrap(true).setFontWeight("bold")
  sheetCopyTemplateActOfReconciliationFile.getRange('A'+ Number(lastRow + 7)).setValue("Від ФОП "+  DIRECTOR_NAME_SHORT)
  sheetCopyTemplateActOfReconciliationFile.getRange('E'+ Number(lastRow + 7) + ':H'+ Number(lastRow + 7)).merge().setValue("Від "+  contragent).setWrap(true)

  sheetCopyTemplateActOfReconciliationFile.getRange('A'+ Number(lastRow + 9) + ":B" + Number(lastRow + 9)).setBorder(null, null, true, null, null, null)
  sheetCopyTemplateActOfReconciliationFile.getRange('E'+ Number(lastRow + 9) + ":F" + Number(lastRow + 9)).setBorder(null, null, true, null, null, null)
  sheetCopyTemplateActOfReconciliationFile.getRange('A'+ Number(lastRow + 11) + ":B" + Number(lastRow + 11)).setBorder(null, null, true, null, null, null)
  sheetCopyTemplateActOfReconciliationFile.getRange('E'+ Number(lastRow + 11) + ":F" + Number(lastRow + 11)).setBorder(null, null, true, null, null, null)
  sheetCopyTemplateActOfReconciliationFile.getRange('A'+ Number(lastRow + 13)).setValue("М.П.")
  sheetCopyTemplateActOfReconciliationFile.getRange('E'+ Number(lastRow + 13)).setValue("М.П.")
  
  SpreadsheetApp.flush()

  return {
    "ss": ssCopyTemplateActOfReconciliationFile
  }
}


function moveActOfReconciliationFiles(obj) {
  let folder_name = "АКТЫ СВЕРКИ"
  let folder = getFolders(folder_name)
  
  if (!folder) {
    folder = DriveApp.createFolder(folder_name)
  }
  obj["file"].moveTo(folder)   
}


function getFirstNameCompany() {
  let obj = getObjSpreadsheetApp()
  return obj.act_sheet.getRange(2,1).getValue()
}

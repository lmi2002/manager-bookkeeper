function addCustomMenu() {
  let email = Session.getEffectiveUser().getEmail()
  
  let permition = EMAIL_PERMITION_LIST.some(function(item){
    return email == item 
  })

  if (permition) {
    SpreadsheetApp.getUi()
      .createMenu('Скрипты')
      .addItem("Первый счет","OpenFormDialog")
      .addItem("Оплата Счета","PaymentInvoiceToJouranl")
      .addItem("Запрос физ. лицо","OpenFormDialogEmailLawyerPerson")
      .addItem("Запрос юр. лицо","OpenFormDialogEmailLawyerCompany")
      .addToUi();
  }
  else {
    SpreadsheetApp.getUi()
      .createMenu('Скрипты')
      .addItem("Первый счет","OpenFormDialog")
      .addItem("Запрос","SendEmailLawyer")
      .addToUi()
  }    
}

function OpenFormDialog() {
  
  var obj = getObjSpreadsheetApp()
  
  if (obj['act_sheet'].getName() == 'СВОБОДНЫЕ') {
    var contragent = obj.values_list[0][11]
    var contragentDetails = obj.values_list[0][22]
    var addDays = new Date(obj.values_list[0][1].getFullYear(), obj.values_list[0][1].getMonth(), obj.values_list[0][1].getDate() + 30 * Number(obj.values_list[0][7])).getTime()
    var finishDate = new Date(obj.values_list[0][8]).getTime()

    if (contragent && contragentDetails) {

      if (Number(obj.values_list[0][3]) * Number(obj.values_list[0][7]) + Number(obj.values_list[0][5]) == Number(obj.values_list[0][6])) {

        if (addDays == finishDate) {
      
          var t = HtmlService.createTemplateFromFile('form_dialog')
            .evaluate()
            .setSandboxMode(HtmlService.SandboxMode.IFRAME)
            .setWidth(300)
            .setHeight(350)
     
          SpreadsheetApp.getUi()
            .showModalDialog(t, "Выставить счет");
        }    
        else {
          SpreadsheetApp.getUi().alert("Даты не совпадают!")
        }    
      }
      else {
        SpreadsheetApp.getUi().alert("Суммы не совпадают!")
      }  
    }
    else {
      SpreadsheetApp.getUi().alert('Не заполнено значением столбец "Контрагент" или "Реквизиты контрагента"')
    }   
  }
  else {
     SpreadsheetApp.getUi().alert("Перейдите на лист СВОБОДНЫЕ!")
  }
}

function RouterInvoiceAct(obj) {
  
  var date_invoice = obj.date_invoice
  var date_act = obj.date_act
  var status_period = obj.status_period
  var status_send_email = obj.status_send_email
  var status_delivery = obj.status_delivery
  var send_journal = obj.send_journal
  var obj_ss = getObjSpreadsheetApp()

 if (date_invoice && date_act) {
    obj_ss['act_sheet'].getRange(obj_ss['act_range'].getRow(), 14).setValue('Скрипт запущен')
    dict_invoice = Invoice(date_invoice, status_period, status_delivery, obj_ss)
    dict_act = Act(date_invoice, date_act, status_period, status_delivery, obj_ss)
    if (status_send_email) {
      sendInvoiceAndAct(dict_invoice, dict_act)
    }
    if (send_journal) {
      addInformInvoiceToInvoiceJournal(date_invoice, dict_invoice, obj_ss)
    }

    obj_ss['act_sheet'].getRange(obj_ss['act_range'].getRow(), 14).setValue('Счет выставлен')
  }
  else {
    var ui = SpreadsheetApp.getUi()
    ui.alert("Заполните реквизиты контрагента или дату на форме и повторите снова.")
  }
}

function OpenFormDialogPaymentJournal() {

  var t = HtmlService.createTemplateFromFile('form_dialog_payment_journal')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(350)
    .setHeight(300)

  SpreadsheetApp.getUi()
    .showModalDialog(t, "Ввод суммы оплаты счета в журнал");  
}


function OpenFormDialogPaymentJournalCreateInvoice() {

  var t = HtmlService.createTemplateFromFile('form_dialog_payment_journal_create_invoice')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(350)
    .setHeight(300)

  SpreadsheetApp.getUi()
    .showModalDialog(t, "Ввод суммы оплаты счета в журнал") 
}

function PaymentInvoiceToJouranl() {

  let objSpreadsheetApp = getObjSpreadsheetApp()
  let sheetActive = objSpreadsheetApp.act_sheet
  let ui = SpreadsheetApp.getUi()

  if (sheetActive.getSheetName() == "Журнал") {
    let listData = objSpreadsheetApp.values_list
    let numInvoice = listData[0][3]
    let code = listData[0][1]

    if (code) {
      let sum = getInvoiceAggregatorSummSQL(numInvoice)

      if (sum < listData[0][5]) {

        if (sum == 0) {
          OpenFormDialogPaymentJournal()
        }
        else {    
          OpenFormDialogPaymentJournalCreateInvoice()
        }
      }
      else {
        ui.alert("Счет полностью оплачен!")
      }
    }  
    else {
      ui.alert("Заполните код ЄДРПОУ/ИНН по счету " + numInvoice)
    }  
  }        
  else {
    ui.alert("Перейдите на лист Журнал!")
  }
}

function writePaymentInvoiceToJouranl(obj) {

  try {

    var ui = SpreadsheetApp.getUi()
    let objSpreadsheetApp = getObjSpreadsheetApp()
    let sheetActive = objSpreadsheetApp.act_sheet
    let activeSpreadsheetID = objSpreadsheetApp.act_ss.getId()
    let activeRange = objSpreadsheetApp.act_range
    let listData = objSpreadsheetApp.values_list
    let numRow = activeRange.getRow()
    let lastColumn = sheetActive.getLastColumn()
    let sum = getInvoiceAggregatorSummSQL(listData[0][3])
    let payment_sum = Number(obj.payment_sum)
    let payment_date = obj.payment_date
    // let payment_sum = Number("500")
    // let payment_date = "2021-12-14"
    let rangeInvoice = "Журнал!D1:D"
    let rangeSort = sheetActive.getRange(2,1, sheetActive.getLastRow(), lastColumn)
    let bool = true
    let totalSum = payment_sum + sum
  
    if (listData[0][10] == "") {

      sheetActive.getRange(numRow, 10).setValue(payment_date)
      sheetActive.getRange(numRow, 11).setValue(payment_sum)
      
      if (payment_sum  >= listData[0][5]) {
        sheetActive.getRange(numRow, 1, 1, lastColumn).setBackground("#FFF2CC")
        makePaymentToSheetAvailable(listData)
        let listFilterNumInvoice = getNumInvoiceFilterFromSheetJournalSQL(listData[0][3])
        addPaymentInvoiceToBookkeeperJournal(listFilterNumInvoice)
        addPaymentInvoiceToBookkeeperSheetPayment(listFilterNumInvoice, totalSum)
        ui.alert(getMessageAboutPayment(payment_sum, listData[0][5], totalSum))
      }
      else {
        makePartialPaymentToSheetAvailable(listData)
        ui.alert("Оплата счета прошла упешно. Счет оплачен частично.")
      }
    }  
    else {

      if (totalSum >= listData[0][5]) {

        sheetActive.insertRows(numRow, 1)
        for (let i = 0;  i < listData[0].length; i++) {
          sheetActive.getRange(numRow, i + 1).setValue(listData[0][i])
        }
        sheetActive.getRange(numRow, 10).setValue(payment_date)
        sheetActive.getRange(numRow, 11).setValue(payment_sum)
        // sheetActive.getRange(numRow, 12).setValue(numContract)

        // Ожидание заполнения суммой ячейки
        let iter = 0
        while (bool) {
          iter ++
          try {
            if (sheetActive.getRange(numRow, 11).getValue() == payment_sum) {
            bool = false
            }
            else {
              if (iter == 1000) {
                throw new Error("Нет внесенной суммы счета в ячейке. Программа закончила работу с ошибкой!")
              }
            }
          }   
          catch(error) {
            ui.alert(error.message)
            sheetActive.deleteRow(numRow)
          }
        }

        let listInvoiceJournal = readRange(activeSpreadsheetID,rangeInvoice)

        for (let i = 0; i < listInvoiceJournal.length; i ++) {
          if (listInvoiceJournal[i] == listData[0][3]) {
            sheetActive.getRange(i + 1, 1, 1, lastColumn).setBackground("#FFF2CC")
          }  
        }
        makePaymentToSheetAvailable(listData)
        let listFilterNumInvoice = getNumInvoiceFilterFromSheetJournalSQL(listData[0][3])
        addPaymentInvoiceToBookkeeperJournal(listFilterNumInvoice)
        addPaymentInvoiceToBookkeeperSheetPayment(listFilterNumInvoice, totalSum)
        ui.alert(getMessageAboutPayment(payment_sum, listData[0][5], totalSum))

      }
      else {
        sheetActive.insertRows(numRow, 1)
        for (let i = 0;  i < listData[0].length; i++) {
          sheetActive.getRange(numRow, i + 1).setValue(listData[0][i])
        }
        sheetActive.getRange(numRow, 10).setValue(payment_date)
        sheetActive.getRange(numRow, 11).setValue(payment_sum)
        makePartialPaymentToSheetAvailable(listData)
        ui.alert("Оплата счета прошла упешно. Счет оплачен частично.")
      }
    }

    rangeSort.sort([{column: 5, ascending: true}, {column: 4, ascending: true}])
  }
  catch(e) {
    ui.alert("Программа завершилась с ошибкой. Перепроверьте все данные! \n Ошибка: " + e.message + "\n" + e.name + "\n" + e.stack)
    
  }  
}

function OpenFormDialogEmailLawyerPerson() {
  try {
    let objSpreadsheetApp = getObjSpreadsheetApp()
    let sheetActive = objSpreadsheetApp.act_sheet
    let list = objSpreadsheetApp.values_list[0]
    let sum = Number(list[2]) * Number(list[7]) + Number(list[4])
    let sum_delivery = Number(list[6])
    let addDays = new Date(list[1].getFullYear(), list[1].getMonth(), list[1].getDate() + 30 * Number(list[7])).getTime()
    let finishDate = new Date(list[8]).getTime()

    if (sheetActive.getSheetName() == "СВОБОДНЫЕ") {
      if (sum == sum_delivery) {
        if (addDays == finishDate) {
          var t = HtmlService.createTemplateFromFile('form_dialog_email_lawyer_person')
            .evaluate()
            .setSandboxMode(HtmlService.SandboxMode.IFRAME)
            .setWidth(350)
            .setHeight(400)

          SpreadsheetApp.getUi()
            .showModalDialog(t, "Запрос Физлицо")
        }
        else {
          SpreadsheetApp.getUi().alert("Даты не совпадают!")
        }    
      }
      else {
        SpreadsheetApp.getUi().alert("Суммы не совпадают!")
      }    
    }
    else {
      SpreadsheetApp.getUi().alert("Перейдите на лист Свободные!")
    }
  }  
  catch(e) {
    SpreadsheetApp.getUi().alert(e.message)
  }      
}

function OpenFormDialogEmailLawyerCompany() {
  try {
    let objSpreadsheetApp = getObjSpreadsheetApp()
    let sheetActive = objSpreadsheetApp.act_sheet
    let list = objSpreadsheetApp.values_list[0]
    let sum = Number(list[3]) * Number(list[7]) + Number(list[5])
    let sum_delivery = Number(list[6])
    let addDays = new Date(list[1].getFullYear(), list[1].getMonth(), list[1].getDate() + 30 * Number(list[7])).getTime()
    let finishDate = new Date(list[8]).getTime()

    if (sheetActive.getSheetName() == "СВОБОДНЫЕ") {
      if (sum == sum_delivery) {
        if (addDays == finishDate) {
          var t = HtmlService.createTemplateFromFile('form_dialog_email_lawyer_company')
            .evaluate()
            .setSandboxMode(HtmlService.SandboxMode.IFRAME)
            .setWidth(350)
            .setHeight(400)

          SpreadsheetApp.getUi()
            .showModalDialog(t, "Запрос Юрлицо")
        }
        else {
          SpreadsheetApp.getUi().alert("Даты не совпадают!")
        }
      }
      else {
        SpreadsheetApp.getUi().alert("Суммы не совпадают!")
      }     
    }
    else {
      SpreadsheetApp.getUi().alert("Перейдите на лист Свободные!")
    }
  }  
  catch(e) {
    SpreadsheetApp.getUi().alert(e.message)
  }      
}

function runSendPersonLetter(obj) {
  let objSpreadsheetApp = getObjSpreadsheetApp()
  let sheetActive = objSpreadsheetApp.act_sheet
  let rangeActive = objSpreadsheetApp.act_range
  let list = objSpreadsheetApp.values_list[0]
  let comment = obj.comment

  sendPersonLetterBody(comment, list)
  sheetActive.getRange(rangeActive.getRow(), 11).setValue(comment)
  sheetActive.getRange(rangeActive.getRow(), 17).setValue("запрос нал")
}

function runSendCompanyLetter(obj) {
  let objSpreadsheetApp = getObjSpreadsheetApp()
  let sheetActive = objSpreadsheetApp.act_sheet
  let rangeActive = objSpreadsheetApp.act_range
  let list = objSpreadsheetApp.values_list[0]
  let comment = obj.comment

  sendCompanyLetterBody(obj, list)
  sheetActive.getRange(rangeActive.getRow(), 11).setValue(comment)
  sheetActive.getRange(rangeActive.getRow(), 17).setValue("запрос юр")
}

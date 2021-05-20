function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Скрипты')
      .addItem("Первый счет","OpenFormDialog")
      .addToUi();
}

function OpenFormDialog() {
  
  var obj = getObjSpreadsheetApp()
  
  if (obj['act_sheet'].getName() == 'СВОБОДНЫЕ') {
    var contragent = obj.values_list[0][11]

    if (contragent) {
      
      var t = HtmlService.createTemplateFromFile('form_dialog')
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(300)
      .setHeight(350)
     
      SpreadsheetApp.getUi()
        .showModalDialog(t, "Выставить счет"); 
    }
    else {
      var ui = SpreadsheetApp.getUi()
      ui.alert('Заполните значением столбец "Контрагент"')
    }   
  }
  else {
     var ui = SpreadsheetApp.getUi()
     ui.alert("Перейдите на лист СВОБОДНЫЕ!")
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

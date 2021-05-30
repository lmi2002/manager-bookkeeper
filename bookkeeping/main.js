//Настраивать Google Table Файл -> Настройки таблицы Часовой пояс (GMT+02:00)Moscow-01-Kaliningrad

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = 
      [
        {name : "Провести оплату",functionName : "PromBox"},
        {name : "Выставить счет",functionName : "OpenFormDialog"},
        {name : "Оплата Счета",functionName : "PaymentInvoiceToJouranl"}
      ]
  sheet.addMenu("Скрипты", entries);
}


function PromBox(){

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetActive = ss.getActiveSheet()
//  var ss_manager_oplati = SpreadsheetApp.openById('1E3ocje0AdmhgvBTjre524ETOduY8XEw4Yd_g0SJ1mU0').getSheetByName('ОПЛАТЫ') // id "ДИРЕКТОР_version"
  if (sheetActive.getSheetName() == "Данные") {
    var sh_data = ss.getSheetByName("Данные");
    var sh_oplata = ss.getSheetByName("Оплата");
    var range = sh_data.getActiveRange();
    var range_sort = sh_data.getRange(2,1, sh_data.getLastRow()-1, sh_data.getLastColumn())
     
    
    
    // Получаю массив значений из элементов (Код, Дата Отгрузки, Нал, БЕЗНАЛ за 3 дней БЕЗ НДС, Конец периода 30 дней, Контрагент)
    var arr = sh_data.getRange(range.getRow(), 1, 1, 8).getValues();
    
    var code = arr[0][0]
    var nal = arr[0][2]
    var beznal = arr[0][3]
    var days_30 = arr[0][4]
    var contragent = arr[0][5]
    var status = arr[0][7]
    
    // Вызов Promt
    var ui = SpreadsheetApp.getUi();
    var promt = ui.prompt("Провести оплату", "Введите дату оплаты счета. \n Код: " + code  + "\n Контрагент: " +  contragent, ui.ButtonSet.OK_CANCEL);   
    var button = promt.getSelectedButton()
        
    
    
    // Добавляет данные в лист Оплата при нажатии кнопки ОК
    if (button == ui.Button.OK){
      
      sh_data.getRange(range.getRow(),8).setValue("Оплатили")
     
      var date = Utilities.formatDate(days_30, "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'")
      
      var d = new Date(date) // Создали объект даты
     
      sh_data.getRange(range.getRow(),5).setValue(new Date(d.getFullYear(), d.getMonth(), d.getDate() + 30)) // Добавили 30 дней к дате и изменили в ячейке.
     
      //Сортировка
      range_sort.sort({column: 5, ascending: true});       
    
      var rows_oplata = sh_oplata.getLastRow(); // кол-во строк
     
      // Добавляю данные по ячейкам в лист Оплата
      sh_oplata.getRange(rows_oplata + 1, 1,1).setValue(code).setBackground('#DEE9FC')
      sh_oplata.getRange(rows_oplata + 1, 2,1).setValue(nal).setBackground('#DEE9FC')
      sh_oplata.getRange(rows_oplata + 1, 3,1).setValue(beznal).setBackground('#DEE9FC')
      sh_oplata.getRange(rows_oplata + 1, 4,1).setValue(promt.getResponseText()).setBackground('#DEE9FC')
      sh_oplata.getRange(rows_oplata + 1, 5,1).setValue(contragent).setBackground('#DEE9FC')
      sh_oplata.getRange(rows_oplata + 1, 6,1).setValue(d)
      sh_oplata.getRange(rows_oplata + 1, 7,1).setValue(status)
    }
   
    // Отменяет действие при нажатии кнопки Отмена
    else if (button = ui.Button.CANCEL){
      ui.alert("Проводка отменена")      
    }
  }  
  else {
     SpreadsheetApp.getUi().alert("Перейдите на лист Данные!")
  }     
}

// Автозаполнение столбца B лист Оплата, при изменении столбца С
function subtractPercentage() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var range = ss.getActiveRange()
  var id_col = range.getColumn()
  if (ss.getActiveSheet().getName() == 'Оплата' && id_col == 3) {
    var dim = range.getValue()
    ss.getActiveSheet().getRange(range.getRow(), range.getColumn() - 1).setValue(Math.round(dim*0.9091))
  }   
}

function onEdit() {
  subtractPercentage()
}



function OpenFormDialog() {
  
  var obj = getObjSpreadsheetApp()

  if (obj['act_sheet'].getName() == 'Данные') {
    var contragent_details = obj.values_list[0][14]
    var contragent_code = obj.values_list[0][15]
    if (contragent_details && validationIsNumber(contragent_code)) {
      var t = HtmlService.createTemplateFromFile('form_dialog')
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(300)
      .setHeight(330)
     
      SpreadsheetApp.getUi()
        .showModalDialog(t, "Выставить счет");
    }
    else {
      var ui = SpreadsheetApp.getUi()
      ui.alert("Заполните реквизиты или проверьте код контрагента")
    }      
  }
  else {
     var ui = SpreadsheetApp.getUi()
     ui.alert("Перейдите на лист Данные!")
  }
}

function RouterInvoiceAct(obj) {
  
  var date_invoice = obj.date_invoice
  var date_act = obj.date_act
  var status_period = obj.status_period
  var status_send_email = obj.status_send_email
  var send_journal = obj.send_journal
  var obj_ss = getObjSpreadsheetApp()

  if (date_invoice && date_act) {
    obj_ss['act_sheet'].getRange(obj_ss['act_range'].getRow(), 8).setValue('Скрипт запущен')
    dict_invoice = Invoice(date_invoice, status_period, obj_ss)
    dict_act = Act(date_invoice, date_act,status_period, obj_ss)
    if (status_send_email) {
      sendInvoiceAndAct(dict_invoice, dict_act)
    }
    if (send_journal) {
      addInformInvoiceToInvoiceJournal(date_invoice, dict_invoice, obj_ss)
    }

    obj_ss['act_sheet'].getRange(obj_ss['act_range'].getRow(), 8).setValue('Счет выставлен')
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
    ui.alert("Перейдите на лист Журнал!")
  }
}

function writePaymentInvoiceToJouranl(obj) {

  let ui = SpreadsheetApp.getUi()
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
  let rangeInvoice = "Журнал!D1:D"
  let rangeSort = sheetActive.getRange(2,1, sheetActive.getLastRow(), lastColumn)
  let bool = true
 
  if (listData[0][10] == "") {

    sheetActive.getRange(numRow, 10).setValue(payment_date)
    sheetActive.getRange(numRow, 11).setValue(payment_sum)
    
    
    if (payment_sum  >= listData[0][5]) {
      sheetActive.getRange(numRow, 1, 1, lastColumn).setBackground("#B7E1CD")
      makePaymentToData(listData,payment_sum)
      ui.alert("Оплата счета на сумму " + payment_sum + " прошла упешно. Счет полность оплачен. Сумма по счету составляет " +( payment_sum + sum))
    }
    else {
      ui.alert("Оплата счета прошла упешно. Счет оплачен частично.")
    }
  }  
  else {

    if (payment_sum + sum >= listData[0][5]) {

      sheetActive.insertRows(numRow, 1)
      for (let i = 0;  i < listData[0].length; i++) {
        sheetActive.getRange(numRow, i + 1).setValue(listData[0][i])
      }
      sheetActive.getRange(numRow, 10).setValue(payment_date)
      sheetActive.getRange(numRow, 11).setValue(payment_sum)

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
            sheetActive.getRange(i + 1, 1, 1, lastColumn).setBackground("#B7E1CD")
          }  
      }

      makePaymentToData(listData,payment_date)
      ui.alert("Оплата счета на сумму " + payment_sum + " прошла упешно. Счет полность оплачен. Сумма по счету составляет " +( payment_sum + sum))

    }
    else {
      sheetActive.insertRows(numRow, 1)
      for (let i = 0;  i < listData[0].length; i++) {
        sheetActive.getRange(numRow, i + 1).setValue(listData[0][i])
      }
      sheetActive.getRange(numRow, 10).setValue(payment_date)
      sheetActive.getRange(numRow, 11).setValue(payment_sum)
      ui.alert("Оплата счета прошла упешно. Счет оплачен частично.")
    }
  }
  rangeSort.sort([{column: 5, ascending: true}, {column: 4, ascending: true}])
}

function makePaymentToData(listDataJournal,paymentDate) {

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh_data = ss.getSheetByName("Данные");
  let sh_oplata = ss.getSheetByName("Оплата");
  let listData = readRange(speadsheetBookkeeper.id, speadsheetBookkeeper.rangeSheetData)
  for (let i = 0; i < listData.length; i++ ) {
    if (listData[i][0] == listDataJournal[0][2]) {
      var numRow = i + 2
      break
    }
  }
  
  let range_sort = sh_data.getRange(2,1, sh_data.getLastRow()-1, sh_data.getLastColumn())
     
    
    
   // Получаю массив значений из элементов (Код, Дата Отгрузки, Нал, БЕЗНАЛ за 3 дней БЕЗ НДС, Конец периода 30 дней, Контрагент)
  let arr = sh_data.getRange(numRow, 1, 1, 8).getValues();
    
  let code = arr[0][0]
  let nal = arr[0][2]
  let beznal = arr[0][3]
  let days_30 = arr[0][4]
  let contragent = arr[0][5]
  let status = arr[0][7]
    
   
  sh_data.getRange(numRow,8).setValue("Оплатили")
     
  let date = Utilities.formatDate(days_30, "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'")
      
  let d = new Date(date) // Создали объект даты
     
  sh_data.getRange(numRow,5).setValue(new Date(d.getFullYear(), d.getMonth(), d.getDate() + 30)) // Добавили 30 дней к дате и изменили в ячейке.
     
  //Сортировка
  range_sort.sort({column: 5, ascending: true});       
    
  let rows_oplata = sh_oplata.getLastRow(); // кол-во строк
     
  // Добавляю данные по ячейкам в лист Оплата
  sh_oplata.getRange(rows_oplata + 1, 1,1).setValue(code).setBackground('#DEE9FC')
  sh_oplata.getRange(rows_oplata + 1, 2,1).setValue(nal).setBackground('#DEE9FC')
  sh_oplata.getRange(rows_oplata + 1, 3,1).setValue(beznal).setBackground('#DEE9FC')
  sh_oplata.getRange(rows_oplata + 1, 4,1).setValue(paymentDate).setBackground('#DEE9FC')
  sh_oplata.getRange(rows_oplata + 1, 5,1).setValue(contragent).setBackground('#DEE9FC')
  sh_oplata.getRange(rows_oplata + 1, 6,1).setValue(d)
  sh_oplata.getRange(rows_oplata + 1, 7,1).setValue(status)  
}

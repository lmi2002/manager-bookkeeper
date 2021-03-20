//Настраивать Google Table Файл -> Настройки таблицы Часовой пояс (GMT+02:00)Moscow-01-Kaliningrad

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = 
      [
        {name : "Провести оплату",functionName : "PromBox"},
        {name : "Выставить счет",functionName : "OpenFormDialog"}
      ]
  sheet.addMenu("Скрипты", entries);
}


function PromBox(){

  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var ss_manager_oplati = SpreadsheetApp.openById('1E3ocje0AdmhgvBTjre524ETOduY8XEw4Yd_g0SJ1mU0').getSheetByName('ОПЛАТЫ') // id "ДИРЕКТОР_version"
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
    
    var t = HtmlService.createTemplateFromFile('form_dialog')
     .evaluate()
     .setSandboxMode(HtmlService.SandboxMode.IFRAME)
     .setWidth(300)
     .setHeight(280)
     
    SpreadsheetApp.getUi()
      .showModalDialog(t, "Выставить счет");  
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
  var contragent_details_not_empty = true
  var obj_ss = getObjSpreadsheetApp()
  var values_list = obj_ss.values_list

  if (values_list[0][14].length == 0) {
    contragent_details_not_empty = false
  }

  if (date_invoice && date_act && contragent_details_not_empty) {
    obj_ss['act_sheet'].getRange(obj_ss['act_range'].getRow(), 8).setValue('Скрипт запущен')
    dict_invoice = Invoice(date_invoice, status_period, obj_ss)
    dict_act = Act(date_invoice, date_act,status_period, obj_ss)
    if (status_send_email) {
      sendInvoiceAndAct(dict_invoice, dict_act)
    }
      
    obj_ss['act_sheet'].getRange(obj_ss['act_range'].getRow(), 8).setValue('Счет выставлен')
  }
  else {
    var ui = SpreadsheetApp.getUi()
    ui.alert("Заполните реквизиты контрагента или дату на форме и повторите снова.")
  }
} 

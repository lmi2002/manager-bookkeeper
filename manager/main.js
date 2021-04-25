//Настраивать Google Table Файл -> Настройки таблицы Часовой пояс (GMT+02:00)Moscow-01-Kaliningrad
// Отладка команда debugger


var spreadsheet_nal = {
  
     'id': ID_SPREADSHEET_NAL,
     'range': 'Оплата!A2:I',
}

var spreadsheet_manager = {
  
     'id': SpreadsheetApp.getActiveSpreadsheet().getId(),
     'range_sheet_oplaty': 'Оплаты!A1',
     'range': 'Оплаты!A1:H'
}

var spreadsheet_bookkeeping = {
    
    'id': ID_BOOKKEPEPING,
    'range': 'Оплата!A2:H',

}

function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Скрипты')
      .addItem("Сдать в безнал", "RentBeznal")
      .addItem("Сдать в нал", "RentNal")
      .addItem("Вернуть","Refund")
      .addItem("Получить оплаты","GetPaymentDetails")
      .addItem("Первый счет","OpenFormDialog")
      .addToUi();
}


function GetDateEnd(spreadsheet_id, sheet_name, value) {
  var spreadsheet_id = spreadsheet_id
  var sheet_name = sheet_name
  var value = value  
  
  var ss = SpreadsheetApp.openById(spreadsheet_id).getSheetByName(sheet_name)
  var last_row = ss.getLastRow() + 1
  for (var i = 2; i < last_row; i++) {
    if (ss.getRange(i, 1).getValue() == value) {
      return ss.getRange(i, 5).getValue()
    }  
  }  
}


function RentBeznal(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bookkeeping = SpreadsheetApp.openById(ID_BOOKKEPEPING).getSheetByName('Данные') // id "Бухгалтерия_version"
  var range_code_bookkeeping = bookkeeping.getRange(2, 1, bookkeeping.getLastRow())
  var values_code_bookkeeping = range_code_bookkeeping.getValues()
  var len = values_code_bookkeeping.length
  var not_duplicate = true
  
  var sh_data = ss.getSheetByName("СВОБОДНЫЕ");
  var rent_beznal = ss.getSheetByName("СДАНО БЕЗНАЛ");
  var range = sh_data.getActiveRange();
  
  // Получаю диапазон выделенной ячейки
  var range_data = sh_data.getRange(range.getRow(), 1, 1, sh_data.getLastColumn())
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("ЛИСТ " + sh_data.getName(),'Переместить код: '+ range_data.getValues()[0][0] + " в лист " + rent_beznal.getName() + "?", ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
  
    for (var i = 0; i < len; i++ ){
      if(values_code_bookkeeping[i] == range_data.getValues()[0][0]){
        ui.alert("Код: " + range_data.getValues()[0][0] + " существует в книге Бухгалтерия в листе " + bookkeeping.getName());
        not_duplicate = false
        return not_duplicate
       }       
    }
    
    if (not_duplicate) {
  
      //CopyTo
      range_data.copyTo(rent_beznal.getRange(rent_beznal.getLastRow() + 1, 1, 1, rent_beznal.getLastColumn()))
  
      //AddToBookkeeping
      var bookkeeping_last_row = bookkeeping.getLastRow() + 1
      var values_range = range_data.getValues()
      var code = values_range[0][0]
      var date_start = values_range[0][1]
      var nal_30day = values_range[0][2]
      var beznal_30day = values_range[0][3]
      var date_end = values_range[0][8]
      var notes_of_manager = values_range[0][9]
      var partner = values_range[0][11]
      var couse = values_range[0][12]
      var contact = values_range[0][14] 
      var dop_contact = values_range[0][15]
      var email = values_range[0][17]
      var address = values_range[0][18]
    
      bookkeeping.getRange(bookkeeping_last_row,1).setValue(code)
      bookkeeping.getRange(bookkeeping_last_row,2).setValue(date_start)
      bookkeeping.getRange(bookkeeping_last_row,3).setValue(nal_30day)
      bookkeeping.getRange(bookkeeping_last_row,4).setValue(beznal_30day)
      bookkeeping.getRange(bookkeeping_last_row,5).setValue(date_end)
      bookkeeping.getRange(bookkeeping_last_row,6).setValue(partner)
      bookkeeping.getRange(bookkeeping_last_row,7).setValue(couse)
      bookkeeping.getRange(bookkeeping_last_row,8).setValue("Оплатили")
      bookkeeping.getRange(bookkeeping_last_row,9).setValue(notes_of_manager)
      bookkeeping.getRange(bookkeeping_last_row,11).setValue(contact)
      bookkeeping.getRange(bookkeeping_last_row,12).setValue(dop_contact)
      bookkeeping.getRange(bookkeeping_last_row,10).setValue(email)
      bookkeeping.getRange(bookkeeping_last_row,14).setValue(address)  
  
      sh_data.deleteRow(range.getRow())
     }
   }  
}

function RentNal(){


  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheet_nal = SpreadsheetApp.openById(ID_SPREADSHEET_NAL).getSheetByName('Данные') // id "Наличные_version"
  var range_code_spreadsheet_nal = spreadsheet_nal.getRange(2, 1, spreadsheet_nal.getLastRow())
  var values_code_spreadsheet_nal = range_code_spreadsheet_nal.getValues()
  var len = values_code_spreadsheet_nal.length
  var not_duplicate = true
  
  var sh_data = ss.getSheetByName("СВОБОДНЫЕ");
  var rent_nal = ss.getSheetByName("СДАНО НАЛ");
  var range = sh_data.getActiveRange();
  
  // Получаю диапазон выделенной ячейки
  var range_data = sh_data.getRange(range.getRow(), 1, 1, sh_data.getLastColumn())
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("ЛИСТ " + sh_data.getName(),'Переместить код: '+ range_data.getValues()[0][0] + " в лист " + rent_nal.getName() + "?", ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
  
    for (var i = 0; i < len; i++ ){
      if(values_code_spreadsheet_nal[i] == range_data.getValues()[0][0]){
        ui.alert("Код: " + range_data.getValues()[0][0] + " существует в книге Наличные в листе " + spreadsheet_nal.getName());
        not_duplicate = false
        return not_duplicate
       }       
    }
    
    if (not_duplicate) {
  
      //CopyTo
      range_data.copyTo(rent_nal.getRange(rent_nal.getLastRow() + 1, 1, 1, rent_nal.getLastColumn()))
  
      //AddToSpreadsheetNal
      var spreadsheet_nal_last_row = spreadsheet_nal.getLastRow() + 1
      var values_range = range_data.getValues()
      var code = values_range[0][0]
      var date_start = values_range[0][1]
      var nal_30day = values_range[0][2]
      var beznal_30day = values_range[0][3]
      var date_end = values_range[0][8]
      var notes_of_manager = values_range[0][9]
      var partner = values_range[0][11]
      var couse = values_range[0][12]
      var contact = values_range[0][14] 
      var dop_contact = values_range[0][15]
      var email = values_range[0][17]
      var address = values_range[0][18]
    
      spreadsheet_nal.getRange(spreadsheet_nal_last_row,1).setValue(code)
      spreadsheet_nal.getRange(spreadsheet_nal_last_row,2).setValue(date_start)
      spreadsheet_nal.getRange(spreadsheet_nal_last_row,3).setValue(nal_30day)
      spreadsheet_nal.getRange(spreadsheet_nal_last_row,4).setValue(beznal_30day)
      spreadsheet_nal.getRange(spreadsheet_nal_last_row,5).setValue(date_end)
      spreadsheet_nal.getRange(spreadsheet_nal_last_row,6).setValue(contact)
      spreadsheet_nal.getRange(spreadsheet_nal_last_row,7).setValue(couse)
      spreadsheet_nal.getRange(spreadsheet_nal_last_row,8).setValue("Оплатили")
      spreadsheet_nal.getRange(spreadsheet_nal_last_row,9).setValue(notes_of_manager)
      spreadsheet_nal.getRange(spreadsheet_nal_last_row,11).setValue(dop_contact)
      spreadsheet_nal.getRange(spreadsheet_nal_last_row,12).setValue(address)
      spreadsheet_nal.getRange(spreadsheet_nal_last_row,10).setValue(email) 
  
      sh_data.deleteRow(range.getRow())
     }
   }  
}


function Refund(){

  

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var spreadsheet_active;
  
  if (sheet.getName() == "СДАНО НАЛ" || sheet.getName() == "СДАНО БЕЗНАЛ") {
     
    switch(sheet.getName()) {
      case "СДАНО НАЛ":
        spreadsheet_active = SpreadsheetApp.openById(ID_SPREADSHEET_NAL).getSheetByName('Данные');// id "Наличные_version"
        break;
      case "СДАНО БЕЗНАЛ":
        spreadsheet_active = SpreadsheetApp.openById(ID_BOOKKEPEPING).getSheetByName('Данные'); //id "Бухгалтерия_version"
        break;
      }
      
    var range_code_spreadsheet_active = spreadsheet_active.getRange(2, 1, spreadsheet_active.getLastRow())
    var values_code_spreadsheet_active = range_code_spreadsheet_active.getValues()
    var len = values_code_spreadsheet_active.length
  

    var sh_data = ss.getSheetByName("СВОБОДНЫЕ");
    var arhive = ss.getSheetByName("АРХИВ");
    var range = sheet.getActiveRange();
  
    // Получаю диапазон выделенной ячейки
    var range_data = sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn())
    
    var code = range_data.getValues()[0][0]
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert("ЛИСТ " + sheet.getName(),'Переместить код: '+ code + " в лист " + sh_data.getName() + "?", ui.ButtonSet.YES_NO);

    if (response == ui.Button.YES) {
    
    //Изменяем в файле Директор по ключу столбца "код" информацию в столбце "конец срока"
    
    switch(sheet.getName()) {
      case "СДАНО НАЛ":
        var date_end = GetDateEnd(ID_SPREADSHEET_NAL, 'Данные', code)
        sheet.getRange(range.getRow(), 9).setValue(date_end)
        break;
      case "СДАНО БЕЗНАЛ":
        var date_end = GetDateEnd(ID_BOOKKEPEPING, 'Данные', code)
        sheet.getRange(range.getRow(), 9).setValue(date_end)
        break;
     }
     
      //CopyTo
      range_data.copyTo(sh_data.getRange(sh_data.getLastRow() + 1, 1, 1, sh_data.getLastColumn()))
      arhive.insertRowBefore(2);
      range_data.copyTo(arhive.getRange(2, 1, 1, arhive.getLastColumn()))
  
      //Удаляет строку с файла Бухгалтерия 
      for (var i = 0; i < len; i++ ){
        if(values_code_spreadsheet_active[i] == range_data.getValues()[0][0]){
          spreadsheet_active.deleteRow(i + 2)
          break
         }
      }
     
      sheet.deleteRow(range.getRow())
    }
  }
  
}

function GetPaymentDetails() {

  var sheet_oplaty = SpreadsheetApp.openById(spreadsheet_manager['id']).getSheetByName('ОПЛАТЫ')
  
  var bookkeeping_filter_data_sum = 0
  var nal_filter_data_vlad_sum = 0
  var nal_filter_data_sum = 0
  
  
  var bookkeeping_data = readRange(spreadsheet_bookkeeping['id'], spreadsheet_bookkeeping['range'])
  var bookkeeping_filter_data = getFilterDataAddFirstEmpty(bookkeeping_data, 7, 't')

  if (bookkeeping_filter_data.length > 0) {
  
    bookkeeping_filter_data_sum = sum(bookkeeping_filter_data, 2)
  
    var bookkeeping_cells_list = getCells(bookkeeping_data, 7, 't', 'Оплата!H')
    
    var range_1 = rangeConcat('ОПЛАТЫ!A', sheet_oplaty.getLastRow() + 1, 'H', sheet_oplaty.getLastRow() + bookkeeping_filter_data.length)
    
    appendToMultipleRanges(bookkeeping_filter_data, spreadsheet_manager['id'], spreadsheet_manager['range_sheet_oplaty']) 
    
    sheet_oplaty.getRange(range_1).setBackground('#FBEC5D')
    
  
    updateToMultipleRanges([['t']], spreadsheet_bookkeeping['id'], bookkeeping_cells_list)
  }  
    
  
  var nal_data = readRange(spreadsheet_nal['id'], spreadsheet_nal['range'])
  var nal_filter_data = getFilterData(nal_data, 8, 't')
   
  
  if (nal_filter_data.length > 0) {
    
    nal_filter_data_vlad_sum = sum(nal_filter_data, 0)
    nal_filter_data_sum = sum(nal_filter_data, 2)
    
    var nal_cells_list = getCells(nal_data, 8, 't', 'Оплата!I')
    
    var range_2 = rangeConcat('ОПЛАТЫ!A', sheet_oplaty.getLastRow() + 1, 'H', sheet_oplaty.getLastRow() + nal_filter_data.length)
    
    appendToMultipleRanges(nal_filter_data, spreadsheet_manager['id'], spreadsheet_manager['range_sheet_oplaty'])
    
    
    sheet_oplaty.getRange(range_2).setBackground('#ADDFAD')
    
    updateToMultipleRanges([['t']], spreadsheet_nal['id'], nal_cells_list)
     
  }
  
  
  if (bookkeeping_filter_data_sum > 0 || nal_filter_data_vlad_sum > 0 || nal_filter_data_sum > 0) {
    
    var date = getNowDate()
    
    var nal_sum = bookkeeping_filter_data_sum + nal_filter_data_sum
    
    var total = nal_filter_data_vlad_sum + nal_sum

    appendToMultipleRanges([[date, nal_filter_data_vlad_sum, nal_sum, total]], spreadsheet_manager['id'], spreadsheet_manager['range_sheet_oplaty'])
    
  }  
}

function OpenFormDialog() {
  
  var obj = getObjSpreadsheetApp()
  
  if (obj['act_sheet'].getName() == 'СВОБОДНЫЕ') {
    
    var t = HtmlService.createTemplateFromFile('form_dialog')
     .evaluate()
     .setSandboxMode(HtmlService.SandboxMode.IFRAME)
     .setWidth(300)
     .setHeight(310)
     
    SpreadsheetApp.getUi()
      .showModalDialog(t, "Выставить счет");  
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
  var obj_ss = getObjSpreadsheetApp()

 if (date_invoice && date_act) {
    obj_ss['act_sheet'].getRange(obj_ss['act_range'].getRow(), 14).setValue('Скрипт запущен')
    dict_invoice = Invoice(date_invoice, status_period, status_delivery, obj_ss)
    dict_act = Act(date_invoice, date_act, status_period, status_delivery, obj_ss)
    if (status_send_email) {
      sendInvoiceAndAct(dict_invoice, dict_act)
    }
    obj_ss['act_sheet'].getRange(obj_ss['act_range'].getRow(), 14).setValue('Счет выставлен')
  }
  else {
    var ui = SpreadsheetApp.getUi()
    ui.alert("Заполните реквизиты контрагента или дату на форме и повторите снова.")
  }
} 

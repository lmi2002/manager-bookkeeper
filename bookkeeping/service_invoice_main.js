function Invoice(date) {
  
  var date = date
  
  if (date) {
    var dict_copy_invoice_file  = copyInvoiceFile(date)
    var file_dict = exportSpreadsheetToXlsx(dict_copy_invoice_file, 'xlsx');
    var dict_move_files =  moveFiles(file_dict, date)
    sendInvoiceToEmail(dict_move_files)
    var obj = getObjSpreadsheetApp()
    obj['act_sheet'].getRange(obj['act_range'].getRow(), 8).setValue('Счет выставлен')
    return true
  }
  else {
    var ui = SpreadsheetApp.getUi()
    ui.alert("Вы не заполнили дату. Повторите снова!")
  }
  
}
function getStrNowDay() {
  return Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd")
}

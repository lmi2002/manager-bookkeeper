function Invoice(date) {
  
  var date = date
  
  if (date) {
    var dict_save_invoice_format_xlsx  = saveInvoiceFormatXlsx(date)
    var dict_move_files =  moveFiles(dict_save_invoice_format_xlsx, date)
    sendInvoiceToEmail(dict_move_files)
    var obj = getObjSpreadsheetApp()
    obj['act_sheet'].getRange(obj['act_range'].getRow(), 8).setValue('Счет выставлен')
  }
  else {
    var ui = SpreadsheetApp.getUi()
    ui.alert("Вы не заполнили дату. Повторите снова!")
  }
  
}
function getStrNowDay() {
  return Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd")
}

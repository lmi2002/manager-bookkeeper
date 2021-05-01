function Invoice(date_invoice, status_period, status_delivery, obj_ss) {
  
  var date_invoice = date_invoice
  var status_period = status_period
  var status_delivery = status_delivery
  var obj_ss = obj_ss
  var dict_copy_invoice_file  = copyInvoiceFile(date_invoice,status_period, status_delivery, obj_ss)
  var file_dict = exportSpreadsheetToXlsx(dict_copy_invoice_file, 'xlsx');
  return moveInvoiceFiles(file_dict, date_invoice)
}

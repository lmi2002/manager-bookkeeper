function Invoice(date_invoice, status_period, obj_ss) {
  
  var date_invoice = date_invoice
  var status_period = status_period
  var values_list = obj_ss.values_list
  var dict_copy_invoice_file  = copyInvoiceFile(date_invoice,status_period, values_list)
  var file_dict = exportSpreadsheetToXlsx(dict_copy_invoice_file, 'xlsx');
  return moveInvoiceFiles(file_dict, date_invoice)
}

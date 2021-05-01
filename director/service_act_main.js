function Act(date_invoice, date_act,status_period, status_delivery, obj_ss)  {
  
  var date_invoice = date_invoice
  var date_act = date_act
  var status_period = status_period
  var status_delivery = status_delivery
  var obj_ss = obj_ss
  var dict_copy_act_file  = copyActFile(date_invoice, date_act, status_period, status_delivery, obj_ss)
  var file_dict = exportSpreadsheetToXlsx(dict_copy_act_file, 'xlsx');
  return moveActFiles(file_dict)   
}

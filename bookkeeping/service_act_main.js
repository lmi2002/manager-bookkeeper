function Act(date_invoice, date_act,status_period, obj_ss)  {
  
  var date_invoice = date_invoice
  var date_act = date_act
  var status_period = status_period
  var values_list = obj_ss.values_list
  var dict_copy_act_file  = copyActFile(date_invoice, date_act, status_period, values_list)
  var file_dict = exportSpreadsheetToXlsx(dict_copy_act_file, 'xlsx');
  return moveActFiles(file_dict)   
}

function moveActFiles(dict) {
  var dict = dict
  var folder_name = 'АКТЫ ' + dict['contragent']
  var folder = getFolders(folder_name)
  
  if (!folder) {
    folder = DriveApp.createFolder(folder_name)
  }
  dict['file'].moveTo(folder)
  return dict      
}

function copyActFile(date_invoice, date_act, status_period, status_delivery, obj_ss) {
 
  var date_invoice = new Date(date_invoice)
  var date_act = new Date(date_act)
  var status_period = status_period
  var status_delivery = status_delivery
  var values_list = obj_ss.values_list

  var date_invoice_dd_mm_yy = Utilities.formatDate(date_invoice, "GMT", "dd-MM-yy")
  var date_act_dd_mm_yyyy = Utilities.formatDate(date_act, "GMT", "dd.MM.yyyy")

  var start_date = values_list[0][1]
  var start_date_dd_mm_yyyy =  new Date(start_date.getFullYear(), start_date.getMonth(), start_date.getDate() + 1)
  var start_date_format = Utilities.formatDate(start_date_dd_mm_yyyy, "GMT", "dd.MM.yyyy")

  var finish_date = values_list[0][8]
  var finish_date_dd_mm_yyyy = Utilities.formatDate(finish_date, "GMT", "dd.MM.yyyy")
   
  var code = values_list[0][0]
  var contragent = values_list[0][11]
  var beznal = values_list[0][3]
  var sum_contract = values_list[0][6]
  var sd = sum_contract * 1
  sd.toLocaleString()

  var act_num = 'Акт № ' + code + "/" + date_invoice_dd_mm_yy + "/1"
  var contract = obj_ss['act_sheet'].getRange(obj_ss['act_range'].getRow(), 13).getValue()
  var contract_str = 'Ми, що нижче підписалися, ВИКОНАВЕЦЬ - ФОП Гордєєв Родіон Вікторович з однієї сторони, та ЗАМОВНИК -'  + contragent + ', з другої сторони, склали цей акт про те, що згідно  договору ' + contract + ' наступні роботи (послуги):'
  
     
  var name_service = "Оренда вагона будівельного згідно договору за період " + start_date_format + " - " +  finish_date_dd_mm_yyyy

  var ss_template_act = SpreadsheetApp.openById(ID_TEMPLATE_FIRST_ACT)
  var ss_copy_act = ss_template_act.copy('Акт 1 ' + code + "/" + date_invoice_dd_mm_yy + "/" + contragent)
  var sh_ss_copy_act = ss_copy_act.getSheets()[0]
  sh_ss_copy_act.getRange('E1').setValue(act_num)
  sh_ss_copy_act.getRange('D3').setValue('складений ' + date_act_dd_mm_yyyy + " року")
  sh_ss_copy_act.getRange('A5').setValue(contract_str)
  if (status_period) {
    sh_ss_copy_act.getRange('B8').setValue(name_service)
  }
  else {
    sh_ss_copy_act.getRange('B8').setValue("Оренда вагона будівельного згідно договору")
  }

  if (status_delivery) {
    sh_ss_copy_act.getRange('I8').setValue(beznal * values_list[0][7])
    sh_ss_copy_act.getRange('J9').setValue(values_list[0][5])
    sh_ss_copy_act.getRange('J10').setValue(sum_contract)
    sh_ss_copy_act.getRange('A12').setValue("Ціна робіт (послуг) складає " + sd.toLocaleString() + ",00 грн")
    sh_ss_copy_act.getRange('A14').setValue("Разом " + sd.toLocaleString() + ",00 (" + NumberInWords(sum_contract) + ") без ПДВ")
    sh_ss_copy_act.getRange('G19').setValue(contragent)
  }
  else {
    sh_ss_copy_act.hideRow(sh_ss_copy_act.getRange('A9:Z9'))
    sh_ss_copy_act.getRange('I8').setValue(sum_contract)
    sh_ss_copy_act.getRange('J9').setValue(sum_contract)
    sh_ss_copy_act.getRange('A12').setValue("Ціна робіт (послуг) складає " + sd.toLocaleString() + ",00 грн")
    sh_ss_copy_act.getRange('A14').setValue("Разом " + sd.toLocaleString() + ",00 (" + NumberInWords(sum_contract) + ") без ПДВ")
    sh_ss_copy_act.getRange('G19').setValue(contragent)
  }
    
  SpreadsheetApp.flush()
  return {
          'ss': ss_copy_act,
          'contragent': contragent,
          "act_num": act_num
  }
}

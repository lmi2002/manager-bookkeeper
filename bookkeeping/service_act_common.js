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

function copyActFile(date_invoice, date_act, status_period, values_list) {
 
  var date_invoice = new Date(date_invoice)
  var date_act = new Date(date_act)
  var status_period = status_period
  var values_list = values_list

  var date_invoice_dd_mm_yy = Utilities.formatDate(date_invoice, "GMT", "dd-MM-yy")
  var date_act_dd_mm_yyyy = Utilities.formatDate(date_act, "GMT", "dd.MM.yyyy")

   
  var code = values_list[0][0]
  var contragent = values_list[0][5]
  var beznal = values_list[0][3]

  // Formating date
  var days_30 = new Date(values_list[0][4])
  var d = new Date(days_30.getFullYear(), days_30.getMonth(), days_30.getDate() + 1 )
  var days_30_format = Utilities.formatDate(d, "GMT", "dd.MM.yyyy")

  var add_30_days = new Date(days_30.getFullYear(), days_30.getMonth(), days_30.getDate() + 30)
  var add_30_days_format = Utilities.formatDate(add_30_days, "GMT", "dd.MM.yyyy")

  var contragent_details = values_list[0][14]
  var act_num = 'Акт № ' + code + "/" + date_invoice_dd_mm_yy

  var contract_str = 'Ми, що нижче підписалися, ВИКОНАВЕЦЬ - ФОП ' + DIRECTOR_NAME_ENTIRE + ' з однієї сторони, та ЗАМОВНИК -'  + contragent + ', з другої сторони, склали цей акт про те, що згідно  договору ' + values_list[0][6] + ' наступні роботи (послуги):'
  
     
  var name_service = "Оренда вагона будівельного за період 30 календарних днів (" + days_30_format + " - " +  add_30_days_format + ")"

  var ss_template_act = SpreadsheetApp.openById(ID_TEMPLATE_ACT)
  var ss_copy_act = ss_template_act.copy('Акт ' + code + "/" + date_invoice_dd_mm_yy + "/" + contragent)
  var sh_ss_copy_act = ss_copy_act.getSheets()[0]
  sh_ss_copy_act.getRange('E1').setValue(act_num)
  sh_ss_copy_act.getRange('D3').setValue('складений ' + date_act_dd_mm_yyyy + " року")
  sh_ss_copy_act.getRange('A5').setValue(contract_str)
  sh_ss_copy_act.getRange('I8').setValue(beznal)
  if (status_period) {
    sh_ss_copy_act.getRange('B8').setValue(name_service)
  }
  else {
    sh_ss_copy_act.getRange('B8').setValue("Оренда вагона будівельного")
  }
  sh_ss_copy_act.getRange('A13').setValue("Разом " + beznal + ",00 (" + NumberInWords(beznal) + ") без ПДВ")
  sh_ss_copy_act.getRange('G18').setValue(contragent_details)
    
  SpreadsheetApp.flush()
  return {
          'ss': ss_copy_act,
          'contragent': contragent,
          "act_num": act_num
  }
}

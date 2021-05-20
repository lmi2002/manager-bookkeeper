function moveInvoiceFiles(dict, date) {
  var dict = dict
  var date = date
  var year = Utilities.formatDate(new Date(date), "GMT", "yyyy")
  var month = Utilities.formatDate(new Date(date), "GMT", "MM")
  var folder_name = 'Счета ' + getTextMonth(month, 'folder') + " " + year
  var folder = getFolders(folder_name)
  
  if (!folder) {
    folder = DriveApp.createFolder(folder_name)
  }
  dict['file'].moveTo(folder)
  return dict        
}

function getTextMonth(month, str_doc) {

  var str_month = {
    "01": {
      "folder":"январь",
      "invoice":"січня"
    },
    "02": {
      "folder":"февраль",
      "invoice":"лютого"
    },
    "03": {
      "folder":"март",
      "invoice":"березня"
    },
    "04":{
      "folder":"апрель",
      "invoice":"квітня"
    },
    "05": {
      "folder":"май",
      "invoice":"травня"
    },
    "06": {
      "folder":"июнь",
      "invoice":"червня"
    },
    "07": {
      "folder":"июль",
      "invoice":"липня"
    },
    "08": {
      "folder":"август",
      "invoice":"серпня"
    },
    "09": {
      "folder":"сентябрь",
      "invoice":"вересня"
    },
    "10": {
      "folder":"октябрь",
      "invoice":"жовтня"
    },
    "11": {
      "folder":"ноябрь",
      "invoice":"листопада"
    },
    "12": {
      "folder":"декабрь",
      "invoice":"грудня"
    }
  }
  return str_month[month][str_doc]
}

function copyInvoiceFile(date_invoice, status_period, status_delivery, obj_ss) {
  
  var date_invoice = new Date(date_invoice)
  var status_period = status_period
  var status_delivery = status_delivery
  var obj_ss = obj_ss
  var values_list = obj_ss.values_list
  
  var dd_mm = Utilities.formatDate(date_invoice, "GMT", "dd-MM")
  var dd = Utilities.formatDate(date_invoice, "GMT", "dd")
  var mm = Utilities.formatDate(date_invoice, "GMT", "MM")
  var yyyy = Utilities.formatDate(date_invoice, "GMT", "yyyy")

  var date_invoice_dd_mm_yyyy = Utilities.formatDate(date_invoice, "GMT", "dd.MM.yyyy")

  var start_date = values_list[0][1]
  var start_date_dd_mm_yyyy =  new Date(start_date.getFullYear(), start_date.getMonth(), start_date.getDate() + 1)
  var start_date_format = Utilities.formatDate(start_date_dd_mm_yyyy, "GMT", "dd.MM.yyyy")

  var finish_date = values_list[0][8]
  var finish_date_dd_mm_yyyy = Utilities.formatDate(finish_date, "GMT", "dd.MM.yyyy")
   
  var code = values_list[0][0]
  var contragent = values_list[0][11]
  var email = values_list[0][17]
  var beznal = values_list[0][3]
  var sum_contract = values_list[0][6]
  var contract = '№  ' + code + "/" + dd_mm + " від " + date_invoice_dd_mm_yyyy + " року"
  var contract_str = 'Договір оренди ' + contract
  var invoice_num = 'Рахунок № ' + code + "/" + dd_mm
  var name_service = "Оренда вагона будівельного згідно договору за період "+ start_date_format + " - " +  finish_date_dd_mm_yyyy

  var ss_template_invoice = SpreadsheetApp.openById(ID_TEMPLATE_FIRST_INVOICE)
  var ss_copy_invoice = ss_template_invoice.copy('Счет 1 ' + code + "/" + dd_mm + "/" + contragent)
  var sh_ss_copy_invoice = ss_copy_invoice.getSheets()[0]
  sh_ss_copy_invoice.getRange('D9').setValue(contragent)
  sh_ss_copy_invoice.getRange('D14').setValue(invoice_num)
  sh_ss_copy_invoice.getRange('D12').setValue(contract_str)
  sh_ss_copy_invoice.getRange('D15').setValue('від ' + dd + " " + getTextMonth(mm,'invoice') + " " + yyyy + " р.")
  if (status_period) {
    sh_ss_copy_invoice.getRange('C18').setValue(name_service)
  }
  else {
    sh_ss_copy_invoice.getRange('C18').setValue("Оренда вагона будівельного згідно договору")
  }
  if (status_delivery) {
    sh_ss_copy_invoice.getRange('G18').setValue(beznal * values_list[0][7])
    sh_ss_copy_invoice.getRange('H19').setValue(values_list[0][5])
    sh_ss_copy_invoice.getRange('H20').setValue(sum_contract)  
    sh_ss_copy_invoice.getRange('D22').setValue(NumberInWords(sum_contract))
  }
  else {
    sh_ss_copy_invoice.getRange('A19:I19').deleteCells(SpreadsheetApp.Dimension.ROWS)
    sh_ss_copy_invoice.getRange('G18').setValue(sum_contract)
    sh_ss_copy_invoice.getRange('H19').setValue(sum_contract) 
    sh_ss_copy_invoice.getRange('D21').setValue(NumberInWords(sum_contract))
  }

  obj_ss['act_sheet'].getRange(obj_ss['act_range'].getRow(), 13).setValue(contract)
    
  SpreadsheetApp.flush()
  return {
          'ss': ss_copy_invoice,
          "invoice_num": invoice_num,
          "email": email,
          "code": code
  }
}

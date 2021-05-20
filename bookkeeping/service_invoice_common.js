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

function copyInvoiceFile(date_invoice,status_period, values_list) {
  
  var date_invoice = new Date(date_invoice)
  var status_period = status_period
  var values_list = values_list
  
  var dd_mm = Utilities.formatDate(date_invoice, "GMT", "dd-MM")
  var dd = Utilities.formatDate(date_invoice, "GMT", "dd")
  var mm = Utilities.formatDate(date_invoice, "GMT", "MM")
  var yyyy = Utilities.formatDate(date_invoice, "GMT", "yyyy")
  // var yy = Utilities.formatDate(date_invoice, "GMT", "yy")

  // Formating date
  var days_30 = new Date(values_list[0][4])
  var d = new Date(days_30.getFullYear(), days_30.getMonth(), days_30.getDate() + 1 )
  var days_30_format = Utilities.formatDate(d, "GMT", "dd.MM.yyyy")

  var add_30_days = new Date(days_30.getFullYear(), days_30.getMonth(), days_30.getDate() + 30)
  var add_30_days_format = Utilities.formatDate(add_30_days, "GMT", "dd.MM.yyyy")
   
  var code = values_list[0][0]
  var contragent = values_list[0][5]
  var email = values_list[0][9]
  var beznal = values_list[0][3]
  var contract_str = 'Договір оренди ' + values_list[0][6]
  // var invoice_num = 'Рахунок № ' + code + "/" + dd_mm - yy
  var invoice_num = 'Рахунок № ' + code + "/" + dd_mm
  var name_service = "Оренда вагона будівельного за період 30 календарних днів (" + days_30_format + " - " +  add_30_days_format + ")"

  var ss_template_invoice = SpreadsheetApp.openById(ID_TEMPLATE_INVOICE)
  var ss_copy_invoice = ss_template_invoice.copy('Счет ' + code + "/" + dd_mm + "/" + contragent)
  var sh_ss_copy_invoice = ss_copy_invoice.getSheets()[0]
  sh_ss_copy_invoice.getRange('D9').setValue(contragent)
  sh_ss_copy_invoice.getRange('D14').setValue(invoice_num)
  sh_ss_copy_invoice.getRange('D12').setValue(contract_str)
  sh_ss_copy_invoice.getRange('D15').setValue('від ' + dd + " " + getTextMonth(mm,'invoice') + " " + yyyy + " р.")
  sh_ss_copy_invoice.getRange('G18').setValue(beznal)
  sh_ss_copy_invoice.getRange('D21').setValue(NumberInWords(beznal))
  if (status_period) {
    sh_ss_copy_invoice.getRange('C18').setValue(name_service)
  }
  else {
    sh_ss_copy_invoice.getRange('C18').setValue("Оренда вагона будівельного")
  }

    
  SpreadsheetApp.flush()
  return {
          'ss': ss_copy_invoice,
          "invoice_num": invoice_num,
          "email": email,
          "contract_str": contract_str
  }
}

function addInformInvoiceToInvoiceJournal(date_invoice, dict_invoice, obj_ss) {
  var date_invoice = new Date(date_invoice)
  var date_invoice_format = Utilities.formatDate(date_invoice, "GMT", "dd.MM.yyyy")
  var obj_ss = obj_ss
  var ss = obj_ss.act_ss
  var invoice_num = dict_invoice.invoice_num.slice(10)
  var values_list = obj_ss.values_list
  var code = values_list[0][0]
  var beznal = values_list[0][3]
  var contragent = values_list[0][5]
  var contragent_code = values_list[0][15]
  var num_contract = values_list[0][6]


  var sheet_journal = ss.getSheetByName("Журнал")
  var last_row = sheet_journal.getLastRow() + 1
  sheet_journal.getRange(last_row, 1).setValue(contragent)
  sheet_journal.getRange(last_row, 2).setValue(contragent_code)
  sheet_journal.getRange(last_row, 3).setValue(code)
  sheet_journal.getRange(last_row, 4).setValue(invoice_num)
  sheet_journal.getRange(last_row, 5).setValue(date_invoice_format)
  sheet_journal.getRange(last_row, 6).setValue(beznal)
  sheet_journal.getRange(last_row, 7).setValue(getStrNowDay_1())
  sheet_journal.getRange(last_row, 12).setValue(num_contract)
}

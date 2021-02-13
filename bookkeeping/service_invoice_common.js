function moveFiles(dict, date) {
  var dict = dict
  var date = date
  var year = Utilities.formatDate(new Date(date), "GMT", "yyyy")
  var month = Utilities.formatDate(new Date(date), "GMT", "MM")
  var file = DriveApp.getFileById(dict['id']);
  var folder_name = 'Счета ' + getTextMonth(month, 'folder') + " " + year
  var folder = getFolders(folder_name)
  
  if (!folder) {
    folder = DriveApp.createFolder(folder_name)
  }
  file.moveTo(folder)
  return {
          "file":file,
          "invoice_num": dict['invoice_num'],
          "email": dict['email']
         }
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

function saveInvoiceFormatXlsx(date) {
  
  var new_date = new Date(date)
  
  var dd_mm = Utilities.formatDate(new_date, "GMT", "dd-MM")
  var dd_mm_yyyy = Utilities.formatDate(new_date, "GMT", "dd.MM.yyyy")
  var dd = Utilities.formatDate(new_date, "GMT", "dd")
  var mm = Utilities.formatDate(new_date, "GMT", "MM")
  var yyyy = Utilities.formatDate(new_date, "GMT", "yyyy")
   
  var values_list = getObjSpreadsheetApp()['values_list']
  var code = values_list[0][0]
  var contragent = values_list[0][5]
  var email = values_list[0][9]
  var beznal = values_list[0][3]
  var contract_str = 'Договір оренди ' + values_list[0][6]

  var ss_template_invoice = SpreadsheetApp.openById(ID_TEMPLATE_INVOICE)
  var ss_copy_invoice = ss_template_invoice.copy('Счет ' + code + "/" + dd_mm + "/" + contragent)
  var sh_ss_copy_invoice = ss_copy_invoice.getSheets()[0]
  sh_ss_copy_invoice.getRange('D9').setValue(contragent)
  sh_ss_copy_invoice.getRange('D14').setValue('Рахунок № ' + code + "/" + dd_mm)
  sh_ss_copy_invoice.getRange('D12').setValue(contract_str)
  sh_ss_copy_invoice.getRange('D15').setValue('від ' + dd + " " + getTextMonth(mm,'invoice') + " " + yyyy + " р.")
  sh_ss_copy_invoice.getRange('G18').setValue(beznal)
  sh_ss_copy_invoice.getRange('D21').setValue(NumberInWords(beznal))
    
  SpreadsheetApp.flush()
  return {
          "id": ss_copy_invoice.getId(),
          "invoice_num": 'Рахунок № ' + code + "/" + dd_mm,
          "email": email,
          "contract_str": contract_str
  }
}

function getFolders(folderName) {      
  var folders = DriveApp.getFolders();
  
  while (folders.hasNext()) {
    var folder = folders.next();
     if(folderName == folder.getName()) {         
       return folder;
     }
   }
  return null;
}

function sendInvoiceToEmail(dict) {

  GmailApp.sendEmail(dict['email'],"Бытовки Харьков " + dict['invoice_num'], 'Пожалуйста посмотрите прикрепленный файл.', {
    attachments: [dict['file'].getAs('application/pdf')],
    htmlBody: getLetterBody(),
    name: 'Бытовки Харьков'
    })
}

function getObjSpreadsheetApp() {
  
  var act_ss = SpreadsheetApp.getActiveSpreadsheet();
  var act_range = act_ss.getActiveRange()
  var act_sheet = act_ss.getActiveSheet()
  var values_list = act_sheet.getRange(act_range.getRow(), 1, 1, 10).getValues();
  
  return {
    "act_ss": act_ss,
    "act_range": act_range,
    "act_sheet": act_sheet,
    "values_list": values_list
  }
}


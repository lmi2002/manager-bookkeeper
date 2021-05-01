function getStrNowDay() {
  return Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd")
}

function getFinishDate() {
  var finish_date = getObjSpreadsheetApp().values_list[0][8]
  return Utilities.formatDate(new Date(finish_date), "GMT", "yyyy-MM-dd")
}

function getObjSpreadsheetApp() {
  
  var act_ss = SpreadsheetApp.getActiveSpreadsheet();
  var act_range = act_ss.getActiveRange()
  var act_sheet = act_ss.getActiveSheet()
  var values_list = act_sheet.getRange(act_range.getRow(), 1, 1, act_sheet.getLastColumn()).getValues();
  
  return {
    "act_ss": act_ss,
    "act_range": act_range,
    "act_sheet": act_sheet,
    "values_list": values_list
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

function exportSpreadsheetToXlsx(dict, type) {
  /* globals __SNIPPETS__TYPES__EXPORT__SHEET__ */
  const type_ = __SNIPPETS__TYPES__EXPORT__SHEET__[type];
  const url = Drive.Files.get(dict['ss'].getId()).exportLinks[type_];
  const blob = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
    },
  })
  var file = DriveApp.createFile(blob).setName(dict['ss'].getName() + '.' + type)
  DriveApp.getFileById(dict['ss'].getId()).setTrashed(true)
  dict['file'] = file
  return dict
}


(function(scope) {
  const TYPES = {
    'application/x-vnd.oasis.opendocument.spreadsheet':
      'application/x-vnd.oasis.opendocument.spreadsheet',
    'application/vnd.oasis.opendocument.spreadsheet':
      'application/vnd.oasis.opendocument.spreadsheet',
    'ods': 'application/x-vnd.oasis.opendocument.spreadsheet',
    'text/tab-separated-values': 'text/tab-separated-values',
    'tsv': 'text/tab-separated-values',
    'application/pdf': 'application/pdf',
    'pdf': 'application/pdf',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'text/csv': 'text/csv',
    'csv': 'text/csv',
    'application/zip': 'application/zip',
    'zip': 'application/zip',
  };
  scope.__SNIPPETS__TYPES__EXPORT__SHEET__ = TYPES;
})(this);


function  sendInvoiceAndAct(dict_invoice, dict_act) {
  GmailApp.sendEmail(dict_invoice['email'],"Счёт и акт на бытовку " + dict_invoice['code'], 'Пожалуйста посмотрите прикрепленный файл.', {  
    attachments: [dict_invoice['ss'].getAs('application/pdf'),dict_act['ss'].getAs('application/pdf')],
    htmlBody: getLetterBody(),
    name: 'Бытовки Харьков', 
    bcc: OFFICE_EMAIL
  })
}

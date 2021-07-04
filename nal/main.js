//Настраивать Google Table Файл -> Настройки таблицы Часовой пояс (GMT+02:00)Moscow-01-Kaliningrad
// Отладка команда debugger

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = 
      [
        {name : "Провести оплату",functionName : "OpenMyDialog"}
      ]
  sheet.addMenu("Скрипты", entries);
}



function Payment(person, date){

  var person = person
  var date_oplata = date
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh_data = ss.getSheetByName("Данные");
  var sh_oplata = ss.getSheetByName("Оплата");
  var range = sh_data.getActiveRange();
  var range_sort = sh_data.getRange(2,1, sh_data.getLastRow()-1, sh_data.getLastColumn())   
    
   // Получаю массив значений из элементов (Код, Дата Отгрузки, Нал, БЕЗНАЛ за 3 дней БЕЗ НДС, Конец периода 30 дней, Контрагент)
   var arr = sh_data.getRange(range.getRow(), 1, 1, 8).getValues();
      
   var code = arr[0][0]
   var nal = arr[0][2]
   var beznal = arr[0][3]
   var days_30 = arr[0][4]
   var status = arr[0][7]
    
      
   sh_data.getRange(range.getRow(),8).setValue("Оплатили")
     
   var date = Utilities.formatDate(days_30, "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'")
      
   var d = new Date(date) // Создали объект даты
     
   sh_data.getRange(range.getRow(),5).setValue(new Date(d.getFullYear(), d.getMonth(), d.getDate() + 30)) // Добавили 30 дней к дате и изменили в ячейке.
     
   //Сортировка
   range_sort.sort({column: 5, ascending: true});       
    
   var rows_oplata = sh_oplata.getLastRow(); // кол-во строк
     
   // Добавляю данные по ячейкам в лист Оплата
   sh_oplata.getRange(rows_oplata + 1, 2,1).setValue(code)
     
   if (person == "director") {
     sh_oplata.getRange(rows_oplata + 1, 3,1).setValue(nal)
     }
   else{
     sh_oplata.getRange(rows_oplata + 1, 1,1).setValue(nal)
     }
   sh_oplata.getRange(rows_oplata + 1, 4,1).setValue(beznal)
   sh_oplata.getRange(rows_oplata + 1, 5,1).setValue(date_oplata)
   sh_oplata.getRange(rows_oplata + 1, 7,1).setValue(d)
   sh_oplata.getRange(rows_oplata + 1, 8,1).setValue(status)             
}

function OpenMyDialog() {
  //Open a dialog
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (ss.getActiveSheet().getName() == "Данные") {
     
    var htmlDlg = HtmlService.createHtmlOutputFromFile('Index')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setWidth(300)
        .setHeight(150);
    SpreadsheetApp.getUi()
        .showModalDialog(htmlDlg, "Провести оплату");
   }
   else {
     var ui = SpreadsheetApp.getUi()
     ui.alert("Перейдите на лист Данные!")
   }
        
};

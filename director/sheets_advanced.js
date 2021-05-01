// Получение данных с листа - api
function readRange(spreadsheet_id, range) {  
  return Sheets.Spreadsheets.Values.get(spreadsheet_id, range).values
}


// Добавление содержимого в конец листа - api
function appendToMultipleRanges(data, spreadsheet_id, range) {
  
  var valueInputOption = ["USER_ENTERED"]
  var valueRange = Sheets.newRowData()
  valueRange.values = data
         
  Sheets.Spreadsheets.Values.append(valueRange, spreadsheet_id, range, {valueInputOption: valueInputOption});
   
}


// Обновление данными диапазона или ячейки - api
function updateToMultipleRanges(range_values, spreadsheet_id, cell_list) {

  var data = []
  
  for (var i = 0; i < cell_list.length; i++) {
    
    var obj = {
                'range': cell_list[i],
                'majorDimension': 'ROWS',
                'values': range_values
               }
     data.push(obj)      
  }
  
  
  var request = {
    'valueInputOption': 'USER_ENTERED',
    'data': data
   }

   
 Sheets.Spreadsheets.Values.batchUpdate(request, spreadsheet_id)  

}


// Получение данных в лист с определенным условием
function getFilterData(data, col, filter) {
  var list = []
  for (var i = 0; i < data.length; i++) {
    if (data[i][col] != filter){
      list.push(data[i])
    }
  }
  return list
}

function getFilterDataAddFirstEmpty(data, col, filter) {
  var list = []
  for (var i = 0; i < data.length; i++) {
    if (data[i][col] != filter){
      data[i].unshift("")
      list.push(data[i])
    }
  }
  return list
}


// Получение номеров ячеек в лист с определенным условием
function getCells(data, col, filter, adress) {
  var list = []
  for (var i = 0; i < data.length; i++) {
    if (data[i][col] != filter) {
      list.push(adress + (i + 2))
    }
  }
  return list
}

// Суммирование по конкректному индексу в массиве
function sum(arr, index) {
  return arr.reduce(function(sum, current) {
    if (current[index].length > 0) {
      var num = current[index].replace(/\s+/g, '').trim()
      num = Number(num.replace(',', '.'))
      return sum + num
     }
     return sum
   }, 0)
}


function rangeConcat(firstcol, firstindex, secondcol, secondindex) {

  return firstcol + firstindex + ':' + secondcol + secondindex
}

function getNowDate(){
  return Utilities.formatDate(new Date(), "GMT", "dd.MM.yyyy")
}

function NumberInWords(str_num) {
  
  var str_num = String(str_num)
  
  var len = str_num.length
        
  if (len == 1) {       
    var str = getUnit(str_num)
  }

  if (len == 2) {
    var str = getDicker(str_num) 
  }
       
  if (len == 3) {
    var str = getCentesimal(str_num) 
  }
       
  if (len == 4) {
    var str = getMillesimal(str_num) 
  }
       
  if (len == 5) {
    var str = getDickerMillesimal(str_num) 
  }
       
  return str + " " + getStrHryvnia(str_num.substr(-2, 2)) + " 00 коп"
}

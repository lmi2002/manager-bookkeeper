function getUnit(str_num) {
    return dnum[str_num]['ed']
}

function getDicker(str_num) {

  if (str_num.endsWith("0")) {
    return dnum[str_num[0]]['ds']
  }
  else if (str_num.startsWith("1")) {       
    return dnum[str_num]['ds']
  }
  else if (str_num.startsWith("0")) {
    return dnum[str_num[1]]['ed']
  }
  else {
    return dnum[str_num[0]]['ds'] + " " + dnum[str_num[1]]['ed']
  }
}

function getCentesimal(str_num) {

  if (str_num.endsWith("00")) {
    return dnum[str_num[0]]['sot']
  }
  else if (str_num.startsWith("0")) {
    return getDicker(str_num.substr(1, 2))
  }
  else {
     return dnum[str_num[0]]['sot'] + " " +  getDicker(str_num.substr(1, 2))
  }
}

function getMillesimal(str_num) {

  if (str_num.endsWith("000")) {
    return dnum[str_num[0]]['tch']
  }
  else {
     return dnum[str_num[0]]['tch'] + " " +  getCentesimal(str_num.substr(1, 3))
  }
}

function getDickerMillesimal(str_num) {

  if (str_num.endsWith("0000")) {
    return dnum[str_num[0]]['ds'] + " тисяч"
  }
  else if (str_num.startsWith("1") && str_num.endsWith("000")) {
    return dnum[str_num.substr(0, 2)]['tch']
  }
  else if (str_num.startsWith("1")) {
    return dnum[str_num.substr(0, 2)]['tch'] + " " + getCentesimal(str_num.substr(2, 3))
  }
  else if (str_num.substr(1, 1) == "0") {
    return dnum[str_num[0]]['ds'] + " тисяч " + getCentesimal(str_num.substr(2, 3))
  }
  else {
     return dnum[str_num[0]]['ds'] + " " + getMillesimal(str_num.substr(1, 4))
  }
}

function getStrHryvnia(num) {
  
  if (num.endsWith("1") && !num.startsWith("11")) {
    return "гривня"
  }
  else if (num.startsWith("1")) {
    return "гривень"
  }
  else if (num.endsWith("2") || num.endsWith("3") || num.endsWith("4")) {
    return "гривні"
  }
  else {
    return "гривень"
  }
}

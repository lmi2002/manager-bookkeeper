function getLetterBody() {
  return "<p>Добрий день!</p>" + 
         "<p>Просимо Вас сплатити рахунок ФОП " + DIRECTOR_NAME_SHORT + " за оренду будівельного вагонa</p>" +
         "<p>з повагою, бухгалтер - " + BOOKKEEPER_NAME + "</p>" + 
         "<p>--</p>" + 
         "<p>ФОП " + DIRECTOR_NAME_SHORT + "</p>"
}

function getTextActOfReconciliation_1(periodDate, contragent) {
  return "взаємних розрахунків за період: " + periodDate + "\n" +
  "між ФОП " + DIRECTOR_NAME_SHORT + "\n" +
  "та " + contragent + "\n" +
  "за договором Основний договір"
}

function getTextActOfReconciliation_2(contragent) {
  return "Ми, що нижче підписалися, ФОП " + DIRECTOR_NAME_ENTIRE + ", з одного боку, та директор " + contragent + "____________________________________________, з іншого боку, склали даний акт звірки у тому, що стан взаємних розрахунків за даними обліку наступний:"
}

function getTextActOfReconciliation_3(finishDate, sum, contragent) {
  let text = ""

  if (sum > 0) {
    text = "на " + finishDate + "р. заборгованість на користь  ФОП " +  DIRECTOR_NAME_SHORT +  " становить " + sum + ",00 грн"
  }
  else if (sum < 0) {
    text = "на " + finishDate + "р. заборгованість на користь " +  contragent + " становить " +  Math.abs(sum) + ",00 грн"
  }
  else if (sum == 0) {
    text = "на " + finishDate + "р. заборгованість відсутня"
  }
  return text
}

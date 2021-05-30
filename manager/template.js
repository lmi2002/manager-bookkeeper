function getLetterBody() {
  return "<p>Добрий день!</p>" + 
         "<p>Просимо Вас сплатити рахунок ФОП Гордєєв Р.В. за оренду будівельного вагонa</p>" +
         "<p>з повагою, бухгалтер - Уханова Ольга</p>" + 
         "<p>--</p>" + 
         "<p>ФОП Гордєєв Р.В.</p>" 
}

function getPersonLetterBody(comment, list) {
  let sum = Number(list[2]) * Number(list[7]) + Number(list[4])
  let date = getStrDay_1(list[1])
  let commentHtml = "<p>" + comment + "</p>"

  if (comment) {
    commentHtml = "<br>" + "<p>" + comment + "</p>" +"<br>"
  }
  
  return "<p>Добрий день!</p>" + 
        "<p>Сдаем бытовку " + list[0] +" физлицу на " + list[7] + " мес</p>" +
        "<p>" + list[2] +  " грн - аренда в мес</p>" +
        "<p>" + + list[4] + " грн - доставки</p>" +
        "<p>всего по договору: " + sum + " грн</p>" +
        commentHtml +
        "<p>Контакт: " + list[11] + "</p>" +
        "<p>Адрес объекта: " + list[18] + "</p>" +
        "<p>Дата доставки: " + date + "</p>" +
        "<br>" +
        "<p>Родион</p>"
}

function getCompanyLetterBody(obj, list) {
  let comment = obj.comment
  let delivery = obj.delivery
  let sum = Number(list[3]) * Number(list[7]) + Number(list[5])
  let date = getStrDay_1(list[1])
  let commentHtml = "<p></p>"
  let deliveryHtml = "<p></p>"

  if (comment) {
    commentHtml = "<br>" + "<p>" + comment + "</p>" +"<br>"
  }
  if(delivery) {
    deliveryHtml = "<br>" + "<p>Отдельно укажите пожалуйста стоимость доставки: " + list[5] + "</p>" +"<br>"
  }
  
  return "<p>Добрий день!</p>" + 
        "<p>Сдаем бытовку " + list[0] +" на " + list[7] + " мес</p>" +
        "<p>контрагент: " + list[11] + "</p>" +
        "<br>" +
        "<p>Номер договора: " + list[12] +  "</p>" +
        "<p>сумма по договору: " + sum + " грн</p>" +
        "<p>Стоимость следующего периода (30 дней): " + list[3] + " грн</p>" +
        deliveryHtml +
        commentHtml +
        "<p>Контакт: " + list[14] + "<br>" + list[17] + "</p>" +
        "<p>Адрес объекта: " + list[18] + "</p>" +
        "<p>Дата доставки: " + date + "</p>" +
        "<br>" +
        "<p>Родион</p>"
}

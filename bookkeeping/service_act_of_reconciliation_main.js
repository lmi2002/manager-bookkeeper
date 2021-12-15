function ActOfReconciliation(objFromForm) {
  
  let startDate = objFromForm.start_date
  let finishDate = objFromForm.finish_date
  // let startDate = '2021-09-01'
  // let finishDate = '2021-12-30'
  let contragent = getFirstNameCompany()
  let formatStartDate = Utilities.formatDate(new Date(startDate), "GMT", "dd.MM.yyyy")
  let formatFinishDate = Utilities.formatDate(new Date(finishDate), "GMT", "dd.MM.yyyy")

  // SQL запросы
  let invoicesForPeriod = getInvoiceToSheetActOfReconciliationForPeriodSQL(startDate, finishDate)
  let sumInvoicesPrevPeriod = getInvoicesPrevPeriodSummSQL(startDate)
  let sumPaidInvoicesPrevPeriod = getPaidInvoicesPrevPeriodSummSQL(startDate)
  let sumInvoicesPeriod = getInvoicesPeriodSummSQL(startDate, finishDate)
  let sumPaidInvoicesPeriod = getPaidInvoicesPeriodSummSQL(startDate, finishDate)

  let balancePrev = sumInvoicesPrevPeriod - sumPaidInvoicesPrevPeriod
  let balanceCurrent = sumInvoicesPeriod - sumPaidInvoicesPeriod

  objAOR = {
    "contragent": contragent,
    "formatStartDate": formatStartDate,
    "formatFinishDate": formatFinishDate,
    "invoicesForPeriod": invoicesForPeriod,
    "balancePrev": balancePrev,
    "balanceCurrent": balanceCurrent,
    "sumInvoicesPeriod": sumInvoicesPeriod,
    "sumPaidInvoicesPeriod": sumPaidInvoicesPeriod
  }

  let dictCopyActOfReconciliationFile  = copyActOfReconciliationFile(objAOR)
  let dictFileXlsx = exportSpreadsheetToXlsx(dictCopyActOfReconciliationFile, 'xlsx')
  moveActOfReconciliationFiles(dictFileXlsx)
  SpreadsheetApp.getUi().alert(dictFileXlsx["ss"].getName() + " сформирован")
}

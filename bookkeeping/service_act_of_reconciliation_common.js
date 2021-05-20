function differenceAmountInvoice() {
  let obj = getObjSpreadsheetApp()
  let numInvoice = obj.values_list[0][3]
  let sum = getInvoiceAggregatorSummSQL(numInvoice)
  return obj.values_list[0][5] - sum
}

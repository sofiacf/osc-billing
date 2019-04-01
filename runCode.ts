function onOpen(_e: any) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Accounting")
    .addItem('Run invoices', 'runInvoices')
    .addItem('Run payroll', 'runPayroll')
    .addToUi();
}

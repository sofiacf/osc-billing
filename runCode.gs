function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Accounting")
    .addItem('Run invoices', 'runInvoices')
    .addItem('Run payroll', 'runInvoices')
    .addToUi();
}

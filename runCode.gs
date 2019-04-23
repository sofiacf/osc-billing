var format, menuName = 'ACCOUNTING';
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ACCOUNTING').addSubMenu(ui.createMenu('SETUP').addItem('BILLING', 'getBilling').addItem('PAYROLL', 'getPayroll'))
  .addSeparator().addItem('RUN','runReports').addItem('POST', 'postReports').addToUi();
}
function getBilling() {
  format = billing;
  setup();
}
function getPayroll() {
  format = payroll;
  setup();
}


var format, menuName = 'ACCOUNTING';
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ACCOUNTING').addSubMenu(ui.createMenu('SETUP').addItem('BILLING', 'getBilling').addItem('PAYROLL', 'getPayroll'))
  .addSeparator().addItem('RUN','runReports').addItem('POST', 'postReports').addToUi();
}
function getBilling() {
  var billing = {
    name: 'BILLING',
    headers: ['ACTIVE CLIENTS', 'STATUS', 'BILL NUMBER', 'TOTAL'],
    subject: {name: 'CLIENTS', column: 3}
  };
  format = billing;
  setup();
}
function getPayroll() {
  var payroll = {name: 'PAYROLL', subject: {name: 'COURIERS', column: 12}};
  format = payroll;
  setup();
}


var format, menuName = 'ACCOUNTING';
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ACCOUNTING').addSubMenu(ui.createMenu('START').addItem('BILLING', 'getBilling').addItem('PAYROLL', 'getPayroll'))
  .addSeparator().addItem('RERUN','runReports').addItem('POST', 'postReports').addToUi();
}

function getBilling() {
  format = {
    name: 'BILLING',
    headers: ['ACTIVE CLIENTS', 'STATUS', 'BILL NUMBER', 'TOTAL'],
    subject: {name: 'CLIENTS', column: 3},
    menus: [{name: 'Run bills', functionName: 'runBills'},
            {name: 'Post bills', functionName: 'postBills'},
            {name: 'Reset', functionName: 'reset' }]
  };
  setFormat();
  showSidebar();
  runReports();
}

function getActiveSubjects(){
  const inputFile = DriveApp.getFilesByName('OSC MASTER INPUT').next();
  const inputSheet = SpreadsheetApp.open(inputFile).getSheets()[1];
  const input = inputSheet.getDataRange().getValues().slice(3);
  const subjects = input.map(function (el){return el[format.subject.column];});
  return subjects.filter(function (e,i,a){return (i == a.indexOf(e));}).sort();
}
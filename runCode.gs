// Compiled using ts2gas 1.6.2 (TypeScript 3.3.4000)
var exports = exports || {};
var module = module || { exports: exports };
function onOpen(_e) {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Accounting")
        .addItem('Run invoices', 'runInvoices')
        .addItem('Run payroll', 'runPayroll')
        .addToUi();
}

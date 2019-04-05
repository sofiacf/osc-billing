function onOpen(e) {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("Accounting")
        .addItem('Run invoices', 'runInvoices')
        .addItem('Run payroll', 'runPayroll')
        .addToUi();
}

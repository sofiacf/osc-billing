function onOpen(e) {
  if (e && e.authMode == ScriptApp.AuthMode.NONE) addSetupMenu();
  else { //Config or add full menu based on properties (fails in AuthMode.NONE)
    checkConfig()
  }
}
function createMenu() { //Creates generic menu labeled 'Accounting'
  return SpreadsheetApp.getUi().createMenu('Accounting');
}
function addSetupMenu() { //Adds menu with single 'Setup' item
  let menu = createMenu();
  menu.addItem('Setup', 'configure').addToUi();
}
function addFullMenu() { //Adds full menu w/ new run, and collections/payroll
  let menu = createMenu();
  menu.addItem('New run', 'createRun')
    .addSeparator()
    .addItem('View collections', 'viewCollections')
    .addItem('View payroll', 'viewPayroll')
    .addToUi();
}
function checkConfig() { //Runs config or adds full menu based on properties
  let properties = PropertiesService.getDocumentProperties();
  let configured = properties.getProperty('configured');
  if (configured) {
    addFullMenu();
  } else {
    addSetupMenu();
    configure();
  }
}

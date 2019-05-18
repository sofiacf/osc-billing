var execution_modes = {
  TESTING: 'testing',
  PRODUCTION: 'production'
}
function run() {
  const wkbk: WorkbookManager = new WorkbookManager();
  let scriptProperties = PropertiesService.getScriptProperties();
  let mode = scriptProperties.getProperty('execution_mode');
  switch (mode) {
    case execution_modes.TESTING:
      test();
      break;
    case execution_modes.PRODUCTION:
      wkbk.doRun();
  }
}
function test() {
  Logger.log('test mode');
}
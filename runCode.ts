var execution_modes = {
  TESTING: 'testing',
  PRODUCTION: 'production',
  ERROR: 'error'
}
function run() {
  const wkbk = new WorkbookManager();
  let scriptProperties = PropertiesService.getScriptProperties();
  let mode = scriptProperties.getProperty('execution_mode');
  switch (mode) {
    case execution_modes.TESTING:
      test();
      break;
    case execution_modes.PRODUCTION:
      wkbk.doRun();
      break;
    default:
      scriptProperties.setProperty('execution_mode', execution_modes.ERROR);
      Logger.log('mode', mode, 'is unknown/failing.');
      test();
  }
}
function test() {
  let cnfg = new ConfigurationManager();
  Logger.log(cnfg.checkValues('fields'));
}

var execution_modes = {
  TESTING: 'testing',
  PRODUCTION: 'production',
  ERROR: 'error'
}
function run() {
  let settings = Run.getSettings();
  let data = Run.getData(settings.period, settings.format);
  let folder = Run.getFolder(settings.period, settings.format);

  let scriptProperties = PropertiesService.getScriptProperties();
  let mode = scriptProperties.getProperty('execution_mode');
  switch (mode) {
    case execution_modes.TESTING:
    case execution_modes.PRODUCTION:
      Run.run(settings.action, data, folder);
      break;
    default:
      scriptProperties.setProperty('execution_mode', execution_modes.ERROR);
      Logger.log('mode', mode, 'is unknown/failing.');
  }
}

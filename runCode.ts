var execution_modes = {
  TESTING: 'testing',
  PRODUCTION: 'production',
  ERROR: 'error'
}
function run() {
  let settings = DataManager.getSettings();
  let data = DataManager.getData(settings.period, settings.format);
  let folderName = DataManager.getFolderName(settings.period, settings.format);
  let folder = FileManager.getFolder(folderName, settings.format.id);

  let scriptProperties = PropertiesService.getScriptProperties();
  let mode = scriptProperties.getProperty('execution_mode');
  switch (mode) {
    case execution_modes.TESTING:
    case execution_modes.PRODUCTION:
      FileManager.runStatements(settings.action, folder, data, settings.date);
      break;
    default:
      scriptProperties.setProperty('execution_mode', execution_modes.ERROR);
      Logger.log('mode', mode, 'is unknown/failing.');
  }
}

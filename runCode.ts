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
  let testClient1 = {
    name: 'thing',
    billedThisMonth: false
  }
  let testClient2 = {
    name: 'thing2',
    billedThisMonth: false
  }
  let testObj = [testClient1, testClient2];
  let testData = JSON.stringify(testObj);
  let userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('clients', testData);
  let result = userProperties.getProperty('clients');
  Logger.log(JSON.parse(result)[0].name);
  userProperties.deleteAllProperties();
}
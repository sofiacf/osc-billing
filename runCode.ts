//Global variables, format set in runCode since depends on config value
var ss = SpreadsheetApp.getActiveSpreadsheet();
var dash = ss.getSheetByName("DASH");

function run() {
  var config: any[] = dash.getSheetValues(1,2,3,1).map(x=>x[0]);
  var f = getFormat(config[0]);
  refresh(f);
  var date = getDate(config[2]);
  var actives: any[][] = dash.getSheetValues(2,4,-1,3);
  runReports(f, actives, date);
}

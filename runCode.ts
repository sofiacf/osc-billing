//Global variables, format set in runCode since depends on config value
var ss = SpreadsheetApp.getActiveSpreadsheet();
var dash = ss.getSheetByName("DASH");

function run() {
  const config: any[] = dash.getSheetValues(1,2,3,1).map(x=>x[0]);
  const f = getFormat(config[0]);
  refresh(f);
  const date = getDate(config[2]);
  const actives: any[][] = dash.getSheetValues(2,4,-1,3);
  const subjects = getSubjects(f, actives, date);
  runReports(f, subjects, date);
}

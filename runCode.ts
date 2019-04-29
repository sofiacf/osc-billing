//Global variables, format set in runCode since depends on config value
var f: { name: string; sc: number; total: number[]; subject: string};
var ss = SpreadsheetApp.getActiveSpreadsheet();
var dash = ss.getSheetByName("DASH");
var date = dash.getRange("B8").getValue();
var items, knowns, numSubs, runFolder;

function run() {
  //retrieves format object (subject and total columns)
  f = getFormat();
  runFolder = DriveApp.getFilesByName(f.name + " " + date);
  items = getItems();
  numSubs = Object.keys(items).length;
  knowns = getKnowns();
  //refresh data based on selection; all refreshes input reruns all reports
  refresh();
  post();
}

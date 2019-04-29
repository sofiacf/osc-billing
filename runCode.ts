//Global variables, format set in runCode since depends on config value
var f: { sc: number; total: number[]; subject: string};
var ss = SpreadsheetApp.getActiveSpreadsheet();
var dash = ss.getSheetByName("DASH");
var items, knowns, numSubs;

function run() {
  //retrieves format object (subject and total columns)
  f = getFormat();
  items = getItems();
  numSubs = Object.keys(items).length;
  knowns = getKnowns();
  //refresh data based on selection; all refreshes input reruns all reports
  refresh();
  post();
}

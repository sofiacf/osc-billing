//Global variables, format set in runCode since depends on config value
var f: { name: string; sc: number; total: number[]; subject: string};
var ss = SpreadsheetApp.getActiveSpreadsheet();
var dash = ss.getSheetByName("DASH");
var date = Utilities.formatDate(dash.getRange("B8").getValue(), "GMT-5", "M/d/yy");
var items, knowns, numSubs, runFolder;

function run() {
  //retrieves format object (subject and total columns)
  f = getFormat();
  var dir = DriveApp.getFoldersByName(f.name).next();
  if (!dir.getFoldersByName(f.name + " " + date).hasNext()) {
    dir.createFolder(f.name + " " + date)
  }
  runFolder =  dir.getFoldersByName(f.name + " " + date).next();
  items = getItems();
  numSubs = Object.keys(items).length;
  knowns = getKnowns();
  //refresh data based on selection; all refreshes input reruns all reports
  refresh();
  post();
}

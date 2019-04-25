var f: { sc: number; total: number[]; subject: string};
var ss = SpreadsheetApp.getActiveSpreadsheet();
function run() {
  const config = {};
  //stores dashboard selections in config object
  ss.getSheetByName("DASHBOARD").getNamedRanges().forEach(function(r){
    config[r.getName()] = r.getRange().getValue();
  });
  //retrieves format object (subject and total columns)
  f = format(config["format"]);
  //refresh data based on selection; all refreshes input reruns all reports
  refresh(config["refresh"]);
  post(config["post"]);
}

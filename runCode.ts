var f: { sc: number; total: number[]}, ss = SpreadsheetApp.getActiveSpreadsheet();
function run() {
  const config = {};
  const dash = ss.getSheetByName("DASHBOARD");

  //stores dashboard selections in config object
  dash.getNamedRanges().forEach(function(r){config[r.getName()] = r.getRange().getValue();});

  //retrieves format object (subject and total columns)
  f = format(config["format"]);

  //refresh data based on selection; all refreshes input reruns all reports
  refresh(config["refresh"]);


}

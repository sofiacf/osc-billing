function getFormat(format: string) {//returns format object (billing or payroll)
  var billing = { name: "BILLING", sc: 3, tc: 'priceColumn', sub: "CLIENTS" };
  var payroll = { name: "PAYROLL", sc: 12, tc: 'payoutColumn', sub: "COURIERS" };
  //TODO: Add data validation by value (eg "BILLING" & "SWITCH TO PAYROLL")
  return format == "BILLING" ? billing : payroll;
}
function getDate(date: Date) { return Utilities.formatDate(date, "GMT-5", "M/d/yy"); }

function refresh(f: { sc: number, tc: string, sub: string }) {
  var state = dash.getRange("B2").getValue();
  if (state == "NONE") return;
  const input = ss.getSheetByName("INPUT");
  if (state == "INPUT") { //refresh "INPUT" sheet and dash
    const master = DriveApp.getFilesByName('OSC MASTER INPUT').next();
    const data = SpreadsheetApp.open(master).getSheets()[1].getSheetValues(4, 2, -1, -1);
    const rows = data.length;
    input.clear().getRange(1, 1, rows, data[0].length).setValues(data).sort(f.sc);
    ss.setNamedRange('priceColumn', input.getRange(1, 13, rows));
    ss.setNamedRange('payoutColumn', input.getRange(1, 14, rows));
  }
  ss.setNamedRange('subs', ss.getSheetByName(f.sub).getRange("A2:A"));
  ss.setNamedRange('tc', ss.getRangeByName(f.tc));
  ss.setNamedRange('sc', input.getRange(1, f.sc, input.getLastRow()));
  dash.getRange("E2:E").clear().clearDataValidations();
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(["RUN", "PRINT", "POST", "SKIP"], true).build();
  var numRows: any = dash.getRange("B8").getValue();
  dash.getRange(2, 5, numRows).setDataValidation(rule);
}

function getKnowns(f: { sub: string }) {//pull data from subject sheet
  const data: any[][] = ss.getSheetByName(f.sub).getDataRange().getValues();
  const props = data.shift();
  const knowns = {};
  data.forEach(d => {
    knowns[d[0]] = {};
    props.forEach((p, i) => knowns[d[0]][p] = d[i]);
  });
  return knowns;
}
function getItems(f: { sc: number }, actives: any[]) {//pull data from input
  const items = {};
  actives.forEach(a => items[a[0]] = []);
  const input: any[][] = ss.getSheetByName("INPUT").getDataRange().getValues();
  for (var i = 0; i < input.length; i++) items[input[i][f.sc - 1]].items.push(input[i]);
  return items;
}
function Subject(sub: any[], items: {}, knowns: {}, date: string) {
  this.state = sub[1];
  this.total = sub[2];
  this.items = items[sub[0]];
  this.data = knowns[sub[0]];
  this.file = sub[0] + " - # " + (knowns[sub[0]]['statementNum'] || 1) + " - " + date;
  this.folder = DriveApp.getFolderById(knowns[sub[0]['folder']]);
  this.print = function(runFolder: GoogleAppsScript.Drive.Folder) {
    var file = runFolder.getFilesByName(this.file).next();
    var tmp = file.makeCopy(this.file + "tmp_pdf_copy");
    let url = tmp.getUrl(), id = SpreadsheetApp.open(file).getSheetId();
    let x = 'export?exportFormat=pdf&format=pdf&fitw=true&portrait=false&gridlines=false&gid=' + id;
    url = url.replace('edit?usp=drivesdk', '') + x;
    let tkn = ScriptApp.getOAuthToken();
    let r = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + tkn } });
    r.getBlob().setName(this.file);
    tmp.setTrashed(true);
    DriveApp.getFolderById(this.folder).createFile(r);
  }
}
function getSubjects(f: { sc: number, sub: string }, actives: any[][], date: any) {
  const items = getItems(f, actives), knowns = getKnowns(f);
  return actives.map(a => Subject(a, items, knowns, date));
}

function getRunFolder(f: { name: string }, date: string) {
  var dir = DriveApp.getFoldersByName(f.name).next();
  var rfName = f.name + " " + date;
  if (!dir.getFoldersByName(rfName).hasNext()) dir.createFolder(rfName);
  return dir.getFoldersByName(rfName).next();
}
function runReports(f: { name: string }, subs: {}, date: string) {
  for (var s in subs) {
    var sub = subs[s];
    var runFolder = getRunFolder(f, date);
    if (!runFolder.getFilesByName(sub['file']).hasNext()) {
      DriveApp.getFileById(subs[s]['template']).makeCopy(runFolder);
      subs[s].state = "PRINT";
    }
    else { Logger.log(s + "file found"); }
  }
}

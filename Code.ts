function getFormat(format: string) {//returns format object (billing or payroll)
  var billing = {name: "BILLING", sc: 3, tc: 'priceColumn', sub: "CLIENTS" };
  var payroll = {name: "PAYROLL", sc: 12, tc: 'payoutColumn', sub: "COURIERS"};
  //TODO: Add data validation by value (eg "BILLING" & "SWITCH TO PAYROLL")
  return format == "BILLING" ? billing : payroll;
}
function getDate(date: Date){return Utilities.formatDate(date, "GMT-5", "M/d/yy");}

function refresh(f: {sc: number, tc: string, sub: string}){
  var state = dash.getRange("B2").getValue();
  if (state == "NONE") return;
  const input = ss.getSheetByName("INPUT");
  if (state == "INPUT") { //refresh "INPUT" sheet and dash
    const data = SpreadsheetApp.open(DriveApp.getFilesByName('OSC MASTER INPUT').next()).getSheets()[1].getSheetValues(4, 2, -1, -1);
    const rows = data.length;
    input.clear().getRange(1, 1, rows, data[0].length).setValues(data).sort(f.sc);
    ss.setNamedRange('priceColumn', input.getRange(1, 13, rows));
    ss.setNamedRange('payoutColumn', input.getRange(1, 14, rows));
  }
  ss.setNamedRange('subs', ss.getSheetByName(f.sub).getRange("A2:A"));
  ss.setNamedRange('tc', ss.getRangeByName(f.tc));
  ss.setNamedRange('sc', input.getRange(1, f.sc, input.getLastRow()));
  dash.getRange("E2:E").clear().clearDataValidations();
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(["RUN","PRINT","POST","SKIP"], true).build();
  dash.getRange(2,5,dash.getSheetValues(2,4,-1,1).length).setDataValidation(rule);
}

function getRunFolder(f: {name: string}, date: string){
  var dir = DriveApp.getFoldersByName(f.name).next();
  var rfName = f.name + " " + date;
  if (!dir.getFoldersByName(rfName).hasNext()) dir.createFolder(rfName);
  return dir.getFoldersByName(rfName).next();
}
function runReports(f: {sc: number, name: string, sub: string}, actives: any[][], date: string) {
  var subs = {};
  actives.forEach(sub=>subs[sub[0]] = {state: sub[1], total: sub[2], items: []});
  const subData: any[][] = ss.getSheetByName(f.sub).getSheetValues(1,1,-1,-1);
  const props = subData.shift();
  subData.forEach(d=>{if (subs[d[0]]) props.forEach((p,i)=>subs[d[0]][p] = d[i])});
  const input: any[][] = ss.getSheetByName("INPUT").getSheetValues(1,1,-1,-1);
  for (var i=0; i<input.length; i++) subs[input[i][f.sc-1]].items.push(input[i]);
  var runFolder = getRunFolder(f, date);
  for (var s in subs) {
    var sub = subs[s];
    sub['file'] = s + " - # " + (sub['statementNum'] || 1) + " - " + date;
    if (!runFolder.getFilesByName(subs[s]['file']).hasNext()) {
      Logger.log(s + "didn't have file");
    }
    else {Logger.log(s + "file found");}
  }
}
function print(sub: {name: string, file: string}, runFolder: any) {
  var folder = DriveApp.getFolderById(sub['folder']);
  var file = folder.getFilesByName(sub.file).next();
  var tmp = file.makeCopy(sub.file+"tmp_pdf_copy");
  let url = tmp.getUrl(), id = SpreadsheetApp.open(file).getSheetId();
  let x = 'export?exportFormat=pdf&format=pdf&fitw=true&portrait=false&gridlines=false&gid=' + id;
  url = url.replace('edit?usp=drivesdk', '');
  let tkn = ScriptApp.getOAuthToken();
  let r = UrlFetchApp.fetch(url + x,{headers: {'Authorization': 'Bearer ' + tkn}});
  r.getBlob().setName(sub.file);
  tmp.setTrashed(true);
  runFolder.createFile(r);
}
function test(){
  var format = getFormat("BILLING");
  var actives = dash.getSheetValues(2,4,-1,3);
  var date = "4/27/19"
  runReports(format, actives, date);
}

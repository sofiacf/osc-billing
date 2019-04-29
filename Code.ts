//Setup functions for global variables, run every time
function getFormat() {
  return (dash.getRange("B1").getValue() == "BILLING" ?
    {name: "BILLING", sc: 3, total: [12], subject: "CLIENTS" }
    : { name: "PAYROLL", sc: 12, total: [14, 13], subject: "COURIERS" });
}
function getItems(){ //Pull charges into object indexed by subject
  const input = ss.getSheetByName("INPUT").getSheetValues(1,1,-1,-1);
  const items = {};
  const col = f.sc - 1; //Since input is read as array, decrease subject column
  for (var i = 0; i < input.length; i++){
    items[input[i][col]] = (items[input[i][col]] || []).concat(input[i]);
  }
  return items;
}
function getKnowns(){
  const data: any[][] = ss.getSheetByName(f.subject).getSheetValues(1,1,-1,-1);
  const props = data.shift();
  const knowns = {};
  data.forEach(function (d) {
    knowns[d[0]] = {};
    props.forEach(function (p,i) {knowns[d[0]][p] = d[i]})
  });
  return knowns;
}
function runReports(){
  var actives = dash.getRange(2,4,numSubs).getValues().map(x=>x[0]);
  actives.forEach(function(sub: string){
    var name = sub + " - # " + (knowns[sub]["statementNum"] || 1) + " - " + date;
    if (runFolder.getFilesByName(name).hasNext()) {
      //do stuff
    }
  });
}

function refresh(){
  var state = dash.getRange("B2").getValue();
  if (state == "NONE") return;
  if (state == "DASH") { //reset dash formulas
    dash.getRange("D2").setFormula('=ARRAYFORMULA(UNIQUE(subColumn))');
    dash.getRange("F2").setFormula('=ARRAYFORMULA(IF(B1="BILLING",SUMIF(subColumn,UNIQUE(subColumn),priceColumn),SUMIF(subColumn,UNIQUE(subColumn),payoutColumn)))');
    ss.getRange("E2:E").clear().clearDataValidations();
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(["RERUN","POST","SKIP"], true).build();
    dash.getRange(2,5,numSubs).setDataValidation(rule);
    return;
  }
  if (state = "ALL") { //pull data from OSC Master, name ranges
    const master = DriveApp.getFilesByName('OSC MASTER INPUT').next();
    const data = SpreadsheetApp.open(master).getSheets()[1].getSheetValues(4, 2, -1, -1);
    const inputSheet = ss.getSheetByName("INPUT").clear();
    inputSheet.getRange(1, 1, data.length, data[0].length).setValues(data).sort(f.sc);
    const numRows = data.length;
    ss.setNamedRange('subColumn', inputSheet.getRange(1, f.sc, numRows));
    ss.setNamedRange('priceColumn', inputSheet.getRange(1, 13, numRows)),
    ss.setNamedRange('payoutColumn', inputSheet.getRange(1, 14, numRows));
  }
  SpreadsheetApp.flush();
  runReports();
}

function post() {
  //if set to all (or selected?), saves pdfs and updates subs/collections sheet
  var state = dash.getRange("B3").getValue();
  if (state == "NONE") return;
  if (state == "SELECTED") {
    var actives = dash.getRange(2,4,numSubs).getValues().map(x=>x[0]);
    actives.forEach(function(sub: string){
      var name = sub + " - #" + knowns[sub]["statementNum"] + " - " + date;
      var folder = DriveApp.getFolderById(knowns[sub]["folder"]);
      var file = folder.getFilesByName(name).next();
      var tmp = file.makeCopy(name+"tmp_pdf_copy");
      let url = tmp.getUrl(), id = SpreadsheetApp.open(file).getSheetId();
      let x = 'export?exportFormat=pdf&format=pdf&fitw=true&portrait=false&gridlines=false&gid=' + id;
      url = url.replace('edit?usp=drivesdk', '');
      let tkn = ScriptApp.getOAuthToken();
      let r = UrlFetchApp.fetch(url + x,{headers: {'Authorization': 'Bearer ' + tkn}});
      r.getBlob().setName(name);
      tmp.setTrashed(true);
      runFolder.createFile(r);
    });
  }
}

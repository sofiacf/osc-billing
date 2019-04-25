function setNamedRanges(){
  var input = ss.getSheetByName("INPUT");
  var dash = ss.getSheetByName("DASHBOARD");
  var namedRanges = {
    input: input.getRange(1,1,-1-1),
    subColumn: input.getRange(1, f.sc, -1),
    priceColumn: input.getRange(1, 13, -1),
    payoutColumn: input.getRange(1, 14, -1),
    actives: dash.getRange(2, 4, -1),
    totals: dash.getRange(2, 4, -1, f.total.length)
  }
  Object.keys(namedRanges).forEach(e => ss.setNamedRange(e,namedRanges[e]));
}
function refreshActives(){
  var actives = getActives();
  ss.getRangeByName('actives').setValues(actives.map(x=>[x]));
  ss.getRangeByName('numStatements').setValue(actives.length);
}
function refreshInput(){
  const master = DriveApp.getFilesByName('OSC MASTER INPUT').next();
  const data = SpreadsheetApp.open(master).getSheets()[1].getSheetValues(4, 2, -1, -1);
  const inputSheet = ss.getSheetByName("INPUT").clear()
  inputSheet.getRange(1, 1, data.length, data[0].length).setValues(data).sort(f.sc);
}
function refresh(state: string){
  if (state == "NONE") return;
  if (state == "ALL") {
    refreshInput();
    setNamedRanges();
    refreshActives();
  }
}
function post(state: string) {
  //if set to all (or selected?), saves pdfs and updates subs/collections sheet
  if (state == "NONE") return;
  if (state == "ALL") {
    var actives = getActives();
    var knowns = getKnowns();
    var activesData = actives.map(x=>knowns[x]);
    activesData.forEach(function(a){
      var fileName = "test";
      DriveApp.getFolderById(a["folder"]).getFilesByName(fileName);
    })
  }
}
function format(state: string): { sc: number; total: number[]; subject: string}{
  return (state == "BILLING") ? {sc: 3, total: [12], subject: "CLIENTS"}
    : {sc: 12, total: [14,13], subject: "COURIERS"};
}
function getItems(){
  //Collects all rows of OSC Master Input sorted by subject
  const input: any[][] = ss.getRangeByName('input').getValues();
  const items = {};
  //Since input is read as array, decrease subject column
  const col = f.sc - 1;
  for (var i = 0; i < input.length; i++){
    items[input[i][col]] = (items[input[i][col]] || []).concat(input[i]);
  }
  return items;
}
function getActives(){
  return Object.keys(getItems());
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
function setup(){
  const items = getItems();
  const actives = getActives();
  const knowns = getKnowns();
}

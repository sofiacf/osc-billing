function refreshInput(){
  const input = SpreadsheetApp.open(DriveApp.getFilesByName('OSC MASTER INPUT').next()).getSheets()[1].getSheetValues(4,2,-1,-1);
  ss.getSheetByName("INPUT").clear().getRange(1, 1, input.length, input[0].length).setValues(input).sort(f.sc);
}
function refresh(state: string){
  if (state == "NONE") return;
  if (state == "ALL") refreshInput();
}
function format(state: string): { sc: number; total: number[]}{
  var format = (state == "BILLING") ? {sc: 3, total: [12]} : {sc: 12, total: [14,13]};
  return format;
}
function getItems(){
  const input: any[][] = ss.getSheetByName("INPUT").getSheetValues(1,1,-1,-1);
  const items = {};
  for (var i = 0; i < input.length; i++){
    items[input[i][f.sc]] = (items[input[i][f.sc]] || []).concat(input[i]);
  }
  return items;
}
function getKnowns(){
  const data: any[][] = ss.getSheetByName("SUBJECTS").getSheetValues(1,1,-1,-1);
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
  const actives = Object.keys(items);
  const knowns = getKnowns();
  Logger.log(Object.keys(knowns));
}

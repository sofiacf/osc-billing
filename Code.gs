var ss = SpreadsheetApp.getActiveSpreadsheet();
var billing = {name: 'BILLING', subject: 'CLIENTS', columns: {subject: 3, total: 13},
    headers: ['ACTIVE CLIENTS', 'STATUS', 'BILL NUMBER', 'TOTAL']};
var payroll = {name: 'PAYROLL', subject: {name: 'COURIERS', column: 12}};
var formats = [billing, payroll];
function setup(){
  //Clear sheet
  const setup = ss.getSheetByName("SETUP").clear();
  const input = ss.getSheetByName("INPUT").clear();
  ss.getSheets().forEach(function(sheet){if (["SETUP", "INPUT", 'CLIENTS', 'COURIERS'].indexOf(sheet.getName()) < 0) ss.deleteSheet(sheet);});
  //Calculate initial setup
  const data = SpreadsheetApp.open(DriveApp.getFilesByName('OSC MASTER INPUT').next()).getSheets()[1].getSheetValues(1,1,-1,-1).slice(3);
  const actives = data.map(function (el){
    return el[format.columns.subject];}).filter(function (e,i,a){
    return (i == a.indexOf(e));}).sort();
  const knowns = ss.getSheetByName(format.subject).getSheetValues(1, 1, -1, -1);
  const knownids = knowns.map(function(el){return el[0];});
  const fields = format.headers.map(function(el){return knowns[0].indexOf(el);});
  const values = [format.headers].concat(actives.map(function(sub){
    var arr = [sub].concat(["READY"]), index = knownids.indexOf(sub);
    if (index < 0) return arr.concat(new Array(format.headers.length-2));
    else return arr.concat(fields.slice(2).map(function(el){return el > -1 ? knowns[index][el] : "";}));
  }));
  //Format spreadsheet
  ss.getSheetByName("INPUT").clear().showSheet().getRange(1, 1, data.length, data[0].length).setValues(data);
  setup.getRange(1,1, values.length, format.headers.length).setValues(values)
  setup.getRange(1,1,1,format.headers.length).setFontWeight('bold');
}
function runReports(){
  const setup = ss.getSheetByName('SETUP');
  const actives = setup.getSheetValues(1, 1, -1, -1);
  format = (actives[0][0] == "ACTIVE CLIENTS") ? billing : payroll;
  const working = actives.filter(function(el){return el[1] == 'READY'});
  working.forEach(function(sub){if (ss.getSheetByName(sub[0]) == null) ss.insertSheet(sub[0])});
  const data = ss.getSheetByName("INPUT").getSheetValues(1, 1, -1, -1);
  data.forEach(function(row){ss.getSheetByName(row[format.columns.subject]).appendRow(row);});
}


//
//function postInvoices(){
//  const summary = SpreadsheetApp.getActiveSheet(), CIdLength = parseInt(summary.getRange('A29').getValue()),
//      invDate = summary.getRange('B3').getValue(), CIdList = summary.getRange(30, 1, CIdLength).getValues(),
//      CInfoSheet = SpreadsheetApp.open(DriveApp.getFilesByName('Client Billing Information').next()).getSheets()[0],
//      CInfoIds = CInfoSheet.getRange(2, 1, CInfoSheet.getLastRow()-2).getValues();
//  for (var i=0; i<CIdList.length; i++) CIdList[i] = CIdList[i][0];
//  for (var i=0; i<CInfoIds.length; i++) {
//    if (CIdList.includes(CInfoIds[i][0])) {
//      var dateCell = CInfoSheet.getRange('F'+(i+2));
//      dateCell.setValue(invDate);
//      var invNumCell = CInfoSheet.getRange('G'+(i+2));
//      invNumCell.setValue(parseInt(invNumCell.getValue())+1);
//    }
//  }
//}
//
////function Setup(){
////  }
////  const folder = createFolder();
////  function createSummary(){
////
////    const clientData = clientInfoSheet.getRange(2, 1, clientInfoSheet.getLastRow()-2,8).getValues();
////    var clients = new Array;
////    function getInputClients(){
////    }
////    const activeClients = getActiveClients();
////    function getExistingClients(){
////      var existingClients = new Array;
////      for (i=1; i<clientData.length; i++){
////          existingClients[i] = clientData[i][0];
////      }
////      return existingClients;
////    }
////    const existingClients = getExistingClients();
////    function getNewClients(){
////      var newClients = new Array;
////      var newClientArray = [[], ['Enter name'], ['Enter attn. (optional)'], ['Enter address'], ['Enter city, state zip'], [0],['NEW']];
////      for (i=0; i<activeClients.length; i++){
////        if (!existingClients.includes(activeClients[i])){
////          newClientArray[0] = activeClients[i];
////          newClients.push(newClientArray);
////        }
////      }
////      return newClients;
////    }
////  }
////    this.clientIdList = new Array;
////    var newClients = getNewClients();
////    function matchClients(){
////      var clients = new Array;
////      for (f=0; f<activeClients.length; f++){
////        if (existingClients.includes(activeClients[f])){
////          var clientIndex = existingClients.indexOf(activeClients[f]);
////          var clientInfoArray = clientData[clientIndex].slice(0,8);
////          var thisClient = new Client(clientInfoArray, charges);
////          if (thisClient.total <1){
////            newClients.push(clientInfoArray);
////            break;
////          } else {
////            clients.push(thisClient)

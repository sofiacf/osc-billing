function reset(){
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  workbook.getSheetByName("SETUP").clear();
  workbook.getSheets().forEach(function(sheet){if (sheet.getName() != "SETUP") sheet.hideSheet();});
  setup();
}
function getActiveSubjects(){
  const inputFile = DriveApp.getFilesByName('OSC MASTER INPUT').next();
  const inputSheet = SpreadsheetApp.open(inputFile).getSheets()[1];
  const input = inputSheet.getDataRange().getValues().slice(3);
  const subjects = input.map(function (el){return el[format.subject.column];});
  return subjects.filter(function (e,i,a){return (i == a.indexOf(e));}).sort();
}
function setup(){
  //Setup values
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const actives = getActiveSubjects(format.name);
  const knowns = workbook.getSheetByName(format.subject.name).getDataRange().getValues();
  const knownids = knowns.map(function(el){return el[0];});
  const matchColumns = format.headers.map(function(el){return knowns[0].indexOf(el);});
  //Update setup sheet
  const setup = workbook.getSheetByName('SETUP');
  const values = [format.headers].concat(actives.map(function(sub){
    var arr = [sub].concat(["READY"]), index = knownids.indexOf(sub);
    if (index < 0) return arr.concat(new Array(format.headers.length-2));
    else return arr.concat(matchColumns.slice(2).map(function(el){return el > -1 ? knowns[index][el] : "";}));    
  }));
  setup.getRange(1,1, values.length, format.headers.length).setValues(values);
  setup.getRange(1,1,1,format.headers.length).setFontWeight('bold');
  actives.forEach(function(sub){
    if (workbook.getSheets().some(function(sheet){return sheet.getName() == sub})) {
      workbook.getSheetByName(sub).showSheet();
    } else {workbook.insertSheet(sub);}
  });
}
function runReports(){
  const actives = getActiveSubjects(format.name);
  const workbook = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SETUP');
  const working = workbook.getRange(2,1,format.headers.length, actives.length).getValues();
  const readys = working.filter(function(el){return el[1] == 'READY'});
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
function Client(cInfo, lineItems){
  const cFolder = DriveApp.getFoldersByName("CLIENTS").next(), name = cInfo[1], items = [], x = cInfo[7],
      cinfo = (cInfo[2]) ? [[name], [cInfo[2]], [cInfo[3]], [cInfo[4]]] : [[name], [cInfo[3]], [cInfo[4]], [""]];
  var total = 0, fName, folder, numItems, invInfo, invNum = (cInfo[6]+1); 
  this.id = cInfo[0];
  for (var i=0; i<lineItems.length; i++) if (lineItems[i].id == this.id) items.push(lineItems[i]);
  numItems = items.length;
  for (var j=0; j<numItems; j++) total += items[j].amount;
  this.total = total;
  for (var i=0; i<numItems; i++) items[i] = items[i].lineItem;
  invInfo = [[today], [invNum], [billPeriod], [total]];
  fName = this.id + " - # " + invNum + x + " - " + today;
  if (!DriveApp.getFoldersByName(this.id).hasNext()) cFolder.createFolder(this.id);
  folder = DriveApp.getFoldersByName(this.id).next();
  this.generateInvoice = function(){
    var t, file, itemRange, priceRange, dateRange, isNixon = (this.id == "NIXON"), tName = (isNixon) ? "Nixon Template" : "Template";
    t = DriveApp.getFilesByName(tName).next().getId();
    DriveApp.getFileById(t).makeCopy(fName, folder);
    file = SpreadsheetApp.open(DriveApp.getFilesByName(fName).next()).getSheets()[0];
    var invInfoRange = (isNixon) ? file.getRange("H4:H7") : file.getRange("F4:F7"),
        cInfoRange = file.getRange("A11:A14"), totalRange = (isNixon) ? file.getRange("I18") : file.getRange("G18");
    invInfoRange.setValues(invInfo);
    cInfoRange.setValues(cinfo);
    totalRange.setValue(total);
    SpreadsheetApp.flush();
    file.insertRows(16, numItems); //Add rows for line items
    itemRange = file.getRange(16, 1, numItems, items[0].length);
    itemRange.setValues(items);
    SpreadsheetApp.flush();
    priceRange = (isNixon) ? file.getRange(16, 9, numItems) : file.getRange(16, 7, numItems);
    priceRange.setNumberFormat('$0.00');
    itemRange.setHorizontalAlignment('general-center').setFontSize(10).setWrap(true);
    dateRange = file.getRange(16, 1, numItems);
    dateRange.setNumberFormat('MM/dd/yy');
    dateRange.setHorizontalAlignment('general-left');
    SpreadsheetApp.flush();
    var invFolder = DriveApp.getFoldersByName(name), ss = DriveApp.getFilesByName(fName).next(), pdfName = ss.getName();
    DriveApp.getFileById(ss.getId()).makeCopy(this.id + "tmp_pdf_copy", folder);
    var pdfSheet = DriveApp.getFilesByName(this.id + "tmp_pdf_copy").next(), url = pdfSheet.getUrl(),
        url_ext = 'export?exportFormat=pdf&format=pdf' + '&fitw=true' + '&portrait=false' + '&gridlines=false' + '&gid=' + file.getSheetId();
    url = url.replace("edit?usp=drivesdk", '');
    var token = ScriptApp.getOAuthToken(), response = UrlFetchApp.fetch(url + url_ext, {headers: {'Authorization': 'Bearer ' + token}}), 
        blob = response.getBlob().setName(pdfName), newFile = DriveApp.getFoldersByName("Invoices Date: " + today).next().createFile(blob);
    pdfSheet.setTrashed(true);
    billsCreated+=1;
  }
}
function Charge(id, itemArray){
  this.id = id,
    this.lineItem = (id == "NIXON") ? [itemArray[0]].concat(itemArray.slice(4, 11).concat(itemArray[12]))
    : [itemArray[0], itemArray[4], itemArray[5], itemArray[6], itemArray[7], itemArray[8], itemArray[12]];
  this.amount = (isNaN(parseFloat(itemArray[12]))) ? 0 : parseFloat(itemArray[12]);
}
function Setup(){
  const master = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1], lastRow = master.getLastRow(),
      mData = master.getRange(4, 2, lastRow-4, 13).getValues(), invFolder = DriveApp.getFoldersByName("INVOICES").next(),
      fName = "Invoices Date: " + today;
  billPeriod = String(master.getSheetName()).toUpperCase();
  if (DriveApp.getFoldersByName(fName).hasNext()) DriveApp.getFoldersByName(fName).next().setTrashed(true);
  invFolder.createFolder(fName);
  const f = DriveApp.getFoldersByName(fName).next(), sTemplate = DriveApp.getFilesByName("Billing Summary Template").next(), sName = "Inv. Report - " + today;
  sTemplate.makeCopy(sName, f);
  this.summary = DriveApp.getFilesByName(sName).next(), this.charges = [];
  for (var i=0; i<mData.length; i++) this.charges.push(new Charge(mData[i][2], mData[i]));
  function getClients(charges){
    const clients = [], mClients = [], savedClients = [], newClients = [], 
        cInfoSheet = SpreadsheetApp.open(DriveApp.getFilesByName("Client Billing Information").next()).getSheets()[0], 
        cData = cInfoSheet.getRange(2, 1, cInfoSheet.getLastRow()-2,8).getValues();
    for (var i=0; i<mData.length; i++) if (!mClients.includes(mData[i][2])) mClients.push(mData[i][2]);
    for (var i=0; i<cData.length; i++) savedClients[i] = cData[i][0];
    for (var i=0; i<mClients.length; i++) if (!savedClients.includes(mClients[i])) newClients.push([
      [mClients[i]], ["Name"], ["Attn. (optional)"], ["Address"], ["City, State Zip"], [0],["x??"], ["ID not found in Client Info Sheet"]]);
    function matchClients(){
      for (var f=0; f<mClients.length; f++){
        if (!savedClients.includes(mClients[f])) continue;
        var cIndex = savedClients.indexOf(mClients[f]), cInfo = cData[cIndex].slice(0,8), thisClient = new Client(cInfo, charges);
        if (thisClient.total < 1) newClients.push(cInfo);
        else clients.push(thisClient);
      }
    }
    matchClients();
    return [clients, newClients];
  }
  var clientsArray = getClients(this.charges);
  this.clients = clientsArray[0], this.newClients = clientsArray[1], this.clientIdList = [];
  for (var g = 0; g < clientsArray[0].length; g++) this.clientIdList.push([clientsArray[0][g].id]);
}
function makeSummary(run){
  var runSummary = SpreadsheetApp.open(run.summary).getSheets()[0], expectedTotal = 0;
  for (i=0; i<run.charges.length; i++) expectedTotal+= run.charges[i].amount;
  runSummary.getRange("B3:B5").setValues([[today],[billsCreated],[expectedTotal]]);
  runSummary.getRange(10, 1, run.newClients.length, run.newClients[0].length).setValues(run.newClients);
  runSummary.getRange(30, 1, run.clientIdList.length).setValues(run.clientIdList);
  runSummary.getRange("A29").setValue(run.clientIdList.length);
}
function generateInvoices(run){
  for (r=0; r<run.clients.length; r++) run.clients[r].generateInvoice();
}
var billPeriod, billsCreated = 0;
function Client(infoArray, lineItems){
  const clientsFolder = DriveApp.getFoldersByName("CLIENTS").next();
  var name = infoArray[1], total = 0, items = [], itemArrays = [], suffix = infoArray[7], fileName, folder, numItems,
      info = (infoArray[2]) ? [[name], [infoArray[2]], [infoArray[3]], [infoArray[4]]] : [[name], [infoArray[3]], [infoArray[4]], [""]], invNum = infoArray[6]+1;
  this.id = infoArray[0];
  for (var i=0; i<lineItems.length; i++) if (lineItems[i].id == this.id) items.push(lineItems[i]);
  numItems = items.length;
  for (var j=0; j<numItems; j++) total += items[j].amount;
  var invInfo = [[today], [invNum], [String(billPeriod).toUpperCase()], [total]];
  this.total = total;
  for (var k=0; k<numItems; k++) itemArrays.push(items[k].lineItem);
  fileName = this.id + " - # " + invNum + suffix + " - " + today;
  if (!DriveApp.getFoldersByName(this.id).hasNext()) clientsFolder.createFolder(this.id);
  folder = DriveApp.getFoldersByName(this.id).next();
  this.generateInvoice = function(){
    var template, file, lineItemRange, priceColumnRange, dateColumnRange, isNixon = (this.id == "NIXON");
    template = (isNixon) ? DriveApp.getFilesByName("Nixon Template").next().getId() : DriveApp.getFilesByName("Template").next().getId();
    DriveApp.getFileById(template).makeCopy(fileName, folder); //Copy template file
    file = SpreadsheetApp.open(DriveApp.getFilesByName(fileName).next()).getSheets()[0];
    var invInfoRange = (isNixon) ? file.getRange("H4:H7") : file.getRange("F4:F7"), clientInfoRange = file.getRange("A11:A14"), totalRange = (isNixon) ? file.getRange("I18") : file.getRange("G18"); //Cells for invoice data
    invInfoRange.setValues(invInfo);
    clientInfoRange.setValues(info);
    totalRange.setValue(total);
    SpreadsheetApp.flush();
    file.insertRows(16, numItems); //Add rows for line items
    lineItemRange = file.getRange(16, 1, numItems, itemArrays[0].length);
    lineItemRange.setValues(itemArrays);
    SpreadsheetApp.flush();
    priceColumnRange = (isNixon) ? file.getRange(16, 9, numItems) : file.getRange(16, 7, numItems);
    priceColumnRange.setNumberFormat('$0.00');
    lineItemRange.setHorizontalAlignment('general-center').setFontSize(10).setWrap(true);
    dateColumnRange = file.getRange(16, 1, numItems);
    dateColumnRange.setHorizontalAlignment('general-left');
    dateColumnRange.setNumberFormat('MM/dd/yy');
    SpreadsheetApp.flush();
    var invFolder = DriveApp.getFoldersByName(name), ss = DriveApp.getFilesByName(fileName).next(), pdfName = ss.getName();
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
  const OSCMaster = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1], lastMasterRow = OSCMaster.getLastRow(),
      monthlyData = OSCMaster.getRange(4, 2, lastMasterRow-4, 13).getValues(),
      invoiceFolder = DriveApp.getFoldersByName("INVOICES").next(),folderName = "Invoices Date: " + today;
  billPeriod = OSCMaster.getSheetName();
  if (DriveApp.getFoldersByName(folderName).hasNext()) DriveApp.getFoldersByName(folderName).next().setTrashed(true); //Delete existing invoice folder
  invoiceFolder.createFolder(folderName); //Make a new folder
  const folder = DriveApp.getFoldersByName(folderName).next(), summaryTemplate = DriveApp.getFilesByName("Billing Summary Template").next(), summaryName = "Inv. Report - " + today;
  summaryTemplate.makeCopy(summaryName, folder);
  this.summary = DriveApp.getFilesByName(summaryName).next(), this.charges = [];
  for (var i=0; i<monthlyData.length; i++) this.charges.push(new Charge(monthlyData[i][2], monthlyData[i]));
  function getClients(charges){
    const clients = [], inputClients = [], activeClients = [], existingClients = [], newClients = [], 
        clientInfoSheet = SpreadsheetApp.open(DriveApp.getFilesByName("Client Billing Information").next()).getSheets()[0], 
        clientData = clientInfoSheet.getRange(2, 1, clientInfoSheet.getLastRow()-2,8).getValues();
    for (var i=0; i<monthlyData.length; i++) if (monthlyData[i][2]) inputClients.push(monthlyData[i][2]);
    for (var i=0; i<inputClients.length; i++) if (!activeClients.includes(inputClients[i])) activeClients.push(inputClients[i]);
    for (var i=0; i<clientData.length; i++) existingClients[i] = clientData[i][0];
    for (var i=0; i<activeClients.length; i++) if (!existingClients.includes(activeClients[i])) newClients.push([
      [activeClients[i]], ["Name"], ["Attn. (optional)"], ["Address"], ["City, State Zip"], [0],["Suffix??"], ["NEW"]]);
    function matchClients(){
      for (var f=0; f<activeClients.length; f++){
        if (!existingClients.includes(activeClients[f])) continue;
        var clientIndex = existingClients.indexOf(activeClients[f]), clientInfoArray = clientData[clientIndex].slice(0,8), thisClient = new Client(clientInfoArray, charges);
        if (thisClient.total < 1) newClients.push(clientInfoArray);
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
  makeSummary(run);
}
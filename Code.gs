//Universal Variables - Monthly Disp/Billing/Payroll Sheet
//TO DO: Summary sheet, post button, rerun for new clients, uppercase?, DATE ORDER WTF??
var OSCMaster = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
var monthlyData = OSCMaster.getRange(4, 2, OSCMaster.getLastRow()-1, 13).getValues();
var today = Utilities.formatDate(new Date(), "GMT-4", "M/d/yy");
var billPeriod = OSCMaster.getSheetName();
var invoiceFolder = DriveApp.getFoldersByName("INVOICES").next();
var folderName = "Invoices Date: " + today;
var summaryName = "Inv. Report - " + today;
function setupFolder(){
  function setupSummary(folder){
    var template = DriveApp.getFilesByName("Billing Summary Template").next().getId();
    DriveApp.getFileById(template).makeCopy(summaryName, DriveApp.getFolderById(folder));
  }
  if (DriveApp.getFoldersByName(folderName).hasNext()){ //Delete existing folder if one exists
    DriveApp.getFoldersByName(folderName).next().setTrashed(true);
  }
  invoiceFolder.createFolder(folderName); //Make a new folder
  var folder = DriveApp.getFoldersByName(folderName).next().getId();
  setupSummary(folder);
}
function createSummary(){
  var summary = DriveApp.getFilesByName(summaryName).next();
}
function setupClients(){
  var clients = [];
  var clientsFolder = DriveApp.getFoldersByName("CLIENTS").next();
  var OSCClientSheet = SpreadsheetApp.open(DriveApp.getFilesByName("Client Billing Information").next()).getSheets()[0];
  function Client(infoArray){
    this.id = infoArray[0];
    this.name = infoArray[1];
    if (infoArray[2]){
      this.info = [[this.name], [infoArray[2]], [infoArray[3]], [infoArray[4]]]; //name, attn, address, city
    } else {
      this.info = [[this.name], [infoArray[3]], [infoArray[4]], ['']]; //name, address, city, and a blank box
    }
    this.name = infoArray[1];
    this.invNum = infoArray[6] + 1;
    this.suffix = infoArray[7];
    this.fileName = this.id + " - # " + this.invNum + this.suffix + " - " + today;
    if (!DriveApp.getFoldersByName(this.id).hasNext()){
      clientsFolder.createFolder(this.id);
    }
    this.folder = DriveApp.getFoldersByName(this.id).next();
    this.generateInvoice = function(){
      var template = DriveApp.getFilesByName("Template").next().getId();
      DriveApp.getFileById(template).makeCopy(this.fileName, this.folder);
      var file = SpreadsheetApp.open(DriveApp.getFilesByName(this.fileName).next()).getSheets()[0];
      var items = [];
      var total = 0;
      var inputItems = OSCMaster.getRange(4, 2, OSCMaster.getLastRow()-4, 13).getValues().sort();
      for (i=0; i<inputItems.length; i++){
        if (inputItems[i][2] == this.id){ //Add items
          items.push([
            inputItems[i][0], //date
            inputItems[i][4], //address
            inputItems[i][5], //company
            inputItems[i][6], //type
            inputItems[i][7], //add charges
            inputItems[i][8], //charge code
            inputItems[i][12]]); //price
          total += inputItems[i][12];
        }
      }
      file.insertRows(16, items.length);
      file.getRange(16, 1, items.length, items[0].length).setValues(items).setHorizontalAlignment('general-center').setFontSize(10).setWrap(true);
      SpreadsheetApp.flush();
      file.getRange("F4:F7").setValues([[today], [this.invNum], [String(billPeriod).toUpperCase()], [total]]); //Add info
      file.getRange("A11:A14").setValues(this.info);
      var priceColumn = file.getRange(16, 7, items.length); //Format sheet
      priceColumn.setNumberFormat('$0.00');
      var dateColumn = file.getRange(16, 1, items.length);
      dateColumn.setHorizontalAlignment('general-left');
      file.getRange(priceColumn.activate().getLastRow()+3, 7).setValue(total);
      file.getRange(16, 1, items.length).setHorizontalAlignment('general-left');
      SpreadsheetApp.flush();
      var ss = DriveApp.getFilesByName(this.fileName).next();
      var pdfName = ss.getName();
      DriveApp.getFileById(ss.getId()).makeCopy(this.id + "tmp_pdf_copy", this.folder);
      var pdfSheet = DriveApp.getFilesByName(this.id + "tmp_pdf_copy").next();
      var url = pdfSheet.getUrl();
      url = url.replace("edit?usp=drivesdk", '');
      var url_ext = 'export?exportFormat=pdf&format=pdf' + '&fitw=true' + '&portrait=false' + '&gridlines=false' + '&gid=' + file.getSheetId();
      var token = ScriptApp.getOAuthToken();
      var response = UrlFetchApp.fetch(url + url_ext, {headers: {'Authorization': 'Bearer ' + token}});
      var blob = response.getBlob().setName(pdfName);
      var newFile = DriveApp.getFoldersByName(folderName).next().createFile(blob);
      pdfSheet.setTrashed(true);
      return;
    }
  }
  var inputClients = OSCMaster.getRange(4, 4, OSCMaster.getLastRow()-4).getValues().sort(); //Add unique ID to activeClients
  var activeClients = [];
  activeClients.push(inputClients[0][0]);
  for (i=1; i<inputClients.length; i++){
    if (inputClients[i][0] != inputClients[i-1][0]){
      activeClients.push(inputClients[i][0]);
    }
  }
  var existingClients = OSCClientSheet.getRange(2, 1, OSCClientSheet.getLastRow()-2,8).getValues(); //Create Client w/ info or push to new
  var newClients = [];
  for (i=0; i<activeClients.length; i++){
    var match = false;
    for (j=0;j<existingClients.length; j++){
      if (activeClients[i] == existingClients[j][0]){
        match = true;
        if (existingClients[j][5]){
          var thisClient = new Client(existingClients[j].slice(0,8));
          clients.push(thisClient);
        }
        else {
          newClients.push(existingClients[j][0]);
        }
        break;
      }
    }
    if (!match) {
      newClients.push(activeClients[i]);
      //        SpreadsheetApp.flush();
      //        OSCClientSheet.getRange(OSCClientSheet.getLastRow()+1, 1).setValues([[activeClients[i]]]);
    }
  }
  return clients;
}
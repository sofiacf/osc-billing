//Universal Variables - Monthly Disp/Billing/Payroll Sheet
//TO DO: Summary sheet, post button, rerun for new clients, files go to correct folder, uppercase?, DATE ORDER WTF??
var OSCMaster = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
var monthlyData = OSCMaster.getRange(4, 2, OSCMaster.getLastRow()-1, 13).getValues();
var today = Utilities.formatDate(new Date(), "GMT-4", "M/d/yy");
var billPeriod = OSCMaster.getSheetName();

function setupFolder(){
    var folderName = "Test Invoices Date: " + today;
    var summaryName = "Test Inv. Report - " + today;
    var invoiceFolder = DriveApp.getFoldersByName("COPY OF INVOICES").next();
    function setupSummary(folder){
      var template = DriveApp.getFilesByName("Billing Summary Template").next().getId();
      DriveApp.getFileById(template).makeCopy(summaryName, DriveApp.getFolderById(folder));
    }
    //Delete existing folder if one exists
    if (DriveApp.getFoldersByName(folderName).hasNext()){
      DriveApp.getFoldersByName(folderName).next().setTrashed(true);
    }
    //Make a new folder
    invoiceFolder.createFolder(folderName);
    var folder = DriveApp.getFoldersByName(folderName).next().getId();
    setupSummary(folder);
  }
function setupClients(){
  var clients = [];
  var clientsFolder = DriveApp.getFoldersByName("CLIENTS").next();
  var OSCClientSheet = SpreadsheetApp.open(DriveApp.getFilesByName("Copy of Client Billing Information").next()).getSheets()[0];
    function Client(infoArray){
      this.id = infoArray[0];
      this.name = infoArray[1];
      this.attn = infoArray[2];
      this.address = infoArray[3]
      this.city = infoArray[4];
      this.invNum = infoArray[6] + 1;
      this.suffix = infoArray[7];
      this.fileName = this.id + " - # " + this.invNum + this.suffix + " - " + today;
      if (!DriveApp.getFoldersByName(this.id).hasNext()){
        clientsFolder.createFolder(this.id);
      }
      this.folder = DriveApp.getFoldersByName(this.id).next();
      this.generateInvoice = function(){
        var template = DriveApp.getFilesByName("Copy of Template").next().getId();
        DriveApp.getFileById(template).makeCopy(this.fileName, this.folder);
        var file = SpreadsheetApp.open(DriveApp.getFilesByName(this.fileName).next()).getSheets()[0];
        //Add items
        var items = [];
        var total = 0;
        var inputItems = OSCMaster.getRange(4, 2, OSCMaster.getLastRow()-4, 13).getValues().sort();
        for (i=0; i<inputItems.length; i++){
          if (inputItems[i][2] == this.id){
            items.push([
              inputItems[i][0], //date
              inputItems[i][4], //address
              inputItems[i][5], //company
              inputItems[i][6], //type
              inputItems[i][7], //add charges
              inputItems[i][8], //charge code
              inputItems[i][12]]); //price
            total += inputItems[i][12];
            Logger.log(inputItems[i][0]);
          }
        }
        
        file.insertRows(16, items.length);
        file.getRange(16, 1, items.length, items[0].length).setValues(items).setHorizontalAlignment('general-center').setFontSize(10).setWrap(true);
        SpreadsheetApp.flush();
        //Add info
        file.getRange("F4:F7").setValues([[today], [this.invNum], [String(billPeriod).toUpperCase()], [total]]);
        if (this.attn){
          var clientInfo = [[this.name], [this.attn], [this.address], [this.city]];
        }
        else{
          var clientInfo = [[this.name], [this.address], [this.city], ['']];
        }
        file.getRange("A11:A14").setValues(clientInfo);
        SpreadsheetApp.flush();
        //Format sheet
        var priceColumn = file.getRange(16, 7, items.length);
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
        var response = UrlFetchApp.fetch(url + url_ext, {
                                         headers: {
                                         'Authorization': 'Bearer ' + token
                                         }
                                         });
        var blob = response.getBlob().setName(pdfName);
        var newFile = this.folder.createFile(blob);
        pdfSheet.setTrashed(true);
        return;
      }
    }
    //Remove duplicate client IDs from input sheet, strip down to just client IDs to get active client ID list
    var inputClients = OSCMaster.getRange(4, 4, OSCMaster.getLastRow()-4).getValues().sort();
    var activeClients = [];
    activeClients.push(inputClients[0][0]);
    for (i=1; i<inputClients.length; i++){
      if (inputClients[i][0] != inputClients[i-1][0]){
        activeClients.push(inputClients[i][0]);
      }
    }
    //Check active client IDs against client info sheet, pull details or mark new client
    var existingClients = OSCClientSheet.getRange(2, 1, OSCClientSheet.getLastRow()-2,8).getValues();
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
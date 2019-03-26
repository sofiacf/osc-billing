if (!Array.prototype.includes) {
  Object.defineProperty(Array.prototype, 'includes', {
    value: function(searchElement, fromIndex) {
      if (this == null) throw new TypeError('"this" is null or not defined');
      var o = Object(this), len = o.length >>> 0;
      if (len === 0) return false;
      var n = fromIndex | 0, k = Math.max(n >= 0 ? n : len - Math.abs(n), 0);
      function sameValueZero(x, y) {
        return x === y || (typeof x === 'number' && typeof y === 'number' && isNaN(x) && isNaN(y));
      }
      while (k < len) {
        if (sameValueZero(o[k], searchElement)) return true;
        k++;
      }
      return false;
    }
  });
}
var format,
    today = Utilities.formatDate(new Date(), "GMT-5", "MM/dd/yy"),
    master = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1],
    period = master.getName(),
    data = master.getDataRange().getValues();
function runInvoices(){
  var billing = {
    subDetailSheet: DriveApp.getFilesByName("CLIENT DATA").next(),
    directory: "BILLING", subs: "CLIENTS", sColumn: 3,
    template: function(sub){
      return (sub == "NIXON") ? DriveApp.getFilesByName("NIXON TEMPLATE").next()
      : DriveApp.getFilesByName("BILLING TEMPLATE").next();},
    sheetName: function(sub, det){return sub+" - # "+det.invNum+" - "+today;},
    Item: function(charge) {
      var isNixon = (charge[3] == "NIXON"), line = [],
      els = [0,1,0,0,0,1,1,1,1,1,isNixon,isNixon,0,1];
      for (var i=0; i<els.length; i++)
        if (els[i]) line.push(charge[[i]]);
      return {line: line, amount: charge[13]};},
    Detail: function(det){
      return {
        address: (det[2]) ? [[det[1]], [det[2]], [det[3]], [det[4]]]
              : [[det[1]], [det[3]], [det[4]],[""]],
        invNum: [(det[6] + 1) + det[7]]
      };},
    formatSheet: function(sub, dets, charges, sheet){
      var s = SpreadsheetApp.open(sheet).getSheets()[0], total = 0, items = [];
      for (var i = 0; i < charges.length; i++) {
        total += charges[i].amount;
        items.push(charges[i].line);
      }
      s.insertRows(16, items.length -1 || 1);
      SpreadsheetApp.flush();
      var rs = [s.getRange("A11:A14"), s.getRange(16,1, items.length, items[0].length),
               s.getRange((sub == "NIXON") ? "H4:H7" : "F4:F7"),
               s.getRange(17+items.length, items[0].length)];
      var summary = [[today], dets.invNum, [period], [total]];
      var vals = [dets.address, items, summary, [[total]]];
      for (var i = 0; i < rs.length; i++) rs[i].setValues(vals[i]);
      s.getRange(16,items[0].length, items.length).setNumberFormat('$0.00');
      rs[1].setFontSize(10).setWrap(true);
      SpreadsheetApp.flush();
    }
  };
  format = billing;
  run();
}
function runPayroll(){
  var payroll = {
    subDetailSheet: DriveApp.getFilesByName("RIDER DATA").next(),
    directory: "PAYROLL", subs: "RIDERS", sColumn: 12,
    template: function(na){
      return DriveApp.getFilesByName("PAYROLL TEMPLATE").next();},
    sheetName: function(sub, na){return sub+" Payroll Report: " + today;},
    Item: function(charge) {
      var els = [1, 3, 5, 6, 7, 8, 9, 12, 13, 14], line = [];
      for (var i=0; i<els.length; i++) line.push(charge[els[i]]);
      return {line: line, amount: charge[14]};},
    Detail: function(info){
      return {
        address: (info[2]) ? [[info[1]], [info[2]], [info[3]], [info[4]]]
              : [[info[1]], [info[3]], [info[4]],[""]],
        invNum: [(info[6] + 1) + info[7]]
      };},
    formatSheet: function(sub, details, charges, sheet){
      var s = SpreadsheetApp.open(sheet).getSheets()[0], total = 0, items = [];
      for (var i = 0; i < charges.length; i++) {
        total += charges[i].amount;
        items.push(charges[i].line);
      }
      s.insertRows(16, items.length -1 || 1);
      SpreadsheetApp.flush();
      var ranges = [s.getRange("A11:A14"), //address
                    s.getRange(16,1, items.length, items[0].length), //items
                    s.getRange((sub == "NIXON") ? "H4:H7" : "F4:F7"), //inv info
                    s.getRange(17+items.length, items[0].length)]; //total
      var summary = [[today], details.invNum, [period], [total]];
      var values = [details.address, items, summary, [[total]]];
      for (var i = 0; i < ranges.length; i++) ranges[i].setValues(values[i]);
      s.getRange(16,items[0].length, items.length).setNumberFormat('$0.00');
      ranges[1].setFontSize(10).setWrap(true);
      SpreadsheetApp.flush();
    }
  }
  format = payroll;
  run(payroll);
}
function run(){
  var newSubjects = new Array,
      folder = getRunFolder(),
      subjects = getSubs(),
      charges = getCharges(subjects),
      details = getDetails(subjects);
  for (var i = 0; i < subjects.length; i++){
    if (!details[i]) {
      newSubjects.push(subjects[i]);
      continue;
    }
    var subject = subjects[i],
      scharges = charges[i],
      sdetails = details[i];
    var sheet = getSheet(subject, sdetails, scharges);
    printPDF(subject, sheet, folder);
  }
}

//   var menu = SpreadsheetApp.getUi().createAddonMenu();
//   if (e && e.authMode == ScriptApp.AuthMode.NONE) {
//     menu.addItem('Run invoices', 'runInvoices');
//     menu.addItem('Run payroll', 'runPayroll');
//   }
//   menu.addToUi();
// }
function getRunFolder(){
  const f = format.directory, fName = f + " " + today;
  try {DriveApp.getFoldersByName(fName).next().setTrashed(true);}
  catch(err) {Logger.log("No folder to delete");}
  finally {DriveApp.getFoldersByName(f).next().createFolder(fName);}
  return DriveApp.getFoldersByName(fName).next();
}
function getSubs(){
  var subs = new Array;
  for (var i = 4; i < data.length; i++) {
    var datum = data[i][format.sColumn];
    if (!subs.includes(datum)) subs.push(datum);
  }
  return subs;
}
function getDetails(subs){
  var details = new Array,
    infoSheet = SpreadsheetApp.open(format.subDetailSheet).getSheets()[0],
    sInfo = infoSheet.getDataRange().getValues();
  sInfo.shift();
  for (var i = 0; i<subs.length; i++) {
    details[subs.indexOf(sInfo[i][0])] = format.Detail(sInfo[i]);
    if (!sInfo[i][sInfo[i].length-1]) {
      if (DriveApp.getFoldersByName(sInfo[i][0]).hasNext()){
        var r = infoSheet.getRange(i+2, sInfo[i].length);
        r.setValue(DriveApp.getFoldersByName(sInfo[i][0]).next().getId());
      }
    }
  }
  return details;
}
function getCharges(subs) {
  var charges = new Array;
  for (var i = 4; i<data.length; i++) {
    var charge = format.Item(data[i]),
        index = subs.indexOf(data[i][format.sColumn]);
    if (charges[index]) charges[index].push(charge);
    else charges[index] = [charge];
  }
  return charges;
}
function getSheet(sub, details, charges){
  if (details[details.length-1]) {
    var folder = DriveApp.getFolderById(details[details.length]);
  }
  else {
    if (!DriveApp.getFoldersByName(sub).hasNext())
      DriveApp.getFoldersByName(format.subs).next().createFolder(sub);
    var folder = DriveApp.getFoldersByName(sub).next();
  }
  const template = format.template(sub), name = format.sheetName(sub, details);
  template.makeCopy(name, folder);
  var sheet = DriveApp.getFilesByName(name).next();
  format.formatSheet(sub, details, charges, sheet);
  return sheet;
}
function printPDF(sub, sheet, folder) {
  var id = SpreadsheetApp.open(sheet).getSheetId();
  sheet.makeCopy(sub + "tmp_pdf_copy");
  var tempCopy = DriveApp.getFilesByName(sub + "tmp_pdf_copy").next();
  var url = tempCopy.getUrl();
  var url_ext = 'export?exportFormat=pdf&format=pdf&fitw=true&portrait=false&gridlines=false&gid=' + id;
  url = url.replace('edit?usp=drivesdk', '');
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url + url_ext, {headers: {'Authorization': 'Bearer ' + token}});
  var blob = response.getBlob().setName(sheet.getName());
  var newFile = folder.createFile(blob);
  tempCopy.setTrashed(true);
}

//function Setup(){
//const sTemplate = DriveApp.getFilesByName("Billing Summary Template").next(), sName = "Inv. Report - " + today;
//  sTemplate.makeCopy(sName, f);
//  this.summary = DriveApp.getFilesByName(sName).next(), this.charges = [];
//  for (var i=0; i<mData.length; i++) this.charges.push(new Charge(mData[i][2], mData[i]));
//function makeSummary(run){
//  var runSummary = SpreadsheetApp.open(run.summary).getSheets()[0], expectedTotal = 0;
//  for (i=0; i<run.charges.length; i++) expectedTotal+= run.charges[i].amount;
//  runSummary.getRange("B3:B5").setValues([[today],[billsCreated],[expectedTotal]]);
//  runSummary.getRange(10, 1, run.newClients.length, run.newClients[0].length).setValues(run.newClients);
//  runSummary.getRange(30, 1, run.clientIdList.length).setValues(run.clientIdList);
//  runSummary.getRange("A29").setValue(run.clientIdList.length);
//}
//function generateInvoices(run){
//  for (r=0; r<run.clients.length; r++) run.clients[r].generateInvoice();
//}

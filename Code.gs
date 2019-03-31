//WARN ING!!
//Before implementing, you MUST edit client information file.
//Remove header row and move "attn" line after city/state!
//FINISH WRITING UPDATESUBJECT SEE ROW 100 ADD INDIVIDUAL RATHER THAN BATCH
//OR MAKE NEW LIST TO SAVE TIME OPENING FOLDER
var findFolders = function(f){return DriveApp.getFoldersByName(f);},
    findFiles = function(f){return DriveApp.getFilesByName(f);},
    noFolder = function(name){return !findFolders(name).hasNext()},
    noFile = function(name){return !findFiles(name).hasNext()},
    getData = function(s){return s.getDataRange().getValues();},
    openSheet = function(s,i){return SpreadsheetApp.open(s).getSheets()[i]},
var stp = {
  today: Utilities.formatDate(new Date(), "GMT-5", "MM/dd/yy"),
  master: SpreadsheetApp.getActiveSpreadsheet().getSheets()[1],
  get period(){return this.master.getName()},
  get data(){return getData(this.master).slice(4)},
  subjects: function(config){
    var f = config, sc = f.sc, si = f.sinfo, ss = {}, S = f.Subject;
    si.forEach(function(x){ss[x[0]] = S(x)});
    this.data.forEach(function(x){
      if (!ss.hasOwnProperty(x[sc])) ss[x[sc]] = S([x[sc]]);
      ss[x[sc]].items.push(x[sc].item(x));
    });
    return ss;
  }
};
function run(config){
  var f = config, day = stp.today, subs = stp.subjects(f), data = stp.data;
  function folder() {
    var dir = f.directory, n = dir + " " + day;
    if (findFolders(n).hasNext()) findFolders(n).next().setTrashed(true);
    findFolders(dir).next().createFolder(n);
    return findFolders(n).next();
  };
  var sheets = [], billed = [], newSubs = [];
  for (var sub in subs) {
    if !sub.items continue;
    sheets.push(sub.summary);
    billed.push(sub);
    if (sub.isNew) newSubs.push(sub);
  };
  function printPdf(summary){
    var s = summary, n = s.getName(), rf = folder();
    var id = SpreadsheetApp.open(s).getSheetId();
    s.makeCopy(n + "tmp_pdf_copy");
    var tmp, url, x, tkn, response, blob, newFile;
    tmp = findFiles(n + "tmp_pdf_copy").next(), url = tmp.getUrl();
    x = 'export?exportFormat=pdf&format=pdf&fitw=true&portrait=false&gridlines=false&gid=' + id;
    url = url.replace('edit?usp=drivesdk', ''), tkn = ScriptApp.getOAuthToken();
    var r = UrlFetchApp.fetch(url + x,{headers: {'Authorization': 'Bearer ' + tkn}});
    var blob = r.getBlob().setName(n), newFile = rf.createFile(blob);
    tmp.setTrashed(true);
  }
  sheets.forEach(printPdf);
}
var billing = {
  data: findFiles("BILLING DATA").next(), folder: findFolders("CLIENTS").next(),
  sc: 3, sinfo: getData(openSheet(this.data,0)),
  Subject: function(s){
    var nix = (s[0] == "NIXON"), isNew = (s.length<2);
    if (isNew) s = updateSubject(s);
    var id = s[0], folder = DriveApp.getFolderById(s[1]),
    form = DriveApp.getFileById(s[2]), invoiceNumber = [(s[3]+1)+s[4]];
    var last = nix ? 12 : 10, fname = id + "- #" + invoiceNumber;
    return {
      id: id, folder: folder, form: form, items: [], isNew: isNew, fname: fname,
      num: invoiceNumber, lastBillDate: [s[5]],
      item: function(i){return {amount: i[13], line: i[1].concat(i.slice(5,last))}},
      get lines(){return this.items.map(function(x){x.line});}
      get total(){return [this.items.reduce(function(a,b){a+b.amount})];},
      get summary(){
        var c = openSheet(this.form.makeCopy(this.fname, this.folder),0), rs, vs;
        var l = this.lines.length, w = this.lines[0].length, p = stp.period;
        c.insertRows(16,l);
        SpreadsheetApp.flush();
        rs = [c.getRange(16,1,l,w), c.getRange(4,w-1,4), c.getRange(17+l,w)];
        vs = [this.lines, [stp.today, this.num, p, this.total], this.total];
        rs.forEach(setValues(vs));
        c.getRange(16, w, l).setNumberFormat('$0.00')
        rs[0].setFontSize(10).setWrap(true);
        SpreadsheetApp.flush();
        return c;
      }
    }
  }
}
function runInvoices(){
  run(billing);
}
class Subject{
  constructor(a){
    this.id = a[0], this.folder = a[1], this.form = a[2];
    this.item
  }
}
function runPayroll(){
  var payroll = {
    data: findFiles("PAYROLL DATA"), sc: 12,
    Subject: function(s){
      var isNew = (s.length < 2);
      if (isNew) s = updateSubject();
      return {
        id: s[0], folder: s[1], form: s[2], items: [], isNew: isNew
        fname: s[0] + " Payroll Report: " + stp.day;},
        item: function(i) {
          var els = [1, 3, 5, 6, 7, 8, 9, 12, 13, 14], line = [];
          els.forEach(function(x){line.push(i[x]));
          return {line: line, amount: arr[14]};}
        },
        get total(){return [this.items.reduce(function(a,b){a+b.amount})];},
        get summary(){

          if (items>1) s.insertRows(16, items.length -1);
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
      }},
  }
  run(payroll);
}
var config = billing;
function updateSubject(s){
  var f = config, d = f.data;
}
function updateSubjects(){
  //WRITE MIGRATE FUNCTION FOR SUBJECT INFO SHEET, IMPLEMENT THERE
  var f = config, d = openSheet(f.sSheet,0),
  subs = d.getRange(1,1,d.getLastRow()).getValues();
  var sSheet = (noFile(f.nss)) ? SpreadsheetApp.create(f.nss).getSheets()[0]
  : openSheet(findFiles(f.nss).next(),0);
  sSheet.getRange(1,1,subs.length).setValues(subs);
  var data = getData(d), ids = [], runInfo = [];
  for (var i=0; i<data.length; i++){
    var a = data[i], s = a[0] tn = a[0] + " TEMPLATE";
    var fl = (noFolder(s) ? f.folder.createFolder(s) : findFolders(s).next();
    var t = (noFile(tn)) ? f.form.makeCopy(tn, fl) : findFiles(tn).next();
    ids.push([[fl.getId()], [t.getId()]]);
    runInfo.push([[a[6]],[a[7]],[a[5]]]);
    f.addressRange(openSheet(t,0)).setValues(f.address(a));
  }
  sSheet.getRange(1,2,subs.length,2).setValues(ids);
  sSheet.getRange(1,3,subs.length,3).setValues(runInfo);
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

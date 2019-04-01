//Remove header row and move "attn" line after city/state!
//FINISH UPDATESUBJECT or MAKE NEW LIST TO SAVE TIME OPENING FOLDER
function printPdf(s: GoogleAppsScript.Drive.File, f: GoogleAppsScript.Drive.Folder){
  let n = s.getName(), tmp = s.makeCopy(n + "tmp_pdf_copy");
  let url = tmp.getUrl(), id = open(s).getSheetId();
  let x = 'export?exportFormat=pdf&format=pdf&fitw=true&portrait=false&gridlines=false&gid=' + id;
  url = url.replace('edit?usp=drivesdk', '');
  let tkn = ScriptApp.getOAuthToken();
  let r = UrlFetchApp.fetch(url + x,{headers: {'Authorization': 'Bearer ' + tkn}});
  r.getBlob().setName(n);
  f.createFile(r);
  tmp.setTrashed(true);
}
function findFolders(n: string){return DriveApp.getFoldersByName(n);}
function findFiles(n: string){return DriveApp.getFilesByName(n);}
function noFolder(n: string){return !findFolders(n).hasNext()}
function noFile(n: string){return !findFiles(n).hasNext()}
function open(s: GoogleAppsScript.Drive.File){return SpreadsheetApp.open(s)}
const today = Utilities.formatDate(new Date(), "GMT-5", "MM/dd/yy");
const master = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
let period = master.getName(), data = master.getDataRange().getValues().slice(4);
class Format{
  n: string; ss: GoogleAppsScript.Drive.File;
  folder: GoogleAppsScript.Drive.Folder; st: GoogleAppsScript.Drive.File;
  si: Object[][]; sc: number; fn: string; sub: any;
  constructor(n: string, sc: number, sub: CallableFunction){
    this.n = n, this.ss = findFiles(this.n + " DATA").next(), this.sc = sc;
    this.sub = sub, this.folder = findFolders(n).next(),
    this.si = open(this.ss).getSheets()[0].getDataRange().getValues();
    this.st = findFiles(this.n + " SUMMARY TEMPLATE").next();
  }
  subjects(){
    let ss = {};
    this.si.forEach((x: Array<any>) =>{ss[x[0]] = new this.sub(x)});
    data.forEach((x: Array<any>) => {
      if (!ss.hasOwnProperty(x[this.sc])) ss[x[this.sc]] = updateSub(x);
      ss[x[this.sc]].items.push(ss[x[this.sc]].item([x]));
    });
    return ss;
  }
  run(){
    function folder() {
      var n = this.fName + " " + today;
      if (!noFolder(n)) findFolders(n).next().setTrashed(true);
      return this.folder.createFolder(n);
    }
    var subs = this.subjects(), billed = [];
    for (let sub in subs) {
      printPdf(subs[sub].summary(), folder());
      billed.push(subs[sub]);
    };
  }
}
abstract class Subject{
  id: string; items: Array<Array<any>>; amounts: Array<number>;
  folder: GoogleAppsScript.Drive.Folder; form: GoogleAppsScript.Drive.File;
  constructor(a: Array<any>){
    this.id = a[0], this.folder = DriveApp.getFolderById(a[1]);
    this.form = DriveApp.getFileById(a[2]), this.items = [], this.amounts = [];
  }
  abstract item(i: Array<any>): Array<any>;
  total(){return this.amounts.reduce((a,b)=>a+b);}
}
class Client extends Subject{
  nix: boolean; num: Array<number>; fn: string; lastBillDate: Array<any>;
  constructor(a: any[]){
    super(a);
    this.nix = (a[0] == "NIXON"), this.num = [a[3]+1+a[4]],
    this.fn = a[0] + "- #" + this.num, this.lastBillDate = [a[5]];
  }
  item(i: any[]) {return [i[1].concat(i.slice(5,this.nix ? 12:10))]}
  summary() {
    let c = open(this.form.makeCopy(this.fn, this.folder)).getSheets()[0];
    let l = this.items.length, w = this.items[0].length;
    c.insertRows(16,l);
    SpreadsheetApp.flush();
    let rs = [c.getRange(16,1,l,w), c.getRange(4,w-1,4), c.getRange(17+l,w)];
    let vs: Array<Array<any>> = [this.items, [[today], this.num, [period], this.total], [this.total]];
    rs.forEach((x,i) => x.setValues(vs[i]));
    c.getRange(16, w, l).setNumberFormat('$0.00')
    rs[0].setFontSize(10).setWrap(true);
    SpreadsheetApp.flush();
    return c;
  }
}
class Rider extends Subject{
  item(i: any[]): any[] {return [1,3,5,6,7,8,9,12,13,14].map(x => i[x])}
  constructor(a: any[]){super(a)}
}
let billing = new Format("BILLING", 3, (x: any[])=> new Client(x));
let payroll = new Format("PAYROLL", 12, (x: any[]) => new Rider(x));
function updateSub(s: any[]){Logger.log(s);}// FIXME
// function updateSubjects(){
  //WRITE MIGRATE FUNCTION FOR SUBJECT INFO SHEET, IMPLEMENT THERE
//   var f = config, d = openSheet(f.sSheet,0),
//   subs = d.getRange(1,1,d.getLastRow()).getValues();
//   var sSheet = (noFile(f.nss)) ? SpreadsheetApp.create(f.nss).getSheets()[0]
//   : openSheet(findFiles(f.nss).next(),0);
//   sSheet.getRange(1,1,subs.length).setValues(subs);
//   var data = getData(d), ids = [], runInfo = [];
//   for (var i=0; i<data.length; i++){
//     var a = data[i], s = a[0] tn = a[0] + " TEMPLATE";
//     var fl = (noFolder(s) ? f.folder.createFolder(s) : findFolders(s).next();
//     var t = (noFile(tn)) ? f.form.makeCopy(tn, fl) : findFiles(tn).next();
//     ids.push([[fl.getId()], [t.getId()]]);
//     runInfo.push([[a[6]],[a[7]],[a[5]]]);
//     f.addressRange(openSheet(t,0)).setValues(f.address(a));
//   }
//   sSheet.getRange(1,2,subs.length,2).setValues(ids);
//   sSheet.getRange(1,3,subs.length,3).setValues(runInfo);
// }

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

function printPdf(s: GoogleAppsScript.Drive.File, f: GoogleAppsScript.Drive.Folder){
  let n = s.getName(), tmp = s.makeCopy(n + "tmp_pdf_copy");
  let url = tmp.getUrl(), id = SpreadsheetApp.open(s).getSheetId();
  let x = 'export?exportFormat=pdf&format=pdf&fitw=true&portrait=false&gridlines=false&gid=' + id;
  url = url.replace('edit?usp=drivesdk', '');
  let tkn = ScriptApp.getOAuthToken();
  let r = UrlFetchApp.fetch(url + x,{headers: {'Authorization': 'Bearer ' + tkn}});
  r.getBlob().setName(n);
  f.createFile(r);
  tmp.setTrashed(true);
}
const today = Utilities.formatDate(new Date(), "GMT-5", "MM/dd/yy");
const master = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
let period = master.getName(), data: Array<Array<any>> = master.getDataRange().getValues().slice(4);
class Format{
  n: string; f: GoogleAppsScript.Drive.Folder; s: Function; h: string[];
  sf: GoogleAppsScript.Drive.Folder; summary: GoogleAppsScript.Drive.File;
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet; si: any[][]; sc: number;
  as: Array<string>; ns: Array<string>; sitem: Function;
  t: GoogleAppsScript.Drive.File; sn: string;
  constructor(n: string, sc: number, sn: string, sitem: Function){
    this.n = n; this.f = DriveApp.getFoldersByName(n).next();
    this.t = this.f.getFilesByName("TEMPLATE").next(); this.sc = sc;
    this.ss = SpreadsheetApp.open(this.f.getFilesByName("DATA").next());
    this.sitem = sitem; this.sf = this.f.getFoldersByName(sn).next();
    this.si = this.ss.getDataRange().getValues(); this.h = this.si[0];
    this.as = data.map(x=>x[sc]).filter((x,i,a)=>(i==a.indexOf(x)));; this.sn = sn;
    this.ns = this.as.filter((x)=>this.si.map(x=>x[0]).indexOf(x)<0);
    this.summary = this.f.getFilesByName("SUMMARY TEMPLATE").next();
  }
  subs(){
    let l = this.h.length, a = new Array(l-3); a.forEach(x=>[]);
    function update(s:string){
      let sf = this.sf, tn = "${s} TEMPLATE";
      let folders = sf.getFoldersByName(s);
      let f = (folders.hasNext()) ? folders.next() : sf.createFolder(s);
      let files = sf.getFilesByName(tn);
      let t = files.hasNext() ? files.next() : this.t.makeCopy(tn,sf)
      return [[s],[f.getId()],[t.getId()]];
    }
    let os:any[][] = this.ss.getDataRange().getValues(), ns = this.ns;
    let updatens = ns.map(x=>update(x).concat(a));
    os.shift();
    let updateks = os.map((x)=>update(x[0]).concat(x.slice(3)));
    let newsi = updatens.concat(updateks).sort();
    this.ss.getSheets()[0].getRange(2,1,newsi.length,l).setValues(newsi);
    SpreadsheetApp.flush();
    return new Map(newsi.map(x=>[x[0], new Map(this.h.map((y,j)=>[y,x[j]]))]));
  }
  items(){
    let as = this.as, items = new Array(as.length).forEach(x=>[]);
    data.forEach(x=>items[this.as.indexOf(x[this.sc])].push(this.sitem(x)));
    return items;
  }
  run(){
    let n = this.n + " " + today, old = this.f.getFoldersByName(n);
    if (old.hasNext()) old.next().setTrashed(true);
    let f = this.f.createFolder(n), items = this.items();
    let subs = this.as.map((x)=>new Subject(subs[x], items[x]));
    subs.forEach((x:Subject)=>x.run(f));
    let s = SpreadsheetApp.open(this.summary.makeCopy("SUMMARY",f));
    subs.forEach((x:Subject)=>x.info.push([this.ns.indexOf(x.id) > -1 ? "new" : "ran"]));
    let r = s.getSheets()[0].getRange(10,1,this.as.length,subs[0].info.length);
    r.setValues(subs.map((x:Subject)=>x.info));
    s.getSheets()[0].getRange(2,2).setValue(subs.length);
  }
}
class Subject{
  id: string; items: any[][]; info: any[][]; fn: string;
  f: GoogleAppsScript.Drive.Folder; t: GoogleAppsScript.Drive.File;
  constructor(a: Map<string,any>, items: any[][]){
    this.id = a["id"], this.f = DriveApp.getFolderById(a["fix"]);
    this.t = DriveApp.getFileById(a["t"]); this.items = items;
    this.info = [[today]].concat(a["n"] ? [a["n"]+1+a["s"], [period]] : [[period]]);
    this.fn = this.id + "- " + (a["n"]) ? "#" + this.info[1] : period;
  }
  run(folder: GoogleAppsScript.Drive.Folder){
    let sheet = this.t.makeCopy(this.fn,this.f)
    let c = SpreadsheetApp.open(sheet).getSheets()[0];
    let l = this.items.length, w = this.items[0].length;
    c.insertRows(16,l);
    SpreadsheetApp.flush();
    c.getRange(16,1,l,w).setValues(this.items).setFontSize(10).setWrap(true);
    c.getRange(4,2-1,this.info.length).setValues(this.info);
    c.getRange(16, w, l).setNumberFormat('$0.00');
    SpreadsheetApp.flush();
    printPdf(sheet,folder);
  }
}
function runInvoices(){
  function item(i:any[]){return [[i[1]].concat(i.slice(5,i[3] == "NIXON" ? 12:10))]}
  let billing = new Format("BILLING", 3, "CLIENTS", item);
  billing.run();
}
function runPayroll(){
  function item(i:any[]){return [i[1]].concat([i[3]], i.slice(5,10), i.slice(12,15))}
  let payroll = new Format("PAYROLL", 12, "RIDERS", item);
  payroll.run();
}

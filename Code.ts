function printPdf(s: GoogleAppsScript.Drive.File, f: GoogleAppsScript.Drive.Folder){
  let n = s.getName(), tmp = s.makeCopy(n + "tmp_pdf_copy");
  let url = tmp.getUrl(), id = SpreadsheetApp.open(s).getSheetId();
  let x = 'export?exportFormat=pdf&format=pdf&fitw=true&portrait=false&gridlines=false&gid=' + id;
  url = url.replace('edit?usp=drivesdk', '');
  let tkn = ScriptApp.getOAuthToken();
  let r = UrlFetchApp.fetch(url + x,{headers: {'Authorization': 'Bearer ' + tkn}});
  r.getBlob().setName(n);
  tmp.setTrashed(true);
  return f.createFile(r);
}
const today = Utilities.formatDate(new Date(), "GMT-5", "MM/dd/yy");
const master = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
let bp = master.getName();
let input: Array<Array<any>> = master.getDataRange().getValues().slice(3);
var Format = function(name:string, subjectColumn:number, subjectName:string, itemFunction: Function){
  const NAME = name;
  const SN = subjectName;
  const ITEM = itemFunction;
  const FOLDER = DriveApp.getFoldersByName(name).next();
  const getFirstEl = function(arr){return (!arr[1].length) ? undefined : arr[0];}
  const getEsIn = function(vals,check){return vals.filter(function(e) {return(check[check.indexof(e)]>-1)});}
  const getSC = function(arr){return arr.map(function(e){return e[subjectColumn];});}
  const getUniq = function(dup){return dup.filter(function(e,i,a){return (i==a.indexOf(e));});}
  const addEmptyEls = function(arr,length){
      var _arr = arr.slice(0), _l = length;
      return _arr.map(function(e){
        var len = e.length;
        for (var i = len - _l; i<_l-1; i++) e[i] = [ ];
        return e;
      });
    }
  const SUBS = function(){
    const iSubs = getSC(input);
    const activeSubs = getUniq(iSubs);
    const dataS = SpreadsheetApp.open(FOLDER.getFilesByName("DATA").next());
    const data:any[][] = dataS.getDataRange().getValues().slice(0);
    const headers = data.slice(0,1);
    const width = headers.length;
    data.shift();
    const dataIds = input.map(getFirstEl);
    const newSubs = getEsIn(activeSubs,dataIds);
    const toRefresh = data.filter(function(e){return (e[1] == e[2]) ? true : false;}); //Fid should != Tid
    if (!newSubs.length && !toRefresh.length){//exit if none to update
      const subs = data.filter(function(x){return (activeSubs.indexOf(x[0]) > -1) ? true : false;});
      return subs;
    }
    const newSubArrs = addEmptyEls(newSubs, width);
    const updates = (newSubs.length) ? toRefresh.concat(newSubArrs).slice(0) : toRefresh.slice(0);
    const uIds: string[]= updates.map(getFirstEl); //ids of updating subs
    const subsF = FOLDER.getFoldersByName(SN).next(); //open subs folder
    const uFs = uIds.map(function(x){//find or make Fs of updating subs
      let folders = subsF.getFoldersByName(x);
      return (folders.hasNext()) ? folders.next() : subsF.createFolder(x);});
    const fT =  FOLDER.getFilesByName("TEMPLATE").next(); //open generic template
    const uTids = uFs.map(function(x,i){//find or make Tids of updating subs
      let tn = uIds[i] + " TEMPLATE", fs = x.getFilesByName(tn);
      return (fs.hasNext()) ? fs.next().getId() : fT.makeCopy(tn).getId();});
    const uFids = uFs.map(function(x){return x.getId()}); //Fids of updating subs
    const newData = data.slice(0); //copy array of data sheet values
    updates.forEach(function(x,i){//update newdata with Fids & Tids
      x[1] = uFids[i];
      x[2] = uTids[i];
      let index = dataIds.indexOf(x[0]);
      if (index < 0) newData.push(x);
      else newData[index] = x;
    });
    newData.sort();
    let newSData = newData.slice(0);
    let height = newSData.unshift(headers);
    let S = dataS.getSheets()[0];
    S.getRange(1,1,data.length,data[0].length).clearContent();
    S.getRange(1,1,height, width).setValues(newSData);
    return newData.filter(function(x){return activeSubs.indexOf(x[0]) > -1;});
  }
  this.RUN = function(){
    const name = NAME + " " + today;
    const oldF = FOLDER.getFoldersByName(name);
    if (oldF.hasNext()) oldF.next().setTrashed(true);
    const folder = FOLDER.createFolder(name);
    const subs = SUBS().slice(0);
    const sids = subs.map(getFirstEl);
    const array = new Array(sids.length);
    const items = addEmptyEls(array,0);
    input.forEach(function(x){items[sids.indexOf(x[subjectColumn])].push(ITEM(x));});
    const rsubs = subs.map(function(x,i){const subject = new Subject(x,items[i]); return subject;});
    rsubs.forEach(function(x){x.run(folder);});
    const summary = FOLDER.getFilesByName("SUMMARY TEMPLATE").next();
    const s = SpreadsheetApp.open(summary.makeCopy("SUMMARY",folder));
    const r = s.getSheets()[0].getRange(10,1,rsubs.length,rsubs[0].info.length);
    r.setValues(rsubs.map(function(x){return x.info;}));
    s.getSheets()[0].getRange(2,2).setValue(rsubs.length);
  }
}
var Subject = function(a: any[], items: any[][]){
  this.fid = a[1]; this.tid = a[2]; this.num = (a[4]) ? (a[4]+1)+a[5] : 0;
  this.info = (this.num) ? [[today],[this.num],[bp]] : [[today],[bp]];
  this.items = items; this.fn = a[0] +"- "+(this.num)?("# "+ this.num) : bp;
  this.go = function(folder: GoogleAppsScript.Drive.Folder){
    const f = DriveApp.getFolderById(this.fid);
    const sheet = DriveApp.getFileById(this.tid).makeCopy(this.fn,f);
    let c = SpreadsheetApp.open(sheet).getSheets()[0];
    let l = this.items.length, w = this.items[0].length;
    c.insertRows(16,l);
    SpreadsheetApp.flush();
    c.getRange(16,1,l,w).setValues(this.items).setFontSize(10).setWrap(true);
    c.getRange(4,w-1,this.info.length).setValues(this.info);
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

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
var findFolders = DriveApp.getFoldersByName,
    findFiles = DriveApp.getFilesByName,
    noFolder = function(name){return !findFolders(name).hasNext()},
    noFile = function(name){return !findFiles(name).hasNext()},
    getData = function(s){return s.getDataRange().getValues();},
    openSheet = function(s,i){return SpreadsheetApp.open(s.getSheets[i])},
    newFolder = function(name, dest){
      dest.createFolder(name);
      return dest.findFolders(name).next();
    },
    newFile = function(name, dest){
      dest.createFile(name, dest);
      return dest.getFilesByName(name).next();
    },
    newCopy = function(o, name, dest){
      o.makeCopy(name, dest);
      return findFiles(name).next();
    };
var stp = {
  today: Utilities.formatDate(new Date(), "GMT-5", "MM/dd/yy"),
  master: SpreadsheetApp.getActiveSpreadsheet().getSheets()[1],
  get period(){return this.master.getName()},
  get data(){return getData(this.master).slice(4)},
  subjects: function(config){
    var subs = new Array, d = this.data, sc = config.sc;
    for (var i = 4; i < d.length; i++) {
      var s = config.Subject(d[i][sc]);
      if (!subs.includes(s)) subs.push(s);
    };
    subs.sort();
    return subs;
  },
  indexedSubInfo: function(config){
    var info = [], ss = openSheet(config.ss,0), o = getData(ss);
    o.shift();
    for (var i = 0; i<o.length; i++) info[subs.indexOf(o[i][0])] = o[i];
    return info;
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
  function details() {return stp.indexedSubInfo.map(f.Detail);};
  function charges(){
    var c = [], d = data, sc = f.sColumn;
    for (var i = 0; i<d.length; i++)
      c[subs.indexOf(d[i][sc])] = (c[subs.indexOf(d[i][sc])] || []).concat(d[i]);
    return c;
  };
  function printPdf(r){
    var n = r.name, s = r.format(r.copy), folder = this.folder(),
      id = SpreadsheetApp.open(s).getSheetId();
    s.makeCopy(n + "tmp_pdf_copy");
    var tmp = findFiles(n + "tmp_pdf_copy").next();
    var url = tmp.getUrl();
    var x = 'export?exportFormat=pdf&format=pdf&fitw=true&portrait=false&gridlines=false&gid=' + id;
    url = url.replace('edit?usp=drivesdk', '');
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url + x, {headers: {'Authorization': 'Bearer ' + token}});
    var blob = response.getBlob().setName(n);
    var newFile = folder.createFile(blob);
    tempCopy.setTrashed(true);
  }
  var newSubs = [], ds = details(), cs = charges();
  for (var i=0; i<subs.length; i++){
    if (ds[i]) {
      Logger.console.log(subs[i]);
      // var statement = f.Statement(ss[i], ds[i], cs[i]);
      // printPdf(statement.format(statement.copy);
    }
    else newSubs.push(ss[i]);
  }
}
function updateSubjects(config){
  var f = config, subs = stp.subjects(f), ss = f.sSheet, sf = f.sFolder,
  tmp = f.template, d = getData(openSheet(ss,0));
  for (var i = 0; i < subs.length; i++) {
    var s = subs[i];
    var fl = findFolders(s) || newFolder(s,sf);
    var t = findFiles(s + " TEMPLATE") || newCopy(tmp, s + " TEMPLATE", fl);
    var o = openSheet(t,0);
    if (r.getRange(4,1,4).getValues()) return;
    for (var i = 0; i < d.length; i++){
      if (d[i][0] != s) continue;
      var detail = f.Detail(d[i]);
      o.getRange(4,1,4).setValues(detail.address);
      o.getRange(4,2).setValue(detail.invNum);
      if (!d[i]) Logger.log(s);
    };
  };
}
function runInvoices(){
  var billing = {
    nss: "NEW CLIENT DATA",
    directory: "BILLING", sc: 3, sFolder: findFolders("CLIENTS").next(),
    sSheet: DriveApp.getFilesByName("CLIENT DATA").next(),
    template: DriveApp.getFilesByName("BILLING TEMPLATE").next(),
    Subject: function(name){return {name: name}},
    Item: function(arr){return {amount: arr[13],
      line: arr[1].concat(arr.slice(5, (arr[3]=="NIXON") ? 12 : 10), arr[13])};
    },
    Detail: function(arr){return {folderId: arr[arr.length-1],
      invNum: [(arr[6] + 1) + arr[7]],
      address: (arr[2]) ? arr.slice(1,5) : arr[1].concat(arr[3],arr[4],arr[2])}
    },
    Statement: function(subjects, details, charges){
      var s = this.Subject(subjects), d = this.Detail(details),
      c = charges.map(this.Item), fid = d.folderId;
      if (!fid && !findFolders(sub.name).hasNext())
          findFolders("CLIENTS").next().createFolder(s.name);
      fid = fid || DriveApp.getFoldersByName(s.name).next().getId();
      return {
        name: [s.name, "- #", d.invNum, "-", today].join(" "),
        template: DriveApp.getFilesByName(s.templateName).next(),
        get copy() {
          return newCopy(this.template, this.name, DriveApp.getFolderById(fid));
        },
        total: [c.reduce(function(a,b){return a+b})], items: c.map(function(x){return x.line}),
        format: function(sheet){
          var s = SpreadsheetApp.open(sheet).getSheets()[0],
          l = this.items.length, w = this.items[0].length;
          s.insertRows(16,l);
          SpreadsheetApp.flush();
          var summary = [[today], d.invNum, [period], this.total],
          vals = [d.address, this.items, summary, [this.total]],
          rs = [s.getRange(4, 1, 4), s.getRange(16,1,l,w),
            s.getRange(4, w-1, 4), s.getRange(17+l,w)];
          for (var i = 0; i < rs.length; i++) rs[i].setValues(vals[i]);
          s.getRange(16, w, l).setNumberFormat('$0.00');
          rs[1].setFontSize(10).setWrap(true);
          SpreadsheetApp.flush();
          return sheet;
        }
      }
    }
  };
  run(billing);
}
function runPayroll(){
  var payroll = {
    directory: "PAYROLL", sColumn: 12,
    get subsSheet() {return DriveApp.getFilesByName("RIDER DATA").next()},
    Subject: function(name){return {name: name}},
    Item: function(arr) {
      var els = [1, 3, 5, 6, 7, 8, 9, 12, 13, 14], line = [];
      for (var i=0; i<els.length; i++) line.push(charge[els[i]]);
      return {line: line, amount: charge[14]};},
    template: function(na){
      return DriveApp.getFilesByName("PAYROLL TEMPLATE").next();},
    sheetName: function(sub, na){return sub+" Payroll Report: " + today;},
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
      s.insertRows(16, items.length -1 || 2);
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
  run(payroll);
}

function migrate(config){
  var config = f, newSubSheet = SpreadsheetApp.create(f.nss),
  d = openSheet(f.sSheet,0), subs = d.getRange(2,1,d.getLastRow()-1),
  r = openSheet(newSubSheet,0).getRange(1,1,d.length);
  r.setValues(subs)
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

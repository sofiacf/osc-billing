function printPdf(s, folder) {
    const n = s.getName(), tmp = s.makeCopy(n + "tmp_pdf_copy");
    const url = tmp.getUrl(), id = SpreadsheetApp.open(s).getSheetId();
    const x = 'export?exportFormat=pdf&format=pdf&fitw=true&portrait=false&gridlines=false&gid=' + id;
    url = url.replace('edit?usp=drivesdk', '');
    const tkn = ScriptApp.getOAuthToken();
    const r = UrlFetchApp.fetch(url + x, { headers: { 'Authorization': 'Bearer ' + tkn } });
    const blob = r.getBlob().setName(n);
    tmp.setTrashed(true);
    return folder.createFile(blob);
}
var master = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
var bp = master.getName();
var input = master.getDataRange().getValues().slice(3);
function FORMAT(name, subjectColumn, subjectName, itemFunction) {
    const today = Utilities.formatDate(new Date(), "GMT-5", "MM/dd/yy");
    const NAME = name;
    const SN = subjectName;
    const ITEM = itemFunction;
    const FOLDER = DriveApp.getFoldersByName(name).next();
    const getFirstEl = function (arr) { return arr[0]; };
    function getEsIn(vals, check) { const filtered = vals.filter(function (e) { return (check[check.indexOf(e)] > -1) ? true : false }); return filtered; };
    function getSC(arr) { return arr.map(function (e) { return e[subjectColumn]; }); };
    function getUniq(dup) { return dup.filter(function (e, i, a) { return (i == a.indexOf(e)); }); };
    function addEmptyEls(arr, length) {
        const _arr = arr.slice(0);
        const _l = length;
        return (_arr.map(function (e) {
            if (!e) e = [ ];
            var len = e.length;
            for (var i = len - _l; i < _l - 1; i++) e[i] = [];
            return e;
        }));
    };
  return {
    get SUBS(){
        const iSubs = getSC(input).slice(0);
        const activeSubs = getUniq(iSubs);
        const dataS = SpreadsheetApp.open(FOLDER.getFilesByName("DATA").next());
        const data = dataS.getDataRange().getValues().slice(0);
        const headers = data[0];
        const width = headers.length;
        data.shift();
        const dataIds = data.map(getFirstEl);
        const newSubs = getEsIn(activeSubs, dataIds);
        const toRefresh = data.filter(function (e) { return e[1] == e[2] ? true : false; }); //Fid should != Tid
        if (!newSubs.length && !toRefresh.length) { //exit if none to update
            const subs = data.filter(function (x) { return (activeSubs.indexOf(x[0]) > -1) ? true : false; });
            return subs;
        }
        const newSubArrs = addEmptyEls(newSubs, width);
        const updates = (newSubs.length) ? toRefresh.concat(newSubArrs).slice(0) : toRefresh.slice(0);
        const uIds = updates.map(getFirstEl); //ids of updating subs
        const subsF = FOLDER.getFoldersByName(SN).next(); //open subs folder
        const uFs = uIds.map(function (x) {
            const folders = subsF.getFoldersByName(x);
            return (folders.hasNext()) ? folders.next() : subsF.createFolder(x);
        });
        const fT = FOLDER.getFilesByName("TEMPLATE").next(); //open generic template
        const uTids = uFs.map(function (x, i) {
            const tn = uIds[i] + " TEMPLATE", fs = x.getFilesByName(tn);
            return (fs.hasNext()) ? fs.next().getId() : fT.makeCopy(tn).getId();
        });
        const uFids = uFs.map(function (x) { return x.getId(); }); //Fids of updating subs
        const newData = data.slice(0); //copy array of data sheet values
        updates.forEach(function (x, i) {
            x[1] = uFids[i];
            x[2] = uTids[i];
            const index = dataIds.indexOf(x[0]);
            if (index < 0) newData.push(x);
            else newData[index] = x;
        });
        newData.sort();
        const newSData = newData.slice(0);
        const height = newSData.unshift(headers);
        const S = dataS.getSheets()[0];
        S.getRange(1, 1, data.length, data[0].length).clearContent();
        S.getRange(1, 1, height, width).setValues(newSData);
        return (newData.filter(function (x) { return (activeSubs.indexOf(x[0]) > -1) ? true : false;}));
    },
      RUN: function(){
        const name = NAME + " " + today;
        const oldF = FOLDER.getFoldersByName(name);
        if (oldF.hasNext()) oldF.next().setTrashed(true);
        const folder = FOLDER.createFolder(name);
        const subs = this.SUBS.slice(0);      
        function sortLineItems(){
            const sids = subs.map(getFirstEl);
            const array = new Array(sids.length);
            const items = addEmptyEls(array, 0);
            input.forEach(function (x) { items[sids.indexOf(x[subjectColumn])].push(ITEM(x)); });
            return items.slice(0);
        }
        const lineItems = sortLineItems();
        for (var i = 0; i<subs.length; i++){new SUBJECT(subs[i], lineItems[i]).go(folder);}
        const summary = FOLDER.getFilesByName("SUMMARY TEMPLATE").next();
        const s = SpreadsheetApp.open(summary.makeCopy("SUMMARY", folder));
        const r = s.getSheets()[0].getRange(10, 1, subs.length, rsubs[0].info.length);
        r.setValues(rsubs.map(function (x) { return x.info; }));
        s.getSheets()[0].getRange(2, 2).setValue(subs.length);
    }
  }
};
function SUBJECT(FDATA, FSITEMS) {
  var snum = (FDATA[4] > 0) ? (FDATA[4] + 1) + FDATA[5] : false;
  var sname = FDATA[0], sfid = FDATA[1], stid = FDATA[2], sitems = FSITEMS.slice(0);
  var sinfo = (snum) ? [[today],[snum], [bp]] : [[today],[bp]];
  var SFN = (snum) ? Utilities.formatString("%s - %s # %s", sname, snum, today) : Utilities.formatString("%s - %s", sname, bp);
  Logger.log(SFN);
  return {
    items: sitems,
    name: sname,
    folderid: sfid,
    templateid: stid,
    infoVals: sinfo,
    sfn: SFN,
    go: function(RUNFOLDER) {
        const subfolder = DriveApp.getFolderById(this.folderid);
        const file = DriveApp.getFileById(this.templateid).makeCopy(this.sfn, subfolder);
        const sheet = SpreadsheetApp.open(file).getSheets()[0];
        const numRows = this.items.length, numCols = this.items[0].length;
        if (numRows > 1) sheet.insertRows(16, numCols-1);
        SpreadsheetApp.flush();
        sheet.getRange(16, 1, numRows, numCols).setValues(this.items).setFontSize(10).setWrap(true);
        sheet.getRange(4, numCols-1, this.infoVals.length).setValues(this.infoVals);
        sheet.getRange(16, numCols, numRows).setNumberFormat('$0.00');
        SpreadsheetApp.flush();
        printPdf(file, RUNFOLDER);
    }
  }
}
function runInvoices() {
    function billingItem(BITEM) { return [BITEM[1]].concat(BITEM.slice(5, BITEM[3] == "NIXON" ? 12 : 10)).concat(BITEM[13]); }
    const billing = new FORMAT("BILLING", 3, "CLIENTS", billingItem);
    billing.RUN();
}
function runPayroll() {
    const today = Utilities.formatDate(new Date(), "GMT-5", "MM/dd/yy");
    function payrollItem(PITEM) { return [PITEM[1]].concat([PITEM[3]], PITEM.slice(5, 10), PITEM.slice(12, 15))[0]; }
    const payroll = new FORMAT("PAYROLL", 12, "RIDERS", payrollItem);
    payroll.RUN();
}

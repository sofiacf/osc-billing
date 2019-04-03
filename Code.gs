// Compiled using ts2gas 1.6.2 (TypeScript 3.3.4000)
var exports = exports || {};
var module = module || { exports: exports };
function printPdf(s, f) {
    var n = s.getName(), tmp = s.makeCopy(n + "tmp_pdf_copy");
    var url = tmp.getUrl(), id = SpreadsheetApp.open(s).getSheetId();
    var x = 'export?exportFormat=pdf&format=pdf&fitw=true&portrait=false&gridlines=false&gid=' + id;
    url = url.replace('edit?usp=drivesdk', '');
    var tkn = ScriptApp.getOAuthToken();
    var r = UrlFetchApp.fetch(url + x, { headers: { 'Authorization': 'Bearer ' + tkn } });
    r.getBlob().setName(n);
    f.createFile(r);
    tmp.setTrashed(true);
}
var today = Utilities.formatDate(new Date(), "GMT-5", "MM/dd/yy");
var master = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
var period = master.getName(), data = master.getDataRange().getValues().slice(4);
var Format = /** @class */ (function () {
    function Format(n, sc, sn, sitem) {
        var _this = this;
        this.n = n;
        this.f = DriveApp.getFoldersByName(n).next();
        this.sn = sn;
        this.t = this.f.getFilesByName("TEMPLATE").next();
        this.sc = sc;
        this.ss = SpreadsheetApp.open(this.f.getFilesByName("DATA").next());
        this.sitem = sitem;
        this.sf = this.f.getFoldersByName(sn).next();
        this.si = this.ss.getDataRange().getValues();
        this.h = this.si[0];
        this.as = data.map(function (x) { return x[sc]; }).filter(function (x, i, a) { return (i == a.indexOf(x)); }).sort();
        this.ns = this.as.filter(function (x) { return _this.si.map(function (x) { return x[0]; }).indexOf(x) < 0; });
        this.summary = this.f.getFilesByName("SUMMARY TEMPLATE").next();
    }
    Format.prototype.subs = function () {
        var l = this.h.length;
        var template = this.t;
        function update(s) {
            var sf = this.sf, tn = s + " TEMPLATE", folders = sf.getFoldersByName(s);
            var f = (folders.hasNext()) ? folders.next() : sf.createFolder(s);
            var files = f.getFilesByName(tn);
            var t = files.hasNext() ? files.next() : template.makeCopy(tn, f);
            return [[s], [f.getId()], [t.getId()]];
        }
        var os = this.si, ns = this.ns, newsi;
        os.shift();
        var updateos = os.filter(function (x) { return !x[1]; });
        if (!ns.length && !updateos.length) return os;
        if (!ns.length) newsi = updateos.map(function (x) { return (update(x[0]).concat(x.slice(3))); });
        else {
            var updatens = ns.map(function (x) { return (update(x).concat(a)); });
            newsi = updatens.concat(os).sort();
        }
        this.ss.getSheets()[0].getRange(2, 1, newsi.length, l).setValues(newsi);
        SpreadsheetApp.flush();
        return newsi;
    };
    Format.prototype.items = function () {
        var _this = this;
        var as = this.as, items = as.map(function (x) { return []; });
        Logger.log("Items:", items, "as:", as);
        data.forEach(function (x) { items[as.indexOf(x[_this.sc])].push(_this.sitem(x)); });
        return items;
    };
    Format.prototype.run = function () {
        var _this = this;
        var n = _this.n + " " + today, o = _this.f.getFoldersByName(n), as = _this.as, ns = _this.ns;
        var subs = this.subs();
        var subids = subs.map(function (x) { return x[0]; });
        if (o.hasNext())
            o.next().setTrashed(true);
        var f = _this.f.createFolder(n);
        var items = _this.items();
        var rsubs = as.map(function (x, i) { return new Subject(subs[subids.indexOf(x)], items[i]); });
        rsubs.forEach(function (x) { x.run(f); });
        var s = SpreadsheetApp.open(this.summary.makeCopy("SUMMARY", f));
        rsubs.forEach(function (x) { x.info.push(ns.indexOf(x.id) > -1 ? ["new"] : ["ran"]); });
        var r = s.getSheets()[0].getRange(10, 1, rsubs.length, rsubs[0].info.length);
        r.setValues(rsubs.map(function (x) { return x.info; }));
        s.getSheets()[0].getRange(2, 2).setValue(rsubs.length);
    };
    return Format;
}());
var Subject = /** @class */ (function () {
    function Subject(a, items) {
        this.id = a[0];
        this.f = DriveApp.getFolderById(a[1]);
        this.t = DriveApp.getFileById(a[2]);
        this.items = items;
        this.num = (a[4]) ? (a[4] + 1) + a[5] : 0;
        this.info = (this.num) ? [[today], [this.num], [period]] : [[today], [period]];
        Logger.log(this.info);
        this.fn = (this.num) ? a[0] + "- # " + this.info[1] : a[0] + " - " + period;
    }
    Subject.prototype.run = function (folder) {
        var sheet = this.t.makeCopy(this.fn, this.f);
        var c = SpreadsheetApp.open(sheet).getSheets()[0];
        var l = this.items.length, w = this.items[0].length;
        c.insertRows(16, l);
        SpreadsheetApp.flush();
        c.getRange(16, 1, l, w).setValues(this.items).setFontSize(10).setWrap(true);
        c.getRange(4, w - 1, this.info.length).setValues(this.info);
        c.getRange(16, w, l).setNumberFormat('$0.00');
        SpreadsheetApp.flush();
        printPdf(sheet, folder);
    };
    return Subject;
}());
function runInvoices() {
    function item(i) { return [[i[1]].concat(i.slice(5, i[3] == "NIXON" ? 12 : 10))]; }
    var billing = new Format("BILLING", 3, "CLIENTS", item);
    billing.run();
}
function runPayroll() {
    function item(i) { return [i[1]].concat([i[3]], i.slice(5, 10), i.slice(12, 15)); }
    var payroll = new Format("PAYROLL", 12, "RIDERS", item);
    payroll.run();
}

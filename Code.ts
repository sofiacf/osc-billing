class Subject {
  id: string;
  state: string;
  total: number;
  file: GoogleAppsScript.Drive.File = null;
  items = [];
  props = {};
  constructor(sub: any[]) {
    this.id = sub[0];
    this.state = (sub[1] > 0 && sub[2] == 'OK') ? sub[3] || 'RUN' : 'SKIP';
    this.total = sub[1];
  }
}
class Run {
  f: string; date: string;
  subs: Object; clear: boolean;
  rf: GoogleAppsScript.Drive.Folder;
  constructor(f: string, date: Date, subs: Object, clear: boolean) {
    this.f = f;
    this.date = Utilities.formatDate(date, "GMT", "MM/dd/yy");
    this.subs = subs;
    let dir = DriveApp.getFoldersByName(f).next();
    let month = Utilities.formatDate(date, "GMT", "MMM").toUpperCase();
    let runFolderName = month + ' ' + f;
    let find = dir.getFoldersByName(runFolderName);
    if (find.hasNext() && clear) find.next().setTrashed(true);
    this.rf = find.hasNext() ? find.next() : dir.createFolder(runFolderName);
    this.clear = clear;
  }
  checkFiles = () => {
    if (this.clear) return;
    let files = this.rf.getFiles();
    while (files.hasNext()) {
      let file = files.next();
      let name = file.getName();
      if (!this.subs.hasOwnProperty(name)) continue;
      let sub = this.subs[name];
      if (sub.state == 'SKIP') continue;
      if (sub.state == 'RUN') file.setTrashed(true);
      else sub.file = file;
    }
  }
  run = (subs: Subject[]) => {
    let template = DriveApp.getFilesByName('TEMPLATE').next();
    subs.forEach(sub => {
      let tmp = sub.props['template'] == 'default' ? template
        : DriveApp.getFileById(sub.props['template']);
      let ss = tmp.makeCopy(sub.id, this.rf);
      let sheet = SpreadsheetApp.open(ss).getSheets()[0];
      let charges = sub.items.map((i: any[]) =>
        [i[0]].concat(i.slice(4, sub.id == 'NIXON' ? 11 : 9)).concat(i[12]));
      let rows = charges.length;
      let cols = charges[0].length;
      let info = [['test']];
      sheet.insertRows(16, rows - 1 || 1);
      let itemsRange = sheet.getRange(16, 1, rows, cols);
      itemsRange.setValues(charges).setFontSize(10).setWrap(true);
      sheet.getRange(4, cols - 1, info.length).setValues(info);
      sheet.getRange(16, cols, rows).setNumberFormat('$0.00');
      SpreadsheetApp.flush();
      sub.state = 'PRINT';
    })
  }
  print = (subs: Subject[]) => {
    subs.forEach(sub => {
      let tmp = sub.file.makeCopy(sub.id + "tmp_pdf_copy");
      let url = tmp.getUrl();
      let id = SpreadsheetApp.open(sub.file).getSheetId();
      let x = ('export?exportFormat=pdf&format=pdf'
        + '&fitw=true&portrait=false&gridlines=false&gid=' + id);
      url = url.replace('edit?usp=drivesdk', '') + x;
      let tkn = ScriptApp.getOAuthToken();
      let r = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + tkn }
      });
      r.getBlob().setName(sub.id);
      tmp.setTrashed(true);
      this.rf.createFile(r);
      DriveApp.getFolderById(sub.props['folder']).addFile(sub.file);
      this.rf.removeFile(sub.file);
      sub.state = 'POST';
    });
  }
  post = (subs: Subject[]) => {
    subs.forEach(sub => {
      let fn = sub.id + ' - # ' + ((sub.props['number'] || 0) + 1);
      sub.file.setName(fn + ' - ' + this.date);
      sub.state = 'DONE';
    });
  }
  doRun = () => {
    let readyToPost = [], readyToPrint = [], readyToRun = [];
    this.checkFiles();
    let subs = Object.keys(this.subs);
    subs.forEach(s => {
      let sub = this.subs[s];
      if (sub.state == 'SKIP') return;
      if (sub.state == 'RUN' || !sub.file || this.clear) readyToRun.push(sub);
      else if (sub.state == 'POST') readyToPost.push(sub);
      else if (sub.state == 'PRINT') readyToPrint.push(sub);
    });
    this.post(readyToPost);
    this.print(readyToPrint);
    this.run(readyToRun);
  }
  getStates = () => {
    this.doRun();
    return Object.keys(this.subs).map(s => [this.subs[s].state]);
  }
}
class WorkbookManager {
  doRun = () => {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let dash = ss.getSheetByName('DASH');
    let settings: any[][] = dash.getSheetValues(1, 2, 3, 1);
    let input = ss.getSheetByName('INPUT');
    if (settings[1][0] != 'NONE') {
      let m = DriveApp.getFilesByName('OSC MASTER INPUT').next();
      let d = SpreadsheetApp.open(m).getSheets()[1].getSheetValues(4, 2, -1, -1);
      input.clear().getRange(1, 1, d.length, d[0].length).setValues(d);
    }
    SpreadsheetApp.flush();
    let subs = dash.getSheetValues(2, 3, -1, 4);
    let items = input.getSheetValues(1, 1, -1, -1);
    let f = settings[0][0];
    let data = ss.getSheetByName(f).getSheetValues(1, 1, -1, -1);
    let clear = settings[1][0] == 'RUN';
    let run = new Run(f, settings[2][0], this.subs(f, subs, items, data), clear);
    dash.getRange(2, 6, subs.length).setValues(run.getStates());
  }
  subs = (f: string, actives: any[][], items: any[][], data: any[][]) => {
    const map = {};
    actives.forEach(a => map[a[0]] = new Subject(a));
    let sc = f == 'BILLING' ? 2 : 11;
    items.forEach(c => map[c[sc]].items.push(c))
    let info = data.slice(0);
    let ps: string[] = info.shift();
    info.forEach(k => {
      ps.forEach((p, i) => {
        if (map.hasOwnProperty(k[0])) map[k[0]].props[p] = k[i];
      });
    });
    return map;
  }
}

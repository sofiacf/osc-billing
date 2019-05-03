class Subject {
  items = []; props = {};
  state: string; total: number; name: string;
  constructor(sub: any[]) {
    this.name = sub[0];
    this.state = (sub[1] > 0 && sub[2] == 'OK') ? sub[3] || 'RUN' : 'SKIP';
    this.total = sub[1];
  }
}
class Run {
  f: string; date: string;
  subs: Object; clear: boolean;
  subNames: string[];
  rf: GoogleAppsScript.Drive.Folder;
  constructor(f: string, date: Date, subs: Object, clear: boolean) {
    this.f = f;
    this.date = Utilities.formatDate(date, "GMT", "MM/dd/yy");
    this.subs = subs;
    this.subNames = Object.keys(subs);
    let month = Utilities.formatDate(date, "GMT", "MMM").toUpperCase();
    let dir = DriveApp.getFoldersByName(f).next();
    let runFolderName = month + ' ' + f;
    let find = dir.getFoldersByName(runFolderName);
    if (find.hasNext() && clear) find.next().setTrashed(true);
    this.rf = find.hasNext() ? find.next() : dir.createFolder(runFolderName);
  }
  getSubsWithState = (state: string) => {
    return this.subNames.filter(s => this.subs[s].state == state);
  }
  post = () => {
    let subs = this.getSubsWithState('POST');
    let iterator = this.rf.getFilesByType('.pdf');
    while (iterator.hasNext()) {
      let file = iterator.next();
      let name = file.getName();
      if (subs.indexOf(name) > -1) {
        let sub = this.subs[name];
        let fn = name + ' - # ' + ((sub.props['statementNum'] || 0) + 1);
        file.setName(fn + ' - ' + this.date);
        sub.state = 'DONE';
      }
    }
  }
  print = () => {
    let subs = this.getSubsWithState('PRINT');
    let iterator = this.rf.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (iterator.hasNext()) {
      let sheet = iterator.next();
      let name = sheet.getName();
      if (subs.indexOf(name) < 0) {
        this.subs[name].state = 'RUN';
        continue;
      }
      let tmp = sheet.makeCopy(name + "tmp_pdf_copy");
      let url = tmp.getUrl();
      let id = SpreadsheetApp.open(sheet).getSheetId();
      let x = ('export?exportFormat=pdf&format=pdf'
        + '&fitw=true&portrait=false&gridlines=false&gid=' + id);
      url = url.replace('edit?usp=drivesdk', '') + x;
      let tkn = ScriptApp.getOAuthToken();
      let r = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + tkn }
      });
      r.getBlob().setName(name);
      tmp.setTrashed(true);
      this.rf.createFile(r);
      DriveApp.getFolderById(this.subs[name].props['folder']).addFile(sheet);
      this.rf.removeFile(sheet);
      this.subs[name].state = 'POST';
    }
  }
  run = () => {
    let subs = this.getSubsWithState('RUN');
    let files = this.rf.getFiles();
    while (files.hasNext()) {
      let file = files.next();
      if (subs.indexOf(file.getName()) > -1) file.setTrashed(true);
    }
    let template = DriveApp.getFilesByName('TEMPLATE').next();
    for (let s of subs) {
      let sub = this.subs[s];
      let tmp = sub.props['template'] == 'default' ?
        template : DriveApp.getFileById(sub.props['template']);
      let ss = tmp.makeCopy(s, this.rf);
      let sheet = SpreadsheetApp.open(ss).getSheets()[0];
      let items = sub.items;
      let rows = items.length;
      let cols = items[0].length;
      let info = [['test']];
      sheet.insertRows(16, rows - 1 || 1);
      sheet.getRange(16, 1, rows, cols).setValues(items).setFontSize(10).setWrap(true);
      sheet.getRange(4, cols-1, info.length).setValues(info);
      sheet.getRange(16, cols, rows).setNumberFormat('$0.00');
      SpreadsheetApp.flush();
      sub.state = 'PRINT';
    }
  }
  getStates = () => {
    this.post();
    this.print();
    this.run();
    return this.subNames.map(s => [this.subs[s].state]);
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
    let subs = dash.getSheetValues(2, 3, -1, 3);
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
    let sc = f =='BILLING' ? 2 : 11;
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

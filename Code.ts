class Subject {
  items = [];
  props = {};
  state: string; total: number;
  constructor(sub: any[]) {
    this.state = sub[3];
    this.total = sub[1];
  }
  runReport(runFolder: GoogleAppsScript.Drive.Folder, date: string) {
    let statementNum = (this.props['statementNum'] || 0) + 1;
    let file = `${this.props['id']} - ${'# ' + statementNum} - ${date}`;
    if (!runFolder.getFilesByName(file).hasNext()) {
      DriveApp.getFileById(this.props['template']).makeCopy(file, runFolder);
      this.state = 'PRINT';
    }
  }
}
class Runner {
  static doRun = (f: string, day: string, subs: {}, clear: boolean) => {
    let dir = DriveApp.getFoldersByName(f).next();
    let search = dir.getFoldersByName(f + ' ' + day);
    if (search.hasNext()) search.next().setTrashed(clear);
    let rf = (search.hasNext()) ? search.next() : dir.createFolder(f + ' ' + day);
    for (let sub in subs) subs[sub].runReport(rf);
  }
}
class WorkbookManager {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  dash = this.ss.getSheetByName('DASH');
  settings: any[][] = this.dash.getSheetValues(1, 2, 3, 1);
  input = this.ss.getSheetByName('INPUT');
  refresh = () => {
    let f = this.settings[0][0];
    if (this.settings[1][0] != 'NONE') {
      let m = DriveApp.getFilesByName('OSC MASTER INPUT').next();
      let d = SpreadsheetApp.open(m).getSheets()[1].getSheetValues(4, 2, -1, -1);
      let r = this.input.clear().getRange(1, 1, d.length, d[0].length);
      r.setValues(d).sort(f == 'BILLING' ? 3 : 12);
    }
    let date = Utilities.formatDate(this.settings[2][0], "GMT-5", "M/d/yy");
    Runner.doRun(f, date, this.getSubjects(f), this.settings[1][0]);
  }
  getSubjects = (f: string) => {
    const map = {};
    this.dash.getSheetValues(2, 4, -1, 3).forEach((s: any[]) => {
      if (s[1] > 0 && s[2] == 'OK' && s[3] != 'SKIP') map[s[0]] = new Subject(s);
    });
    let charges: any[][] = this.input.getSheetValues(1, 1, -1, -1);
    for (let c of charges) { map[c[f == 'BILLING' ? 2 : 11]].items.push(c); }
    let knowns = this.ss.getSheetByName(f).getSheetValues(1, 1, -1, -1);
    let props = knowns.shift();
    knowns = knowns.filter((k: any[]) => map.hasOwnProperty(k[0]));
    knowns.forEach((k: any[]) => {
      if (map.hasOwnProperty(k[0])) props.forEach((p, i) => map[k[0]].props[p] = k[i]);
    });
    return map;
  }
}

// class Runner {
//   static getRunFolder (f: string, date: string) {
//     var dir = DriveApp.getFoldersByName(f).next();
//     var rfName = f + " " + date;
//     if (!dir.getFoldersByName(rfName).hasNext()) dir.createFolder(rfName);
//     return dir.getFoldersByName(rfName).next();
//   }
//   static print(subs) {
//     subs.forEach(sub => {
//       var file = this.runFolder.getFilesByName(sub.file).next();
//       var tmp = file.makeCopy(sub.file + "tmp_pdf_copy");
//       let url = tmp.getUrl(), id = SpreadsheetApp.open(file).getSheetId();
//       let x = 'export?exportFormat=pdf&format=pdf&fitw=true&portrait=false&gridlines=false&gid=' + id;
//       url = url.replace('edit?usp=drivesdk', '') + x;
//       let tkn = ScriptApp.getOAuthToken();
//       let r = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + tkn } });
//       r.getBlob().setName(sub.file);
//       tmp.setTrashed(true);
//       DriveApp.getFolderById(sub.folder).createFile(r);
//     });
//   }
// }

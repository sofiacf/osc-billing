class Subject {
  id: string;
  state: string;
  file: GoogleAppsScript.Drive.File = null;
  items = [];
  props = {};
  constructor(sub: any[]) {
    this.id = sub[0];
    this.state = (sub[1] > 0 && sub[2] == 'OK') ? sub[3] || 'RUN' : 'SKIP';
  }
}
class Run {
  f: string;
  date: Date;
  period: string;
  subs: {};
  rf: GoogleAppsScript.Drive.Folder;
  constructor(f: string, date: Date, period: string) {
    this.f = f;
    this.period = period;
    this.date = date;
  }
  setupFolder = () => {
    let dir = DriveApp.getFoldersByName(this.f).next();
    let folderName = this.period + ' ' + this.f;
    let find = dir.getFoldersByName(folderName);
    return find.hasNext() ? find.next() : dir.createFolder(folderName);
  }
  setupFiles = (action: string) => {
    if (action == 'RESET') return this.rf.setTrashed(true);
    let files = this.rf.getFiles();
    while (files.hasNext()) {
      let file = files.next();
      let name = file.getName();
      if (!this.subs.hasOwnProperty(name)) continue;
      let sub = this.subs[name];
      if (sub.state == 'RUN') file.setTrashed(true);
      else if (action == 'POST' && sub.state != 'POST') file.setTrashed(true);
      else sub.file = file;
    }
  }
  doRun = (subs: {}, action: string) => {
    this.subs = subs;
    this.rf = this.setupFolder();
    this.setupFiles(action);
    if (action == 'RESET') return;
    let readyToPost = [], readyToPrint = [], readyToRun = [];
    Object.keys(subs).forEach(s => {
      let sub: Subject = subs[s];
      if (sub.state == 'RUN') readyToRun.push(sub);
      if (sub.state == 'PRINT') readyToPrint.push(sub);
      if (sub.state == 'POST') readyToPost.push(sub);
    });
    if (action == 'POST') return this.post(readyToPost);
    this.run(readyToRun);
    this.print(readyToPrint);
  }
  run = (subs: Subject[]) => {
    let template = DriveApp.getFilesByName('TEMPLATE').next();
    subs.forEach(sub => {
      let tmp = sub.props['template'] == 'default' ? template
        : DriveApp.getFileById(sub.props['template']);
      let ss = tmp.makeCopy(sub.id, this.rf);
      let sheet = SpreadsheetApp.open(ss).getSheets()[0];
      let props = sub.props;
      sheet.getNamedRanges().forEach(r => {
        let name = r.getName();
        if (props.hasOwnProperty(name)) r.getRange().setValue(props[name]);
        if (this.hasOwnProperty(name)) r.getRange().setValue(this[name]);
      });
      let charges = sub.items.map((i: any[]) => {
        let ar = i.slice(0, sub.id == 'NIXON' ? 11 : 9).concat(i[12]);
        ar.splice(1, 3);
        return ar;
      });
      let rows = charges.length;
      let cols = charges[0].length;
      sheet.insertRows(16, rows);
      let itemsRange = sheet.getRange(16, 1, rows, cols);
      itemsRange.setValues(charges).setFontSize(10).setWrap(true);
      sheet.getRange(16, cols, rows).setNumberFormat('$0.00');
      SpreadsheetApp.flush();
      sub.state = 'PRINT';
    })
  }
  print = (subs: Subject[]) => {
    subs.forEach(sub => {
      let url = sub.file.getUrl().replace('edit?usp=drivesdk', '');
      let options = {
        headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
      }
      let x = 'export?exportFormat=pdf&format=pdf&size=letter'
        + '&portrait=false'
        + '&fitw=true&gridlines=false&gid=0';
      let r = UrlFetchApp.fetch(url + x, options);
      let blob = r.getBlob().setName(sub.id);
      this.rf.createFile(blob);
      // DriveApp.getFolderById(sub.props['folder']).addFile(sub.file);
      // this.rf.removeFile(sub.file);
      sub.state = 'POST';
    });
  }
  post = (subs: Subject[]) => {
    subs.forEach(sub => {
      let fn = sub.id + ' - # ' + sub.props['number'] + 1;
      sub.file.setName(fn + ' - ' + this.date);
      sub.state = 'DONE';
    });
  }
}
class WorkbookManager {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  dash = this.ss.getSheetByName('DASH');
  settings: any[][] = this.dash.getSheetValues(1, 2, 4, 1);
  f = this.settings[0][0];
  date = this.settings[3][0];
  period = this.settings[2][0];
  action = this.settings[1][0];
  itemSheet = this.ss.getSheetByName(this.period);
  dataSheet = this.ss.getSheetByName(this.f == 'BILLING' ? 'CLIENTS' : 'COURIERS');
  doRun = () => {
    let actives = this.dash.getSheetValues(2, 3, -1, 4);
    let items = this.itemSheet.getSheetValues(2, 2, -1, -1);
    let data = this.dataSheet.getSheetValues(1, 1, -1, -1);
    let run = new Run(this.f, this.date, this.period);
    let subs = this.subs(actives, items, data);
    run.doRun(subs, this.action);
    this.updateStates(subs);
    if (this.action != 'POST') return;
    data.forEach((d: any[]) => {
      if (!subs.hasOwnProperty(d[0])) return;
      let sub = subs[d[0]];
      if (sub.state != 'DONE') return;
      d[1] = this.date;
      d[2] = sub.number + 1;
    });
    this.dataSheet.getDataRange().setValues(data);
  }
  subs = (actives: any[][], items: any[][], data: any[][]) => {
    const map = {};
    actives.forEach(a => map[a[0]] = new Subject(a));
    let sc = this.f == 'BILLING' ? 2 : 11;
    items.forEach(c => (map[c[sc]] || { items: [] }).items.push(c));
    let info = data.slice(0);
    let ps: string[] = info.shift();
    info.forEach(k => {
      ps.forEach((p, i) => {
        if (map.hasOwnProperty(k[0])) map[k[0]].props[p] = k[i];
      });
    });
    return map;
  }
  updateStates = (subs: {}) => {
    let states = Object.keys(subs).map(s => [subs[s].state]);
    states.forEach(s => {
      let state = s[0];
      if (this.action == 'POST' && state != 'DONE') s = ['SKIP'];
      if (this.action == 'RESET' && state != 'SKIP') s = ['RUN'];
    });
    this.dash.getRange(2, 6, states.length).setValues(states);
  }
}
class ConfigurationManager {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  readInputFields = () => {
    let input = this.ss.getSheets()[0];
    let fieldsRange = input.getRange(1, 1, 1, input.getLastColumn());
    return fieldsRange.getValues()[0];
  }
  saveProperty = () => {
    let fields = this.readInputFields();
    let userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('fields', JSON.stringify(fields));
  }
  readProperties = () => {
    let userProperties = PropertiesService.getUserProperties();
    let result = userProperties.getProperty('fields');
    userProperties.deleteAllProperties();
    return JSON.parse(result);
  }
}
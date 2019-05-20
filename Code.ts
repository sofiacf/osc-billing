class Subject {
  id: string;
  state: string;
  file: GoogleAppsScript.Drive.File;
  constructor(sub: any[]) {
    this.id = sub[0];
    this.state = (sub[1] > 0 && sub[2] == 'OK') ? sub[3] || 'RUN' : 'SKIP';
  }
}
interface Format {
  id: string;
  subject: string;
  subjectColumn: number;
}
interface Data {
  subjects: Subject[],
  items: {},
  subjectData: {}
}
class SheetManager {
  static ss = SpreadsheetApp.getActiveSpreadsheet();
  static dash = SheetManager.ss.getSheetByName('DASH');
  static settings = SheetManager.dash.getRange(1, 2, 4).getValues();
  static readData = (period: string, subject: string) => {
    let ss = SheetManager.ss;
    return ({
      actives: SheetManager.dash.getSheetValues(2, 3, -1, 4),
      items: ss.getSheetByName(period).getDataRange().getValues(),
      subjectData: ss.getSheetByName(subject).getDataRange().getValues()
    });
  }
  static writeStates = (states: string[]) => {
    SheetManager.dash.getRange(2, 6, states.length).setValues(states.map(s => [s]));
  }
  static updateInvoiceNumbers = () => {
    return;
    {// data.forEach((d: any[]) => {
      //   if (!subs.hasOwnProperty(d[0])) return;
      //   let sub = subs[d[0]];
      //   if (sub.state != 'DONE') return;
      //   d[1] = setup.date;
      //   d[2] = sub.number + 1;
      // });
      // try {
      //   this.dataSheet.getDataRange().setValues(data);
      //   this.ss.toast('Run complete!');
      // }
      // catch(e) {
      //   this.ss.toast('Didn\'t write subject data, probably for the best haha.')
      // }
    }
  }
}
class DataManager {
  static actions = { RESET: 'RESET', POST: 'POST', RUN: 'RUN' };
  static getSettings = () => {
    let formats = {
      BILLING: { id: 'BILLING', subject: 'CLIENTS', subjectColumn: 3 },
      PAYROLL: { id: 'PAYROLL', subject: 'COURIERS', subjectColumn: 12 }
    }
    let settings: any[] = SheetManager.settings.map(x => x[0]);
    return ({
      action: settings[1],
      date: settings[3],
      format: formats[settings[0]],
      period: settings[2]
    });
  }
  static getData = (period: string, format: Format) => {
    let data = SheetManager.readData(period, format.subject);
    let subjects = data.actives.map(sub => new Subject(sub));
    let items: any[][] = data.items.slice(1);
    let subjectData: any[][] = data.subjectData.slice(0);
    let props = subjectData.shift();
    return {
      subjects: subjects,
      items: items.reduce((acc, x) => {
        let subject = x[format.subjectColumn];
        if (!acc[subject]) acc[subject] = [];
        acc[subject].push(x.slice(1));
        return acc;
      }, {}),
      subjectData: subjectData.reduce((acc, x) => {
        let datum = { props: {} };
        props.forEach((prop: string, i: number) => datum.props[prop] = x[i]);
        acc[x[0]] = datum;
        return acc;
      }, {})
    }
  }
  static getFolderName = (period: string, format: Format) => {
    return period + ' ' + format.id;
  }
  userProperties = PropertiesService.getUserProperties();
  setProperty = (property: string, value: any) => {
    let properties = {
      FIELDS: 'fields',
      CLIENTS: 'clients',
      FORMATS: 'formats'
    }
    this.userProperties.setProperty(property, JSON.stringify(value));
    return this.userProperties.getProperty(property);
  }
  checkPropertyValues = (property: string, value = 'RESET') => {
    let userProperties = PropertiesService.getUserProperties();
    return this.userProperties.getProperty(property) || this.setProperty(property, value);
  }
}
class FileManager {
  static getFolder = (name: string, format: string) => {
    let directory = DriveApp.getFoldersByName(format).next();
    let find = directory.getFoldersByName(name);
    return find.hasNext() ? find.next() : directory.createFolder(name);
  }
  static runStatements = (action: string, folder: GoogleAppsScript.Drive.Folder, data: Data, date: Date) => {
    if (action == DataManager.actions.RESET) {
      folder.setTrashed(true);
      return;
    }
    FileManager.run(folder, data);
    FileManager.print(folder, data);
    FileManager.post(folder, data, Utilities.formatDate(date, 'GMT', 'MM/dd/yy'));
  }
  static run = (folder: GoogleAppsScript.Drive.Folder, data: Data) => {
    let template = DriveApp.getFilesByName('TEMPLATE').next();
    let subjects: Subject[] = data.subjects.filter((sub: Subject) => sub.state == 'RUN');
    let subjectData = data.subjectData;
    let items = data.items;
    subjects.forEach(sub => {
      let props = subjectData[sub.id].props;
      if (props['template'] != 'default') {
        try {
          template = DriveApp.getFileById(props['template']);
        }
        catch (e) {
          Logger.log('No template found for', sub);
          data.subjectData[sub.id].props['template'] = 'default';
          template = DriveApp.getFilesByName('TEMPLATE').next();
        }
      }
      let ss = template.makeCopy(sub.id, folder);
      let sheet = SpreadsheetApp.open(ss).getSheets()[0];
      sheet.getNamedRanges().forEach(r => {
        let name = r.getName();
        if (props.hasOwnProperty(name)) r.getRange().setValue(props[name]);
      });
      let charges = items[sub.id].map((i: any[]) => {
        try {
          let ar = i.slice(0, sub.id == 'NIXON' ? 11 : 9).concat(i[12]);
          ar.splice(1, 3);
          return ar;
        }
        catch (e) {
          return;
        }
      });
      let rows = charges.length;
      let cols = charges[0].length;
      sheet.insertRows(16, rows);
      let itemsRange = sheet.getRange(16, 1, rows, cols);
      itemsRange.setValues(charges).setFontSize(10).setWrap(true);
      sheet.getRange(16, cols, rows).setNumberFormat('$0.00');
      SpreadsheetApp.flush();
      sub.state = 'PRINT';
    });
  }
  static print = (folder: GoogleAppsScript.Drive.Folder, data: Data) => {
    let subjects = data.subjects.filter((sub: Subject) => sub.state == 'PRINT');
    subjects.forEach(sub => {
      let files = folder.getFilesByName(sub.id);
      if (!files.hasNext()) {
        sub.state = 'RUN';
        return;
      }
      let file = files.next();
      let url = file.getUrl().replace('edit?usp=drivesdk', '');
      let options = {
        headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
      }
      let x = 'export?exportFormat=pdf&format=pdf&size=letter'
        + '&portrait=false'
        + '&fitw=true&gridlines=false&gid=0';
      let r = UrlFetchApp.fetch(url + x, options);
      let blob = r.getBlob().setName(sub.id);
      folder.createFile(blob);
      // DriveApp.getFolderById(sub.props['folder']).addFile(sub.file);
      // folder.removeFile(sub.file);
      sub.state = 'POST';
    });
  }
  static post = (folder: GoogleAppsScript.Drive.Folder, data: Data, date: string) => {
    let subjects = data.subjects.filter((sub: Subject) => sub.state == 'POST');
    subjects.forEach(sub => {
      let files = folder.getFilesByName(sub.id);
      if (!files.hasNext()) return;
      let file = files.next();
      let props = data.subjectData[sub.id].props;
      let fn = sub.id + ' - # ' + (props['number'] + 1);
      file.setName(fn + ' - ' + date + '.pdf');
      sub.state = 'DONE';
    });
  }
}

// class f {
//     verifyData = (items: {}, run) => {
//         let subjects = Object.keys(items);
//         Logger.log(subjects);
//         // let missingSubjects = Object.keys(items).filter(s => (!subjects[s]));
//         // for (let i = 0; i < missingSubjects.length; i++) {
//         //   let subject = missingSubjects[i];
//         //   let response = ui.prompt('Provide info for ' + subject);
//         //   if (response.getSelectedButton() == ui.Button.OK) {
//         //     Logger.log(response.getResponseText(), subject);
//         //   }
//         // }
//     }
//     static run = (folder: GoogleAppsScript.Drive.Folder, data: Data) => {
//         let template = DriveApp.getFilesByName('TEMPLATE').next();
//         let subjects: Subject[] = data.subjects.filter((sub: Subject) => sub.state == 'RUN');
//         let subjectData = data.subjectData;
//         let items = data.items;
//         subjects.forEach(sub => {
//             let props = subjectData[sub.id].props;
//             if (props['template'] != 'default') {
//                 try {
//                     template = DriveApp.getFileById(props['template']);
//                 }
//                 catch (e) {
//                     Logger.log('No template found for', sub);
//                     data.subjectData[sub.id].props['template'] = 'default';
//                     template = DriveApp.getFilesByName('TEMPLATE').next();
//                 }
//             }
//             let ss = template.makeCopy(sub.id, folder);
//             let sheet = SpreadsheetApp.open(ss).getSheets()[0];
//             sheet.getNamedRanges().forEach(r => {
//                 let name = r.getName();
//                 if (props.hasOwnProperty(name)) r.getRange().setValue(props[name]);
//             });
//             let charges = items[sub.id].map((i: any[]) => {
//                 try {
//                     let ar = i.slice(0, sub.id == 'NIXON' ? 11 : 9).concat(i[12]);
//                     ar.splice(1, 3);
//                     return ar;
//                 }
//                 catch (e) {
//                     return;
//                 }
//             });
//             let rows = charges.length;
//             let cols = charges[0].length;
//             sheet.insertRows(16, rows);
//             let itemsRange = sheet.getRange(16, 1, rows, cols);
//             itemsRange.setValues(charges).setFontSize(10).setWrap(true);
//             sheet.getRange(16, cols, rows).setNumberFormat('$0.00');
//             SpreadsheetApp.flush();
//             sub.state = 'PRINT';
//         });
//     }
//     static print = (folder: GoogleAppsScript.Drive.Folder, data: Data) => {
//         let subjects = data.subjects.filter((sub: Subject) => sub.state == 'PRINT');
//         subjects.forEach(sub => {
//             let files = folder.getFilesByName(sub.id);
//             if (!files.hasNext()) {
//                 sub.state = 'RUN';
//                 return;
//             }
//             let file = files.next();
//             let url = file.getUrl().replace('edit?usp=drivesdk', '');
//             let options = {
//                 headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
//             }
//             let x = 'export?exportFormat=pdf&format=pdf&size=letter'
//                 + '&portrait=false'
//                 + '&fitw=true&gridlines=false&gid=0';
//             let r = UrlFetchApp.fetch(url + x, options);
//             let blob = r.getBlob().setName(sub.id);
//             folder.createFile(blob);
//             // DriveApp.getFolderById(sub.props['folder']).addFile(sub.file);
//             // folder.removeFile(sub.file);
//             sub.state = 'POST';
//         });
//     }
//     static post = (folder: GoogleAppsScript.Drive.Folder, data: Data, date: string) => {
//         let subjects = data.subjects.filter((sub: Subject) => sub.state == 'POST');
//         subjects.forEach(sub => {
//             let files = folder.getFilesByName(sub.id);
//             if (!files.hasNext()) return;
//             let file = files.next();
//             let props = data.subjectData[sub.id].props;
//             let fn = sub.id + ' - # ' + (props['number'] + 1);
//             file.setName(fn + ' - ' + date + '.pdf');
//             sub.state = 'DONE';
//         });
//     }
// }
// class FileManager {
//     folder: GoogleAppsScript.Drive.Folder;
//     template: GoogleAppsScript.Drive.File;
//     constructor(settings) {
//         this.settings = settings;
//         this.folder = this.getRunFolder();
//         this.template = DriveApp.getFilesByName('TEMPLATE').next();
//     }
//     runStatements = (data: { items: {}, subjects: {} }) => {
//         let date = Utilities.formatDate(this.settings.date, 'GMT', 'MM/dd/yy');
//         this.run(data);
//         this.print(data);
//         this.post(data, date);
//     }
//     run = (data: { subjects, items }) => {
//         let subjects = data.subjects.filter((sub: Subject) => sub.state == 'RUN');
//         let items = data.items;
//         subjects.forEach(subject => {
//             let ss = this.copyStatement(subject);
//             let sheet = SpreadsheetApp.open(ss).getSheets()[0];
//             sheet.getNamedRanges().forEach(r => {
//                 let name = r.getName();
//                 if (props.hasOwnProperty(name)) r.getRange().setValue(props[name]);
//             });
//             let charges = items[subject.id].map((i: any[]) => {
//                 try {
//                     let ar = i.slice(0, subject.id == 'NIXON' ? 11 : 9).concat(i[12]);
//                     ar.splice(1, 3);
//                     return ar;
//                 }
//                 catch (e) {
//                     return;
//                 }
//             });
//             let rows = charges.length;
//             let cols = charges[0].length;
//             sheet.insertRows(16, rows);
//             let itemsRange = sheet.getRange(16, 1, rows, cols);
//             itemsRange.setValues(charges).setFontSize(10).setWrap(true);
//             sheet.getRange(16, cols, rows).setNumberFormat('$0.00');
//             SpreadsheetApp.flush();
//             subject.state = 'PRINT';
//         });
//     }
//     copyStatement = (subject: string) => {
//         let template: GoogleAppsScript.Drive.File
//         if (subject.template == 'default') {
//             return template.makeCopy(subject.id, this.folder);
//         }
//         try {
//             template = DriveApp.getFileById(subject.template);
//         }
//         catch (e) {
//             Logger.log('No template found for', subject);
//         }
//         finally {
//             template = template || this.template;
//             return template.makeCopy(subject.id, this.folder);
//         }
//     }
//     print = (data) => {
//         let subjects = data.subjects.filter((sub: Subject) => sub.state == 'PRINT');
//         subjects.forEach(sub => {
//             let files = this.folder.getFilesByName(sub.id);
//             if (!files.hasNext()) return;
//             let file = files.next();
//             let url = file.getUrl().replace('edit?usp=drivesdk', '');
//             let options = {
//                 headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
//             }
//             let x = 'export?exportFormat=pdf&format=pdf&size=letter'
//                 + '&portrait=false'
//                 + '&fitw=true&gridlines=false&gid=0';
//             let r = UrlFetchApp.fetch(url + x, options);
//             let blob = r.getBlob().setName(sub.id);
//             folder.createFile(blob);
//         });
//     }
//     post = (data, date: string) => {
//         let subjects = data.subjects.filter((sub: Subject) => sub.state == 'POST');
//         let posted: Subject[];
//         subjects.forEach(sub => {
//             let files = this.folder.getFilesByName(sub.id);
//             if (!files.hasNext()) return;
//             let file = files.next();
//             let props = data.subjectData[sub.id].props;
//             let fn = sub.id + ' - # ' + (props['number'] + 1);
//             file.setName(fn + ' - ' + date + '.pdf');
//             posted.push(sub);
//         });
//         return posted;
//     }
// }
// class StatementCreator {
//     copyStatement = () => {

//     }
// }
// // Compiled using ts2gas 1.6.2 (TypeScript 3.3.4000)
// var exports = exports || {};
// var module = module || { exports: exports };
// var Run = /** @class */ (function () {
//     function Run(settings) {
//         var _this = this;
//         this.create = function (settings) {
//             var template = { date: settings.date, subjects: [{}], folder: '' };
//             return { BILLING: template, PAYROLL: template };
//         };
//         this.run = function (run, settings) {
//             //THIS IS TO COLLECT ITEMS, THEN RUN AGAINST STORED SUBJECT VALUES TO COMPARE, AND IF ANY ARE MISSING FROM OBJECT, PROMPT USER TO INCLUDE
//             var items = _this.getItems(settings.period, settings.subject);
//             _this.verifyData(items, run);
//         };
//         this.getItems = function (period, subject) {
//             var data = _this.getItemSheet(period).getDataRange().getValues();
//             var col = data.shift().indexOf(subject);
//             return (data.reduce(function (obj, x) {
//                 if (!obj[x[col]])
//                     obj[x[col]] = [];
//                 obj[x[col]].push(x);
//                 return obj;
//             }, {}));
//         };
//         this.getItemSheet = function (period) {
//             var sheet = ss.getSheetByName(period);
//             if (sheet)
//                 return sheet;
//             var message = 'Period not in this workbook. Locate and import?';
//             var result = ui.alert(message, ui.ButtonSet.YES_NO);
//             if (result == ui.Button.YES)
//                 ss.toast('Not yet configured.');
//         };
//         this.verifyData = function (items, run) {
//             var subjects = Object.keys(items);
//             Logger.log(subjects);
//             // let missingSubjects = Object.keys(items).filter(s => (!subjects[s]));
//             // for (let i = 0; i < missingSubjects.length; i++) {
//             //   let subject = missingSubjects[i];
//             //   let response = ui.prompt('Provide info for ' + subject);
//             //   if (response.getSelectedButton() == ui.Button.OK) {
//             //     Logger.log(response.getResponseText(), subject);
//             //   }
//             // }
//         };
//         this.overwrite = function () {
//             var result = ui.alert('Continue existing run?', ui.ButtonSet.YES_NO);
//             return result == ui.Button.NO;
//         };
//         this.updateStates = function () {
//             ss.toast('All done!');
//         };
//         var existing = JSON.parse(properties.getProperty(settings.period));
//         var run = (existing || this.create(settings))[settings.format];
//         this.run(run, settings);
//     }
//     return Run;
// }());
// // Compiled using ts2gas 1.6.2 (TypeScript 3.3.4000)
// var exports = exports || {};
// var module = module || { exports: exports };
// var properties = PropertiesService.getDocumentProperties();
// var ss = SpreadsheetApp.getActiveSpreadsheet();
// var ui = SpreadsheetApp.getUi();
// var Configure = /** @class */ (function () {
//     function Configure() {
//         var _this = this;
//         this.properties = PropertiesService.getDocumentProperties();
//         this.updateSubjects = function (sheet) {
//             var data = _this.wkbk.readSheet(sheet).slice(1);
//             for (var i = 0; i < data.length; i++) {
//                 var subject = {
//                     id: data[i][0],
//                     name: data[i][5],
//                     street: data[i][6],
//                     city: data[i][7],
//                     attn: data[i][8],
//                     suffix: data[i][4]
//                 };
//                 _this.properties.setProperty(subject.id, JSON.stringify(subject));
//             }
//             Logger.log(_this.properties.getProperties());
//         };
//         this.setFormats = function () {
//             var billing = { id: 'BILLING', subject: 'CLIENT' };
//             var payroll = { id: 'PAYROLL', subject: 'COURIER' };
//             _this.properties.setProperty('BILLING', JSON.stringify(billing));
//             _this.properties.setProperty('PAYROLL', JSON.stringify(payroll));
//         };
//     }
//     return Configure;
// }());
// var FileManager = /** @class */ (function () {
//     function FileManager(settings) {
//         var _this = this;
//         this.getDirectory = function () {
//             return DriveApp.getFoldersByName(_this.settings.format).next();
//         };
//         this.getRunFolder = function () {
//             var directory = _this.getDirectory();
//             var name = _this.settings.folderName;
//             var find = directory.getFoldersByName(name);
//             return find.hasNext() ? find.next() : directory.createFolder(name);
//         };
//         this.trashRunFolder = function () {
//             _this.folder.setTrashed(true);
//         };
//         this.runStatements = function (data) {
//             var date = Utilities.formatDate(_this.settings.date, 'GMT', 'MM/dd/yy');
//             _this.run(data);
//             _this.print(data);
//             _this.post(data, date);
//         };
//         this.run = function (data) {
//             var subjects = data.subjects.filter(function (sub) { return sub.state == 'RUN'; });
//             var items = data.items;
//             subjects.forEach(function (subject) {
//                 var ss = _this.copyStatement(subject);
//                 var sheet = SpreadsheetApp.open(ss).getSheets()[0];
//                 sheet.getNamedRanges().forEach(function (r) {
//                     var name = r.getName();
//                     if (props.hasOwnProperty(name))
//                         r.getRange().setValue(props[name]);
//                 });
//                 var charges = items[subject.id].map(function (i) {
//                     try {
//                         var ar = i.slice(0, subject.id == 'NIXON' ? 11 : 9).concat(i[12]);
//                         ar.splice(1, 3);
//                         return ar;
//                     }
//                     catch (e) {
//                         return;
//                     }
//                 });
//                 var rows = charges.length;
//                 var cols = charges[0].length;
//                 sheet.insertRows(16, rows);
//                 var itemsRange = sheet.getRange(16, 1, rows, cols);
//                 itemsRange.setValues(charges).setFontSize(10).setWrap(true);
//                 sheet.getRange(16, cols, rows).setNumberFormat('$0.00');
//                 SpreadsheetApp.flush();
//                 subject.state = 'PRINT';
//             });
//         };
//         this.copyStatement = function (subject) {
//             var template;
//             if (subject.template == 'default') {
//                 return template.makeCopy(subject.id, _this.folder);
//             }
//             try {
//                 template = DriveApp.getFileById(subject.template);
//             }
//             catch (e) {
//                 Logger.log('No template found for', subject);
//             }
//             finally {
//                 template = template || _this.template;
//                 return template.makeCopy(subject.id, _this.folder);
//             }
//         };
//         this.print = function (data) {
//             var subjects = data.subjects.filter(function (sub) { return sub.state == 'PRINT'; });
//             subjects.forEach(function (sub) {
//                 var files = _this.folder.getFilesByName(sub.id);
//                 if (!files.hasNext())
//                     return;
//                 var file = files.next();
//                 var url = file.getUrl().replace('edit?usp=drivesdk', '');
//                 var options = {
//                     headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
//                 };
//                 var x = 'export?exportFormat=pdf&format=pdf&size=letter'
//                     + '&portrait=false'
//                     + '&fitw=true&gridlines=false&gid=0';
//                 var r = UrlFetchApp.fetch(url + x, options);
//                 var blob = r.getBlob().setName(sub.id);
//                 folder.createFile(blob);
//             });
//         };
//         this.post = function (data, date) {
//             var subjects = data.subjects.filter(function (sub) { return sub.state == 'POST'; });
//             var posted;
//             subjects.forEach(function (sub) {
//                 var files = _this.folder.getFilesByName(sub.id);
//                 if (!files.hasNext())
//                     return;
//                 var file = files.next();
//                 var props = data.subjectData[sub.id].props;
//                 var fn = sub.id + ' - # ' + (props['number'] + 1);
//                 file.setName(fn + ' - ' + date + '.pdf');
//                 posted.push(sub);
//             });
//             return posted;
//         };
//         this.settings = settings;
//         this.folder = this.getRunFolder();
//         this.template = DriveApp.getFilesByName('TEMPLATE').next();
//     }
//     return FileManager;
// }());
// // Compiled using ts2gas 1.6.2 (TypeScript 3.3.4000)
// var exports = exports || {};
// var module = module || { exports: exports };
// function run() {
//     var settings = new Settings(ss.getRangeByName('SETTINGS').getValues());
//     var run = new Run(settings);
// }
// // Compiled using ts2gas 1.6.2 (TypeScript 3.3.4000)
// var exports = exports || {};
// var module = module || { exports: exports };
// var Settings = /** @class */ (function () {
//     function Settings(settings) {
//         this.period = settings[2][0];
//         this.type = settings[1][0];
//         this.date = settings[3][0];
//         this.format = settings[0][0];
//         this.subject = this.format == 'BILLING' ? 'CLIENT' : 'COURIER';
//     }
//     return Settings;
// }());

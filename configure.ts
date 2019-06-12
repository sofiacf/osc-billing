function configure() {
    initConfig();
    return;
    let configure = {
        folder: (month: string, format: Format): GoogleAppsScript.Drive.Folder => {
            try {
                let folder = utils.get.folder.byId(format.runs[month].folder);
                return folder;
            }
            catch {
                let dir = utils.get.folder.byId(format.folder);
                let name = month + ' ' + format.id;
                let f = dir.getFoldersByName(name).hasNext() ?
                    dir.getFoldersByName(name).next() : dir.createFolder(name);
                return f;
            }
            ;
        },
        data: (a: any[][], subject: string): {} => {
            let col = a.shift().indexOf(subject);
            return (a.reduce((o, x) => {
                o[x[col]] = (o[x[col]] || []).concat(x);
                return o;
            }, {}));
        },
        runs: (format: Format): Object => {
            let sheets = utils.get.sheets();
            for (let i = sheets.length - 1; i > 3; i--) {
                let sheet = sheets[i].getName();
                let folder = configure.folder(sheet, format).getId();
                let data = configure.data(sheets[i].getSheetValues(1, 1, -1, -1), format.subject);
                format.runs[sheet] = { folder, data, subjects: Object.keys(data) };
            }
            return format.runs;
        },
        formats: (id: string): Format => {
            let folder = utils.get.folder.byName(id).getId();
            let subject = id == 'BILLING' ? 'CLIENT' : 'COURIER';
            let subjects = ((ar: any[][]): Subject[] => ar.map(c => id == 'BILLING' ?
                { name: c[5], street: c[6], city: c[7], attn: c[8], sfx: c[4], dates: {} }
                : { name: c[0], dates: {} }))(utils.read.sheet(subject + 'S').slice(1));
            let format: Format = { id, runs: {}, folder, subject, subjects };
            configure.runs(format);
            utils.write.prop(id, format);
            return format;
        },
    };
}
function initConfig(props: GoogleAppsScript.Properties.Properties) {
    let formats = ['billing', 'payroll'];
    let sheetValues = readSourceSpreadsheet(props);
    let subjects = loadSubjects(sheetValues, formats);
}
function readSourceSpreadsheet(props: GoogleAppsScript.Properties.Properties) {
    let ss = loadSourceSpreadsheet(props)
    return (ss.getSheets().reduce((o, s) => {
        o[s.getName()] = s.getSheetValues(1, 1, -1, -1);
        return o;
    }, {}));
}
function loadSubjects(sheetValues: Object, formats: string[]) {
    for (let format in formats) {
        let data = loadSubjectData(sheetValues, format);
        let headers = data.shift();
        return (data.map(row => {
            let o = {};
            headers.forEach(header => o[header] = row);
            return o;
        }));
    }
    SpreadsheetApp.getUi().prompt('')
}
function loadSubjectData(values: Object, format: string): any[][] {
    let subject = format == 'billing' ? 'CLIENTS' : 'COURIERS';
    if (values.hasOwnProperty(subject)) return values[subject];
    let message = 'Enter sheet name for ' + subject;
    while (true) {
        let name = SpreadsheetApp.getUi().prompt(message).getResponseText();
        if (values.hasOwnProperty(name)) return values[name];
    }
}
function loadSourceSpreadsheet(props: GoogleAppsScript.Properties.Properties): GoogleAppsScript.Spreadsheet.Spreadsheet {
    if (!props.getProperty('sheet')) showPicker();
    return SpreadsheetApp.openById(props.getProperty('sheet'));
}
function showPicker() {
    var html = HtmlService.createHtmlOutputFromFile('dialog.html')
        .setWidth(600)
        .setHeight(425)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, 'Select source spreadsheet');
}
function saveSheet(id) {
    PropertiesService.getDocumentProperties().setProperty('sheet', id);
}
function getOAuthToken() {
    DriveApp.getRootFolder();
    return ScriptApp.getOAuthToken();
}
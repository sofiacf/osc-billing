let utils = (() => {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let properties = PropertiesService.getDocumentProperties();
  let ui = SpreadsheetApp.getUi();
  return {
    read: {
      sheet: (n: string): any[] => ss.getSheetByName(n).getSheetValues(1, 1, -1, -1),
      range: (r: string): any[] => ss.getRangeByName(r).getValues(),
      prop: (p: string): any => {
        try {
          return JSON.parse(properties.getProperty(p));
        }
        catch {
          return null;
        }
        ;
      },
      alert: (m: string) => ui.alert(m, ui.ButtonSet.YES_NO) == ui.Button.YES,
      prompt: (m: string) => ui.prompt(m, ui.ButtonSet.OK_CANCEL),
    },
    write: {
      prop: (k: string, v: {}) => properties.setProperty(k, JSON.stringify(v)),
    },
    erase: {
      prop: (p: string) => properties.deleteProperty(p),
      folder: (folder: GoogleAppsScript.Drive.Folder) => folder.setTrashed(true),
      props: () => properties.deleteAllProperties()
    },
    get: {
      folder: {
        byName: (name: string) => DriveApp.getFoldersByName(name).next(),
        byId: (id: string) => DriveApp.getFolderById(id)
      },
      sheets: (): GoogleAppsScript.Spreadsheet.Sheet[] => ss.getSheets(),
    }
  };
})();

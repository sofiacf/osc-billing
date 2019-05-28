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

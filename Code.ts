function process(setup: Setup) {
  let process = {
    reset: (run: Run) => {
      let files = utils.get.folder.byId(run.folder).getFiles();
      while (files.hasNext()) files.next().setTrashed(true);
      if (utils.read.alert('Erase all properties?')) utils.erase.props();
    },
    subjects: (format: Format, subjects: string[]) => {
      let saved = Object.keys(format.subjects);
      let missing = subjects.filter(s => saved.indexOf(s) < 0);
      let properties = Object.keys(format.subjects[saved[0]]);
      for (let i = 0; i < missing.length; i++) {
        let info: Subject;
        for (let j = 0; j < properties.length; j++) {
          let p = properties[j];
          info[p] = utils.read.prompt('Provide ' + p + ' for ' + missing[i]);
        }
        return info;
      }
    },
  }
  if (setup.config.action == 'RESET') process.reset(setup.run);
  else {
    let data = process.subjects(setup.run, setup.run.subject);
    Logger.log(data);
  }
}
interface Setup {
  config: Config; format: Format; run: any;
}
interface Config {
  month: string, action: 'RUN' | 'RESET' | 'POST', date: Date, type: 'BILLING' | 'PAYROLL';
}
interface Run {
  data: {}; subjects: string[]; folder: string;
}
interface Format {
  id: string; folder: string; subject: string; runs: {}; subjects: Subject[];
}
interface Subject {
  name: any; street?: any; city?: any; attn?: any; sfx?: any; dates?: {};
}

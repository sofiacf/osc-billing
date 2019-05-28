function initialize(): Setup {
    let init = {
        config: ((): Config => {
            let a = utils.read.range('SETTINGS').map(x => x[0]);
            return { month: a[2], action: a[1], date: a[3], type: a[0] };
        })(),
        format: (type: 'BILLING' | 'PAYROLL'): Format => utils.read.prop(type) || configure.formats(type),
        run: (format: Format, month: string): Run => format.runs[month] || configure.runs(format)[month]
    };
    let config = init.config;
    let format = init.format(config.type);
    let run = init.run(format, config.month);
    return { config, format, run };
}

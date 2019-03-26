function runInvoices(){
  billing.run();
}
function runPayroll(){
  run(payroll);
}
function run(format){
  var newSubjects = new Array,
      folder = getRunFolder(),
      subjects = getSubs(),
      charges = getCharges(subjects),
      details = getDetails(subjects);
  for (var i = 0; i < subjects.length; i++){
    if (!details[i]) {
      newSubjects.push(subjects[i]);
      continue;
    }
    var subject = subjects[i],
      scharges = charges[i],
      sdetails = details[i];
    var sheet = getSheet(subject, sdetails, scharges);
    printPDF(subject, sheet, folder);
  }
}

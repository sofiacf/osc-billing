//Remove header row and move "attn" line after city/state!
//Check if there's a template. If there is, copy a new one. If not, make a new one with name.
//Check if the template has address, etc. If not, add to sheet. Remove line.
function findFolders(n){return DriveApp.getFoldersByName(n);}
function findFiles(n){return DriveApp.getFilesByName(n);}
function noFolder(n){return !findFolders(n).hasNext()}
function noFile(n){return !findFiles(n).hasNext()}
function newIfNoFolder(n, dest){
  var search = dest.getFoldersByName(n);
  return search.hasNext() ? search.next() : dest.createFolder(n);
}
function copyIfNoFile(n, dest){
  var search = dest.getFilesByName();
  return search.hasNext() ? search.next() : dest.createFile(n);
}
var sinfoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
var sf = DriveApp.getFoldersByName("CLIENTS").next();
var st = "BILLING TEMPLATE";
var Client = function(array){
  var a = array[0];
  this.id = a[0], this.name = a[1], this.attn = a[2], this.address = a[3];
  this.city = a[4], this.lastBillDate = a[5], this.lastBillNum = a[6];
  this.suffix = a[7];
  this.tn = a[0] + " TEMPLATE";
  this.folder = function(){return newIfNoFolder(this.id, sf);}
  this.template = function(){return copyIfNoFile(this.tn, st, this.folder());}
}
function updateSubjects(){
  for (var i=1; i<sinfoSheet.getLastRow(); i++){
    var cells = "A" + i + ":I" + i;
    var line = sinfoSheet.getRange(cells);
    var data = line.getValues();
    var client = new Client(data);
    var folder = client.folder();
  }
}

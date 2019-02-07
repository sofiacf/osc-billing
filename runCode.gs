function runInvoices() {
  setupFolder();
  var clients = setupClients();
  for (r=0; r<clients.length; r++){
    clients[r].generateInvoice();
  }
};
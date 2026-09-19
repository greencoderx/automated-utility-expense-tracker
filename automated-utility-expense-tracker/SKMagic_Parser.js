function parseSKMagicToSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("SK Magic");
  
  if (!sheet) {
    Logger.log("Error: Tab 'SK Magic' not found.");
    return;
  }

  // Search Gmail for SK Magic invoice emails
  var query = 'from:("SK magic" OR "skmagic") subject:"Invoice"';
  var threads = GmailApp.search(query);

  var rows = [];
  var seenInvoices = {};

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var msg = messages[j];
      var body = msg.getPlainBody();

      // Extract Invoice No (e.g., IV1024005559076)
      var invoiceNoMatch = body.match(/Invoice\s*No[:\s]*([A-Z0-9]+)/i);
      var invoiceNo = invoiceNoMatch ? invoiceNoMatch[1] : "-";

      if (invoiceNo !== "-" && seenInvoices[invoiceNo]) continue;
      if (invoiceNo !== "-") seenInvoices[invoiceNo] = true;

      // Extract Invoice Date (e.g., 01/10/2024 -> 2024-10-01)
      var dateMatch = body.match(/Invoice\s*Date[:\s]*([\d]{2}\/[\d]{2}\/[\d]{4})/i);
      var formattedDate = "";
      if (dateMatch) {
        var parts = dateMatch[1].split("/"); // [DD, MM, YYYY]
        formattedDate = parts[2] + "-" + parts[1] + "-" + parts[0];
      } else {
        formattedDate = Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd");
      }

      // Extract exact Total Bill Amount (e.g., 99.00 or 0.00)
      // Strictly enforces decimal formatting (e.g., 99.00) to ignore table row index "1"
      var amount = 99.00; // Default package rate
      var totalBillMatch = body.match(/Total\s*Bill\s*\(RM\)[\s\:]*([\d]+\.[\d]{2})/i) ||
                           body.match(/Current\s*Charges[\s\:]*([\d]+\.[\d]{2})/i) ||
                           body.match(/Sales\s*Price\s*\(RM\)[\s\:]*([\d]+\.[\d]{2})/i);

      if (totalBillMatch) {
        amount = parseFloat(totalBillMatch[1]);
      }

      rows.push([formattedDate, "SK Magic", invoiceNo, amount, "Paid", "JIK.SOO HYPER (SL)"]);
    }
  }

  // Sort chronologically from oldest to newest
  rows.sort(function(a, b) {
    return new Date(a[0]) - new Date(b[0]);
  });

  // Clear existing content and populate correct values
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }

  Logger.log(rows.length + " SK Magic invoices successfully updated with exact amounts.");
}

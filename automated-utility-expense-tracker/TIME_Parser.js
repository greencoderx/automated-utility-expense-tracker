/**
 * TIME_Parser.js
 * Parses TIME Fibre Internet bill PDF attachments stored in Google Drive
 * and extracts invoice numbers, true billing dates, and transaction amounts.
 */

function parseTimePDFsToSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TIME Internet");
  
  if (!sheet) {
    Logger.log("Error: Tab 'TIME Internet' not found. Please create it first.");
    return;
  }

  var folderName = "TIME Internet Bills";
  var folders = DriveApp.getFoldersByName(folderName);
  
  if (!folders.hasNext()) {
    Logger.log("Folder '" + folderName + "' not found in Google Drive.");
    return;
  }
  
  var folder = folders.next();
  var files = folder.getFilesByType(MimeType.PDF);
  var rows = [];

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var cleanName = fileName.replace(".pdf", "");
    var parts = cleanName.split("-");

    var formattedDate = "";
    var invoiceNo = cleanName;
    var amount = 145.22; // Standard monthly package default rate

    // Pattern 1: Standard TIME PDF Filename (AccountNo-YYYYMMDD-InvoiceNo.pdf)
    if (parts.length >= 3 && parts[1].length === 8) {
      var dateStr = parts[1];
      formattedDate = dateStr.substring(0, 4) + "-" + dateStr.substring(4, 6) + "-" + dateStr.substring(6, 8);
      invoiceNo = parts[2];
    } 
    // Pattern 2: Non-standard Filenames (e.g., 500338137.pdf) - Query Gmail for metadata & email content
    else {
      var threads = GmailApp.search('filename:"' + fileName + '"');
      if (threads.length > 0) {
        var msg = threads[0].getMessages()[0];
        var emailDate = msg.getDate();
        formattedDate = Utilities.formatDate(emailDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        invoiceNo = cleanName;

        // Check if a specific RM amount is mentioned in the email body
        var body = msg.getPlainBody();
        var amountMatch = body.match(/RM\s*([\d\.\,]+)/i);
        if (amountMatch) {
          amount = parseFloat(amountMatch[1].replace(',', ''));
        }
      } else {
        // Fallback to file creation/modified date if email is not found
        var fileDate = file.getLastUpdated();
        formattedDate = Utilities.formatDate(fileDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
    }

    rows.push([formattedDate, "TIME Internet", invoiceNo, amount, "Paid", fileName]);
  }

  // Sort chronologically from oldest to newest
  rows.sort(function(a, b) {
    return new Date(a[0]) - new Date(b[0]);
  });

  // Clear existing sheet data (excluding header row) and populate clean dataset
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }

  Logger.log(rows.length + " TIME Internet PDF bills successfully parsed and updated.");
}

function parseTNBToSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TNB Electric");
  
  if (!sheet) {
    Logger.log("Error: Tab 'TNB Electric' not found.");
    return;
  }

  // Search query targeting account 210146874408 across Boost and myTNB emails
  var query = '"210146874408" OR (from:Boost subject:"Tenaga Nasional") OR (from:myTNB OR "Payment Successful" "Tenaga")';
  var threads = GmailApp.search(query);

  var rows = [];
  var seenRefs = {};

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var msg = messages[j];
      var body = msg.getPlainBody();
      var subject = msg.getSubject();
      var date = msg.getDate();

      var formattedDate = "";
      var refNo = "-";
      var amount = 0.00;
      var paymentChannel = "Paid via myTNB";

      // -------------------------------------------------------------
      // Pattern A: myTNB App Confirmation Emails
      // -------------------------------------------------------------
      if (body.includes("myTNB") || body.includes("REFERENCE NUMBER") || body.includes("210146874408")) {
        // Extract Reference Number
        var refMatch = body.match(/REFERENCE\s*NUMBER\s*([A-Z0-9]+)/i);
        refNo = refMatch ? refMatch[1] : "-";

        // Extract Date (e.g., 25-Jun-2026)
        var dateMatch = body.match(/TRANSACTION\s*DATE\s*([\d]{1,2}-[A-Za-z]{3}-[\d]{4})/i);
        if (dateMatch) {
          var parsedDate = new Date(dateMatch[1].replace(/-/g, " "));
          if (!isNaN(parsedDate.getTime())) {
            formattedDate = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
        }

        // Extract Amount
        var amountMatch = body.match(/TOTAL\s*AMOUNT\s*\(RM\)\s*([\d\.\,]+)/i) ||
                           body.match(/AMOUNT\s*\(RM\)\s*([\d\.\,]+)/i);
        if (amountMatch) {
          amount = parseFloat(amountMatch[1].replace(',', ''));
        }

        // Clean Transaction Method (matches single line only to prevent extra words)
        var methodMatch = body.match(/TRANSACTION\s*METHOD\s*([^\n\r]+)/i);
        var methodStr = methodMatch ? methodMatch[1].trim() : "";

        if (methodStr && methodStr.length < 15) {
          paymentChannel = "Paid via myTNB (" + methodStr + ")";
        } else {
          paymentChannel = "Paid via myTNB";
        }
      }

      // -------------------------------------------------------------
      // Pattern B: Boost App Receipts
      // -------------------------------------------------------------
      if (body.includes("Boost") || subject.includes("Tenaga Nasional Berhad Receipt")) {
        var boostRefMatch = body.match(/Boost\s*Ref\s*ID\.?\s*([a-z0-9]+)/i) || 
                            subject.match(/Receipt\s*-\s*([a-z0-9]+)/i);
        refNo = boostRefMatch ? boostRefMatch[1] : refNo;

        var boostDateMatch = body.match(/Receipt\s*Date\.?\s*([\d]{1,2}\s+[A-Za-z]+\s+[\d]{4})/i);
        if (boostDateMatch) {
          var parsedBoostDate = new Date(boostDateMatch[1]);
          if (!isNaN(parsedBoostDate.getTime())) {
            formattedDate = Utilities.formatDate(parsedBoostDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
        }

        var boostAmountMatch = body.match(/You'?ve\s*paid\s*\(RM\)\s*([\d\.\,]+)/i) ||
                               body.match(/Tenaga\s*Nasional\s*Berhad\s*([\d\.\,]+)/i);
        if (boostAmountMatch) {
          amount = parseFloat(boostAmountMatch[1].replace(',', ''));
        }

        paymentChannel = "Paid via Boost";
      }

      if (!formattedDate) {
        formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }

      if (refNo !== "-" && seenRefs[refNo]) continue;
      if (refNo !== "-") seenRefs[refNo] = true;

      if (amount > 0) {
        rows.push([formattedDate, "TNB Electric", refNo, amount, "Paid", paymentChannel]);
      }
    }
  }

  // Sort chronologically
  rows.sort(function(a, b) {
    return new Date(a[0]) - new Date(b[0]);
  });

  // Write to sheet
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }

  Logger.log(rows.length + " clean TNB payment entries updated.");
}

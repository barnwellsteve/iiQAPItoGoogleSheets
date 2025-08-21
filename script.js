function getTicketInformation() {
  var subdomain = "";
  var siteGuid  = "";
  var apiToken  = "";

  var PAGE_SIZE = 20;   // I always seem to be capped at 20
  var MAX_PAGES = 1000;

  var base = "https://" + subdomain + ".incidentiq.com/api/v1.0/tickets";

  // Either make a sheet or this will make one for you based on the name
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("iiQAPI2") || ss.insertSheet("iiQAPI2");
  sheet.getRange("A2:I").clearContent(); // This will leave the header row alone

  var headers = {
    "Authorization": "Bearer " + apiToken,
    "SiteId": siteGuid,
    "Client": "ApiClient",
    "Accept": "application/json",
    "Content-Type": "application/json"
  };

  // tiny helper so H/I write real dates
  var asDate = function (v) { if (!v) return ""; var d = new Date(v); return isNaN(d) ? "" : d; };

  var nextRow = 2;
  for (var p = 0; p < MAX_PAGES; p++) {
    // Use $s (size), $p (page), $o (order field), $d (direction)
    var url = base + "?$s=" + PAGE_SIZE + "&$p=" + p + "&$o=TicketCreatedDate&$d=Descending";

    // iiQ’s tickets endpoint still expects POST, but it honors the query params above
    var resp = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ Schema: "All" }), // minimal, harmless body
      headers: headers,
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      throw new Error("HTTP " + resp.getResponseCode() + ": " + resp.getContentText().slice(0, 800));
    }

    var parsed = JSON.parse(resp.getContentText());
    var items  = Array.isArray(parsed) ? parsed : (parsed.Items || parsed.items || []);
    if (!items || !items.length) break;

    var rows = items.map(function (t) {
      return [
        t.Subject || "",
        t.Priority || "",                         
        t.WorkflowStep?.StatusName || "",
        t.For?.Name || "",
        t.AssignedToUser?.Name || "",
        t.Location?.Name || "",
        t.LocationRoom?.Name || "",
        asDate(t.CreatedDate || t.TicketCreatedDate),
        asDate(t.ClosedDate)
      ];
    });

    sheet.getRange(nextRow, 1, rows.length, 9).setValues(rows);
    sheet.getRange(nextRow, 8, rows.length, 2).setNumberFormat("yyyy-mm-dd hh:mm");
    nextRow += rows.length;

    // stop if we’ve hit the last page
    if (items.length < PAGE_SIZE) break;
    Utilities.sleep(150); // Pause shortly to let the API rest
  }
}

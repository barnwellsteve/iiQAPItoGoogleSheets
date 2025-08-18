function getTicketInformation2() {
  // === CONFIG ===
  var subdomain = ""; // whatever comes before the .incidentiq.com site
  var siteGuid  = ""; // your GUID - get it from Developer Tools
  var apiToken  = ""; // you iiQ API token

  var url = "https://" + subdomain + ".incidentiq.com/api/v1.0/tickets";
  var PAGE_SIZE = 20; // I had no luck going larger than 20 per page
  var MAX_PAGES = 100; // multiply the page size and this, that's how many tickets get returned

  // Sheet prep
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("iiQAPI2"); // make this sheet before running
  sheet.getRange("A2:I").clearContent(); // keep header row if you have one

  var nextRow = 2;

  for (var page = 0; page < MAX_PAGES; page++) {
    var body = {
      RequestOptions: {
        Paging: { PageIndex: page, PageSize: PAGE_SIZE },
        Sort:   [{ Field: "TicketCreatedDate", Descending: false }]
        // not actually sure if the created date sort does anything
      }
    };

    var options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(body),
      headers: {
        "Authorization": "Bearer " + apiToken,
        "SiteId": siteGuid,
        "Client": "ApiClient",
        "Accept": "application/json"
      },
      muteHttpExceptions: true
    };

    var resp = UrlFetchApp.fetch(url, options);
    var code = resp.getResponseCode();
    if (code !== 200) throw new Error("HTTP " + code + ": " + resp.getContentText().slice(0, 800));

    var parsed = JSON.parse(resp.getContentText());
    var items  = Array.isArray(parsed) ? parsed : (parsed.Items || parsed.items || []);
    if (!items || !items.length) break;


    // These were the 9 fields that mattered to me
    var rows = items.map(function (t) {
      return [
        t.Subject || "",                          // A: Ticket (subject)
        t.Priority || "",                         // B: Priority
        t.WorkflowStep?.StatusName || "",         // C: Status
        t.For?.Name || "",                        // D: Requested For
        t.AssignedToUser?.Name || "",             // E: Assigned To
        t.Location?.Name || "",                   // F: Location
        t.LocationRoom?.Name || "",               // G: Room
        t.CreatedDate,                            // H: Submitted
        t.ClosedDate                              // I: Ticket Closed Date
      ];
    });

    sheet.getRange(nextRow, 1, rows.length, 9).setValues(rows);
    // Clean date formatting for H:I
    sheet.getRange(nextRow, 8, rows.length, 2).setNumberFormat("yyyy-mm-dd hh:mm");

    nextRow += rows.length;

    // If this page returned fewer than PAGE_SIZE, we're done
    if (items.length < PAGE_SIZE) break;

    Utilities.sleep(150); // be polite, small pause between pages
  }

  // Optional: tidy columns
  // sheet.autoResizeColumns(1, 9);
}
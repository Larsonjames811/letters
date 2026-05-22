// Google Apps Script — landing page tracker.
// Logs every page view and CTA click to a Google Sheet.
//
// SETUP (10 min):
//   1. Create a new Google Sheet. Name it e.g. "Letter Campaign Tracker".
//   2. In the sheet: Extensions > Apps Script. Delete any boilerplate code,
//      paste this whole file in, and save.
//   3. Deploy > New deployment > type "Web app".
//      - Execute as: Me
//      - Who has access: Anyone
//      - Click Deploy, authorize when prompted, copy the /exec URL.
//   4. Paste that /exec URL into template.html (replace {{TRACKER_URL}}).
//   5. Sanity check: open the /exec URL in a browser with
//      ?prospect=test&event=view appended. A row should appear in the
//      "events" tab of your sheet.
//
// Sheet layout (auto-created on first run):
//   events tab     — every ping, one row per scan/click
//   prospects tab  — rolling summary: scans, clicks, first-seen, last-seen
//
// If you redeploy after editing, choose "Manage deployments" and update the
// existing deployment (don't create a new one), so the URL stays the same.

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var events = ss.getSheetByName('events') || initEventsSheet(ss);
    var prospects = ss.getSheetByName('prospects') || initProspectsSheet(ss);

    var prospect = (e && e.parameter && e.parameter.prospect) || 'unknown';
    var event = (e && e.parameter && e.parameter.event) || 'view';
    var now = new Date();
    var ua = (e && e.parameter && e.parameter.ua) || '';

    events.appendRow([now, prospect, event, ua]);
    updateProspectSummary(prospects, prospect, event, now);

    // 1x1 transparent gif — so this works when called from <img>/new Image().
    var gif = Utilities.newBlob(
      Utilities.base64Decode('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'),
      'image/gif'
    );
    return ContentService
      .createTextOutput(gif.getDataAsString())
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput('err: ' + err.message);
  }
}

function initEventsSheet(ss) {
  var s = ss.insertSheet('events');
  s.appendRow(['timestamp', 'prospect', 'event', 'user_agent']);
  s.setFrozenRows(1);
  return s;
}

function initProspectsSheet(ss) {
  var s = ss.insertSheet('prospects');
  s.appendRow(['prospect', 'views', 'cta_clicks', 'first_seen', 'last_seen']);
  s.setFrozenRows(1);
  return s;
}

function updateProspectSummary(sheet, prospect, event, now) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === prospect) {
      var row = i + 1;
      if (event === 'view') sheet.getRange(row, 2).setValue((data[i][1] || 0) + 1);
      if (event === 'cta_click') sheet.getRange(row, 3).setValue((data[i][2] || 0) + 1);
      sheet.getRange(row, 5).setValue(now);
      return;
    }
  }
  sheet.appendRow([
    prospect,
    event === 'view' ? 1 : 0,
    event === 'cta_click' ? 1 : 0,
    now,
    now
  ]);
}

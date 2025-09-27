/**
 * Google Apps Script Web App for Yacht Calendar
 * One endpoint returns events for two calendars between start/end dates in a given TZ.
 * Response format:
 * {
 *   tz: "Asia/Dubai",
 *   boats: [
 *     { id:"A", name:"Yacht A", events:[ {start:"2025-09-22T08:30:00+04:00", end:"2025-09-22T11:00:00+04:00"}, ... ] },
 *     { id:"B", name:"Yacht B", events:[ ... ] }
 *   ]
 * }
 */

function doGet(e) {
  try {
    var tz = (e && e.parameter && e.parameter.tz) || 'Asia/Dubai';
    var startYMD = e && e.parameter && e.parameter.start;
    var endYMD = e && e.parameter && e.parameter.end;
    if (!startYMD || !endYMD) {
      return json_({ error: 'Missing start or end (YYYY-MM-DD)' });
    }

    var range = makeRangeFromYMD(startYMD, endYMD, tz);
    var boats = getBoatsConfig(); // [{id,name,calId}]

    var out = { tz: tz, boats: [] };

    boats.forEach(function(b) {
      var cal = CalendarApp.getCalendarById(b.calId);
      if (!cal) return;

      var events = cal.getEvents(range.start, range.end).map(function(ev) {
        var s = ev.getStartTime();
        var f = ev.getEndTime();
        if (ev.isAllDayEvent && ev.isAllDayEvent()) {
          // Normalize all-day to 00:00–23:59 in TZ
          var sYMD = Utilities.formatDate(s, tz, 'yyyy-MM-dd');
          var fYMD = Utilities.formatDate(f, tz, 'yyyy-MM-dd');
          s = parseTZBoundary(sYMD, tz, true);
          // For all‑day end, Apps Script often sets next-day midnight; clamp to end of previous day
          var endDay = new Date(parseTZBoundary(fYMD, tz, true).getTime() - 1);
          f = endDay;
        }
        return { start: formatISO(s, tz), end: formatISO(f, tz) };
      });

      out.boats.push({ id: b.id, name: b.name || b.id, events: events });
    });

    return json_(out);
  } catch (err) {
    return json_({ error: String((err && err.message) || err) });
  }
}

// ---------------- Config ----------------

// Option A: Hardcode here (fallback if Script Properties are not set)
var BOATS_HARDCODED = [
  // Replace with real Calendar IDs or set Script Properties (recommended)
  { id: 'A', name: 'Yacht A', calId: 'PUT_CALENDAR_ID_A_HERE' },
  { id: 'B', name: 'Yacht B', calId: 'PUT_CALENDAR_ID_B_HERE' }
];

// Option B (recommended): Script Properties
// Keys: CAL_A_ID, CAL_A_NAME, CAL_B_ID, CAL_B_NAME
function getBoatsConfig() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var aId = props.CAL_A_ID, aName = props.CAL_A_NAME || 'Yacht A';
  var bId = props.CAL_B_ID, bName = props.CAL_B_NAME || 'Yacht B';
  if (aId && bId) {
    return [
      { id: 'A', name: aName, calId: aId },
      { id: 'B', name: bName, calId: bId }
    ];
  }
  return BOATS_HARDCODED;
}

// ---------------- Helpers ----------------

// Format ISO with timezone offset, e.g. 2025-09-22T08:30:00+04:00
function formatISO(date, tz) {
  return Utilities.formatDate(date, tz, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

// Build range from YYYY-MM-DD strings in a timezone
function makeRangeFromYMD(startYMD, endYMD, tz) {
  var s = parseTZBoundary(startYMD, tz, true);
  var e = parseTZBoundary(endYMD, tz, false);
  return { start: s, end: e };
}

// Return Date for start/end of day in specified tz
function parseTZBoundary(ymd, tz, isStart) {
  var parts = ymd.split('-');
  var y = Number(parts[0]), m = Number(parts[1]) - 1, d = Number(parts[2]);
  // Construct local boundary in tz by adjusting from a UTC anchor with tz offset
  var baseUtc = new Date(Date.UTC(y, m, d, isStart ? 0 : 23, isStart ? 0 : 59, isStart ? 0 : 59, isStart ? 0 : 999));
  var offsetStr = Utilities.formatDate(baseUtc, tz, 'Z'); // "+0400"
  var sign = offsetStr[0] === '-' ? -1 : 1;
  var hh = Number(offsetStr.slice(1, 3));
  var mm = Number(offsetStr.slice(3, 5));
  var offsetMin = sign * (hh * 60 + mm);
  var msUtc = baseUtc.getTime() - offsetMin * 60000;
  return new Date(msUtc);
}

// JSON response wrapper
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Optional: simple ping route for testing deployments
function doGet_ping() {
  return json_({ ok: true, now: new Date().toISOString() });
}


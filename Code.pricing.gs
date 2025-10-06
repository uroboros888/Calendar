/**
 * Google Apps Script Web App for Yacht Calendar (Pricing-aware version)
 *
 * Adds the ability to fetch pricing configuration from the "Pricing" sheet
 * of the bound spreadsheet. Supports three modes via the `mode` query
 * parameter:
 *   - (default) events: only calendar events
 *   - pricing: only pricing configuration
 *   - combined: events + pricing in a single payload
 */

function doGet(e) {
  var mode = (e && e.parameter && e.parameter.mode) || 'events';
  var callback = e && e.parameter && e.parameter.callback;
  try {
    if (mode === 'pricing') {
      var pricingOnly = loadPricingConfig_();
      return json_(pricingOnly, callback);
    }

    var tz = (e && e.parameter && e.parameter.tz) || 'Asia/Dubai';
    var startYMD = e && e.parameter && e.parameter.start;
    var endYMD = e && e.parameter && e.parameter.end;
    if (!startYMD || !endYMD) {
      return json_({ error: 'Missing start or end (YYYY-MM-DD)' }, callback);
    }

    var range = makeRangeFromYMD(startYMD, endYMD, tz);
    var boats = getBoatsConfig();
    var pricing = (mode === 'combined') ? loadPricingConfig_() : null;
    var nameOverrides = pricing ? pricing._nameById : null;
    var eventsPayload = buildEventsPayload_(boats, range, tz, nameOverrides);

    if (mode === 'combined') {
      delete pricing._nameById;
      return json_({
        tz: tz,
        range: { start: startYMD, end: endYMD },
        boats: eventsPayload.boats,
        pricing: pricing
      }, callback);
    }

    return json_(eventsPayload, callback);
  } catch (err) {
    return json_({ error: String((err && err.message) || err) }, callback);
  }
}

function buildEventsPayload_(boats, range, tz, nameOverrides) {
  var out = { tz: tz, boats: [] };
  boats.forEach(function (b) {
    var cal = CalendarApp.getCalendarById(b.calId);
    if (!cal) return;
    var events = cal.getEvents(range.start, range.end).map(function (ev) {
      var s = ev.getStartTime();
      var f = ev.getEndTime();
      if (ev.isAllDayEvent && ev.isAllDayEvent()) {
        var sYMD = Utilities.formatDate(s, tz, 'yyyy-MM-dd');
        var fYMD = Utilities.formatDate(f, tz, 'yyyy-MM-dd');
        s = parseTZBoundary(sYMD, tz, true);
        var endDay = new Date(parseTZBoundary(fYMD, tz, true).getTime() - 1);
        f = endDay;
      }
      return { start: formatISO(s, tz), end: formatISO(f, tz) };
    });
    var name = (nameOverrides && nameOverrides[b.id]) || b.name || b.id;
    out.boats.push({ id: b.id, name: name, events: events });
  });
  return out;
}

function loadPricingConfig_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Pricing');
  if (!sheet) {
    throw new Error('Sheet "Pricing" not found');
  }
  var dataRange = sheet.getDataRange();
  var raw = dataRange.getValues();
  var display = dataRange.getDisplayValues();
  var height = raw.length;
  var boatHeaderRow = findRowIndex_(display, 'boat_id');
  if (boatHeaderRow === -1) {
    throw new Error('Header "boat_id" not found in Pricing sheet');
  }
  var boats = [];
  for (var r = boatHeaderRow + 1; r < height; r++) {
    var id = trim_(display[r][0]);
    if (!id) {
      break;
    }
    var boat = {
      id: id,
      name: trim_(display[r][1]) || id,
      baseRate: normalizeNumber_(parseNumber_(raw[r][2], display[r][2])),
      minRate: normalizeNumber_(parseNumber_(raw[r][3], display[r][3])),
      maxRate: normalizeNumber_(parseNumber_(raw[r][4], display[r][4])),
      roundTo: normalizeNumber_(parseNumber_(raw[r][5], display[r][5]))
    };
    boats.push(boat);
  }
  if (!boats.length) {
    throw new Error('No boats configured in Pricing sheet');
  }
  var bandHeaderRow = findRowIndex_(display, 'band');
  var bands = [];
  if (bandHeaderRow !== -1) {
    var bandHeader = display[bandHeaderRow];
    var bandBoatCols = extractMultiplierColumns_(bandHeader);
    for (var i = bandHeaderRow + 1; i < height; i++) {
      var label = trim_(display[i][0]);
      if (!label) {
        if (isRowEmpty_(display[i])) {
          break;
        }
        continue;
      }
      var startTime = parseTimeCell_(raw[i][1], display[i][1]);
      var endTime = parseTimeCell_(raw[i][2], display[i][2]);
      var multipliers = {};
      bandBoatCols.forEach(function (col) {
        var val = parseNumber_(raw[i][col.index], display[i][col.index]);
        if (isFinite(val) && !isNaN(val)) {
          multipliers[col.id] = val;
        }
      });
      bands.push({
        label: label,
        start: startTime.text,
        end: endTime.text,
        startMinutes: startTime.minutes,
        endMinutes: endTime.minutes,
        multipliers: multipliers
      });
    }
  }
  var dowHeaderRow = findRowIndex_(display, 'dow');
  var dowMultipliers = {};
  if (dowHeaderRow !== -1) {
    var dowHeader = display[dowHeaderRow];
    var dowBoatCols = extractMultiplierColumns_(dowHeader);
    for (var d = dowHeaderRow + 1; d < height; d++) {
      var dowKey = normalizeDow_(display[d][0]);
      if (!dowKey) {
        if (isRowEmpty_(display[d])) {
          break;
        }
        continue;
      }
      var rowObj = {};
      dowBoatCols.forEach(function (col) {
        var val = parseNumber_(raw[d][col.index], display[d][col.index]);
        if (isFinite(val) && !isNaN(val)) {
          rowObj[col.id] = val;
        }
      });
      dowMultipliers[dowKey] = rowObj;
    }
  }
  var busyHeaderRow = findRowIndex_(display, 'busyfrom%');
  var busyLevels = [];
  if (busyHeaderRow !== -1) {
    var busyHeader = display[busyHeaderRow];
    var busyBoatCols = extractMultiplierColumns_(busyHeader);
    for (var b = busyHeaderRow + 1; b < height; b++) {
      var fromVal = parsePercent_(raw[b][0], display[b][0]);
      if (fromVal === null) {
        if (isRowEmpty_(display[b])) {
          break;
        }
        continue;
      }
      var toVal = parsePercent_(raw[b][1], display[b][1]);
      var rowMultipliers = {};
      busyBoatCols.forEach(function (col) {
        var val = parseNumber_(raw[b][col.index], display[b][col.index]);
        if (isFinite(val) && !isNaN(val)) {
          rowMultipliers[col.id] = val;
        }
      });
      busyLevels.push({
        from: fromVal,
        to: toVal,
        multipliers: rowMultipliers,
        comment: trim_(display[b][busyHeader.length - 1] || '')
      });
    }
    busyLevels.sort(function (a, b) {
      return (a.from || 0) - (b.from || 0);
    });
  }
  var props = PropertiesService.getScriptProperties().getProperties();
  var openTime = props.OPEN_TIME || '08:00';
  var closeTime = props.CLOSE_TIME || '24:00';
  var slotMins = Number(props.SLOT_MINS || 60);
  var tz = props.CALENDAR_TZ || 'Asia/Dubai';
  var defaultRoundTo = props.DEFAULT_ROUND_TO ? Number(props.DEFAULT_ROUND_TO) : null;
  if ((!defaultRoundTo || isNaN(defaultRoundTo)) && boats.length) {
    for (var idx = 0; idx < boats.length; idx++) {
      var rt = boats[idx].roundTo;
      if (rt && !isNaN(rt)) {
        defaultRoundTo = Number(rt);
        break;
      }
    }
  }
  var nameMap = {};
  boats.forEach(function (b) {
    nameMap[b.id] = b.name;
  });
  return {
    timezone: tz,
    open: openTime,
    close: closeTime,
    slotMins: slotMins,
    boats: boats,
    bands: bands,
    dowMultipliers: dowMultipliers,
    busyLevels: busyLevels,
    defaultRoundTo: defaultRoundTo,
    updatedAt: new Date().toISOString(),
    _nameById: nameMap
  };
}

function findRowIndex_(rows, headerValue) {
  var target = String(headerValue || '').toLowerCase();
  for (var i = 0; i < rows.length; i++) {
    var first = String(rows[i][0] || '').toLowerCase();
    if (first === target) {
      return i;
    }
  }
  return -1;
}

function extractMultiplierColumns_(headerRow) {
  var cols = [];
  for (var c = 0; c < headerRow.length; c++) {
    var header = String(headerRow[c] || '');
    var match = header.match(/mult\s*(.+)/i);
    if (match) {
      cols.push({ id: trim_(match[1]), index: c });
    }
  }
  return cols;
}

function parseNumber_(rawValue, displayValue) {
  if (rawValue instanceof Date) {
    return rawValue.getHours() + rawValue.getMinutes() / 60;
  }
  if (typeof rawValue === 'number') {
    return rawValue;
  }
  var str = trim_(displayValue || rawValue);
  if (!str) return NaN;
  str = str.replace(',', '.');
  var num = Number(str);
  return isNaN(num) ? NaN : num;
}

function parseTimeCell_(rawValue, displayValue) {
  if (rawValue instanceof Date) {
    var h = rawValue.getHours();
    var m = rawValue.getMinutes();
    return { text: pad2_(h) + ':' + pad2_(m), minutes: h * 60 + m };
  }
  if (typeof rawValue === 'number') {
    var totalMin = Math.round(rawValue * 24 * 60);
    var hh = Math.floor(totalMin / 60);
    var mm = totalMin % 60;
    return { text: pad2_(hh) + ':' + pad2_(mm), minutes: totalMin };
  }
  var str = trim_(displayValue || rawValue);
  if (!str) {
    return { text: '00:00', minutes: 0 };
  }
  var parts = str.split(':');
  var hh = Number(parts[0]) || 0;
  var mm = Number(parts[1]) || 0;
  var minutes = hh * 60 + mm;
  return { text: pad2_(hh) + ':' + pad2_(mm), minutes: minutes };
}

function parsePercent_(rawValue, displayValue) {
  var num = parseNumber_(rawValue, displayValue);
  if (isNaN(num)) return null;
  if (num > 1) {
    num = num / 100;
  }
  return Math.max(0, Math.min(1, num));
}

function normalizeDow_(value) {
  var str = trim_(value).toLowerCase();
  if (!str) return null;
  var map = {
    'monday': 'monday', 'mon': 'monday', 'понедельник': 'monday',
    'tuesday': 'tuesday', 'tue': 'tuesday', 'вторник': 'tuesday',
    'wednesday': 'wednesday', 'wed': 'wednesday', 'среда': 'wednesday',
    'thursday': 'thursday', 'thu': 'thursday', 'четверг': 'thursday',
    'friday': 'friday', 'fri': 'friday', 'пятница': 'friday',
    'saturday': 'saturday', 'sat': 'saturday', 'суббота': 'saturday',
    'sunday': 'sunday', 'sun': 'sunday', 'воскресенье': 'sunday'
  };
  return map[str] || null;
}

function isRowEmpty_(row) {
  for (var i = 0; i < row.length; i++) {
    if (trim_(row[i])) return false;
  }
  return true;
}

function trim_(value) {
  return String(value || '').trim();
}

function normalizeNumber_(value) {
  if (value === null || value === undefined) return null;
  if (value === '') return null;
  if (typeof value === 'number' && !isNaN(value)) return value;
  var num = Number(value);
  return isNaN(num) ? null : num;
}

function pad2_(n) {
  return ('0' + n).slice(-2);
}

/** Existing helper functions from Code.gs duplicated for standalone file */
function formatISO(date, tz) {
  return Utilities.formatDate(date, tz, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function makeRangeFromYMD(startYMD, endYMD, tz) {
  var s = parseTZBoundary(startYMD, tz, true);
  var e = parseTZBoundary(endYMD, tz, false);
  return { start: s, end: e };
}

function parseTZBoundary(ymd, tz, isStart) {
  var parts = ymd.split('-');
  var y = Number(parts[0]), m = Number(parts[1]) - 1, d = Number(parts[2]);
  var baseUtc = new Date(Date.UTC(y, m, d, isStart ? 0 : 23, isStart ? 0 : 59, isStart ? 0 : 59, isStart ? 0 : 999));
  var offsetStr = Utilities.formatDate(baseUtc, tz, 'Z');
  var sign = offsetStr[0] === '-' ? -1 : 1;
  var hh = Number(offsetStr.slice(1, 3));
  var mm = Number(offsetStr.slice(3, 5));
  var offsetMin = sign * (hh * 60 + mm);
  var msUtc = baseUtc.getTime() - offsetMin * 60000;
  return new Date(msUtc);
}

function json_(obj, callback) {
  var text = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(String(callback) + '(' + text + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(text)
    .setMimeType(ContentService.MimeType.JSON);
}

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

var BOATS_HARDCODED = [
  { id: 'A', name: 'Yacht A', calId: 'PUT_CALENDAR_ID_A_HERE' },
  { id: 'B', name: 'Yacht B', calId: 'PUT_CALENDAR_ID_B_HERE' }
];

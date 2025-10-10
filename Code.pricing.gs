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
  var params = (e && e.parameter) || {};
  try {
    if (mode === 'pricing') {
      var pricingOnly = loadPricingConfig_({
        sheetId: params.sheetId || params.sheet || null,
        user: params.user || params.u || ''
      });
      try {
        logVisit_({
          sheetId: params.sheetId || params.sheet || null,
          mode: 'pricing',
          tz: (e && e.parameter && e.parameter.tz) || (PropertiesService.getScriptProperties().getProperty('CALENDAR_TZ') || 'Asia/Dubai'),
          start: params.start || '',
          end: params.end || '',
          test: String(params.test||'').toLowerCase()==='1' || String(params.env||'').toLowerCase()==='test',
          client: params.client || '',
          uid: params.uid || '',
          ua: params.ua || '',
          user: params.user || params.u || ''
        });
      } catch (ignore) {}
      return json_(pricingOnly, callback);
    }

    var tz = (e && e.parameter && e.parameter.tz) || 'Asia/Dubai';
    var startYMD = e && e.parameter && e.parameter.start;
    var endYMD = e && e.parameter && e.parameter.end;
    if (!startYMD || !endYMD) {
      return json_({ error: 'Missing start or end (YYYY-MM-DD)' }, callback);
    }

    var range = makeRangeFromYMD(startYMD, endYMD, tz);
    // Determine test mode: query param wins; else read from Config sheet (useTestCalendars)
    var testParam = String(params.test||'').toLowerCase()==='1' || String(params.env||'').toLowerCase()==='test';
    var testFromSheet = false;
    if (!testParam) {
      try {
        var ssCfg = openPricingSpreadsheet_(params.sheetId || params.sheet || null);
        var cfgMap = readConfigMap_(ssCfg);
        testFromSheet = cfgMap && cfgMap.usetestcalendars === true;
      } catch (err) {}
    }
    var boats = getBoatsConfig({
      testMode: testParam || testFromSheet,
      calA: params.calA || null,
      calB: params.calB || null,
      calAName: params.calAName || null,
      calBName: params.calBName || null
    });
    var pricing = (mode === 'combined') ? loadPricingConfig_({
      sheetId: params.sheetId || params.sheet || null,
      user: params.user || params.u || ''
    }) : null;
    var nameOverrides = pricing ? pricing._nameById : null;
    var eventsPayload = buildEventsPayload_(boats, range, tz, nameOverrides);

    try {
      logVisit_({
        sheetId: params.sheetId || params.sheet || null,
        mode: mode,
        tz: tz,
        start: startYMD,
        end: endYMD,
        test: String(params.test||'').toLowerCase()==='1' || String(params.env||'').toLowerCase()==='test',
        client: params.client || '',
        uid: params.uid || '',
        ua: params.ua || '',
        user: params.user || params.u || '',
        boats: boats
      });
    } catch (ignore) {}

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

function loadPricingConfig_(opts) {
  var propsAll = PropertiesService.getScriptProperties().getProperties();
  var pricingSheetId = (opts && opts.sheetId) || propsAll.PRICING_SHEET_ID;
  var ss = pricingSheetId ? SpreadsheetApp.openById(pricingSheetId) : SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Pricing');
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
  var tz = props.CALENDAR_TZ || 'Asia/Dubai';
  // Defaults from script properties
  var openTime = props.OPEN_TIME || '08:00';
  var closeTime = props.CLOSE_TIME || '24:00';
  var slotMins = Number(props.SLOT_MINS || 60);
  // Optional overrides from Config sheet
  try {
    var cfgMap = readConfigMap_(ss);
  if (cfgMap) {
      if (cfgMap.open_hm) openTime = cfgMap.open_hm;
      if (cfgMap.close_hm) closeTime = cfgMap.close_hm;
      if (typeof cfgMap.slotmins === 'number' && !isNaN(cfgMap.slotmins)) slotMins = Number(cfgMap.slotmins);
    }
  } catch (err) {}
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
  // Determine pricing method per user (dynamic|normal) from Users sheet
  var userId = opts && opts.user ? String(opts.user).trim() : '';
  var userPolicy = readUserPolicy_(ss, userId);
  // Load special dates from sheet 'Special_DT' if present
  var specialDates = loadSpecialDates_(ss, boats, tz);
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
    specialDates: specialDates,
    // UI/config hints from Config sheet (if present)
    useTestCalendars: (cfgMap && cfgMap.usetestcalendars) === true,
    showBusyDebug: (cfgMap && cfgMap.showbusydebug) === true,
    occupancyMode: (cfgMap && cfgMap.occupancymode) || null,
    occupancyAroundDays: (cfgMap && (cfgMap.occupancyarounddays !== undefined)) ? cfgMap.occupancyarounddays : null,
    occupancyHorizonDays: (cfgMap && (cfgMap.occupancyhorizondays !== undefined)) ? cfgMap.occupancyhorizondays : null,
    occupancyFarMultiplier: (cfgMap && (cfgMap.occupancyfarmultiplier !== undefined)) ? cfgMap.occupancyfarmultiplier : null,
    dowMinOcc: (cfgMap && (cfgMap.dowminocc !== undefined)) ? cfgMap.dowminocc : null,
    pricingMethod: userPolicy && userPolicy.pricingMethod ? userPolicy.pricingMethod : (userId? 'normal':'normal'),
    updatedAt: new Date().toISOString(),
    _nameById: nameMap
  };
}

function openPricingSpreadsheet_(sheetIdOpt) {
  var propsAll = PropertiesService.getScriptProperties().getProperties();
  var pricingSheetId = sheetIdOpt || propsAll.PRICING_SHEET_ID;
  return pricingSheetId ? SpreadsheetApp.openById(pricingSheetId) : SpreadsheetApp.getActive();
}

function readConfigMap_(ss) {
  try {
    var sh = ss.getSheetByName('Config');
    if (!sh) {
      // case-insensitive fallback: find sheet named 'config'
      var sheets = ss.getSheets();
      for (var i = 0; i < sheets.length; i++) {
        if (String(sheets[i].getName() || '').toLowerCase() === 'config') { sh = sheets[i]; break; }
      }
    }
    if (!sh) return null;
    var range = sh.getDataRange();
    var raw = range.getValues();
    var display = range.getDisplayValues();
    var out = {};
    for (var r = 0; r < raw.length; r++) {
      var key = String(display[r][0] || '').trim();
      if (!key || key.toLowerCase() === 'param') continue;
      var k = key.toLowerCase();
      var rv = raw[r][1];
      var dv = display[r][1];
      if (k === 'open' || k === 'close') {
        var tt = parseTimeCell_(rv, dv);
        out[k === 'open' ? 'open_hm' : 'close_hm'] = tt.text;
        continue;
      }
      if (k === 'slotmins' || k === 'occupancyarounddays') {
        var num = parseNumber_(rv, dv);
        if (!isNaN(num)) out[k] = Number(num);
        continue;
      }
      if (k === 'occupancyhorizondays' || k === 'occupancyfarmultiplier') {
        var num2 = parseNumber_(rv, dv);
        if (!isNaN(num2)) out[k] = Number(num2);
        continue;
      }
      if (k === 'dowminocc' || k === 'dowminocc%') {
        var p = parsePercent_(rv, dv);
        if (p !== null) out['dowminocc'] = p;
        continue;
      }
      if (k === 'usetestcalendars' || k === 'showbusydebug') {
        var s = String(dv || rv || '').toLowerCase();
        out[k] = (s === 'true' || s === '1' || s === 'yes');
        continue;
      }
      if (k === 'occupancymode') {
        var val = String(dv || rv || '').trim().toLowerCase();
        if (val) out[k] = val;
        continue;
      }
    }
    return out;
  } catch (err) {
    return null;
  }
}

function readUserPolicy_(ss, userId) {
  try {
    if (!userId) return { pricingMethod: 'normal' };
    var sh = ss.getSheetByName('Users');
    if (!sh) return { pricingMethod: 'normal' };
    var data = sh.getDataRange().getValues();
    if (!data || data.length < 2) return { pricingMethod: 'normal' };
    var header = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
    function col(name) { var idx = header.indexOf(name.toLowerCase()); return idx >= 0 ? idx : -1; }
    var cId = col('id');
    var cUser = col('user');
    var cPricing = col('pricing');
    var found = null;
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var idMatch = (cId>=0 && String(row[cId]).trim() === userId);
      var userMatch = (cUser>=0 && String(row[cUser]).trim().toLowerCase() === userId.toLowerCase());
      if (idMatch || userMatch) { found = row; break; }
    }
    if (!found) return { pricingMethod: 'normal' };
    var allowDynamic = false;
    if (cPricing >= 0) {
      var val = String(found[cPricing]).trim().toLowerCase();
      allowDynamic = (val === 'true' || val === '1' || val === 'yes');
    }
    return { pricingMethod: allowDynamic ? 'dynamic' : 'normal' };
  } catch (e) {
    return { pricingMethod: 'normal' };
  }
}

function loadSpecialDates_(ss, boats, tz) {
  try {
    var sh = ss.getSheetByName('Special_DT');
    if (!sh) {
      // case-insensitive fallback
      var sheets = ss.getSheets();
      for (var i = 0; i < sheets.length; i++) {
        if (String(sheets[i].getName() || '').toLowerCase() === 'special_dt') { sh = sheets[i]; break; }
      }
    }
    if (!sh) return {};
    var range = sh.getDataRange();
    var raw = range.getValues();
    var display = range.getDisplayValues();
    if (!raw || !raw.length) return {};
    // Header indices by name
    var header = raw[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
    function col(name) { var idx = header.indexOf(name.toLowerCase()); return idx >= 0 ? idx : -1; }
    var cDate = col('date');
    var cName = col('name date');
    var cBoat = col('boat');
    var cMin = col('min order');
    var cMorning = col('morning');
    var cDay = col('day');
    var cSunset = col('sunset');
    var cNight = col('night');
    var c24h = col('24h');
    var c12h = col('12h');
    var cStart = col('start');
    var cEnd = col('end');
    var cType = col('type');
    var cNote = col('note');
    var mapByDate = {};
    var boatIds = boats.map(function (b) { return b.id; });
    var boatNames = {};
    boats.forEach(function (b) { boatNames[(b.name||'').toLowerCase()] = b.id; });
    for (var r = 1; r < raw.length; r++) {
      var row = raw[r];
      var disp = display[r];
      var date = parseDateCell_(row[cDate], disp[cDate], tz);
      if (!date) continue;
      var name = cName>=0 ? String(row[cName]||'').trim() : '';
      var boatRaw = cBoat>=0 ? String(row[cBoat]||'').trim() : '';
      var boatId = normalizeBoatId_(boatRaw, boatIds, boatNames);
      var minOrder = cMin>=0 ? normalizeNumber_(parseNumber_(row[cMin], row[cMin])) : null;
      function numAt(ci){ return (ci>=0)? normalizeNumber_(parseNumber_(row[ci], row[ci])) : null; }
      var morning = numAt(cMorning);
      var day = numAt(cDay);
      var sunset = numAt(cSunset);
      var night = numAt(cNight);
      var p24 = numAt(c24h);
      var p12 = numAt(c12h);
      function hmAt(ci){
        if(ci<0) return null;
        var rv = row[ci];
        var dv = disp && disp[ci];
        var isEmpty = (rv===null || rv===undefined || rv==='') && (!dv || String(dv).trim()==='');
        if(isEmpty) return null;
        var t = parseTimeCell_(rv, dv);
        // if explicitly 00:00 but cell looked empty, we already returned null above
        return t.text;
      }
      var startHM = hmAt(cStart);
      var endHM = hmAt(cEnd);
      var type = cType>=0 ? String(row[cType]||'').trim().toLowerCase() : '';
      var note = cNote>=0 ? String(row[cNote]||'').trim() : '';
      var entry = {
        name: name,
        boat: boatId, // null means applies to all
        minOrder: minOrder,
        bands: { morning: morning, day: day, sunset: sunset, night: night },
        p24h: p24, p12h: p12,
        start: startHM, end: endHM,
        type: type, note: note
      };
      if (!mapByDate[date]) mapByDate[date] = { items: [] };
      mapByDate[date].items.push(entry);
    }
    return mapByDate;
  } catch (err) {
    return {};
  }
}

function normalizeBoatId_(val, knownIds, nameMap) {
  var s = String(val||'').trim().toLowerCase();
  if (!s) return null;
  // direct id match
  for (var i=0;i<knownIds.length;i++){ if (String(knownIds[i]||'').toLowerCase()===s) return knownIds[i]; }
  // by name map
  if (nameMap && nameMap[s]) return nameMap[s];
  // simple aliases
  var alias = { 'vd':'A', 'van dutch':'A', 'vandutch':'A', 'mc':'B', 'monte carlo':'B' };
  if (alias[s]) return alias[s];
  return null;
}

function parseDateCell_(rawValue, displayValue, tz) {
  try {
    if (rawValue instanceof Date) {
      return Utilities.formatDate(rawValue, tz || 'Asia/Dubai', 'yyyy-MM-dd');
    }
    var s = String(displayValue || rawValue || '').trim();
    if (!s) return null;
    // ISO
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    // dd.MM.yyyy or dd/MM/yyyy or dd-MM-yyyy
    var m = s.match(/^(\d{1,2})[\.\/\-](\d{1,2})[\.\/\-](\d{2,4})$/);
    if (m) {
      var dd = ('0' + Number(m[1])).slice(-2);
      var MM = ('0' + Number(m[2])).slice(-2);
      var yyyy = String(m[3]).length === 2 ? ('20' + m[3]) : m[3];
      return yyyy + '-' + MM + '-' + dd;
    }
    // try Date.parse
    var d = new Date(s);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, tz || 'Asia/Dubai', 'yyyy-MM-dd');
    }
  } catch (e) {}
  return null;
}

function logVisit_(info) {
  info = info || {};
  var ss = openPricingSpreadsheet_(info.sheetId || null);
  var lock = LockService.getScriptLock();
  try { lock.tryLock(5000); } catch (e) {}
  try {
    var tz = info.tz || (PropertiesService.getScriptProperties().getProperty('CALENDAR_TZ') || 'Asia/Dubai');
    var now = new Date();
    var tsIso = Utilities.formatDate(now, tz, "yyyy-MM-dd'T'HH:mm:ss");
    var visits = ss.getSheetByName('Visits') || ss.insertSheet('Visits');
    if (visits.getLastRow() === 0) {
      visits.appendRow(['ts', 'tz', 'mode', 'start', 'end', 'test', 'client', 'uid', 'user', 'ua', 'calA', 'calB']);
    }
    var calA = '', calB = '';
    if (info.boats && info.boats.length) {
      for (var i = 0; i < info.boats.length; i++) {
        var b = info.boats[i];
        if (b.id === 'A') calA = b.calId || '';
        if (b.id === 'B') calB = b.calId || '';
      }
    }
    visits.appendRow([tsIso, tz, info.mode || '', info.start || '', info.end || '', info.test ? 1 : 0, info.client || '', info.uid || '', info.user || '', info.ua || '', calA, calB]);

    // Users sheet остаётся мастер-реестром, без автосоздания/изменений —
    // при необходимости можно читать записи вручную отдельной функцией.
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
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
  // normalize thousand separators/spaces and decimal comma
  try { str = String(str).replace(/\u00A0|\s/g, ''); } catch(e) {}
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

function getBoatsConfig(opts) {
  opts = opts || {};
  var props = PropertiesService.getScriptProperties().getProperties();
  var useTest = !!opts.testMode;
  var aId = opts.calA || (useTest ? (props.TEST_CAL_A_ID || props.CAL_A_ID) : props.CAL_A_ID);
  var bId = opts.calB || (useTest ? (props.TEST_CAL_B_ID || props.CAL_B_ID) : props.CAL_B_ID);
  var aName = opts.calAName || (useTest ? (props.TEST_CAL_A_NAME || props.CAL_A_NAME) : props.CAL_A_NAME) || 'Yacht A';
  var bName = opts.calBName || (useTest ? (props.TEST_CAL_B_NAME || props.CAL_B_NAME) : props.CAL_B_NAME) || 'Yacht B';
  var boats = [];
  if (aId) boats.push({ id: 'A', name: aName, calId: aId });
  if (bId) boats.push({ id: 'B', name: bName, calId: bId });
  return boats.length ? boats : BOATS_HARDCODED;
}

var BOATS_HARDCODED = [
  { id: 'A', name: 'Yacht A', calId: 'PUT_CALENDAR_ID_A_HERE' },
  { id: 'B', name: 'Yacht B', calId: 'PUT_CALENDAR_ID_B_HERE' }
];

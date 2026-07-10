// ============================================================
//  engine.js — compute free gaps and format the bulleted list
// ============================================================
//
//  Mirrors the VSTO version's logic:
//   - busy = showAs busy / oof / workingElsewhere (tentative optional)
//   - all-day events ignored
//   - overlaps merged, subtracted from each day's window
//   - gaps below minGap dropped
//   - output: "• Monday 7/13: 10:00am-1:00pm CT, 2:30pm-4:00pm CT"
//

var AvailabilityEngine = (function () {

  // Graph returns date-times as "2026-07-13T11:00:00.0000000" (no zone suffix)
  // already converted into the Prefer zone. Parse as a *floating* local time.
  function parseFloating(dateTimeStr) {
    // strip fractional seconds and any trailing Z; treat as wall-clock
    var s = dateTimeStr.replace("Z", "").split(".")[0];
    var parts = s.split("T");
    var d = parts[0].split("-").map(Number);
    var t = parts[1].split(":").map(Number);
    return new Date(d[0], d[1] - 1, d[2], t[0], t[1], t[2] || 0);
  }

  function isBusy(ev, tentativeIsBusy) {
    switch ((ev.showAs || "").toLowerCase()) {
      case "busy":
      case "oof":
      case "workingelsewhere":
        return true;
      case "tentative":
        return tentativeIsBusy;
      case "free":
      default:
        return false;
    }
  }

  function sameDate(a, b) {
    return a.getFullYear() === b.getFullYear() &&
           a.getMonth() === b.getMonth() &&
           a.getDate() === b.getDate();
  }

  function atTime(dateObj, hhmm) {
    var p = hhmm.split(":").map(Number);
    return new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), p[0], p[1], 0);
  }

  function addDays(d, n) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
  }

  // Split an interval that crosses midnight into per-day pieces.
  function splitByDay(start, end) {
    var pieces = [];
    var cur = start;
    while (!sameDate(cur, end) && cur < end) {
      var midnight = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate() + 1, 0, 0, 0);
      if (midnight >= end) break;
      pieces.push({ start: cur, end: midnight });
      cur = midnight;
    }
    if (cur < end) pieces.push({ start: cur, end: end });
    return pieces;
  }

  function mergeIntervals(list) {
    var sorted = list.slice().sort(function (a, b) { return a.start - b.start; });
    var merged = [];
    for (var i = 0; i < sorted.length; i++) {
      var iv = sorted[i];
      if (merged.length === 0 || iv.start > merged[merged.length - 1].end) {
        merged.push({ start: iv.start, end: iv.end });
      } else if (iv.end > merged[merged.length - 1].end) {
        merged[merged.length - 1].end = iv.end;
      }
    }
    return merged;
  }

  function computeFreeGaps(windowStart, windowEnd, busy, minGapMinutes) {
    var free = [];
    var cursor = windowStart;
    var sorted = busy.slice().sort(function (a, b) { return a.start - b.start; });

    for (var i = 0; i < sorted.length; i++) {
      var bStart = sorted[i].start < windowStart ? windowStart : sorted[i].start;
      var bEnd = sorted[i].end > windowEnd ? windowEnd : sorted[i].end;
      if (bEnd <= windowStart || bStart >= windowEnd) continue;

      if (bStart > cursor) addGap(free, cursor, bStart, minGapMinutes);
      if (bEnd > cursor) cursor = bEnd;
    }
    if (cursor < windowEnd) addGap(free, cursor, windowEnd, minGapMinutes);
    return free;
  }

  function addGap(free, s, e, minGapMinutes) {
    if ((e - s) / 60000 >= minGapMinutes) free.push({ start: s, end: e });
  }

  function fmtTime(d) {
    var h = d.getHours();
    var m = d.getMinutes();
    var ampm = h >= 12 ? "pm" : "am";
    h = h % 12; if (h === 0) h = 12;
    var mm = m < 10 ? "0" + m : "" + m;
    return h + ":" + mm + ampm;
  }

  var DAY_NAMES = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];

  // opt: { startDate:Date, endDate:Date, startTime:"10:00", endTime:"16:00",
  //        tzAbbrev:"CT", minGap:15, tentativeBusy:false, includeWeekends:false }
  // events: array from Graph (already in display zone)
  function buildSummary(events, opt) {
    // bucket busy intervals by day
    var busyByDay = {}; // key: yyyy-mm-dd -> [intervals]

    for (var i = 0; i < events.length; i++) {
      var ev = events[i];
      if (ev.isAllDay) continue;
      if (!isBusy(ev, opt.tentativeBusy)) continue;

      var s = parseFloating(ev.start.dateTime);
      var e = parseFloating(ev.end.dateTime);
      var pieces = splitByDay(s, e);
      for (var j = 0; j < pieces.length; j++) {
        var key = keyOf(pieces[j].start);
        (busyByDay[key] = busyByDay[key] || []).push(pieces[j]);
      }
    }

    for (var k in busyByDay) {
      if (busyByDay.hasOwnProperty(k)) busyByDay[k] = mergeIntervals(busyByDay[k]);
    }

    var lines = [];
    var day = new Date(opt.startDate.getFullYear(), opt.startDate.getMonth(), opt.startDate.getDate());
    var last = new Date(opt.endDate.getFullYear(), opt.endDate.getMonth(), opt.endDate.getDate());

    while (day <= last) {
      var dow = day.getDay();
      var isWeekend = (dow === 0 || dow === 6);
      if (opt.includeWeekends || !isWeekend) {
        var windowStart = atTime(day, opt.startTime);
        var windowEnd = atTime(day, opt.endTime);
        var busy = busyByDay[keyOf(day)] || [];
        var free = computeFreeGaps(windowStart, windowEnd, busy, opt.minGap);

        var label = DAY_NAMES[dow] + " " + (day.getMonth() + 1) + "/" + day.getDate();
        var slots;
        if (free.length === 0) {
          slots = "No availability";
        } else {
          slots = free.map(function (f) {
            return fmtTime(f.start) + "-" + fmtTime(f.end) + " " + opt.tzAbbrev;
          }).join(", ");
        }
        lines.push("\u2022 " + label + ": " + slots);
      }
      day = addDays(day, 1);
    }

    return lines.join("\n");
  }

  function keyOf(d) {
    return d.getFullYear() + "-" + (d.getMonth() + 1) + "-" + d.getDate();
  }

  return { buildSummary: buildSummary };
})();

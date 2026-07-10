// ============================================================
//  app.js — UI wiring, time zones, orchestration
// ============================================================

// IANA zone + short label pairs. Graph wants IANA ("America/Chicago");
// we show the friendly name and append the CT/ET/… label in output.
var ZONES = [
  { iana: "America/Chicago",    name: "Central Time (CT)",        abbrev: "CT"  },
  { iana: "America/New_York",   name: "Eastern Time (ET)",        abbrev: "ET"  },
  { iana: "America/Denver",     name: "Mountain Time (MT)",       abbrev: "MT"  },
  { iana: "America/Phoenix",    name: "Arizona (MST, no DST)",    abbrev: "MST" },
  { iana: "America/Los_Angeles",name: "Pacific Time (PT)",        abbrev: "PT"  },
  { iana: "America/Anchorage",  name: "Alaska Time (AKT)",        abbrev: "AKT" },
  { iana: "Pacific/Honolulu",   name: "Hawaii Time (HT)",         abbrev: "HT"  },
  { iana: "Europe/London",      name: "UK Time (GMT/BST)",        abbrev: "GMT" }
];

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    initUI();
  }
});

function el(id) { return document.getElementById(id); }

function initUI() {
  // populate zones
  var sel = el("timeZone");
  ZONES.forEach(function (z, idx) {
    var o = document.createElement("option");
    o.value = idx;
    o.textContent = z.name;
    sel.appendChild(o);
  });
  sel.value = 0; // Central default

  // default dates: current week Mon–Fri
  var today = new Date();
  var dow = (today.getDay() + 6) % 7; // days since Monday
  var monday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - dow);
  var friday = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + 4);
  el("startDate").value = toInputDate(monday);
  el("endDate").value = toInputDate(friday);

  // wire buttons
  el("btnSignIn").addEventListener("click", onSignIn);
  el("btnGenerate").addEventListener("click", onGenerate);
  el("btnCopyText").addEventListener("click", onCopyText);
  el("btnCopyHtml").addEventListener("click", onCopyHtml);

  // show correct area based on existing session
  if (AvailabilityAuth.getAccount()) {
    showMain();
  } else {
    showSignIn();
  }
}

function toInputDate(d) {
  var m = ("0" + (d.getMonth() + 1)).slice(-2);
  var day = ("0" + d.getDate()).slice(-2);
  return d.getFullYear() + "-" + m + "-" + day;
}

function fromInputDate(str) {
  var p = str.split("-").map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}

function showSignIn() {
  el("signin-area").classList.remove("hidden");
  el("main-area").classList.add("hidden");
}
function showMain() {
  el("signin-area").classList.add("hidden");
  el("main-area").classList.remove("hidden");
}

async function onSignIn() {
  try {
    setStatus("Signing in…");
    await AvailabilityAuth.signIn();
    setStatus("");
    showMain();
  } catch (e) {
    setStatus("Sign-in failed: " + e.message, true);
  }
}

async function onGenerate() {
  try {
    var startDate = fromInputDate(el("startDate").value);
    var endDate = fromInputDate(el("endDate").value);
    var startTime = el("startTime").value;   // "10:00"
    var endTime = el("endTime").value;       // "16:00"
    var zone = ZONES[parseInt(el("timeZone").value, 10)];
    var minGap = parseInt(el("minGap").value, 10) || 15;
    var tentativeBusy = el("tentativeBusy").checked;
    var includeWeekends = el("includeWeekends").checked;

    if (endDate < startDate) { setStatus("Stop Date must be on/after Start Date.", true); return; }
    if (endTime <= startTime) { setStatus("End Time must be after Start Time.", true); return; }

    setStatus("Reading calendar…");

    var token = await AvailabilityAuth.getToken();

    // Build query window: pad a day each side so TZ shifts/cross-midnight are safe.
    var qStart = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() - 1);
    var qEnd = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate() + 2);

    var events = await AvailabilityGraph.getCalendarView(
      token, qStart.toISOString(), qEnd.toISOString(), zone.iana);

    var summary = AvailabilityEngine.buildSummary(events, {
      startDate: startDate,
      endDate: endDate,
      startTime: startTime,
      endTime: endTime,
      tzAbbrev: zone.abbrev,
      minGap: minGap,
      tentativeBusy: tentativeBusy,
      includeWeekends: includeWeekends
    });

    el("output").value = summary;
    setStatus("Done.");
  } catch (e) {
    setStatus("Error: " + e.message, true);
  }
}

function onCopyText() {
  var out = el("output").value;
  if (!out) return;
  navigator.clipboard.writeText(out).then(function () {
    setStatus("Copied to clipboard.");
  });
}

function onCopyHtml() {
  var out = el("output").value;
  if (!out) return;
  var items = out.split("\n").filter(Boolean).map(function (l) {
    return "<li>" + escapeHtml(l.replace(/^\u2022\s*/, "")) + "</li>";
  });
  var html = "<ul>" + items.join("") + "</ul>";

  // Write both HTML and plain text flavors to the clipboard.
  var blobHtml = new Blob([html], { type: "text/html" });
  var blobText = new Blob([out], { type: "text/plain" });
  var item = new ClipboardItem({ "text/html": blobHtml, "text/plain": blobText });
  navigator.clipboard.write([item]).then(function () {
    setStatus("Copied as bulleted list.");
  }).catch(function () {
    // Fallback: plain text only
    navigator.clipboard.writeText(out);
    setStatus("Copied as text (HTML not supported here).");
  });
}

function escapeHtml(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function setStatus(msg, isError) {
  var s = el("status");
  s.textContent = msg || "";
  s.className = "status" + (isError ? " error" : "");
}

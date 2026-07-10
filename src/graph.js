// ============================================================
//  graph.js — pulls calendar events from Microsoft Graph
// ============================================================

var AvailabilityGraph = (function () {

  // Fetches events overlapping [startIso, endIso], expanded for recurrences,
  // returned in the requested IANA time zone via the Prefer header.
  // Uses /calendarView so recurring meetings are expanded to instances.
  async function getCalendarView(token, startIso, endIso, ianaTimeZone) {
    var results = [];
    var base = "https://graph.microsoft.com/v1.0/me/calendarView";
    var params =
      "?startDateTime=" + encodeURIComponent(startIso) +
      "&endDateTime=" + encodeURIComponent(endIso) +
      "&$select=subject,start,end,isAllDay,showAs,type" +
      "&$orderby=start/dateTime" +
      "&$top=100";

    var url = base + params;

    while (url) {
      var resp = await fetch(url, {
        headers: {
          "Authorization": "Bearer " + token,
          // Ask Graph to return start/end already converted to this zone.
          "Prefer": 'outlook.timezone="' + ianaTimeZone + '"'
        }
      });

      if (!resp.ok) {
        var text = await resp.text();
        throw new Error("Graph error " + resp.status + ": " + text);
      }

      var data = await resp.json();
      if (data.value) results = results.concat(data.value);

      // Follow pagination if present.
      url = data["@odata.nextLink"] || null;
    }

    return results;
  }

  return { getCalendarView: getCalendarView };
})();

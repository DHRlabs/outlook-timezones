/* Shared parsing and conversion logic. Used by read.js and commands.js.
   Plain JS, no dependencies. Offset math uses Intl only. */
(function (root, factory) {
  var api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.TZ = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // The one zone table. Written abbreviation (upper case) -> IANA region.
  // Regions, not fixed offsets: EST in July is still New York time (EDT).
  var ZONES = {
    EST: "America/New_York", EDT: "America/New_York", ET: "America/New_York", EASTERN: "America/New_York",
    CST: "America/Chicago", CDT: "America/Chicago", CT: "America/Chicago", CENTRAL: "America/Chicago",
    MST: "America/Denver", MDT: "America/Denver", MT: "America/Denver", MOUNTAIN: "America/Denver",
    PST: "America/Los_Angeles", PDT: "America/Los_Angeles", PT: "America/Los_Angeles", PACIFIC: "America/Los_Angeles",
    AZ: "America/Phoenix", ARIZONA: "America/Phoenix",
    AKST: "America/Anchorage", AKDT: "America/Anchorage", AKT: "America/Anchorage",
    HST: "Pacific/Honolulu",
    UTC: "UTC", GMT: "UTC",
    // World abbreviations with one meaning. These count only when written in capitals.
    CET: "Europe/Paris", CEST: "Europe/Paris", EET: "Europe/Athens", EEST: "Europe/Athens",
    MSK: "Europe/Moscow",
    JST: "Asia/Tokyo", KST: "Asia/Seoul", SGT: "Asia/Singapore", HKT: "Asia/Hong_Kong", PHT: "Asia/Manila",
    WIB: "Asia/Jakarta", WITA: "Asia/Makassar", PKT: "Asia/Karachi",
    AEST: "Australia/Sydney", AEDT: "Australia/Sydney", ACST: "Australia/Adelaide", ACDT: "Australia/Adelaide",
    AWST: "Australia/Perth", NZST: "Pacific/Auckland", NZDT: "Pacific/Auckland",
    ADT: "America/Halifax", NST: "America/St_Johns", NDT: "America/St_Johns", BRT: "America/Sao_Paulo"
  };
  // Table entries that also count in lower case ("4pm est"). Every other entry needs capitals.
  var ANY_CASE = /^(?:E|C|M|P)(?:ST|DT|T)$|^(?:AKST|AKDT|AKT|HST|UTC|GMT|AZ|EASTERN|CENTRAL|MOUNTAIN|PACIFIC|ARIZONA)$/;
  // Abbreviations with more than one meaning (IST: India, Israel, Irish). A time written with one is skipped.
  // WET and WEST are also ordinary words in all-caps text.
  var AMBIGUOUS = ["IST", "BST", "AST", "GST", "ART", "ECT", "SST", "WET", "WEST"];
  // Spelled-out world zone names, any case. Tested before the US names so "Central European Time" is never Chicago.
  var WORLD_NAMES = { CENTRAL: ["CET", "Europe/Paris"], EASTERN: ["EET", "Europe/Athens"] };

  // Labels the add-in generates itself, by region. Other regions fall back to Intl's short name.
  var LABELS = {
    "America/Phoenix": "AZ", "America/New_York": "ET", "America/Chicago": "CT",
    "America/Denver": "MT", "America/Los_Angeles": "PT", "America/Anchorage": "AKT",
    "Pacific/Honolulu": "HST", "UTC": "UTC"
  };

  // Compose targets, in output order.
  var TARGETS = [["CT", "America/Chicago"], ["ET", "America/New_York"]];

  var DAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  var MAX_BAR = 150; // NotificationMessageDetails.message maximum length

  // ---- Intl offset math ----
  var cache = {};
  function dtf(tz) {
    return cache[tz] || (cache[tz] = new Intl.DateTimeFormat("en-US", {
      timeZone: tz, hourCycle: "h23", year: "numeric", month: "numeric", day: "numeric",
      hour: "numeric", minute: "numeric", second: "numeric"
    }));
  }
  function parts(tz, ms) {
    var o = {};
    dtf(tz).formatToParts(new Date(ms)).forEach(function (p) { o[p.type] = p.value; });
    return { y: +o.year, m: +o.month, d: +o.day, h: +o.hour % 24, mi: +o.minute, s: +o.second };
  }
  function offsetMin(tz, ms) {
    ms = Math.floor(ms / 1000) * 1000;
    var p = parts(tz, ms);
    return (Date.UTC(p.y, p.m - 1, p.d, p.h, p.mi, p.s) - ms) / 60000;
  }
  // Wall time in tz -> instant. A wall time skipped by spring-forward resolves forward.
  function zonedToMs(tz, y, m, d, h, mi) {
    var guess = Date.UTC(y, m - 1, d, h, mi);
    var t1 = guess - offsetMin(tz, guess) * 60000;
    var t2 = guess - offsetMin(tz, t1) * 60000;
    if (t2 === t1) return t1;
    var p = parts(tz, t2);
    if (p.h === h && p.mi === mi) return t2;
    p = parts(tz, t1);
    return p.h === h && p.mi === mi ? t1 : Math.max(t1, t2);
  }
  function dow(y, m, d) { return new Date(Date.UTC(y, m - 1, d)).getUTCDay(); }

  function fmt(h, mi) {
    var suffix = h >= 12 ? "pm" : "am";
    var hh = h % 12 === 0 ? 12 : h % 12;
    return hh + (mi ? ":" + (mi < 10 ? "0" : "") + mi : "") + suffix;
  }

  function labelFor(tz, ms) {
    if (LABELS[tz]) return LABELS[tz];
    var p = new Intl.DateTimeFormat("en-US", { timeZone: tz, timeZoneName: "short" })
      .formatToParts(new Date(ms)).filter(function (x) { return x.type === "timeZoneName"; })[0];
    return p ? p.value : tz;
  }

  // ---- Finding times ----
  var MER = "am|pm|a\\.m\\.|p\\.m\\.";
  // Zone after a time. Separator: none or one space, two spaces, "," ";" ":" "/" dashes, or "(" "[" with a close.
  // Groups: 1 open bracket, 2 world name phrase, 3 word, 4 "(Standard|Daylight) Time (zone)" tail.
  var ZONE_AFTER = /^(?:[ \u00a0]{2}|[ \u00a0]?[,;:\/\-\u2013\u2014][ \u00a0]{0,2}|[ \u00a0]?([(\[])|[ \u00a0]?)(?:((central|eastern)[ \u00a0]+european(?:[ \u00a0]+summer)?[ \u00a0]+time(?:[ \u00a0]+zone)?)\b|([A-Za-z]+)((?:[ \u00a0]+(?:standard|daylight))?[ \u00a0]+time(?:[ \u00a0]+zone)?\b)?)/i;
  // A spelled-out name (Eastern, Central...) with no "Time" is a zone at the end of a clause, before
  // punctuation (or the bracket this add-in wrote), or before a function word. Before any other word
  // ("Central Park") it is not read, and the time is left untouched (AGENTS.md zone invariant).
  var NAME_END = /^[ \u00a0]*(?:[.,;:!?()\[\]\/\-\u2013\u2014\r\n]|$|(?:or|and|if|with|for|on|at|is|works|but|so|then|to|please|today|tomorrow|tonight)(?![A-Za-z]))/i;
  var NEAR_TOKEN = /^[^A-Za-z0-9\r\n]{0,4}([A-Za-z]+)/;
  var SKIP = { skip: true };

  // Classify a written word as a zone: region, SKIP (ambiguous), or null (just a word).
  function zoneWord(tok) {
    var u = tok.toUpperCase();
    if (tok === u && AMBIGUOUS.indexOf(u) >= 0) return SKIP;
    return ZONES[u] && (tok === u || ANY_CASE.test(u)) ? ZONES[u] : null;
  }
  // Zone written right after a time: {len, tok, region}, SKIP, or null.
  function zoneAfter(after) {
    var m = ZONE_AFTER.exec(after), world = m && m[2] && WORLD_NAMES[m[3].toUpperCase()];
    var tok = world ? world[0] : m && m[4].toUpperCase(), z = world ? world[1] : m && zoneWord(m[4]), len;
    if (z === SKIP) return SKIP;
    if (z) {
      len = m[0].length;
      if (!world && m[4].length > 4 && !m[5] && !NAME_END.test(after.slice(len))) return SKIP; // "Central Park"
      if (z && m[1]) {
        if (after[len] !== (m[1] === "(" ? ")" : "]")) z = null;
        else len++;
      }
      if (z && /^[+-]\d/.test(after.slice(len))) return SKIP; // offset label like GMT+2
      if (z) return { len: len, tok: tok, region: z };
    }
    // Not cleanly parsed, but a zone word starts right after the time: skip it whole.
    var n = NEAR_TOKEN.exec(after);
    return n && zoneWord(n[1]) ? SKIP : null;
  }
  var LB = "(?<![\\w$.:/@#-])";
  var RE_SRC = [
    // range: 2-3pm, 2pm-3pm, 2:30 to 3:30pm, optional zone
    LB + "(?<rh1>\\d{1,2})(?::(?<rm1>\\d{2}))?(?:\\s?(?<rmer1>" + MER + "))?\\s*(?:[-\\u2013\\u2014]|\\bto\\b)\\s*" +
      "(?<rh2>\\d{1,2})(?::(?<rm2>\\d{2}))?\\s?(?<rmer2>" + MER + ")(?![A-Za-z0-9])",
    // single: 4pm, 4 pm, 4:30 PM, 4p, optional zone
    LB + "(?<sh>\\d{1,2})(?::(?<sm>\\d{2}))?(?:\\s?(?<smer>" + MER + ")|(?<smer1>[ap]))(?![A-Za-z0-9])",
    // 24-hour and noon / midnight: zone required (checked in findTimes)
    LB + "(?<hh>\\d{1,2}):(?<hm>[0-5]\\d)(?![\\w:])",
    "(?<![\\w])(?<nm>noon|midnight)\\b"
  ].join("|");

  var DAYRE = "sunday|monday|tuesday|wednesday|thursday|friday|saturday|sun|mon|tues|tue|wed|thurs|thur|thu|fri|sat";
  var DAY_BEFORE = new RegExp("\\b(" + DAYRE + ")\\b\\.?,?[ \\u00a0]*(?:(?:at|@|from|by|around|on)\\b[ \\u00a0]*)?$", "i");
  var DAY_AFTER = new RegExp("^[ \\u00a0]*,?[ \\u00a0]*(?:on[ \\u00a0]+)?(" + DAYRE + ")\\b(?![ \\u00a0]*,?[ \\u00a0]*(?:(?:at|@|from|by|around)[ \\u00a0]*)?\\d)", "i");
  var DAYIDX = { sun: 0, mon: 1, tue: 2, wed: 3, thu: 4, fri: 5, sat: 6 };

  // Not a clock time: the word just before is a unit or locator ("Suite 4 PM", "Ext 4pm").
  var UNIT_BEFORE = /(?:(?:^|[^A-Za-z])(?:suite|ste|ext|apt|room|rm|unit|no|box|floor|fl|gate|page|pg)\.?|#)[ \u00a0]*$/i;
  // Half of a range whose other end has no meridiem ("2pm-3", "noon - 1pm"): skip the whole range.
  var RANGE_BEFORE = /(?:\d|\bnoon|\bmidnight)[ \u00a0]*(?:[-\u2013\u2014]|\bto)[ \u00a0]*$|\d[ \u00a0]?[ap]\.?m\.?[ \u00a0]*(?:[-\u2013\u2014]|\bto)[ \u00a0]*$/i;
  var RANGE_AFTER = /^[ \u00a0]*(?:[-\u2013\u2014]|to\b)[ \u00a0]*(?:\d{1,2}(?::\d\d)?(?![\d:])(?![ \u00a0]?[ap](?:\.?m|\b))|noon\b|midnight\b)/i;

  // Output this add-in already wrote after a time: an optional label of any shape, then "[3pm CT / 4pm ET]".
  var T1 = "\\d{1,2}(?::\\d{2})?\\s?(?:am|pm)";
  var ITEM = T1 + "(?:\\s?-\\s?" + T1 + ")?\\s+(?:" + TARGETS.map(function (t) { return t[0]; }).join("|") + ")(?:\\s+[+-]\\d+d)?";
  var OWN = new RegExp("^(?:[ \\u00a0]*[^\\s\\[\\]]{1,12})?[ \\u00a0]*\\[\\s*" + ITEM + "(?:\\s*\\/\\s*" + ITEM + ")*\\s*\\]", "i");

  function to24(h, merChar) { return (h % 12) + (merChar === "p" ? 12 : 0); }

  // Short forms that are also words (sat, sun, mon, wed) count only when capitalized.
  function isDay(m) { return m && !(/^(sat|sun|mon|wed)$/.test(m[1])); }
  function dayWord(text, s, e) {
    var m = DAY_BEFORE.exec(text.slice(Math.max(0, s - 30), s));
    if (!isDay(m)) m = DAY_AFTER.exec(text.slice(e, e + 30));
    return isDay(m) ? DAYIDX[m[1].slice(0, 3).toLowerCase()] : null;
  }

  // Returns matches: {s, e, text, h1, mi1, h2, mi2, zoneTok, hasZone, dayIdx}
  // Times with an invalid shape are consumed and dropped, never half-converted.
  function findTimes(text) {
    var re = new RegExp(RE_SRC, "gi"), out = [], m;
    while ((m = re.exec(text))) {
      var g = m.groups, r = { s: m.index, e: m.index + m[0].length, text: m[0], h2: null, mi2: null };
      var z = zoneAfter(text.slice(r.e, r.e + 60));
      if (z === SKIP || (!z && g.sh === undefined && g.rh1 === undefined)) continue;
      if (z) { r.e += z.len; r.text = text.slice(r.s, r.e); }
      var before = text.slice(Math.max(0, r.s - 30), r.s), after = text.slice(r.e, r.e + 40);
      var tokS = text.slice(0, r.s).search(/\S*$/), tokE = r.s + text.slice(r.s).search(/\s|$/);
      if (/[@+]/.test(text.slice(tokS, tokE)) || UNIT_BEFORE.test(before)) continue;
      if (RANGE_BEFORE.test(before) || RANGE_AFTER.test(after)) continue;
      if (g.rh1 !== undefined) {
        var h1 = +g.rh1, h2 = +g.rh2, m1 = g.rm1 ? +g.rm1 : 0, m2 = g.rm2 ? +g.rm2 : 0;
        if (h1 < 1 || h1 > 12 || h2 < 1 || h2 > 12 || m1 > 59 || m2 > 59) continue;
        var c2 = g.rmer2.toLowerCase()[0];
        var c1 = g.rmer1 ? g.rmer1.toLowerCase()[0] : c2;
        r.h1 = to24(h1, c1); r.mi1 = m1; r.h2 = to24(h2, c2); r.mi2 = m2;
        if (!g.rmer1 && r.h1 * 60 + r.mi1 > r.h2 * 60 + r.mi2) { r.h1 = to24(h1, c2 === "p" ? "a" : "p"); }
      } else if (g.sh !== undefined) {
        var h = +g.sh, mi = g.sm ? +g.sm : 0;
        if (h < 1 || h > 12 || mi > 59) continue;
        r.h1 = to24(h, (g.smer || g.smer1).toLowerCase()[0]); r.mi1 = mi;
      } else if (g.hh !== undefined) {
        // 24-hour is unambiguous only for 13-23 or a leading zero (09:00). Else skip: never guess am/pm.
        var hs = g.hh, hv = +hs;
        if (!(hv >= 13 && hv <= 23) && !(hs.length === 2 && hs[0] === "0")) continue;
        r.h1 = hv; r.mi1 = +g.hm;
      } else {
        r.h1 = g.nm.toLowerCase() === "noon" ? 12 : 0; r.mi1 = 0;
      }
      r.hasZone = !!z;
      if (z) { r.zoneTok = z.tok; r.zoneRegion = z.region; }
      r.dayIdx = dayWord(text, r.s, r.e);
      out.push(r);
    }
    return out;
  }

  // Convert match m written in src zone into dst zone, on the message date.
  function conv(m, src, dst, refMs) {
    var p = parts(src, refMs), add = 0;
    if (m.dayIdx !== null) add = (m.dayIdx - dow(p.y, p.m, p.d) + 7) % 7;
    var base = new Date(Date.UTC(p.y, p.m - 1, p.d + add));
    var y = base.getUTCFullYear(), mo = base.getUTCMonth() + 1, d = base.getUTCDate();
    var t1 = zonedToMs(src, y, mo, d, m.h1, m.mi1);
    var q1 = parts(dst, t1);
    var text = fmt(q1.h, q1.mi);
    if (m.h2 !== null) {
      var q2 = parts(dst, zonedToMs(src, y, mo, d, m.h2, m.mi2));
      text += "-" + fmt(q2.h, q2.mi);
    }
    var delta = Math.round((Date.UTC(q1.y, q1.m - 1, q1.d) - Date.UTC(y, mo - 1, d)) / 864e5);
    return {
      text: text, delta: delta, t: t1,
      sameOffset: offsetMin(src, t1) === offsetMin(dst, t1),
      srcText: fmt(m.h1, m.mi1) + (m.h2 !== null ? "-" + fmt(m.h2, m.mi2) : ""),
      dayIdx: m.dayIdx === null ? null : (m.dayIdx + delta + 7) % 7
    };
  }

  function dayShift(delta) { return delta ? " " + (delta > 0 ? "+" : "-") + Math.abs(delta) + "d" : ""; }

  // ---- Read side ----
  function firstName(displayName) {
    var n = String(displayName || "").replace(/"/g, "").trim().replace(/^'+|'+$/g, "");
    if (!n || n.indexOf("@") >= 0) return "";
    if (n.indexOf(",") >= 0) n = n.split(",")[1].trim();
    return (n.split(/\s+/)[0] || "").slice(0, 30);
  }

  // Cut the quoted thread and signature-style tails so old messages' times are ignored.
  function stripQuoted(text) {
    var lines = String(text || "").split(/\r?\n/), out = [];
    for (var i = 0; i < lines.length; i++) {
      var l = lines[i];
      if (/^\s*(-{2,}\s*original message|(from|de|von|da)\s?:\s|on .{5,200}wrote:|_{10,}|sent from my |get outlook for )/i.test(l)) break;
      if (/^\s*on\s.{5,200}$/i.test(l) && /^.{0,200}wrote:\s*$/i.test(lines[i + 1] || "")) break;
      if (/^\s*>/.test(l)) continue;
      out.push(l);
    }
    return out.join("\n");
  }

  // o: {displayName, refMs, deviceTz}. Returns bar text or null.
  function readerBar(text, o) {
    var items = [], seen = {};
    findTimes(stripQuoted(text)).forEach(function (m) {
      if (!m.hasZone || !m.zoneRegion) return; // no zone written: show nothing, never guess
      var c = conv(m, m.zoneRegion, o.deviceTz, o.refMs);
      if (c.sameOffset) return;
      var key = (m.dayIdx === null ? "" : m.dayIdx) + "|" + c.srcText + "|" + m.zoneTok;
      if (seen[key]) return;
      seen[key] = 1;
      items.push({ m: m, c: c });
    });
    if (!items.length) return null;
    var you = labelFor(o.deviceTz, o.refMs), name = firstName(o.displayName);
    var segs = items.map(function (it, i) {
      var m = it.m, c = it.c, first = i === 0;
      var left = (m.dayIdx === null ? "" : DAYS[m.dayIdx] + " ") + m.text.replace(/[\s\u00a0]+/g, " ") + (first && name ? " " + name : "");
      var right = (m.dayIdx !== null && c.delta ? DAYS[c.dayIdx] + " " : "") + c.text + " " + you +
        (first ? " you" : "") + (m.dayIdx === null ? dayShift(c.delta) : "");
      return left + " | " + right;
    });
    // Drop later times rather than truncate mid-time, and say how many were dropped.
    for (var k = segs.length; k > 0; k--) {
      var bar = segs.slice(0, k).join(" · ") + (k < segs.length ? " +" + (segs.length - k) + " more" : "");
      if (bar.length <= MAX_BAR) return bar;
    }
    return null;
  }

  // ---- Compose side ----
  // o: {deviceTz, nowMs, before, after}. Returns the text with each clock time rewritten.
  // before/after: neighboring inline text in the same block (HTML splits text across spans);
  // a time that runs into a neighbor ("4pm" + "EST") is left untouched.
  function rewriteText(text, o) {
    var full = (o.before || "") + text + (o.after || ""), a = (o.before || "").length, z = a + text.length;
    var spans = [], re = /\[[^\[\]]*\]/g, b;
    while ((b = re.exec(full))) spans.push([b.index, b.index + b[0].length]);
    var out = "", pos = a;
    findTimes(full).forEach(function (m) {
      if (m.s < a || m.e > z) return;
      if (spans.some(function (sp) { return m.s >= sp[0] && m.s < sp[1]; })) return; // inside any bracket
      if (OWN.test(full.slice(m.e))) return; // already done by this add-in
      var src = m.hasZone ? m.zoneRegion : o.deviceTz;
      if (!src) return;
      var items = [];
      TARGETS.forEach(function (t) {
        var c = conv(m, src, t[1], o.nowMs);
        if (c.sameOffset) return;                                   // same as the time's own zone
        if (offsetMin(o.deviceTz, c.t) === offsetMin(t[1], c.t)) return; // same as the writer's zone
        items.push(c.text + " " + t[0] + dayShift(c.delta));
      });
      var add = m.hasZone ? "" : " " + labelFor(o.deviceTz, o.nowMs);
      var br = items.length ? " [" + items.join(" / ") + "]" : "";
      out += full.slice(pos, m.e) + add + br;
      pos = m.e;
    });
    return out + full.slice(pos, z);
  }

  return {
    ZONES: ZONES, LABELS: LABELS, TARGETS: TARGETS, MAX_BAR: MAX_BAR,
    findTimes: findTimes, readerBar: readerBar, rewriteText: rewriteText,
    firstName: firstName, stripQuoted: stripQuoted
  };
});

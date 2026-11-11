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
  // A zone string taken from email text (an offset like "-05:00") or from a caller's people list
  // may be nonsense. Every Intl call goes through a zone that passed here, so nothing throws.
  var zoneOk = {};
  function validZone(tz) {
    if (!tz || typeof tz !== "string") return false;
    if (zoneOk[tz] !== undefined) return zoneOk[tz];
    try { new Intl.DateTimeFormat("en-US", { timeZone: tz }); zoneOk[tz] = true; }
    catch (e) { zoneOk[tz] = false; }
    return zoneOk[tz];
  }

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
  // am, pm, a.m., p.m., a.m (any case). A trailing dot counts only in the dotted form, so the
  // sentence period in "do 2pm." stays outside the time. The single letters "4p"/"4a" are separate.
  var MER = "[ap]\\.m\\.?|[ap]m";
  // Zone words, longest first, so a zone glued to the meridiem ("4pmET", "4PMEST") still ends the time.
  var ZONE_KEYS = Object.keys(ZONES).sort(function (a, b) { return b.length - a.length; }).join("|");
  // End of a meridiem: a non-word character, or a zone word written with no space.
  var MER_END = "(?:(?![A-Za-z0-9])|(?=(?:" + ZONE_KEYS + ")(?![A-Za-z])))";
  // Zone after a time. Separator: none or one space, two spaces, "," ";" ":" "/" dashes, or "(" "[" with a close.
  // Groups: 1 open bracket, 2 world name phrase, 3 word, 4 "(Standard|Daylight) Time (zone)" tail.
  var ZONE_AFTER = /^(?:[ \u00a0]{2}|[ \u00a0]?[,;:\/\-\u2013\u2014][ \u00a0]{0,2}|[ \u00a0]?([(\[])|[ \u00a0]?)(?:((central|eastern)[ \u00a0]+european(?:[ \u00a0]+summer)?[ \u00a0]+time(?:[ \u00a0]+zone)?)\b|([A-Za-z]+)((?:[ \u00a0]+(?:standard|daylight))?[ \u00a0]+time(?:[ \u00a0]+zone)?\b)?)/i;
  // A spelled-out name (Eastern, Central...) with no "Time" is a zone at the end of a clause, before
  // punctuation (or the bracket this add-in wrote), or before a function word. Before any other word
  // ("Central Park") it is not read, and the time is left untouched (AGENTS.md zone invariant).
  var NAME_END = /^[ \u00a0]*(?:[.,;:!?()\[\]\/\-\u2013\u2014\r\n]|$|(?:or|and|if|with|for|on|at|is|works|but|so|then|to|please|today|tomorrow|tonight)(?![A-Za-z]))/i;
  var NEAR_TOKEN = /^[^A-Za-z0-9\r\n]{0,4}([A-Za-z]+)/;
  var SKIP = { skip: true };
  // Function words that can close a clause after a spelled-out zone name or a single-letter meridiem.
  var FUNC_WORD = "or|and|if|with|for|on|at|is|works|but|so|then|to|please|today|tomorrow|tonight";

  // Classify a written word as a zone: region, SKIP (ambiguous), or null (just a word).
  function zoneWord(tok) {
    var u = tok.toUpperCase();
    if (tok === u && AMBIGUOUS.indexOf(u) >= 0) return SKIP;
    return ZONES[u] && (tok === u || ANY_CASE.test(u)) ? ZONES[u] : null;
  }
  // "UTC-5", "GMT+1", "UTC-07:00": an exact offset, written as an Intl offset zone.
  // The minus may be a hyphen or a Unicode minus sign (U+2212), which mail clients insert.
  var OFFSET_AFTER = /^[ \u00a0]?([+\-\u2212])(\d{1,2})(?::?([0-5]\d))?(?![\d:])/;
  function offsetZone(sign, hh, mm) {
    if (+hh > 14) return null;
    var z = (sign === "+" ? "+" : "-") + (hh.length < 2 ? "0" : "") + hh + ":" + (mm || "00");
    return validZone(z) ? z : null;
  }

  // Zone written right after a time: {len, tok, region}, SKIP, or null.
  function zoneAfter(after) {
    var m = ZONE_AFTER.exec(after), world = m && m[2] && WORLD_NAMES[m[3].toUpperCase()];
    var tok = world ? world[0] : m && m[4].toUpperCase(), z = world ? world[1] : m && zoneWord(m[4]), len;
    if (z === SKIP) return SKIP;
    if (z) {
      len = m[0].length;
      if (!world && m[4].length > 4 && !m[5] && !NAME_END.test(after.slice(len))) return SKIP; // "Central Park"
      // An offset written on UTC or GMT is the zone: convert from that exact offset.
      if (!world && !m[5] && (tok === "UTC" || tok === "GMT")) {
        var off = OFFSET_AFTER.exec(after.slice(len));
        var reg = off && offsetZone(off[1], off[2], off[3]);
        if (off && !reg) return SKIP;
        // A sign that is not a readable offset ("UTC++5", "UTC+-5") is still an offset the writer
        // meant: the time is skipped whole rather than read as plain UTC.
        if (!off && /^[ \u00a0]?[+\-\u2212]/.test(after.slice(len))) return SKIP;
        if (reg) { z = reg; tok = tok + off[0].replace(/[ \u00a0]/g, ""); len += off[0].length; }
      }
      if (z && m[1]) {
        if (after[len] !== (m[1] === "(" ? ")" : "]")) z = null;
        else len++;
      }
      if (z && /^[+\-\u2212]\d/.test(after.slice(len))) return SKIP; // offset label like GMT+2
      if (z) return { len: len, tok: tok, region: z };
    }
    // Not cleanly parsed, but a zone word starts right after the time: skip it whole.
    var n = NEAR_TOKEN.exec(after);
    return n && zoneWord(n[1]) ? SKIP : null;
  }
  var LB = "(?<![\\w$.:/@#-])";
  // A single-letter meridiem ("4p", "4:30a") counts only when it is glued to the number and is
  // followed by a space and a zone word, the end of a clause, or punctuation. A space before the
  // letter would turn prose into a time ("3 a day", "4 p-values"), and a hyphen-word is never a
  // clause end ("4p-values"), so both are left out.
  var SMER1_END = "(?=[ \\u00a0](?:" + ZONE_KEYS + ")(?![A-Za-z])" +
    "|[ \\u00a0](?:" + FUNC_WORD + ")(?![A-Za-z])" +
    "|[.,;:!?)\\]\\r\\n]|$)";
  // Range separator. "and" counts only after "between" (checked in findTimes).
  var RSEP = "(?:[-\\u2013\\u2014]|\\bto\\b|\\band\\b)";
  var RE_PARTS = [
    // range: 2-3pm, 2pm-3pm, 2:30 to 3:30pm, 2.30-3.30pm, optional zone
    LB + "(?<rh1>\\d{1,2})(?:[:.](?<rm1>\\d{2})(?!\\d))?(?:[ \\u00a0]?(?<rmer1>" + MER + "))?[ \\u00a0]*" + RSEP + "[ \\u00a0]*" +
      "(?<rh2>\\d{1,2})(?:[:.](?<rm2>\\d{2})(?!\\d))?[ \\u00a0]?(?<rmer2>" + MER + ")" + MER_END,
    // single: 4pm, 4 pm, 4:30 PM, 4.30pm, 4p, 4pmET, optional zone
    LB + "(?<sh>\\d{1,2})(?:[:.](?<sm>\\d{2})(?!\\d))?(?:[ \\u00a0]?(?<smer>" + MER + ")" + MER_END +
      "|(?<smer1>[ap])" + SMER1_END + ")",
    // 24-hour clock: zone required (checked in findTimes). Dot minutes need a written meridiem,
    // so a decimal with a unit after it ("18.40 MT", "Revenue grew 14.30 CT") is not a time.
    LB + "(?<hh>\\d{1,2}):(?<hm>[0-5]\\d)(?![\\w:.])",
    // military with a unit word: 1600 hrs, 0930h. Zone required (checked in findTimes).
    LB + "(?<mh>\\d{3,4})[ \\u00a0]?(?<munit>hrs?\\.?|hours|h)(?![A-Za-z])",
    "(?<![\\w])(?<nm>noon|midnight)\\b"
  ];
  var RE_SRC = RE_PARTS.join("|");
  // The same finder with the range branch removed. Used to re-read the span of a range that was
  // rejected ("step 2pm to 3pm", "10am and 2pm"), so every real time inside it is still found.
  var RE_NORANGE = RE_PARTS.slice(1).join("|");

  var DAYRE = "sunday|monday|tuesday|wednesday|thursday|friday|saturday|sun|mon|tues|tue|wed|thurs|thur|thu|fri|sat";
  var DAY_BEFORE = new RegExp("\\b(" + DAYRE + ")\\b\\.?,?[ \\u00a0]*(?:(?:at|@|from|by|around|on)\\b[ \\u00a0]*)?$", "i");
  var DAY_AFTER = new RegExp("^[ \\u00a0]*,?[ \\u00a0]*(?:on[ \\u00a0]+)?(" + DAYRE + ")\\b(?![ \\u00a0]*,?[ \\u00a0]*(?:(?:at|@|from|by|around)[ \\u00a0]*)?\\d)", "i");
  var DAYIDX = { sun: 0, mon: 1, tue: 2, wed: 3, thu: 4, fri: 5, sat: 6 };

  // Not a clock time: the word just before is a unit or locator ("Suite 4 PM", "Ext 4pm").
  var UNIT_BEFORE = /(?:(?:^|[^A-Za-z])(?:suite|ste|ext|apt|room|rm|unit|no|box|floor|fl|gate|page|pg)\.?|#)[ \u00a0]*$/i;
  // Half of a range whose other end has no meridiem ("2pm-3", "noon - 1pm"): skip the whole range.
  var RANGE_BEFORE = /(?:\d|\bnoon|\bmidnight)[ \u00a0]*(?:[-\u2013\u2014]|\bto)[ \u00a0]*$|\d[ \u00a0]?[ap]\.?m\.?[ \u00a0]*(?:[-\u2013\u2014]|\bto)[ \u00a0]*$/i;
  var RANGE_AFTER = /^[ \u00a0]*(?:[-\u2013\u2014]|to\b)[ \u00a0]*(?:\d{1,2}(?::\d\d)?(?![\d:])(?![ \u00a0]?[ap](?:\.?m|\b))|noon\b|midnight\b)/i;
  // "step 3 to 5pm", "option 2-3pm": the first number counts something else, so only the second is a time.
  var RANGE_UNIT_BEFORE = /(?:^|[^A-Za-z])(?:step|steps|phase|part|section|chapter|item|items|line|row|level|version|option|round|grade|tier|week|weeks|day|days|question|slide|page|pages|point|rev|build|tab|column|col)\.?[ \u00a0]*$/i;
  // "and" is a range separator only after "between 2 and 3pm".
  var BETWEEN_BEFORE = /\bbetween[ \u00a0]*$/i;

  // Output this add-in already wrote after a time: an optional label of any shape, then "[3pm CT / 4pm ET]".
  var T1 = "\\d{1,2}(?::\\d{2})?\\s?(?:am|pm)";
  var ITEM = T1 + "(?:\\s?-\\s?" + T1 + ")?\\s+(?:" + TARGETS.map(function (t) { return t[0]; }).join("|") + ")(?:\\s+[+-]\\d+d)?";
  var OWN = new RegExp("^(?:[ \\u00a0]*[^\\s\\[\\]]{1,12})?[ \\u00a0]*\\[\\s*" + ITEM + "(?:\\s*\\/\\s*" + ITEM + ")*\\s*\\]", "i");
  // The same shape, anchored so a bracket group can be recognized where it starts.
  var OWN_AT = new RegExp("^\\[\\s*" + ITEM + "(?:\\s*\\/\\s*" + ITEM + ")*\\s*\\]", "i");
  // Spans of bracket groups this add-in wrote. Only these hide a time; "see [notes: 4pm ET]" does not.
  function ownSpans(text) {
    var spans = [], re = /\[[^\[\]]*\]/g, b;
    while ((b = re.exec(text))) {
      if (OWN_AT.test(b[0])) spans.push([b.index, b.index + b[0].length]);
    }
    return spans;
  }
  function inOwnBracket(spans, i) {
    return spans.some(function (sp) { return i >= sp[0] && i < sp[1]; });
  }

  function to24(h, merChar) { return (h % 12) + (merChar === "p" ? 12 : 0); }

  // Short forms that are also words (sat, sun, mon, wed) count only when capitalized.
  // Full names (thursday, monday) count in any case: no English word collides with them.
  function isDay(m) { return m && !(/^(sat|sun|mon|wed)$/.test(m[1])); }
  function dayWord(text, s, e) {
    var m = DAY_BEFORE.exec(text.slice(Math.max(0, s - 30), s));
    if (!isDay(m)) m = DAY_AFTER.exec(text.slice(e, e + 30));
    return isDay(m) ? DAYIDX[m[1].slice(0, 3).toLowerCase()] : null;
  }

  // Returns matches: {s, e, text, h1, mi1, h2, mi2, zoneTok, hasZone, dayIdx}
  // Times with an invalid shape are consumed and dropped, never half-converted.
  // refMs: the message or draft date, used to tell whether a "4pm ET / 1pm PT" pair is one instant.
  function findTimes(text, refMs) {
    var re = new RegExp(RE_SRC, "gid"), out = [], m;
    var refForPair = refMs || Date.now();
    while ((m = re.exec(text))) {
      var g0 = m.groups, ms = m.index, me = ms + m[0].length;
      if (g0.rh1 !== undefined) {
        // The opening number counts something else ("step 3 to 5pm"), or "and" with no "between":
        // not a range. A written meridiem on the opening number ("step 2pm to 3pm") makes it a real
        // time, so the unit word does not reject it.
        var before0 = text.slice(Math.max(0, ms - 30), ms);
        var sepAnd = /\band\b/i.test(m[0]);
        var unitRej = g0.rmer1 === undefined && RANGE_UNIT_BEFORE.test(before0);
        if (unitRej || (sepAnd && !BETWEEN_BEFORE.test(before0))) {
          // Read every real time before the closing number with the range branch removed, then
          // resume at the closing number itself, so it can still open a range of its own
          // ("4pm and 2pm to 6pm") and no time on either side of the separator is lost.
          var cut = m.indices.groups.rh2[0];
          var sub = new RegExp(RE_NORANGE, "gi"), sm;
          sub.lastIndex = ms;
          while ((sm = sub.exec(text)) && sm.index < cut) handle(sm);
          re.lastIndex = cut;
          continue;
        }
      }
      handle(m);
    }
    function handle(m) {
      var g = m.groups, r = { s: m.index, e: m.index + m[0].length, text: m[0], h2: null, mi2: null };
      var before = text.slice(Math.max(0, r.s - 30), r.s);
      var z = zoneAfter(text.slice(r.e, r.e + 60));
      if (z === SKIP || (!z && g.sh === undefined && g.rh1 === undefined)) return;
      if (z) { r.e += z.len; r.text = text.slice(r.s, r.e); }
      var after = text.slice(r.e, r.e + 40);
      // The word the time sits in, for the glued-token guard below. A slash ends the word the same
      // way a space does, so an offset on the time before it ("4pm UTC+1/noon PT") never reaches
      // across the slash and hides the time after it.
      var tokS = text.slice(0, r.s).search(/[^\s\/]*$/), tokE = r.s + text.slice(r.s).search(/[\s\/]|$/);
      if (/[@+]/.test(text.slice(tokS, tokE)) || UNIT_BEFORE.test(before)) return;
      if (RANGE_BEFORE.test(before) || RANGE_AFTER.test(after)) return;
      if (g.rh1 !== undefined) {
        var h1 = +g.rh1, h2 = +g.rh2, m1 = g.rm1 ? +g.rm1 : 0, m2 = g.rm2 ? +g.rm2 : 0;
        if (h1 < 1 || h1 > 12 || h2 < 1 || h2 > 12 || m1 > 59 || m2 > 59) return;
        var c2 = g.rmer2.toLowerCase()[0];
        var c1 = g.rmer1 ? g.rmer1.toLowerCase()[0] : c2;
        r.h1 = to24(h1, c1); r.mi1 = m1; r.h2 = to24(h2, c2); r.mi2 = m2;
        if (!g.rmer1 && r.h1 * 60 + r.mi1 > r.h2 * 60 + r.mi2) { r.h1 = to24(h1, c2 === "p" ? "a" : "p"); }
      } else if (g.sh !== undefined) {
        var h = +g.sh, mi = g.sm ? +g.sm : 0;
        if (h < 1 || h > 12 || mi > 59) return;
        r.h1 = to24(h, (g.smer || g.smer1).toLowerCase()[0]); r.mi1 = mi;
      } else if (g.hh !== undefined) {
        // 24-hour is unambiguous only for 13-23 or a leading zero (09:00). Else skip: never guess am/pm.
        var hs = g.hh, hv = +hs;
        if (!(hv >= 13 && hv <= 23) && !(hs.length === 2 && hs[0] === "0")) return;
        r.h1 = hv; r.mi1 = +g.hm;
      } else if (g.mh !== undefined) {
        // "1600 hrs": the unit word makes a 24-hour clock plain.
        var ms4 = g.mh.length === 3 ? "0" + g.mh : g.mh, mhv = +ms4.slice(0, 2), mmv = +ms4.slice(2);
        if (mhv > 23 || mmv > 59) return;
        r.h1 = mhv; r.mi1 = mmv;
      } else {
        r.h1 = g.nm.toLowerCase() === "noon" ? 12 : 0; r.mi1 = 0;
      }
      r.hasZone = !!z;
      if (z) { r.zoneTok = z.tok; r.zoneRegion = z.region; }
      r.dayIdx = dayWord(text, r.s, r.e);
      out.push(r);
    }
    // "4pm ET / 1pm PT": the writer already gave both zones for ONE instant, so it is one time
    // written twice. Two different instants ("Options 2pm ET / 3pm CT") are two times, and a bare
    // second side ("4pm ET / 1pm") is neither: it sits next to a zoned time, so it is left alone.
    for (var i = 0; i < out.length; i++) {
      var a = out[i], b = out[i + 1];
      if (!a.hasZone) continue;
      if (b && /^[ \u00a0]*\/[ \u00a0]*$/.test(text.slice(a.e, b.s))) {
        if (!b.hasZone) continue; // a bare second side: handled by the nextToZoned pass below
        if (sameInstant(a, b, refForPair)) { a.pairHead = true; b.pairTail = true; }
        continue;
      }
      // With no space ("4pm ET/3pm CT") the slash blocks the left boundary, so the second side is
      // never a match of its own. Read it straight from the text to compare the two instants.
      var tail = pairTail(text.slice(a.e), a);
      if (tail && sameInstant(a, tail, refForPair)) a.pairHead = true;
    }
    // A time with no zone of its own joined directly to a time that has a zone word is the other
    // half of the writer's own pair, so it is left untouched. The joiner is a slash or one of the
    // choice words (or, and, &, and/or), with at most one filler word (at, maybe, either) between
    // the joiner and the time. The zone word counts even when it was ambiguous and skipped
    // ("4pm IST / 1pm"), and the add-in's own bracket between them is stepped over so a second
    // click sees the same thing as the first.
    // The rule is symmetric: the zoned time may sit on either side of the joiner.
    for (var j = 0; j < out.length; j++) {
      if (out[j].hasZone) continue;
      if (leftIsZonedTime(text.slice(0, out[j].s)) ||
        rightIsZonedTime(text.slice(out[j].e, out[j].e + 80))) out[j].nextToZoned = true;
    }
    return out;
  }

  // The left text ends with "<time><zone word> [own bracket] <joiner>": a zoned time joined to the
  // bare time that follows. The joiner is a slash or a choice word (or, and, &, and/or), with at
  // most one filler word ("or at 5pm", "and maybe 5pm") between the joiner and the bare time.
  // The zone word may be ambiguous (IST), carry an offset (UTC-5), or sit in brackets ("(ET)").
  var OWN_LEFT = new RegExp("[ \\u00a0]*\\[\\s*" + ITEM + "(?:\\s*\\/\\s*" + ITEM + ")*\\s*\\][ \\u00a0]*$", "i");
  // One filler word the writer may put between the joiner and the time, plus "between", which
  // opens a bare range ("4pm ET / between 2 and 3pm").
  var FILLER_LEFT = /(?:^|[ \u00a0])(?:at|maybe|either|between)[ \u00a0]*$/i;
  var JOIN_LEFT = /(?:\band\/or|\bor|\band|&|\/)[ \u00a0]*$/i;
  var LEFT_ZONE = new RegExp("(?:\\d|noon|midnight)[^\\/\\[\\]]{0,45}?(" + ZONE_KEYS + "|" + AMBIGUOUS.join("|") + ")" +
    "(?:[ \\u00a0]+(?:standard|daylight))?(?:[ \\u00a0]+time(?:[ \\u00a0]+zone)?)?" +
    "(?:[+\\-\\u2212]\\d{1,2}(?::?\\d{2})?)?[)\\]]?[ \\u00a0]*$", "i");
  function leftIsZonedTime(left) {
    var s = left.replace(/[ \u00a0]*$/, "").replace(FILLER_LEFT, "").replace(/[ \u00a0]*$/, "");
    var j = JOIN_LEFT.exec(s);
    if (!j) return false;
    s = s.slice(0, j.index).replace(/[ \u00a0]*$/, "");
    s = s.replace(OWN_LEFT, "").replace(/[ \u00a0]*$/, "");
    var m = LEFT_ZONE.exec(s);
    if (!m || !zoneWord(m[1])) return false;
    // Same name rule as the right side: a rejected name ("3pm Central Park or 2pm") is not a zone,
    // so the bare time beside it is not a joined pair.
    return nameAccepted(m[1], s.slice(m.index + m[0].length));
  }

  // The mirror of the above: the text right after the bare time is
  // "[own bracket] <joiner> [filler] <time><zone word>".
  var OWN_RIGHT = new RegExp("^[ \\u00a0]*\\[\\s*" + ITEM + "(?:\\s*\\/\\s*" + ITEM + ")*\\s*\\]", "i");
  var JOIN_RIGHT = /^(?:and\/or|or|and|&|\/)(?![A-Za-z\/])/i;
  var FILLER_RIGHT = /^[ \u00a0]*(?:at|maybe|either)(?![A-Za-z])/i;
  var RIGHT_ZONE = new RegExp("^[ \\u00a0]*(?:between[ \\u00a0]+)?[(\\[]?(?:\\d|noon|midnight)" +
    "[^\\/\\[\\]]{0,45}?[ \\u00a0(]?(" + ZONE_KEYS + "|" + AMBIGUOUS.join("|") + ")(?![A-Za-z])", "i");
  function rightIsZonedTime(right) {
    var s = right.replace(OWN_RIGHT, "").replace(/^[ \u00a0]*/, "");
    var j = JOIN_RIGHT.exec(s);
    if (!j) return false;
    s = s.slice(j[0].length).replace(FILLER_RIGHT, "");
    var m = RIGHT_ZONE.exec(s);
    if (!m || !zoneWord(m[1])) return false;
    // The zone word has to be one the reader would accept. A spelled-out name followed by an
    // ordinary word ("3pm Central Park") is not a zone, so the time beside it is not a joined pair
    // and the bare time keeps the writer's label (AGENTS.md name rule).
    return nameAccepted(m[1], s.slice(m.index + m[0].length));
  }
  // A spelled-out name (over four letters) is a zone only with a "Time" tail, at a clause end,
  // before punctuation, or before a function word. Abbreviations are always accepted.
  function nameAccepted(tok, after) {
    if (tok.length <= 4) return true;
    if (/^(?:[ \u00a0]+(?:standard|daylight))?[ \u00a0]+time(?:[ \u00a0]+zone)?\b/i.test(after)) return true;
    return NAME_END.test(after);
  }

  // The second half of a "4pm ET/3pm CT" pair, read from the text: a match-shaped object or null.
  var PAIR_TAIL = new RegExp("^[ \\u00a0]*\\/[ \\u00a0]*(\\d{1,2})(?::(\\d{2}))?[ \\u00a0]?(" + MER + ")?" +
    "[ \\u00a0]?(" + ZONE_KEYS + ")(?![A-Za-z])", "i");
  function pairTail(after, head) {
    var m = PAIR_TAIL.exec(after);
    if (!m) return null;
    var h = +m[1], mi = m[2] ? +m[2] : 0;
    if (h < 1 || h > 12 || mi > 59) return null;
    // With no meridiem of its own the tail borrows the head's, which is how people write these.
    var c = m[3] ? m[3].toLowerCase()[0] : (head.h1 >= 12 ? "p" : "a");
    var region = zoneWord(m[4]);
    if (!region || region === SKIP) return null;
    return { h1: to24(h, c), mi1: mi, h2: null, mi2: null, dayIdx: head.dayIdx, zoneRegion: region };
  }

  // Two matches are the same instant when both convert to the same moment, ends included.
  function sameInstant(a, b, refMs) {
    if ((a.h2 === null) !== (b.h2 === null)) return false;
    try {
      var ca = conv(a, a.zoneRegion, "UTC", refMs), cb = conv(b, b.zoneRegion, "UTC", refMs);
      if (ca.t !== cb.t) return false;
      if (a.h2 === null) return true;
      return endMs(a, refMs) === endMs(b, refMs);
    } catch (e) { return false; }
  }
  function endMs(m, refMs) {
    return conv({ h1: m.h2, mi1: m.mi2, h2: null, mi2: null, dayIdx: m.dayIdx }, m.zoneRegion, "UTC", refMs).t;
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
  // Display name -> "F.Lastname": Maria Santos -> M.Santos, "Santos, Maria" -> M.Santos,
  // Maria -> Maria, "Maria de la Cruz" -> M.Cruz (last word only when the surname is several words).
  var MAX_NAME = 24;
  // Tags and company tails mail systems bolt onto a display name.
  var TAG = /\[[^\[\]]*\]|\((?:external|ext|extern)\)/ig;
  // Courtesy titles before the first name.
  var TITLE = /^(?:dr|mr|mrs|ms|miss|mx|prof|professor|rev|fr|sr|capt|col|gen|lt|sgt|hon|sir)\.?[ \u00a0]+/i;
  // Suffixes and credentials after the surname. They never become the surname.
  var SUFFIX = /^(?:jr|sr|ii|iii|iv|v|vi|md|do|dds|dvm|rn|np|pa|phd|edd|psyd|jd|esq|cpa|cfa|cfp|pe|pmp|mba|ma|ms|msc|bs|ba|mph|lcsw|cpc|rhit|emeritus|retired|ret)\.?$/i;
  // A company or team tail after a dash or a pipe ("Maria Santos - Acme", "Maria Santos | Acme").
  var ORG_TAIL = /[ \u00a0]*[|\u2502][^|\u2502]*$|[ \u00a0]+[-\u2013\u2014][ \u00a0]+.*$/;
  // The first code point, so an emoji or other astral character is never split in half.
  function firstCodePoint(s) {
    var c = s.codePointAt(0);
    return c === undefined ? "" : String.fromCodePoint(c);
  }
  // Never cut a surname mid-word: the surname alone, else the first MAX_NAME characters cut at a
  // word or hyphen boundary.
  function capName(full, last, cleaned) {
    if (full.length <= MAX_NAME) return full;
    if (last && last.length <= MAX_NAME) return last;
    var src = last || full || cleaned;
    var cut = src.slice(0, MAX_NAME);
    var b = Math.max(cut.lastIndexOf(" "), cut.lastIndexOf("-"), cut.lastIndexOf("'"));
    return (b > 0 ? cut.slice(0, b) : cut) || cleaned.slice(0, MAX_NAME);
  }
  function shortName(displayName) {
    var raw = String(displayName || "").replace(/"/g, "").trim().replace(/^'+|'+$/g, "");
    // Drop an address, a parenthesized company, and [EXTERNAL]-style tags wherever they sit.
    var n = raw.replace(/[ \u00a0]*<[^<>]*>/g, " ").replace(TAG, " ");
    n = n.replace(/[ \u00a0]*\([^()]*\)[ \u00a0]*$/, "").replace(/[ \u00a0]+/g, " ").trim();
    // Stripping left nothing: the display name was only a tag, so echo it rather than an empty bar.
    if (!n) return raw.replace(/[ \u00a0]+/g, " ").trim().slice(0, MAX_NAME);
    if (n.indexOf("@") >= 0) {
      // An email-only display name: the part before @, unchanged.
      if (/\s/.test(n)) return "";
      return n.split("@")[0].slice(0, 30);
    }
    var cleaned = n;
    var first, rest;
    if (n.indexOf(",") >= 0) {
      var halves = n.split(",");
      // "Maria Santos, Jr." is a suffix, not a surname-first form: only "Santos, Maria" flips.
      var tail = (halves[1] || "").trim().replace(/[ \u00a0]*\([^()]*\)[ \u00a0]*$/, "").trim();
      if (SUFFIX.test(tail)) {
        n = halves[0].trim();
      } else {
        n = null;
        rest = halves[0].trim().split(/[ \u00a0]+/);
        first = tail.split(/[ \u00a0]+/)[0] || "";
      }
    }
    if (n !== null) {
      n = n.replace(TITLE, "").replace(ORG_TAIL, "").trim();
      if (!n) return cleaned.slice(0, MAX_NAME);
      var words = n.split(/[ \u00a0]+/);
      // Strip trailing suffixes and credentials ("Jr.", "III", "PhD") so they never read as surname.
      while (words.length > 1 && SUFFIX.test(words[words.length - 1])) words.pop();
      first = words[0];
      rest = words.slice(1);
    }
    while (rest.length > 1 && SUFFIX.test(rest[rest.length - 1])) rest.pop();
    if (!rest.length || !rest[0]) return capName(first || "", "", cleaned) || cleaned.slice(0, MAX_NAME);
    if (!first) {
      var only = rest[rest.length - 1];
      return capName(only, only, cleaned) || cleaned.slice(0, MAX_NAME);
    }
    var last = rest.length > 1 ? rest[rest.length - 1] : rest[0];
    return capName(firstCodePoint(first) + "." + last, last, cleaned) || cleaned.slice(0, MAX_NAME);
  }
  // Kept for callers of 0.1.
  function firstName(displayName) { return shortName(displayName); }

  // ---- One time, formatted for a list of people ----
  // o: {sender: {name, zone}, readerTz, others: [{name, zone}], refMs, level, names}
  // level 0: two names per segment. 1: one name + count. 2: counts only.
  // names false: no names at all (used for the second and later times in the bar).
  function timeForPeople(m, o) {
    var src = m.zoneRegion, refMs = o.refMs, readerTz = o.readerTz, level = o.level || 0;
    var withNames = o.names !== false;
    var cRead = conv(m, src, readerTz, refMs), t = cRead.t;
    var groups = [], byOff = {};
    function group(zone) {
      var off = offsetMin(zone, t), g = byOff[off];
      if (!g) { g = byOff[off] = { off: off, zone: zone, names: [], count: 0 }; groups.push(g); }
      return g;
    }
    var gs = group(src), gr = group(readerTz);
    var senderName = withNames ? shortName(o.sender && o.sender.name) : "";
    gs.count++;
    if (senderName) gs.names.push(senderName);
    gr.count++;
    if (withNames) gr.names.push("you");
    (o.others || []).forEach(function (p) {
      // A person with a missing or unusable zone is left out silently: never throw on a bad zone.
      if (!p || !validZone(p.zone)) return;
      var g = group(p.zone);
      g.count++;
      var nm = withNames ? shortName(p.name) : "";
      if (nm) g.names.push(nm);
    });
    var rest = groups.filter(function (g) { return g !== gs && g !== gr; })
      .sort(function (a, b) { return b.count - a.count || a.off - b.off; });
    var ordered = [gs].concat(gr === gs ? [] : [gr]).concat(rest);
    var extraZones = ordered.length > 4 ? ordered.length - 4 : 0;

    function seg(g) {
      var txt, shift = "";
      if (g === gs) {
        txt = (m.dayIdx === null ? "" : DAYS[m.dayIdx] + " ") + m.text.replace(/[\s\u00a0]+/g, " ");
      } else {
        var c = conv(m, src, g.zone, refMs);
        txt = (m.dayIdx !== null && c.delta ? DAYS[c.dayIdx] + " " : "") + c.text + " " + labelFor(g.zone, t);
        if (m.dayIdx === null) shift = dayShift(c.delta);
      }
      // The rollover marker belongs to the time, so it sits right after the time and its zone
      // label, before any names or counts: "11pm AZ -1d you +1", never "you +1 -1d".
      txt += shift;
      if (!withNames) return txt;
      if (level >= 2) return txt + " +" + g.count;
      var max = level >= 1 ? 1 : 2;
      var shown = g.names.slice(0, max), more = g.count - shown.length;
      if (!shown.length) return txt + (g.count ? " +" + g.count : "");
      return txt + " " + shown.join(", ") + (more > 0 ? " +" + more : "");
    }
    return ordered.slice(0, 4).map(seg).join(" | ") +
      (extraZones ? " | +" + extraZones + " more zones" : "");
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

  // o: {displayName, refMs, deviceTz, others}. Returns bar text or null.
  // others (version 0.2): [{name, zone}] for the other people on the message.
  function readerBar(text, o) {
    var items = [], seen = {}, body = stripQuoted(text);
    if (!validZone(o.deviceTz)) return null;
    var own = ownSpans(body);
    findTimes(body, o.refMs).forEach(function (m) {
      if (!m.hasZone || !m.zoneRegion) return;   // no zone written: show nothing, never guess
      if (m.pairTail) return;                    // "4pm ET / 1pm PT": one time, read from the first zone
      if (!validZone(m.zoneRegion)) return;      // an offset the platform will not take: skip the time
      // A bracket group this add-in wrote holds a time already shown by the time before it. Any
      // other bracket ("see [notes: 4pm ET]") is ordinary prose and is read normally.
      if (inOwnBracket(own, m.s)) return;
      try {
        var c = conv(m, m.zoneRegion, o.deviceTz, o.refMs);
        if (c.sameOffset) return;
        var key = (m.dayIdx === null ? "" : m.dayIdx) + "|" + c.srcText + "|" + m.zoneTok;
        if (seen[key]) return;
        seen[key] = 1;
        items.push(m);
      } catch (e) { /* a zone Intl refuses: skip this time, never break the bar */ }
    });
    if (!items.length) return null;
    var sender = { name: o.displayName, zone: null };
    function build(k, level) {
      var segs = items.slice(0, k).map(function (m, i) {
        return timeForPeople(m, {
          sender: sender, readerTz: o.deviceTz, others: o.others, refMs: o.refMs,
          level: level, names: i === 0
        });
      });
      return segs.join(" · ") + (k < items.length ? " +" + (items.length - k) + " more" : "");
    }
    try {
      // Shorten in order: two names, one name + count, counts only, then drop later times.
      for (var level = 0; level <= 2; level++) {
        var bar = build(items.length, level);
        if (bar.length <= MAX_BAR) return bar;
      }
      for (var k = items.length - 1; k > 0; k--) {
        var b2 = build(k, 2);
        if (b2.length <= MAX_BAR) return b2;
      }
      // A time was found, so the bar is never dropped: fall back to the first time alone, as the
      // written time with the sender and the reader's time. Still too long: echo the short form.
      return lastResort(items[0], sender, o);
    } catch (e) {
      return null;
    }
  }

  // Two segments for one time: "<written time> <sender> | <time> <label> you".
  function lastResort(m, sender, o) {
    var c = conv(m, m.zoneRegion, o.deviceTz, o.refMs);
    var mine = c.text + " " + labelFor(o.deviceTz, c.t) + dayShift(c.delta) + " you";
    var who = shortName(sender.name);
    var written = (m.dayIdx === null ? "" : DAYS[m.dayIdx] + " ") + m.text.replace(/[\s\u00a0]+/g, " ");
    var bar = written + (who ? " " + who : "") + " | " + mine;
    if (bar.length <= MAX_BAR) return bar;
    // Shorten the echoed written time to its own "h[:mm]am/pm ZONE" form.
    var shortWritten = (m.dayIdx === null ? "" : DAYS[m.dayIdx] + " ") + c.srcText + " " + m.zoneTok;
    return (shortWritten + (who ? " " + who : "") + " | " + mine).slice(0, MAX_BAR);
  }

  // ---- Compose side ----
  // o: {deviceTz, nowMs, before, after}. Returns the text with each clock time rewritten.
  // before/after: neighboring inline text in the same block (HTML splits text across spans);
  // a time that runs into a neighbor ("4pm" + "EST") is left untouched.
  function rewriteText(text, o) {
    var full = (o.before || "") + text + (o.after || ""), a = (o.before || "").length, z = a + text.length;
    if (!validZone(o.deviceTz)) return text;
    var spans = [], re = /\[[^\[\]]*\]/g, b;
    while ((b = re.exec(full))) spans.push([b.index, b.index + b[0].length]);
    var out = "", pos = a;
    findTimes(full, o.nowMs).forEach(function (m) {
      if (m.s < a || m.e > z) return;
      if (spans.some(function (sp) { return m.s >= sp[0] && m.s < sp[1]; })) return; // inside any bracket
      if (m.pairHead || m.pairTail) return; // "4pm ET / 1pm PT": the writer gave both zones already
      if (m.nextToZoned) return;            // "4pm ET / 1pm": a bare time beside a zoned one, left alone
      if (OWN.test(full.slice(m.e))) return; // already done by this add-in
      var src = m.hasZone ? m.zoneRegion : o.deviceTz;
      if (!validZone(src)) return;
      try {
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
      } catch (e) { /* a zone Intl refuses: leave this time exactly as written */ }
    });
    return out + full.slice(pos, z);
  }

  return {
    ZONES: ZONES, LABELS: LABELS, TARGETS: TARGETS, MAX_BAR: MAX_BAR,
    findTimes: findTimes, readerBar: readerBar, rewriteText: rewriteText,
    timeForPeople: timeForPeople, shortName: shortName,
    firstName: firstName, stripQuoted: stripQuoted
  };
});

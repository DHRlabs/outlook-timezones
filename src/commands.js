/* Compose side: "Add time zones" button. Rewrites clock times in the draft body, text nodes only. */

// Quoted reply/forward content is left alone. Everything from the first of these markers onward
// (in document order) counts as quoted. Markers are the ones Outlook writes:
//   #appendonsend, #divRplyFwdMsg  Outlook on the web / Windows reply and forward
//   #mail-editor-reference-message-container  new Outlook
//   [id^=OLK_SRC_BODY_SECTION]  Outlook for Mac and classic
//   .gmail_quote, blockquote  generic
var QUOTE_MARKERS = "#appendonsend, #divRplyFwdMsg, #x_divRplyFwdMsg, #mail-editor-reference-message-container, " +
  "[id^='OLK_SRC_BODY_SECTION'], .gmail_quote, blockquote";
// The writer's own signature is also left alone.
var SKIP_BLOCKS = "#Signature, #x_Signature, #ms-outlook-mobile-signature, [id^='Signature']";

// Text in the same block is joined so a time split across spans ("4pm" + "EST") is seen whole.
var BLOCKS = "p, div, li, td, th, h1, h2, h3, h4, h5, h6, pre, blockquote, body, tr, table, ul, ol, dd, dt";
var CONTEXT = 200; // characters of neighboring text passed on each side

// Outlook's own HTML is never re-serialized. The DOM is used only to decide which text gets a
// label; the edits are then spliced into the original source string, so every byte the add-in did
// not add stays exactly as Outlook wrote it: conditional markers (<![if !supportLists]>), named
// entities, a missing <tbody>, the doctype and leading comments all survive untouched.

// Elements whose content the HTML parser reads as raw text: the whole content is one segment and
// is never edited, so a time written inside a stylesheet or a script is left alone.
var RAW_TEXT = { script: 1, style: 1, textarea: 1, title: 1, xmp: 1 };
// The parser ignores whitespace before <html> and before <head>, so those runs have no text node.
var PREAMBLE_TAGS = { html: 1 };

// ---- source tokenizer: no DOM, no regex over the whole body ----
// Returns segments in source order. Only segments with .text are text runs; they line up, one for
// one, with the document's text nodes. .raw marks a run that must not be edited or entity-decoded.
function scanSource(src) {
  var segs = [], i = 0, n = src.length, preamble = true, start, j, c, name, lower, close, tmpl = 0;
  function pushText(from, to) {
    if (to <= from) return;
    var run = src.slice(from, to);
    // Template content is parsed into a separate fragment, so it has no text node in the
    // document and is never a run: it is skipped, like preamble whitespace.
    if (tmpl > 0) { segs.push({ skip: 1, s: from, e: to }); return; }
    if (preamble && !/[^ \t\n\r\f]/.test(run)) { segs.push({ skip: 1, s: from, e: to }); return; }
    segs.push({ text: run, s: from, e: to });
  }
  while (i < n) {
    var lt = src.indexOf("<", i);
    if (lt < 0) { pushText(i, n); break; }
    // A "<" that cannot start markup (stray "<" in prose) is ordinary text.
    c = src.charAt(lt + 1);
    if (!/[a-zA-Z!?\/]/.test(c)) {
      pushText(i, lt + 1); i = lt + 1; continue;
    }
    pushText(i, lt);
    start = lt;
    if (src.substr(lt, 4) === "<!--") {
      // The parser closes a comment on "-->" or "--!>", and "<!-->" / "<!--->" are comments that
      // close right away. Treating any of these as still open would misplace every later run.
      if (src.charAt(lt + 4) === ">") i = lt + 5;
      else if (src.charAt(lt + 4) === "-" && src.charAt(lt + 5) === ">") i = lt + 6;
      else {
        var cm = /--!?>/.exec(src.slice(lt + 4));
        i = cm ? lt + 4 + cm.index + cm[0].length : n;
      }
      segs.push({ markup: 1, s: start, e: i });
      continue;
    }
    // doctype, bogus and conditional markers (<![if !supportLists]>, <![endif]>), and a "</"
    // that is not a real end tag, which the parser reads as a bogus comment: all run to ">".
    if (c === "!" || c === "?" || (c === "/" && !/[a-zA-Z]/.test(src.charAt(lt + 2)))) {
      j = src.indexOf(">", lt + 1);
      i = j < 0 ? n : j + 1;
      segs.push({ markup: 1, s: start, e: i });
      if (c === "/") preamble = false;
      continue;
    }
    // a start or end tag: walk attributes the way HTML does, so a ">" inside a quoted value does
    // not end the tag and a quote that is not a value opener ("title=x'") is an ordinary character.
    var end = c === "/";
    j = lt + (end ? 2 : 1);
    while (j < n && /[a-zA-Z0-9]/.test(src.charAt(j))) j++;
    name = src.slice(lt + (end ? 2 : 1), j).toLowerCase();
    // State: 0 between attributes / in an attribute name, 1 just after "=", 2 in a quoted value,
    // 3 in an unquoted value. A quote opens a value only in state 1.
    var st = 0, q = 0;
    while (j < n) {
      var ch = src.charAt(j);
      if (st === 2) { if (ch === q) st = 0; }
      else if (st === 1) {
        if (ch === '"' || ch === "'") { q = ch; st = 2; }
        else if (ch === ">") { j++; break; }
        else if (!/[ \t\n\r\f]/.test(ch)) st = 3;   // unquoted value begins
      } else if (st === 3) {
        if (ch === ">") { j++; break; }
        if (/[ \t\n\r\f]/.test(ch)) st = 0;          // unquoted value ends at whitespace
      } else {
        if (ch === ">") { j++; break; }
        if (ch === "=") st = 1;
      }
      j++;
    }
    i = j;
    segs.push({ markup: 1, s: start, e: i });
    if (!PREAMBLE_TAGS[name]) preamble = false;
    if (name === "template") { if (end) { if (tmpl > 0) tmpl--; } else tmpl++; }
    if (!end && RAW_TEXT[name]) {
      lower = src.toLowerCase();
      close = lower.indexOf("</" + name, i);
      var body = close < 0 ? n : close;
      if (body > i) segs.push({ text: src.slice(i, body), s: i, e: body, raw: name });
      i = body;
    }
  }
  return segs;
}

// ---- entity decoding ----
// The small table covers what Outlook and Word write. Anything else is decoded by the browser's
// own parser once and remembered, so no named-entity table has to be carried here.
var ENT = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: "\u00a0" };
var entCache = {};
function entityValue(match) {
  if (entCache[match] !== undefined) return entCache[match];
  var m = /^&([a-zA-Z][a-zA-Z0-9]*);$/.exec(match), v = null;
  if (m && ENT[m[1]] !== undefined) v = ENT[m[1]];
  else {
    try {
      var d = new DOMParser().parseFromString("<i>" + match + "</i>", "text/html");
      v = d.body.textContent;
    } catch (e) { v = match; }
  }
  entCache[match] = v;
  return v;
}
var ENT_RE = /&(?:#[0-9]+;?|#[xX][0-9a-fA-F]+;?|[a-zA-Z][a-zA-Z0-9]*;?)/g;
// decode(run) -> plain text. With wantMap, also an index from each decoded character position to
// its offset in the run, so an insertion can never land inside an entity.
function decode(run, wantMap) {
  if (run.indexOf("&") < 0) {
    if (!wantMap) return { text: run };
    var m0 = new Array(run.length + 1);
    for (var k = 0; k <= run.length; k++) m0[k] = k;
    return { text: run, map: m0 };
  }
  var out = "", map = wantMap ? [] : null, pos = 0, mm;
  ENT_RE.lastIndex = 0;
  while ((mm = ENT_RE.exec(run))) {
    var at = mm.index, lit = run.slice(pos, at), val = entityValue(mm[0]);
    for (var a = 0; a < lit.length; a++) { out += lit.charAt(a); if (map) map.push(pos + a); }
    if (val === mm[0]) {
      for (var b = 0; b < val.length; b++) { out += val.charAt(b); if (map) map.push(at + b); }
    } else {
      for (var d2 = 0; d2 < val.length; d2++) { out += val.charAt(d2); if (map) map.push(d2 ? at + mm[0].length : at); }
    }
    pos = at + mm[0].length;
  }
  var tail = run.slice(pos);
  for (var t = 0; t < tail.length; t++) { out += tail.charAt(t); if (map) map.push(pos + t); }
  if (map) map.push(run.length);
  return { text: out, map: map };
}

// The parser normalizes CR LF and a lone CR to a single LF before anything else, but the source
// run still holds the CR bytes. This gives the parser's view of a run plus a map from each
// character of that view back to its byte in the run, so an insertion still lands in the original
// string with every CR preserved. A NUL is not modeled (parsers differ): the caller gives up.
function parserView(run, raw, wantMap) {
  var res;
  if (raw) {
    var m0 = [], z;
    for (z = 0; z <= run.length; z++) m0[z] = z;
    res = { text: run, map: m0 };
  } else res = decode(run, true);
  if (res.text.indexOf("\u0000") >= 0) return null;
  if (res.text.indexOf("\r") < 0) return { text: res.text, map: wantMap ? res.map : null };
  var out = "", map = [], t = res.text, k;
  for (k = 0; k < t.length; k++) {
    var ch = t.charAt(k);
    if (ch === "\r") {
      out += "\n";
      map.push(res.map[k]);
      if (t.charAt(k + 1) === "\n") k++;   // CR LF counts as the one LF just emitted
    } else {
      out += ch;
      map.push(res.map[k]);
    }
  }
  map.push(res.map[t.length]);
  return { text: out, map: wantMap ? map : null };
}

// The add-in only ever adds characters. This finds where, keeping every original character of the
// run byte for byte. Null if the new text is not the old text with insertions.
function insertionsOf(old, next) {
  var ins = [], i = 0, j = 0, buf = "", at = 0;
  while (i < old.length) {
    if (j >= next.length) return null;
    if (next.charAt(j) === old.charAt(i)) {
      if (buf) { ins.push({ at: at, text: buf }); buf = ""; }
      i++; j++;
    } else {
      if (!buf) at = i;
      buf += next.charAt(j); j++;
    }
  }
  if (j < next.length) { if (!buf) at = old.length; buf += next.slice(j); }
  if (buf) ins.push({ at: at, text: buf });
  return ins;
}

// Inserted text is plain ASCII with brackets and slashes, but it is escaped anyway.
function escapeIns(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// ---- final verification ----
// Walks a document in order and lists what must not change: every element with all of its
// attributes, every text value, every comment value. Template content is walked in place.
function profileDoc(root, nodes, textAt) {
  var out = [];
  (function walk(node) {
    for (var c = node.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 1) {
        var at = [], as = c.attributes, k;
        for (k = 0; k < as.length; k++) at.push(as[k].name + "=" + as[k].value);
        at.sort();
        out.push("E:" + c.nodeName + "|" + at.join("\u0001"));
        if (c.content) walk(c.content);
        walk(c);
      } else if (c.nodeType === 3 || c.nodeType === 4) {
        if (nodes) nodes.push(c);
        if (textAt) textAt.push(out.length);
        out.push("T:" + c.nodeValue);
      } else if (c.nodeType === 8) {
        out.push("C:" + c.nodeValue);
      } else if (c.nodeType === 10) {
        out.push("D:" + c.nodeName);
      }
    }
  })(root);
  return out;
}

// The spliced string must parse to the same document as the input with only the intended text
// values changed, on exactly the intended text nodes. Anything else means the splice landed
// somewhere it should not have (inside a tag, an attribute, a comment, an entity) and nothing
// is written.
function spliceIsClean(inDoc, spliced, intended) {
  var inNodes = [], textAt = [], want = profileDoc(inDoc, inNodes, textAt), i;
  for (i = 0; i < inNodes.length; i++) {
    var nv = intended.get(inNodes[i]);
    if (nv !== undefined) want[textAt[i]] = "T:" + nv;
  }
  var got;
  try { got = profileDoc(new DOMParser().parseFromString(spliced, "text/html"), null, null); }
  catch (e) { return false; }
  if (got.length !== want.length) return false;
  for (i = 0; i < want.length; i++) if (got[i] !== want[i]) return false;
  return true;
}

function rewriteHtml(html, deviceTz, nowMs) {
  var source = String(html == null ? "" : html);
  var doc = new DOMParser().parseFromString(source, "text/html");
  var markers = Array.prototype.slice.call(doc.body.querySelectorAll(QUOTE_MARKERS));
  var skips = Array.prototype.slice.call(doc.body.querySelectorAll(SKIP_BLOCKS));

  // Every text node of the whole document, in document order, for the mapping check.
  var all = [], w2 = doc.createTreeWalker(doc, NodeFilter.SHOW_TEXT), t2;
  while ((t2 = w2.nextNode())) all.push(t2);

  // Verify the mapping before anything is written. The parser joins neighboring text runs into one
  // node (a stray "<" in prose, whitespace after </body>), so a node takes one or more runs in
  // order, and the entity-decoded runs must spell the node's value exactly. Anything left over on
  // either side, or any inexact match, means no write at all.
  var segs = scanSource(source), runs = [], s;
  for (var si = 0; si < segs.length; si++) if (segs[si].text !== undefined) runs.push(segs[si]);
  var runOf = new Map(), ri = 0;
  for (var ni = 0; ni < all.length; ni++) {
    var want = all[ni].nodeValue, got = "", taken = [];
    while (ri < runs.length && got.length < want.length) {
      s = runs[ri++];
      var view = parserView(s.text, !!(s.raw && s.raw !== "title" && s.raw !== "textarea"), false);
      if (!view) return null;                       // NUL: the parser's view is not modeled
      taken.push({ run: s, len: view.text.length });
      got += view.text;
    }
    if (got !== want) return null;
    runOf.set(all[ni], taken);
  }
  if (ri !== runs.length) return null;

  // Decide the edits with the existing logic, over the body only.
  var walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
  var parts = [], n;
  while ((n = walker.nextNode())) {
    if (n.nodeType === Node.TEXT_NODE) parts.push({ node: n, text: n.nodeValue, block: n.parentNode.closest(BLOCKS) });
    else if (n.tagName === "BR") parts.push({ node: null, text: "\n", block: n.closest(BLOCKS) });
  }
  var edits = [], bad = false, intended = new Map();
  parts.forEach(function (seg, i) {
    var node = seg.node;
    if (!node || bad) return;
    var taken = runOf.get(node);
    if (!taken || taken.some(function (p) { return p.run.raw; })) return;
    var quoted = markers.some(function (m) {
      return m.contains(node) || (m.compareDocumentPosition(node) & Node.DOCUMENT_POSITION_FOLLOWING);
    });
    if (quoted || skips.some(function (sk) { return sk.contains(node); })) return;
    var before = "", after = "", j;
    for (j = i - 1; j >= 0 && parts[j].block === seg.block && before.length < CONTEXT; j--) before = parts[j].text + before;
    for (j = i + 1; j < parts.length && parts[j].block === seg.block && after.length < CONTEXT; j++) after += parts[j].text;
    var out = TZ.rewriteText(seg.text, { deviceTz: deviceTz, nowMs: nowMs, before: before, after: after });
    if (out === seg.text) return;
    var ins = insertionsOf(seg.text, out);
    if (!ins) { bad = true; return; }
    intended.set(node, out);
    ins.forEach(function (x) {
      // Find which of the node's source runs holds that decoded offset, then the exact byte in it.
      var rest = x.at, p, k;
      for (k = 0; k < taken.length; k++) {
        p = taken[k];
        if (rest <= p.len && (rest < p.len || k === taken.length - 1)) break;
        rest -= p.len;
      }
      if (k >= taken.length) { bad = true; return; }
      var mv = parserView(p.run.text, !!p.run.raw, true);
      if (!mv) { bad = true; return; }
      var off = mv.map[rest];
      if (off === undefined) { bad = true; return; }
      edits.push({ at: p.run.s + off, text: escapeIns(x.text) });
    });
  });
  if (bad || !edits.length) return null;

  // The output is the original source with only those insertions.
  edits.sort(function (a, b) { return a.at - b.at; });
  var out2 = "", prev = 0;
  edits.forEach(function (e) { out2 += source.slice(prev, e.at) + e.text; prev = e.at; });
  out2 += source.slice(prev);
  // Last gate: the result must parse to the same document, with only the intended text changed.
  if (!spliceIsClean(doc, out2, intended)) return null;
  return out2;
}

// A failed read or write leaves the draft as it was and says so on the draft. Called at most once
// per click; a throw anywhere inside still completes the event.
function failed(item, event, done) {
  var fin = function () { if (!done.d) { done.d = 1; event.completed(); } };
  try {
    item.notificationMessages.replaceAsync("otz-error", {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: "Time zones were not added."
    }, fin);
  } catch (e) { fin(); }
}

function addTimeZones(event) {
  var done = { d: 0 };
  var finish = function () { if (!done.d) { done.d = 1; event.completed(); } };
  var item = null;
  var fail = function () { if (item) failed(item, event, done); else finish(); };
  try {
    item = Office.context.mailbox.item;
    var ok = Office.AsyncResultStatus.Succeeded;
    item.body.getAsync(Office.CoercionType.Html, function (r) {
      try {
        if (!r || r.status !== ok) { fail(); return; }
        var html = rewriteHtml(r.value, Intl.DateTimeFormat().resolvedOptions().timeZone, Date.now());
        if (html === null) { finish(); return; }
        item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, function (w) {
          try {
            if (!w || w.status !== ok) fail();
            else finish();
          } catch (e) { finish(); }
        });
      } catch (e) { fail(); }
    });
  } catch (e) { fail(); }
}

// OnMessageSend: the same rewrite the button does, run automatically at Send. A failure of any
// kind still allows the send: the manifest's SendMode is PromptUser, and nothing here is worth
// stopping a message over. The rewrite needs DOMParser, which the classic Outlook on Windows
// JavaScript-only runtime does not have; there the body is left alone and the send goes through.
function onMessageSendHandler(event) {
  var done = 0;
  var allow = function () { if (!done) { done = 1; event.completed({ allowEvent: true }); } };
  try {
    var item = Office.context.mailbox.item, ok = Office.AsyncResultStatus.Succeeded;
    if (typeof DOMParser === "undefined") { allow(); return; }
    item.body.getAsync(Office.CoercionType.Html, function (r) {
      try {
        if (r.status !== ok) { allow(); return; }
        var html = rewriteHtml(r.value, Intl.DateTimeFormat().resolvedOptions().timeZone, Date.now());
        if (html === null) { allow(); return; }   // nothing to add, so nothing is written back
        item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, allow);
      } catch (e) { allow(); }
    });
  } catch (e) { allow(); }
}

Office.onReady(function () {});
Office.actions.associate("addTimeZones", addTimeZones);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

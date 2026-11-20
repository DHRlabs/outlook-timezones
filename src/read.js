/* Read side: keeps one informational bar per time above the open message, and lists the same
   lines in the pinned pane. */
(function () {
  var KEYS = TZ.BAR_KEYS;

  // The pane body: one entry per bar, the reader's time (everything before "<-") large and bold,
  // the rest smaller under it. No bars: exactly "No times in this message". No other text.
  function paint(bars) {
    var host = document.getElementById("otz-list");
    if (!host) return;
    while (host.firstChild) host.removeChild(host.firstChild);
    if (!bars.length) {
      var none = document.createElement("p");
      none.className = "otz-none";
      none.textContent = "No times in this message";
      host.appendChild(none);
      return;
    }
    bars.forEach(function (bar) {
      var cut = bar.indexOf("  <-  ");
      var mine = cut < 0 ? bar : bar.slice(0, cut);
      var rest = cut < 0 ? "" : bar.slice(cut + 6);
      var row = document.createElement("div");
      row.className = "otz-row";
      var big = document.createElement("div");
      big.className = "otz-mine";
      big.textContent = mine;
      row.appendChild(big);
      if (rest) {
        var small = document.createElement("div");
        small.className = "otz-from";
        small.textContent = rest;
        row.appendChild(small);
      }
      host.appendChild(row);
    });
  }

  // The UTC offset of the message's own Date header, in minutes, or null. getAllInternetHeadersAsync
  // is requirement set 1.8 and is missing on older hosts, so it is feature-detected: without it the
  // add-in behaves as it did before, reading only times that carry a written zone.
  // A header read that fails, throws, or never calls back at all still lets the message be
  // handled, with no sender clock (AGENTS.md). The 3 s wait is not the end of the story: it hands
  // over no sender clock so the message is painted now, and a header result that arrives after it
  // is handed over as well, so the same message is repainted with its assumed bars. The handover
  // repeats only when the sender clock actually changed, so nothing repaints for nothing.
  var HEADER_WAIT = 3000;
  function withSenderOffset(item, done) {
    var answered = false, handed = false, last = null;
    function deliver(off) {
      if (handed && off === last) return;
      handed = true;
      last = off;
      done(off);
    }
    if (typeof item.getAllInternetHeadersAsync !== "function") return deliver(null);
    // A host that accepts the call but never answers must not leave the pane on the last message.
    setTimeout(function () { if (!answered) deliver(null); }, HEADER_WAIT);
    try {
      item.getAllInternetHeadersAsync(function (r) {
        var off = null;
        if (r && r.status === Office.AsyncResultStatus.Succeeded) {
          try { off = TZ.senderOffsetFromHeaders(r.value); } catch (e) { off = null; }
        }
        answered = true;
        deliver(off);
      });
    } catch (e) {
      answered = true;
      deliver(null);
    }
  }

  // Paint the pane and set or clear this message's bars. Every path through update ends here, so a
  // message with no times, a failed read, or a thrown error all leave the pane and the bars right.
  function show(item, bars) {
    try { paint(bars); } catch (e) { /* the pane is gone: the bars below still matter */ }
    if (!item || !item.notificationMessages) return;
    KEYS.forEach(function (key, i) {
      try {
        if (i < bars.length) {
          item.notificationMessages.replaceAsync(key, {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: bars[i],
            icon: "icon-16",
            persistent: false
          }, function () {});
        } else {
          // Every key this message is not using is cleared, so no bar from the last message stays.
          // Removing a key that was never set is fine: the failed result is read and dropped.
          item.notificationMessages.removeAsync(key, function () {});
        }
      } catch (e) { /* one key refused: keep going so the rest of the bars are still right */ }
    });
  }

  // Outlook's Text coercion glues HTML block elements together with no separator ("tuesdayOn Fri,
  // Nov 13..."), which hides weekdays and line starts from the parser. So an HTML body is read as
  // HTML and turned into text here: a line break before and after every block element. DOMParser
  // builds a detached document, so nothing from the mail is ever inserted into the live pane.
  var BLOCKS = {};
  ("DIV P BR LI TR TABLE H1 H2 H3 H4 H5 H6 BLOCKQUOTE HR SECTION ARTICLE HEADER FOOTER PRE")
    .split(" ").forEach(function (t) { BLOCKS[t] = true; });
  var SKIPPED = { SCRIPT: true, STYLE: true, HEAD: true };

  function htmlToText(html) {
    var doc = new DOMParser().parseFromString(String(html), "text/html");
    var out = [];
    function walk(node) {
      if (node.nodeType === 3) { out.push(node.nodeValue); return; }
      if (node.nodeType !== 1) return;
      var tag = node.nodeName.toUpperCase();
      if (SKIPPED[tag]) return;
      if (BLOCKS[tag]) out.push("\n");
      for (var c = node.firstChild; c; c = c.nextSibling) walk(c);
      if (BLOCKS[tag]) out.push("\n");
    }
    walk(doc.body || doc.documentElement);
    return out.join("").replace(/\u00a0/g, " ").replace(/\n{3,}/g, "\n\n");
  }

  // The body as text: HTML mail through htmlToText, plain-text mail (or any HTML read that fails)
  // as the Text coercion. cb gets an AsyncResult-shaped { status, value } exactly once.
  function readBody(item, cb) {
    var sent = false;
    function reply(r) { if (!sent) { sent = true; cb(r); } }
    function asText() {
      try { item.body.getAsync(Office.CoercionType.Text, reply); } catch (e) { reply(null); }
    }
    if (typeof item.body.getTypeAsync !== "function") return asText();
    try {
      item.body.getTypeAsync(function (t) {
        if (!t || t.status !== Office.AsyncResultStatus.Succeeded || t.value !== Office.CoercionType.Html) return asText();
        try {
          item.body.getAsync(Office.CoercionType.Html, function (r) {
            if (!r || r.status !== Office.AsyncResultStatus.Succeeded) return asText();
            try {
              reply({ status: r.status, value: htmlToText(r.value) });
            } catch (e) { asText(); }
          });
        } catch (e) { asText(); }
      });
    } catch (e) { asText(); }
  }

  // The body read waits at most this long. A host that accepts getAsync but never answers must
  // still leave the message painted, with its bars cleared (AGENTS.md: the pane is never blank).
  // It covers the body read only: once the body arrives this timer is cleared and the header wait
  // (HEADER_WAIT above) is what decides, so a slow header read never cancels a body already read.
  var BODY_WAIT = 5000;

  function update() {
    var item = Office.context.mailbox.item;
    // Read mode only: a pinned pane also fires for drafts (compose items have no displayReplyForm) and for no selection.
    if (!item || !item.body || !item.notificationMessages || typeof item.displayReplyForm !== "function") {
      show(null, []);
      return;
    }
    // The message the pane is showing. A later callback for a message the user has left is dropped,
    // so a slow read never paints over the message now open.
    function current() { return Office.context.mailbox.item === item; }
    // The waits above are what the message is painted with so far, never the last word: a body or
    // a header result that arrives after its wait is still this message's own answer, so it
    // repaints over the empty state the wait left. Only a callback for a message the user has
    // since left is dropped. The same bars are never painted twice in a row.
    var last = null;
    function finish(bars) {
      if (!current()) return;
      var sig = bars.join("\n");
      if (last === sig) return;
      last = sig;
      show(item, bars);
    }
    var bodyTimer = setTimeout(function () { finish([]); }, BODY_WAIT);
    try {
      readBody(item, function (r) {
        // The body is here, so its wait is over: a body that arrived after the wait still gets its
        // bars, replacing the empty state the wait painted.
        clearTimeout(bodyTimer);
        if (!current()) return;
        // The body could not be read: the message still gets a pane and cleared bars.
        if (!r || r.status !== Office.AsyncResultStatus.Succeeded) {
          finish([]);
          return;
        }
        withSenderOffset(item, function (senderOffsetMin) {
          if (!current()) return; // the user moved on while the headers loaded
          var bars = [];
          try {
            bars = TZ.readerBar(r.value, {
              displayName: item.from && item.from.displayName,
              refMs: item.dateTimeCreated ? item.dateTimeCreated.getTime() : Date.now(),
              deviceTz: Intl.DateTimeFormat().resolvedOptions().timeZone,
              senderOffsetMin: senderOffsetMin,
              others: null // coworker zones arrive in version 0.2 (Graph getSchedule)
            }) || [];
          } catch (e) {
            bars = []; // anything thrown inside the bar logic: no times rather than a stale pane
          }
          finish(bars);
        });
      });
    } catch (e) {
      clearTimeout(bodyTimer);
      finish([]);
    }
  }

  // The pane's own document may still be parsing when Office is ready: paint only once
  // "otz-list" can exist, so the first message is never dropped on a blank pane.
  function whenDomReady(fn) {
    if (document.readyState !== "loading") return fn();
    document.addEventListener("DOMContentLoaded", fn);
  }

  Office.onReady(function () {
    whenDomReady(function () {
      // A host that refuses the handler registration must not cost the open message its bars.
      try {
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, update);
      } catch (e) { /* no change events: the message now open is still painted below */ }
      update();
    });
  });
})();

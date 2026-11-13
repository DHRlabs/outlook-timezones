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
  // The callback runs exactly once: a header read that fails, throws, or never calls back at all
  // still lets the message be handled, with no sender clock (AGENTS.md).
  var HEADER_WAIT = 3000;
  function withSenderOffset(item, done) {
    var fired = false;
    function once(off) {
      if (fired) return;
      fired = true;
      done(off);
    }
    if (typeof item.getAllInternetHeadersAsync !== "function") return once(null);
    // A host that accepts the call but never answers must not leave the pane on the last message.
    setTimeout(function () { once(null); }, HEADER_WAIT);
    try {
      item.getAllInternetHeadersAsync(function (r) {
        var off = null;
        if (r && r.status === Office.AsyncResultStatus.Succeeded) {
          try { off = TZ.senderOffsetFromHeaders(r.value); } catch (e) { off = null; }
        }
        once(off);
      });
    } catch (e) {
      once(null);
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
    try {
      item.body.getAsync(Office.CoercionType.Text, function (r) {
        if (!current()) return;
        // The body could not be read: the message still gets a pane and cleared bars.
        if (!r || r.status !== Office.AsyncResultStatus.Succeeded) {
          show(item, []);
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
          show(item, bars);
        });
      });
    } catch (e) {
      show(item, []);
    }
  }

  Office.onReady(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, update);
    update();
  });
})();

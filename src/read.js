/* Read side: keeps one informational bar per time above the open message. Runs in the pinned pane. */
(function () {
  var KEYS = TZ.BAR_KEYS;

  function update() {
    var item = Office.context.mailbox.item;
    // Read mode only: a pinned pane also fires for drafts (compose items have no displayReplyForm) and for no selection.
    if (!item || !item.body || !item.notificationMessages || typeof item.displayReplyForm !== "function") return;
    item.body.getAsync(Office.CoercionType.Text, function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return;
      if (Office.context.mailbox.item !== item) return; // user moved on
      var bars = TZ.readerBar(r.value, {
        displayName: item.from && item.from.displayName,
        refMs: item.dateTimeCreated ? item.dateTimeCreated.getTime() : Date.now(),
        deviceTz: Intl.DateTimeFormat().resolvedOptions().timeZone,
        others: null // coworker zones arrive in version 0.2 (Graph getSchedule)
      }) || [];
      KEYS.forEach(function (key, i) {
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
      });
    });
  }

  Office.onReady(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, update);
    update();
  });
})();

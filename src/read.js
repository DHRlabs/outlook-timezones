/* Read side: keeps one informational bar above the open message. Runs in the pinned task pane. */
(function () {
  var KEY = "otz-bar";

  function update() {
    var item = Office.context.mailbox.item;
    // Read mode only: a pinned pane also fires for drafts (compose items have no displayReplyForm) and for no selection.
    if (!item || !item.body || !item.notificationMessages || typeof item.displayReplyForm !== "function") return;
    item.body.getAsync(Office.CoercionType.Text, function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return;
      if (Office.context.mailbox.item !== item) return; // user moved on
      var bar = TZ.readerBar(r.value, {
        displayName: item.from && item.from.displayName,
        refMs: item.dateTimeCreated ? item.dateTimeCreated.getTime() : Date.now(),
        deviceTz: Intl.DateTimeFormat().resolvedOptions().timeZone
      });
      if (bar) {
        item.notificationMessages.replaceAsync(KEY, {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: bar,
          icon: "icon-16",
          persistent: false
        });
      } else {
        item.notificationMessages.removeAsync(KEY, function () {}); // no bar yet is fine
      }
    });
  }

  Office.onReady(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, update);
    update();
  });
})();

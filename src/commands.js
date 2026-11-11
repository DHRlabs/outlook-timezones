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

function rewriteHtml(html, deviceTz, nowMs) {
  var doc = new DOMParser().parseFromString(html, "text/html");
  var markers = Array.prototype.slice.call(doc.body.querySelectorAll(QUOTE_MARKERS));
  var skips = Array.prototype.slice.call(doc.body.querySelectorAll(SKIP_BLOCKS));
  var walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
  var segs = [], n, changed = false;
  while ((n = walker.nextNode())) {
    if (n.nodeType === Node.TEXT_NODE) segs.push({ node: n, text: n.nodeValue, block: n.parentNode.closest(BLOCKS) });
    else if (n.tagName === "BR") segs.push({ node: null, text: "\n", block: n.closest(BLOCKS) });
  }
  segs.forEach(function (seg, i) {
    var node = seg.node;
    if (!node) return;
    var quoted = markers.some(function (m) {
      return m.contains(node) || (m.compareDocumentPosition(node) & Node.DOCUMENT_POSITION_FOLLOWING);
    });
    if (quoted || skips.some(function (s) { return s.contains(node); })) return;
    var before = "", after = "", j;
    for (j = i - 1; j >= 0 && segs[j].block === seg.block && before.length < CONTEXT; j--) before = segs[j].text + before;
    for (j = i + 1; j < segs.length && segs[j].block === seg.block && after.length < CONTEXT; j++) after += segs[j].text;
    var out = TZ.rewriteText(seg.text, { deviceTz: deviceTz, nowMs: nowMs, before: before, after: after });
    if (out !== seg.text) { node.nodeValue = out; changed = true; }
  });
  return changed ? doc.body.innerHTML : null;
}

// A failed read or write leaves the draft as it was and says so on the draft.
function failed(item, event) {
  item.notificationMessages.replaceAsync("otz-error", {
    type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
    message: "Time zones were not added."
  }, function () { event.completed(); });
}

function addTimeZones(event) {
  var item = Office.context.mailbox.item, ok = Office.AsyncResultStatus.Succeeded;
  item.body.getAsync(Office.CoercionType.Html, function (r) {
    if (r.status !== ok) { failed(item, event); return; }
    var html = rewriteHtml(r.value, Intl.DateTimeFormat().resolvedOptions().timeZone, Date.now());
    if (html === null) { event.completed(); return; }
    item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, function (w) {
      if (w.status !== ok) failed(item, event);
      else event.completed();
    });
  });
}

Office.onReady(function () {});
Office.actions.associate("addTimeZones", addTimeZones);

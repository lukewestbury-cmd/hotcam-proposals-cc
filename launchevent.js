/*
 * Hotcam – Proposals CC check (Outlook Smart Alerts, OnMessageSend)
 *
 * On Send: if the subject or the new (non-quoted) body text mentions
 * quote / budget / proposal, and proposals@hotcam.tv isn't already a
 * recipient, it adds proposals@ to CC and shows a prompt:
 *   "Send Anyway" -> sends with proposals@ copied in
 *   "Don't Send"  -> back to the draft; remove the CC if not needed and
 *                    send again (it won't nag a second time for that email)
 *
 * Written as plain ES5 callbacks so it runs in classic Outlook for Windows'
 * JavaScript-only runtime as well as New Outlook (Mac/Windows) and the web.
 */

var TARGET = "proposals@hotcam.tv";
var KEYWORDS = /\b(quot(e|es|ation|ations)|budgets?|proposals?)\b/i;
var SESSION_FLAG = "hotcamProposalsPrompted";

// Text above the first reply/forward header, i.e. what you've actually typed.
function stripQuoted(text) {
  var markers = [
    /^[ \t]*From:[ \t]/im,                    // Outlook reply header
    /^[ \t]*-{2,}[ \t]*Original Message/im,   // older Outlook
    /^[ \t]*On .{1,200} wrote:[ \t]*$/im,     // Apple Mail / Gmail style
    /^[ \t]*_{10,}[ \t]*$/m                   // Outlook separator line
  ];
  var cut = text.length;
  for (var i = 0; i < markers.length; i++) {
    var m = markers[i].exec(text);
    if (m && m.index < cut) cut = m.index;
  }
  return text.slice(0, cut);
}

function mentionsKeywords(subject, body) {
  var text = (subject || "") + "\n" + stripQuoted(body || "");
  text = text.replace(/proposals@hotcam\.tv/gi, ""); // the address itself doesn't count
  return KEYWORDS.test(text);
}

function hasTarget(recipientLists) {
  for (var i = 0; i < recipientLists.length; i++) {
    var list = recipientLists[i] || [];
    for (var j = 0; j < list.length; j++) {
      if ((list[j].emailAddress || "").toLowerCase() === TARGET) return true;
    }
  }
  return false;
}

function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;
  var allow = function () { event.completed({ allowEvent: true }); };
  var ok = function (r) { return r.status === Office.AsyncResultStatus.Succeeded; };

  try {
    item.to.getAsync(function (toRes) {
      item.cc.getAsync(function (ccRes) {
        item.bcc.getAsync(function (bccRes) {
          if (!ok(toRes) || !ok(ccRes) || !ok(bccRes)) return allow();
          if (hasTarget([toRes.value, ccRes.value, bccRes.value])) return allow();

          item.subject.getAsync(function (subRes) {
            item.body.getAsync(Office.CoercionType.Text, function (bodyRes) {
              if (!ok(subRes) || !ok(bodyRes)) return allow();
              if (!mentionsKeywords(subRes.value, bodyRes.value)) return allow();

              // Already prompted on this email and the user removed the CC: respect that.
              item.sessionData.getAsync(SESSION_FLAG, function (flagRes) {
                if (ok(flagRes) && flagRes.value) return allow();

                item.cc.addAsync([TARGET], function (addRes) {
                  if (!ok(addRes)) return allow();
                  item.sessionData.setAsync(SESSION_FLAG, "1", function () {
                    event.completed({
                      allowEvent: false,
                      errorMessage:
                        "This looks like a quote, budget or proposal email, so " + TARGET +
                        " has been added to CC.\n\n" +
                        "Send Anyway: send it with proposals@ copied in.\n" +
                        "Don't Send: go back to check it. If proposals@ isn't needed, remove it and send again."
                    });
                  });
                });
              });
            });
          });
        });
      });
    });
  } catch (e) {
    allow(); // never block sending because of an add-in error
  }
}

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
if (typeof module !== "undefined") {
  module.exports = { stripQuoted: stripQuoted, mentionsKeywords: mentionsKeywords, hasTarget: hasTarget, onMessageSendHandler: onMessageSendHandler };
}

/*
 * Hotcam – Proposals CC check (Outlook Smart Alerts, OnMessageSend)
 *
 * On Send it checks for a quote / estimate / budget / proposal in:
 *   - the subject
 *   - the new (non-quoted) text you've typed
 *   - the most recent email in the thread you're replying to (not older history)
 *   - attachment file names (e.g. "Job - Quote - V1.0.pdf")
 * If one matches and proposals@hotcam.tv isn't already a recipient, it adds
 * proposals@ to CC and shows a prompt:
 *   "Send Anyway" -> sends with proposals@ copied in
 *   "Don't Send"  -> back to the draft; remove the CC if not needed and
 *                    send again (it won't nag a second time for that email)
 *
 * Written as plain ES5 callbacks so it runs in classic Outlook for Windows'
 * JavaScript-only runtime as well as New Outlook (Mac/Windows) and the web.
 */

var TARGET = "proposals@hotcam.tv";
// "rate" alone is too broad in broadcast (frame rate, bit rate), so only priced rates count.
var KEYWORDS = /\b(quot(e|es|ed|ing|ation|ations)|budgets?|proposals?|estimates?|pricing|prices?|costings?|rate ?cards?|(day|hire|crew|kit) rates?)\b/i;
var ATTACHMENT_NAMES = /(quot|estimat|proposal|budget|costing|pricing|rate ?card)/i;
var SESSION_FLAG = "hotcamProposalsPrompted";
var MAX_QUOTED_CHARS = 5000;

// Reply/forward header markers, in the forms Outlook, Apple Mail and Gmail write them.
var MARKERS = [
  /^[ \t]*From:[ \t]/gim,                    // Outlook reply header
  /^[ \t]*-{2,}[ \t]*Original Message/gim,   // older Outlook
  /^[ \t]*On .{1,200} wrote:[ \t]*$/gim,     // Apple Mail / Gmail style
  /^[ \t]*_{10,}[ \t]*$/gm                   // Outlook separator line
];
var HEADER_LINE = /^[ \t]*(From|Sent|To|Cc|Bcc|Subject|Date):|^[ \t]*-{2,}[ \t]*Original Message|^[ \t]*On .{1,200} wrote:[ \t]*$|^[ \t]*_{10,}[ \t]*$/i;

function markerPositions(text) {
  var found = [];
  for (var i = 0; i < MARKERS.length; i++) {
    var re = MARKERS[i];
    re.lastIndex = 0;
    var m;
    while ((m = re.exec(text)) !== null) {
      found.push(m.index);
      if (re.lastIndex === m.index) re.lastIndex++;
    }
  }
  return found.sort(function (a, b) { return a - b; });
}

// Text above the first reply/forward header, i.e. what you've actually typed.
function stripQuoted(text) {
  var found = markerPositions(text);
  return found.length ? text.slice(0, found[0]) : text;
}

// Lines of a quoted block that aren't header lines, blank, or nested (">>") quotes.
function contentLines(block) {
  var lines = block.split(/\r?\n/);
  var out = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (/^[ \t]*$/.test(line) || HEADER_LINE.test(line) || /^[ \t]*>[ \t]*>/.test(line)) continue;
    out.push(line);
  }
  return out;
}

// The most recent email in the thread: from the first header down to the next
// header that follows some real content. A "____" line straight before "From:"
// counts as one header, not two.
function latestQuoted(text) {
  var found = markerPositions(text);
  if (!found.length) return "";
  var start = found[0];
  for (var k = 1; k < found.length; k++) {
    var block = text.slice(start, found[k]);
    if (contentLines(block).length) return contentLines(block).join("\n").slice(0, MAX_QUOTED_CHARS);
  }
  return contentLines(text.slice(start, start + MAX_QUOTED_CHARS)).join("\n");
}

function mentionsKeywords(subject, body) {
  body = body || "";
  var text = (subject || "") + "\n" + stripQuoted(body) + "\n" + latestQuoted(body);
  text = text.replace(/proposals@hotcam\.tv/gi, ""); // the address itself doesn't count
  return KEYWORDS.test(text);
}

function attachmentsLookLikeQuote(attachments) {
  for (var i = 0; i < (attachments || []).length; i++) {
    var a = attachments[i] || {};
    if (a.isInline) continue;
    if (ATTACHMENT_NAMES.test(a.name || "")) return true;
  }
  return false;
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
  var finished = false;
  var finish = function (result) {
    if (finished) return;
    finished = true;
    event.completed(result);
  };
  var allow = function () { finish({ allowEvent: true }); };
  var ok = function (r) { return r && r.status === Office.AsyncResultStatus.Succeeded; };

  var block = function () {
    // Add the CC and set the "already asked" flag at the same time.
    var pending = 2, added = false;
    var done = function () {
      if (--pending) return;
      if (!added) return allow();
      finish({
        allowEvent: false,
        errorMessage:
          "This looks like a quote, estimate or proposal email, so " + TARGET +
          " has been added to CC.\n\n" +
          "Send Anyway: send it with proposals@ copied in.\n" +
          "Don't Send: go back to check it. If proposals@ isn't needed, remove it and send again."
      });
    };
    item.cc.addAsync([TARGET], function (r) { added = ok(r); done(); });
    item.sessionData.setAsync(SESSION_FLAG, "1", function () { done(); });
  };

  try {
    // Ask Outlook for everything at once rather than one after another.
    var res = {};
    var calls = {
      to: function (cb) { item.to.getAsync(cb); },
      cc: function (cb) { item.cc.getAsync(cb); },
      bcc: function (cb) { item.bcc.getAsync(cb); },
      subject: function (cb) { item.subject.getAsync(cb); },
      body: function (cb) { item.body.getAsync(Office.CoercionType.Text, cb); },
      flag: function (cb) { item.sessionData.getAsync(SESSION_FLAG, cb); },
      attachments: function (cb) {
        if (typeof item.getAttachmentsAsync !== "function") return cb(null);
        item.getAttachmentsAsync(cb);
      }
    };
    var names = ["to", "cc", "bcc", "subject", "body", "flag", "attachments"];
    var outstanding = names.length;

    var decide = function () {
      if (!ok(res.to) || !ok(res.cc) || !ok(res.bcc)) return allow();
      if (hasTarget([res.to.value, res.cc.value, res.bcc.value])) return allow();
      // Already prompted on this email and the user removed the CC: respect that.
      if (ok(res.flag) && res.flag.value) return allow();

      var keywordHit = ok(res.subject) && ok(res.body) && mentionsKeywords(res.subject.value, res.body.value);
      var attachmentHit = ok(res.attachments) && attachmentsLookLikeQuote(res.attachments.value);
      if (keywordHit || attachmentHit) return block();
      allow();
    };

    for (var i = 0; i < names.length; i++) {
      (function (name) {
        calls[name](function (r) {
          res[name] = r;
          if (--outstanding === 0) {
            try { decide(); } catch (e) { allow(); }
          }
        });
      })(names[i]);
    }
  } catch (e) {
    allow(); // never block sending because of an add-in error
  }
}

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
if (typeof module !== "undefined") {
  module.exports = {
    stripQuoted: stripQuoted,
    latestQuoted: latestQuoted,
    mentionsKeywords: mentionsKeywords,
    attachmentsLookLikeQuote: attachmentsLookLikeQuote,
    hasTarget: hasTarget,
    onMessageSendHandler: onMessageSendHandler
  };
}

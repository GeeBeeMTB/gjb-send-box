/*
 * GJB send box — runs when Send is pressed in Outlook.
 * Decides whether to let the email go, or to stop and offer the "Log time" panel.
 * Self-contained on purpose: classic Outlook loads only this one file.
 */
var GJB = {"jobs": [{"n": "003", "name": "003 Hayes Green, Cadishead", "match": ["Cadishead", "Hayes Green"], "fixedFee": false}, {"n": "005", "name": "005 Rolls Ave, Crewe", "match": ["Rolls Ave", "Crewe", "Grosvenor"], "fixedFee": false}, {"n": "058", "name": "058 Buckshaw Village", "match": ["Buckshaw", "Ordnance"], "fixedFee": false}, {"n": "064", "name": "064 The Paddocks, Upton Rocks", "match": ["Upton Rocks", "Paddocks"], "fixedFee": false}, {"n": "077", "name": "077 Crown Park / Regents Grange, Chester", "match": ["Crown Park", "Regents Grange", "Saighton"], "fixedFee": false}, {"n": "100", "name": "100 Omega 3 South, Warrington", "match": ["Omega"], "fixedFee": false}, {"n": "101", "name": "101 Gibfield, Atherton", "match": ["Gibfield", "Atherton", "Cottonfields"], "fixedFee": false}, {"n": "102", "name": "102 Huyton ph3/ph4", "match": ["Huyton", "Melbury", "St Davids", "Canterbury Park"], "fixedFee": false}, {"n": "108", "name": "108 Richmond Drive, Leigh", "match": ["Richmond"], "fixedFee": false}, {"n": "109", "name": "109 Willow Park, Middleton", "match": ["Willow Park", "Middleton"], "fixedFee": false}, {"n": "110", "name": "110 Castlefields, Runcorn", "match": ["Castlefields", "Runcorn", "Lakeside"], "fixedFee": false}, {"n": "111", "name": "111 West Park / Sweetbriar, Darlington", "match": ["Darlington", "Sweetbriar"], "fixedFee": true}, {"n": "112", "name": "112 Millstone View, Penymynydd", "match": ["Millstone"], "fixedFee": false}, {"n": "113", "name": "113 Hollington Grange, Stoke", "match": ["Hollington"], "fixedFee": false}, {"n": "114", "name": "114 Bucknall Grange, Stoke", "match": ["Bucknall"], "fixedFee": false}, {"n": "801", "name": "801 Keepmoat (client-wide)", "match": ["Keepmoat client-wide", "Keepmoat lights out", "Keepmoat lights-out walkover", "Keepmoat evening walkover"], "fixedFee": false}, {"n": "802", "name": "802 Redrow NW (client-wide)", "match": ["Redrow NW client-wide", "Redrow lights out", "Redrow lights-out walkover", "Redrow evening walkover", "Buckshaw lights out"], "fixedFee": false}, {"n": "803", "name": "803 TW North West (client-wide)", "match": ["TW North West client-wide", "TWNW client-wide", "TWNW lights out", "TW North West evening walkover"], "fixedFee": false}, {"n": "804", "name": "804 TW Manchester (client-wide)", "match": ["TW Manchester client-wide", "TWM client-wide", "TW Manchester lights out"], "fixedFee": false}], "clientDomains": ["keepmoat.com", "redrow.co.uk", "taylorwimpey.com"]};

var GJB_NOTE_MARK = "DELETE BEFORE SENDING";
var GJB_CHASE_CATEGORY = "GJB chase";
var GJB_PANE_BUTTON_ID = "gjbOpenPaneButton";

/* Text Glyn wrote himself: everything above the first quoted earlier email. */
function gjbOwnText(body) {
  var text = body || "";
  var cuts = [
    text.search(/\r?\n\s*From:\s.+\r?\n\s*Sent:\s/i),
    text.search(/\r?\n_{20,}/),
    text.search(/\r?\n-{3,}\s*Original Message\s*-{3,}/i),
    text.search(/\r?\nOn .{5,120} wrote:\s*\r?\n/i)
  ].filter(function (i) { return i >= 0; });
  if (cuts.length === 0) { return text; }
  return text.slice(0, Math.min.apply(null, cuts));
}

function gjbFindJob(subject) {
  var s = (subject || "").toLowerCase();
  var hits = [];
  GJB.jobs.forEach(function (job) {
    // A job number only counts at the very start of the subject ("110 Castlefields ..."),
    // because references such as "S102" or "s104" would otherwise look like job numbers.
    var numberHit = new RegExp("^\\s*((re|fw|fwd):\\s*)*" + job.n + "\\s").test(s);
    var wordHit = job.match.some(function (m) { return s.indexOf(m.toLowerCase()) >= 0; });
    if (numberHit || wordHit) { hits.push(job); }
  });
  return hits.length === 1 ? hits[0] : null;
}

function gjbAnyJobWord(subject) {
  var s = (subject || "").toLowerCase();
  return GJB.jobs.some(function (job) {
    return job.match.some(function (m) { return s.indexOf(m.toLowerCase()) >= 0; });
  });
}

function gjbHasClientRecipient(recipients) {
  return (recipients || []).some(function (r) {
    var addr = ((r && r.emailAddress) || "").toLowerCase();
    return GJB.clientDomains.some(function (d) { return addr.slice(-(d.length + 1)) === "@" + d; });
  });
}

function gjbGet(fn, fallback) {
  return new Promise(function (resolve) {
    try {
      fn(function (result) {
        resolve(result && result.status === Office.AsyncResultStatus.Succeeded ? result.value : fallback);
      });
    } catch (e) {
      resolve(fallback);
    }
  });
}

function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;
  Promise.all([
    gjbGet(function (cb) { item.sessionData.getAsync("gjbDone", cb); }, ""),
    gjbGet(function (cb) { item.subject.getAsync(cb); }, ""),
    gjbGet(function (cb) { item.body.getAsync(Office.CoercionType.Text, cb); }, ""),
    gjbGet(function (cb) { item.to.getAsync(cb); }, []),
    gjbGet(function (cb) { item.cc.getAsync(cb); }, []),
    gjbGet(function (cb) { item.categories.getAsync(cb); }, [])
  ]).then(function (v) {
    var done = v[0] === "1";
    var subject = v[1];
    var noteStill = gjbOwnText(v[2]).indexOf(GJB_NOTE_MARK) >= 0;
    var recipients = (v[3] || []).concat(v[4] || []);
    var isChase = (v[5] || []).some(function (c) { return c && c.displayName === GJB_CHASE_CATEGORY; });

    // The panel has already been completed for this email: let it go.
    if (done) { event.completed({ allowEvent: true }); return; }

    var noteText = "Your internal \"DELETE BEFORE SENDING\" note is still in this email. The people it is going to would see it.";
    var promptUser = Office.MailboxEnums && Office.MailboxEnums.SendModeOverride
      ? Office.MailboxEnums.SendModeOverride.PromptUser : "promptUser";

    // A chase: time is logged automatically, so only the note is checked.
    if (isChase) {
      if (noteStill) {
        event.completed({ allowEvent: false, errorMessage: noteText, sendModeOverride: promptUser });
      } else {
        event.completed({ allowEvent: true });
      }
      return;
    }

    var job = gjbFindJob(subject);
    var looksLikeJobMail = !!job || gjbAnyJobWord(subject) || gjbHasClientRecipient(recipients);

    if (!looksLikeJobMail) {
      if (noteStill) {
        event.completed({ allowEvent: false, errorMessage: noteText, sendModeOverride: promptUser });
      } else {
        event.completed({ allowEvent: true });
      }
      return;
    }

    var message = (job ? job.name + ". " : "") + "Log your time on this before it goes?" + (noteStill ? " " + noteText : "");
    event.completed({
      allowEvent: false,
      errorMessage: message,
      cancelLabel: "Log time",
      commandId: GJB_PANE_BUTTON_ID,
      contextData: JSON.stringify({ job: job ? job.n : "", note: noteStill }),
      sendModeOverride: promptUser
    });
  }).catch(function () {
    // Never trap an email because of a fault in the add-in.
    event.completed({ allowEvent: true });
  });
}

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}

/*
 * GJB send box — the side panel.
 * Pick the job and the time, optionally set a chase, then send.
 * The choice is saved on the email (the add-in's own properties) so the system can act on it.
 */
(function () {
  var NOTE_MARK = "DELETE BEFORE SENDING";
  var TIMES = [[0, "None"], [10, "10 min"], [15, "15 min"], [30, "30 min"], [45, "45 min"], [60, "1 hr"], [90, "1 hr 30"], [120, "2 hr"]];
  var INTERVALS = [2, 3, 5];
  var state = { job: "", minutes: null, chase: false, interval: 2, note: false, sending: false };

  function $(id) { return document.getElementById(id); }

  function jobByNumber(n) {
    return GJB.jobs.filter(function (j) { return j.n === n; })[0] || null;
  }

  function guessJob(subject) {
    var s = (subject || "").toLowerCase();
    var hits = GJB.jobs.filter(function (job) {
      var numberHit = new RegExp("^\\s*((re|fw|fwd):\\s*)*" + job.n + "\\s").test(s);
      var wordHit = job.match.some(function (m) { return s.indexOf(m.toLowerCase()) >= 0; });
      return numberHit || wordHit;
    });
    return hits.length === 1 ? hits[0].n : "";
  }

  function ownText(body) {
    var text = body || "";
    var cuts = [
      text.search(/\r?\n\s*From:\s.+\r?\n\s*Sent:\s/i),
      text.search(/\r?\n_{20,}/),
      text.search(/\r?\n-{3,}\s*Original Message\s*-{3,}/i),
      text.search(/\r?\nOn .{5,120} wrote:\s*\r?\n/i)
    ].filter(function (i) { return i >= 0; });
    return cuts.length ? text.slice(0, Math.min.apply(null, cuts)) : text;
  }

  function render() {
    var job = jobByNumber(state.job);
    var fixedFee = !!(job && job.fixedFee);
    var notAJob = state.job === "none";

    $("job").value = state.job;
    $("fixed-fee-hint").hidden = !fixedFee;
    $("time-group").hidden = notAJob;
    $("times").hidden = fixedFee;
    $("note-banner").hidden = !state.note;
    $("interval-row").hidden = !state.chase;
    $("chase").checked = state.chase;
    $("chase").disabled = notAJob;

    Array.prototype.forEach.call($("times").children, function (b) {
      b.setAttribute("aria-pressed", String(Number(b.dataset.minutes) === state.minutes));
    });
    Array.prototype.forEach.call($("intervals").children, function (b) {
      b.setAttribute("aria-pressed", String(Number(b.dataset.days) === state.interval));
    });

    var ready = notAJob || (!!job && (fixedFee || state.minutes !== null));
    var lines = [];
    if (!state.job) {
      lines.push("Pick the job to send.");
    } else if (notAJob) {
      lines.push("Nothing is logged or filed for this one.");
    } else if (!ready) {
      lines.push("Pick a time to send.");
    } else {
      if (fixedFee) { lines.push("No time logged (fixed fee)."); }
      else if (state.minutes === 0) { lines.push("No time logged."); }
      else { lines.push("Logs " + labelFor(state.minutes) + " to " + job.name + ", dated today."); }
      lines.push("Files the email to the job folder.");
      if (state.chase) { lines.push("Chases every " + state.interval + " working days until there is a reply."); }
    }
    $("summary").textContent = lines.join("\n");
    $("send").disabled = !ready || state.sending;
    $("send").textContent = state.note ? "Send with the note left in" : "Send email";
  }

  function labelFor(minutes) {
    var t = TIMES.filter(function (x) { return x[0] === minutes; })[0];
    return t ? t[1] : minutes + " min";
  }

  function build() {
    var select = $("job");
    var blank = document.createElement("option");
    blank.value = ""; blank.textContent = "Choose a job";
    select.appendChild(blank);
    GJB.jobs.forEach(function (job) {
      var o = document.createElement("option");
      o.value = job.n; o.textContent = job.name;
      select.appendChild(o);
    });
    var none = document.createElement("option");
    none.value = "none"; none.textContent = "Not a job email";
    select.appendChild(none);
    select.addEventListener("change", function () { state.job = select.value; render(); });

    TIMES.forEach(function (t) {
      var b = document.createElement("button");
      b.type = "button"; b.textContent = t[1]; b.dataset.minutes = String(t[0]);
      b.addEventListener("click", function () { state.minutes = t[0]; render(); });
      $("times").appendChild(b);
    });
    INTERVALS.forEach(function (d) {
      var b = document.createElement("button");
      b.type = "button"; b.textContent = String(d); b.dataset.days = String(d);
      b.addEventListener("click", function () { state.interval = d; render(); });
      $("intervals").appendChild(b);
    });
    $("chase").addEventListener("change", function () { state.chase = $("chase").checked; render(); });
    $("send").addEventListener("click", send);
  }

  function call(fn) {
    return new Promise(function (resolve, reject) {
      try {
        fn(function (result) {
          if (result && result.status === Office.AsyncResultStatus.Succeeded) { resolve(result.value); }
          else { reject(new Error(result && result.error ? result.error.message : "failed")); }
        });
      } catch (e) { reject(e); }
    });
  }

  /* What was chosen, as it is stored on the email. */
  function stamp() {
    if (state.job === "none") { return { job: "none" }; }
    var job = jobByNumber(state.job);
    var out = { job: state.job };
    if (!(job && job.fixedFee)) { out.minutes = state.minutes; }
    if (state.chase) { out.chaseEvery = state.interval; }
    return out;
  }

  /* The choice is saved on the email as the add-in's own properties. They stay on Glyn's copy
     (and so on the Sent Items copy) and are not sent to the people the email goes to.
     Categories cannot be used: new Outlook does not let an add-in set them on an email being written. */
  function saveStamp(item, data) {
    return call(function (cb) { item.loadCustomPropertiesAsync(cb); }).then(function (props) {
      props.set("gjbJob", String(data.job));
      props.set("gjbMinutes", data.minutes === undefined ? "" : String(data.minutes));
      props.set("gjbChaseEvery", data.chaseEvery === undefined ? "" : String(data.chaseEvery));
      props.set("gjbStamped", new Date().toISOString());
      return call(function (cb) { props.saveAsync(cb); });
    });
  }

  function withTimeout(promise, ms, what) {
    return new Promise(function (resolve, reject) {
      var t = setTimeout(function () { reject(new Error(what + " took too long")); }, ms);
      promise.then(function (v) { clearTimeout(t); resolve(v); }, function (e) { clearTimeout(t); reject(e); });
    });
  }

  function send() {
    if (state.sending) { return; }
    var item = Office.context.mailbox.item;
    state.sending = true;
    $("status").textContent = "Sending\u2026";
    render();

    withTimeout(saveStamp(item, stamp()), 8000, "Saving the time").then(function () {
      return withTimeout(call(function (cb) { item.sessionData.setAsync("gjbDone", "1", cb); }), 5000, "Marking the email");
    }).then(function () {
      if (typeof item.sendAsync !== "function") {
        state.sending = false;
        $("status").textContent = "Saved. Now press Send in the email.";
        render();
        return null;
      }
      return call(function (cb) { item.sendAsync(cb); });
    }).catch(function (e) {
      state.sending = false;
      $("status").textContent = "That did not go through (" + e.message + "). Your email has not been sent. You can press Send in the email and choose \"Send anyway\".";
      render();
    });
  }

  Office.onReady(function () {
    build();
    var item = Office.context.mailbox.item;
    var fromAlert = new Promise(function (resolve) {
      if (!item || typeof item.getInitializationContextAsync !== "function") { resolve(null); return; }
      item.getInitializationContextAsync(function (r) {
        try { resolve(r.status === Office.AsyncResultStatus.Succeeded && r.value ? JSON.parse(r.value) : null); }
        catch (e) { resolve(null); }
      });
    });
    var subject = call(function (cb) { item.subject.getAsync(cb); }).catch(function () { return ""; });
    var body = call(function (cb) { item.body.getAsync(Office.CoercionType.Text, cb); }).catch(function () { return ""; });

    Promise.all([fromAlert, subject, body]).then(function (v) {
      var ctx = v[0];
      state.job = (ctx && ctx.job) || guessJob(v[1]);
      state.note = ownText(v[2]).indexOf(NOTE_MARK) >= 0;
      render();
    });
  });
})();

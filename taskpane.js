(function () {
  "use strict";

  var Core = PolisherCore;
  var SETTINGS_KEYS = ["apiKey", "baseUrl", "model"];
  var state = { mode: null, lastOpts: null, lastResult: null };

  function $(id) { return document.getElementById(id); }
  function show(el, on) { el.classList.toggle("hidden", !on); }

  function setStatus(msg, kind, busy) {
    var el = $("status");
    el.className = "status" + (kind ? " " + kind : "");
    el.innerHTML = (busy ? '<span class="spinner"></span>' : "") + Core.escapeHtml(msg || "");
  }

  function setBusy(on) {
    ["polishBtn", "suggestBtn", "retryBtn", "replaceBtn", "openReplyBtn"].forEach(function (id) {
      $(id).disabled = on;
    });
  }

  // ---------- Settings storage (Office roaming settings, fallback localStorage) ----------
  function roaming() {
    try { return Office.context && Office.context.roamingSettings ? Office.context.roamingSettings : null; } catch { return null; }
  }

  function loadSettings() {
    var s = {};
    var rs = roaming();
    SETTINGS_KEYS.forEach(function (k) {
      var v = null;
      if (rs) v = rs.get("rp_" + k);
      if (v == null) { try { v = window.localStorage.getItem("rp_" + k); } catch { v = null; } }
      s[k] = v || "";
    });
    if (!s.baseUrl) s.baseUrl = Core.DEFAULTS.baseUrl;
    if (!s.model) s.model = Core.DEFAULTS.model;
    return s;
  }

  function saveSettings(s) {
    return new Promise(function (resolve) {
      SETTINGS_KEYS.forEach(function (k) {
        try { window.localStorage.setItem("rp_" + k, s[k] || ""); } catch { /* storage unavailable */ }
      });
      var rs = roaming();
      if (!rs) return resolve();
      SETTINGS_KEYS.forEach(function (k) { rs.set("rp_" + k, s[k] || ""); });
      rs.saveAsync(function () { resolve(); });
    });
  }

  function fillSettingsForm() {
    var s = loadSettings();
    $("apiKey").value = s.apiKey;
    $("baseUrl").value = s.baseUrl;
    $("model").value = s.model;
  }

  function readSettingsForm() {
    return {
      apiKey: $("apiKey").value.trim(),
      baseUrl: $("baseUrl").value.trim() || Core.DEFAULTS.baseUrl,
      model: $("model").value.trim() || Core.DEFAULTS.model
    };
  }

  function toggleSettings(open) {
    var el = $("settings");
    var willOpen = typeof open === "boolean" ? open : el.classList.contains("hidden");
    show(el, willOpen);
    $("settingsToggle").setAttribute("aria-expanded", String(willOpen));
  }

  // ---------- Office helpers ----------
  function asyncCall(fn) {
    return new Promise(function (resolve, reject) {
      fn(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
        else reject(new Error((r.error && r.error.message) || "Outlook could not complete the request."));
      });
    });
  }

  function htmlToText(html) {
    try {
      var doc = new DOMParser().parseFromString("<div>" + html + "</div>", "text/html");
      doc.querySelectorAll("style,script,head,title").forEach(function (n) { n.remove(); });
      doc.querySelectorAll("br").forEach(function (n) { n.replaceWith("\n"); });
      doc.querySelectorAll("p,div,li,tr,h1,h2,h3,h4,h5,h6").forEach(function (n) { n.append("\n"); });
      return (doc.body.textContent || "").replace(/\u00a0/g, " ").replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
    } catch {
      return Core.htmlToTextBasic(html);
    }
  }

  function getBodyType(item) {
    if (!item.body.getTypeAsync) return Promise.resolve(Office.CoercionType.Html);
    return asyncCall(function (cb) { item.body.getTypeAsync(cb); }).catch(function () { return Office.CoercionType.Html; });
  }

  function fmtRecipients(list) {
    return (list || []).map(function (r) {
      return r.displayName && r.displayName !== r.emailAddress ? r.displayName + " <" + r.emailAddress + ">" : r.emailAddress;
    }).join(", ");
  }

  // Gather draft + thread context in compose mode.
  async function gatherCompose() {
    var item = Office.context.mailbox.item;
    var subject = await asyncCall(function (cb) { item.subject.getAsync(cb); }).catch(function () { return ""; });
    var to = await asyncCall(function (cb) { item.to.getAsync(cb); }).catch(function () { return []; });
    var type = await getBodyType(item);
    var draft = "";
    var thread = "";
    if (type === Office.CoercionType.Html) {
      var html = await asyncCall(function (cb) { item.body.getAsync(Office.CoercionType.Html, cb); });
      var a = Core.analyzeHtml(html);
      draft = htmlToText(html.slice(a.start, a.userEnd));
      if (a.quoteStart !== -1) thread = htmlToText(html.slice(a.quoteStart, a.end));
    }
    if (!thread || !draft) {
      // Fallback: plain-text body and text quote markers.
      var text = await asyncCall(function (cb) { item.body.getAsync(Office.CoercionType.Text, cb); }).catch(function () { return ""; });
      var split = Core.splitQuotedText(text);
      if (!draft) draft = split.userText;
      if (!thread) thread = split.quotedText;
    }
    return { mode: "compose", subject: subject, to: fmtRecipients(to), draft: draft, thread: thread };
  }

  async function gatherRead() {
    var item = Office.context.mailbox.item;
    var text = await asyncCall(function (cb) { item.body.getAsync(Office.CoercionType.Text, cb); });
    var from = item.from ? fmtRecipients([item.from]) : "";
    return {
      mode: "read",
      subject: item.subject || "",
      from: from,
      thread: text,
      draft: $("readNotes").value.trim(),
      extra: $("readExtra").value.trim()
    };
  }

  // ---------- Actions ----------
  async function run(opts) {
    var settings = loadSettings();
    if (!settings.apiKey) {
      toggleSettings(true);
      setStatus("Add your xAI API key in Settings first.", "error");
      $("apiKey").focus();
      return;
    }
    setBusy(true);
    setStatus(opts.mode === "read" ? "Writing a suggested reply..." : "Polishing your reply...", "", true);
    try {
      var result = await Core.polish(opts, settings);
      state.lastOpts = opts;
      state.lastResult = result;
      renderResult(result);
      setStatus("");
    } catch (e) {
      setStatus(e.message || String(e), "error");
    } finally {
      setBusy(false);
    }
  }

  function renderResult(r) {
    var badge = $("modeBadge");
    badge.textContent = r.isSales ? "Sales mode" : "Professional mode";
    badge.className = "badge " + (r.isSales ? "sales" : "pro");
    $("category").textContent = r.category || "";
    $("output").value = r.reply;
    $("notes").textContent = r.notes ? "Check: " + r.notes : "";
    show($("notes"), !!r.notes);
    show($("result"), true);
    $("result").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function onPolish() {
    try {
      var opts = await gatherCompose();
      opts.extra = $("extra").value.trim();
      await run(opts);
    } catch (e) {
      setStatus(e.message || String(e), "error");
    }
  }

  async function onSuggest() {
    try {
      await run(await gatherRead());
    } catch (e) {
      setStatus(e.message || String(e), "error");
    }
  }

  async function onRetry() {
    if (!state.lastOpts) return;
    if (state.mode === "compose") return onPolish();
    return onSuggest();
  }

  async function onReplace() {
    var item = Office.context.mailbox.item;
    var text = $("output").value.trim();
    if (!text) return;
    setBusy(true);
    try {
      var type = await getBodyType(item);
      var kept;
      if (type === Office.CoercionType.Html) {
        var html = await asyncCall(function (cb) { item.body.getAsync(Office.CoercionType.Html, cb); });
        var res = Core.replaceUserPortion(html, Core.textToHtml(text));
        await asyncCall(function (cb) { item.body.setAsync(res.html, { coercionType: Office.CoercionType.Html }, cb); });
        kept = res.preservedTail;
      } else {
        var body = await asyncCall(function (cb) { item.body.getAsync(Office.CoercionType.Text, cb); });
        var split = Core.splitQuotedText(body);
        var idx = split.quotedText ? body.replace(/\r\n/g, "\n").indexOf(split.quotedText) : -1;
        var tail = idx !== -1 ? "\n\n" + body.replace(/\r\n/g, "\n").slice(idx) : "";
        await asyncCall(function (cb) { item.body.setAsync(text + tail, { coercionType: Office.CoercionType.Text }, cb); });
        kept = !!tail;
      }
      setStatus(kept ? "Draft replaced. Your signature / earlier emails below were kept." : "Draft replaced.", "ok");
    } catch (e) {
      setStatus("Could not replace the draft: " + (e.message || e) + " Use Copy and paste instead.", "error");
    } finally {
      setBusy(false);
    }
  }

  function onOpenReply() {
    var text = $("output").value.trim();
    if (!text) return;
    try {
      Office.context.mailbox.item.displayReplyForm(Core.textToHtml(text));
    } catch (e) {
      setStatus("Could not open a reply: " + (e.message || e) + " Use Copy instead.", "error");
    }
  }

  async function onCopy() {
    var text = $("output").value;
    var ok = false;
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(text);
        ok = true;
      }
    } catch { ok = false; }
    if (!ok) {
      var ta = $("output");
      ta.focus();
      ta.select();
      try { ok = document.execCommand("copy"); } catch { ok = false; }
    }
    setStatus(ok ? "Copied to clipboard." : "Select the text and press Ctrl+C / Cmd+C to copy.", ok ? "ok" : "error");
  }

  async function onSaveSettings() {
    var s = readSettingsForm();
    var st = $("settingsStatus");
    await saveSettings(s);
    if (!s.apiKey) { st.className = "status error"; st.textContent = "Saved, but no API key entered."; return; }
    st.className = "status";
    st.innerHTML = '<span class="spinner"></span>Saved. Testing connection...';
    try {
      await Core.polish({ mode: "compose", subject: "Test", draft: "hi thanks for the info, talk soon" }, s);
      st.className = "status ok";
      st.textContent = "Saved. Connection works.";
    } catch (e) {
      st.className = "status error";
      st.textContent = "Saved, but the test failed: " + (e.message || e);
    }
  }

  function onResetSettings() {
    $("baseUrl").value = Core.DEFAULTS.baseUrl;
    $("model").value = Core.DEFAULTS.model;
  }

  function wire() {
    $("settingsToggle").addEventListener("click", function () { toggleSettings(); });
    $("toggleKey").addEventListener("click", function () {
      var k = $("apiKey");
      k.type = k.type === "password" ? "text" : "password";
      this.textContent = k.type === "password" ? "Show" : "Hide";
    });
    $("saveSettings").addEventListener("click", onSaveSettings);
    $("resetSettings").addEventListener("click", onResetSettings);
    $("polishBtn").addEventListener("click", onPolish);
    $("suggestBtn").addEventListener("click", onSuggest);
    $("retryBtn").addEventListener("click", onRetry);
    $("replaceBtn").addEventListener("click", onReplace);
    $("openReplyBtn").addEventListener("click", onOpenReply);
    $("copyBtn").addEventListener("click", onCopy);
  }

  function detectMode() {
    var item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item) return null;
    // In compose mode, subject is an object with getAsync; in read mode it is a string.
    return item.body && typeof item.subject === "object" && item.subject.getAsync ? "compose" : "read";
  }

  function applyMode() {
    state.mode = detectMode();
    show($("composeView"), state.mode === "compose");
    show($("readView"), state.mode === "read");
    show($("unsupported"), !state.mode);
    show($("replaceBtn"), state.mode === "compose");
    show($("replaceHint"), state.mode === "compose");
    show($("openReplyBtn"), state.mode === "read");
    show($("result"), false);
    setStatus("");
  }

  Office.onReady(function () {
    wire();
    fillSettingsForm();
    applyMode();
    if (!loadSettings().apiKey) toggleSettings(true);
    // Pinned task pane: refresh when the user switches to another email.
    try {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, applyMode);
    } catch { /* not supported on this client */ }
  });
})();

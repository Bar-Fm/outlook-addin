/* FM-ENG | Outlook add-in "FM - mail assistant" v1.1 | pilot (single user)
   The API key is kept in Office roamingSettings (per mailbox).
   For an office-wide rollout, move the Claude call to a small server. */

var NL = String.fromCharCode(10);
var DEFAULT_MODEL = "claude-sonnet-4-5"; // לא מאומת שזה המודל המעודכן – לבדוק ב-docs.claude.com ולעדכן בהגדרות

var GLOSSARY = [
  "חברת החשמל (לא חח״י ולא חברת חשמל)", "תב״ע", "היתר בנייה", "כתב כמויות", "אומדן", "מפרט טכני",
  "מתח גבוה", "מתח נמוך", "לוח חשמל ראשי", "לוח קומתי", "חדר שנאים", "חדר מונים", "גומחת מונים",
  "מקדם בו-זמניות", "מפל מתח", "זרם קצר", "הארקה", "הגנה מפני ברק", "גנרטור", "מערכת סולארית",
  "אגירה", "עמדות טעינה לרכב חשמלי", "בקרה", "SCADA", "תאורת כבישים", "תשתיות תקשורת", "רמדור"
];

var BASE_RULES = [
  "אתה עורך לשון במשרד תכנון חשמל ״אפ.אמ. הנדסה בע״מ״ (FM-ENG).",
  "כתוב בעברית תקינה, מקצועית וטבעית – שתישמע כמו מהנדס שכתב בעצמו, לא כמו מכונה. קצר וישיר, בלי ביטויים נבובים.",
  "שמור על כל המידע, המספרים, התאריכים, השמות, מספרי הפרויקטים, הערכים והיחידות בדיוק כפי שהם.",
  "אל תוסיף עובדות, ערכים, סעיפי תקן, מועדים או התחייבויות שלא נכתבו.",
  "אל תוסיף חתימה."
];

var TONES = {
  external: "סגנון: רשמי ומכובד לגורם חיצוני. פתיחה מנומסת, תמצית ברורה, וסיום בשורה 'בכבוד רב,' (בלי שם אחריה).",
  internal: "סגנון: ענייני וברור לעמית במשרד, בלי רשמיות יתר.",
  short: "סגנון: קצר ותמציתי ככל האפשר, בלי לאבד מידע.",
  fix: "תקן שגיאות כתיב, דקדוק ופיסוק בלבד. אל תשנה סגנון, מבנה או ניסוח מעבר לנדרש."
};

var source = { mode: null };
var lastOriginal = "";
var isCompose = false;

function $(id) { return document.getElementById(id); }
function setStatus(t, err) { $("status").textContent = t; $("status").className = "status" + (err ? " err" : ""); }
function ok(r) { return r.status === Office.AsyncResultStatus.Succeeded; }
function settings() { return Office.context.roamingSettings; }

function glossaryText() {
  var extra = (settings().get("fm_glossary") || "").split(NL).map(function (s) { return s.trim(); }).filter(Boolean);
  return "מונחים מקצועיים – לכתוב בדיוק כך: " + GLOSSARY.concat(extra).join("; ") + ".";
}
function systemPrompt(extra) { return BASE_RULES.concat([glossaryText()], extra || []).join(NL); }

/* ---------- Claude ---------- */
function callClaude(system, user, maxTokens) {
  var key = settings().get("fm_api_key") || $("apiKey").value.trim();
  if (!key) return Promise.reject(new Error("חסר מפתח API – הזן אותו ב'הגדרות' ושמור."));
  return fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": key,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true"
    },
    body: JSON.stringify({
      model: settings().get("fm_model") || $("model").value.trim() || DEFAULT_MODEL,
      max_tokens: maxTokens || 2000,
      system: system,
      messages: [{ role: "user", content: user }]
    })
  }).then(function (resp) {
    return resp.json().then(function (data) {
      if (!resp.ok) throw new Error((data && data.error && data.error.message) || resp.status);
      return data.content.filter(function (c) { return c.type === "text"; }).map(function (c) { return c.text; }).join("").trim();
    });
  });
}

/* ---------- Office helpers ---------- */
function getSelection() {
  return new Promise(function (res) {
    if (!isCompose) return res("");
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, function (r) { res(ok(r) ? (r.value.data || "") : ""); });
  });
}
function getBody() {
  return new Promise(function (res, rej) {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (r) { ok(r) ? res(r.value) : rej(r.error); });
  });
}
function getComposeField(field) {
  return new Promise(function (res) {
    var f = Office.context.mailbox.item[field];
    if (!f || !f.getAsync) return res(null);
    f.getAsync(function (r) { res(ok(r) ? r.value : null); });
  });
}
function getAttachments() {
  return new Promise(function (res) {
    var item = Office.context.mailbox.item;
    if (!item.getAttachmentsAsync) return res(null);
    item.getAttachmentsAsync(function (r) { res(ok(r) ? r.value : null); });
  });
}
function esc(t) { return t.split("&").join("&amp;").split("<").join("&lt;").split(">").join("&gt;"); }
/* email drafts: RTL, Arial 11 */
function toHtml(t) {
  return '<div dir="rtl" style="font-family:Arial,sans-serif;font-size:11pt;text-align:right">' + esc(t).split(NL).join("<br>") + "</div>";
}
function busy(ids, on) { ids.forEach(function (id) { $(id).disabled = on; }); }

/* ---------- word diff ---------- */
function tokenize(s) {
  var out = [], cur = "", sp = null;
  for (var i = 0; i < s.length; i++) {
    var ws = s.charCodeAt(i) <= 32;
    if (sp === null || ws === sp) cur += s[i]; else { out.push(cur); cur = s[i]; }
    sp = ws;
  }
  if (cur) out.push(cur);
  return out;
}
function renderDiff(a, b) {
  var A = tokenize(a), B = tokenize(b), n = A.length, m = B.length;
  if (n * m > 400000) { $("diff").textContent = b; return; }
  var L = [];
  for (var i = 0; i <= n; i++) { L.push(new Array(m + 1).fill(0)); }
  for (i = n - 1; i >= 0; i--) for (var j = m - 1; j >= 0; j--)
    L[i][j] = A[i] === B[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
  var box = $("diff"); box.innerHTML = "";
  function add(tag, t) { var e = tag ? document.createElement(tag) : document.createTextNode(t); if (tag) e.textContent = t; box.appendChild(e); }
  i = 0; j = 0;
  while (i < n && j < m) {
    if (A[i] === B[j]) { add(null, A[i]); i++; j++; }
    else if (L[i + 1][j] >= L[i][j + 1]) { add("del", A[i]); i++; }
    else { add("ins", B[j]); j++; }
  }
  while (i < n) add("del", A[i++]);
  while (j < m) add("ins", B[j++]);
}

/* ---------- rewrite ---------- */
function doRewrite() {
  getSelection().then(function (sel) {
    sel = sel.trim(); source.mode = "selection";
    return sel ? sel : getBody().then(function (b) { source.mode = "body"; return b.trim(); });
  }).then(function (text) {
    if (!text) throw new Error("לא נמצא טקסט לניסוח.");
    lastOriginal = text;
    busy(["btnRewrite"], true);
    setStatus(source.mode === "selection" ? "מנסח את הטקסט המסומן..." : "מנסח את כל גוף המייל...");
    return callClaude(systemPrompt([TONES[$("tone").value], "החזר את הטקסט המנוסח בלבד, בלי הקדמה ובלי הסברים."]), "נסח את הטקסט הבא:" + NL + NL + text);
  }).then(function (out) {
    $("result").value = out;
    renderDiff(lastOriginal, out);
    $("rwOut").classList.remove("hidden"); showDiff(true);
    setStatus("מוכן. ירוק = נוסף, אדום = הוסר. אפשר לעבור ל'עריכה' לפני ההחלפה.");
  }).catch(function (e) { setStatus("שגיאה: " + e.message, true); })
    .then(function () { busy(["btnRewrite"], false); });
}
function showDiff(d) {
  $("vDiff").classList.toggle("on", d); $("vEdit").classList.toggle("on", !d);
  $("diff").classList.toggle("hidden", !d); $("result").classList.toggle("hidden", d);
  if (d) renderDiff(lastOriginal, $("result").value);
}
function doReplace() {
  var html = toHtml($("result").value), item = Office.context.mailbox.item;
  var cb = function (r) { setStatus(ok(r) ? "הטקסט הוחלף במייל." : "ההחלפה נכשלה: " + r.error.message, !ok(r)); };
  if (source.mode === "selection") item.body.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, cb);
  else item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, cb);
}

/* ---------- pre-send check ---------- */
function doCheck() {
  busy(["btnCheck"], true); setStatus("בודק...");
  Promise.all([getComposeField("subject"), getComposeField("to"), getComposeField("cc"), getBody(), getAttachments()]).then(function (v) {
    var who = function (list) { return (list || []).map(function (x) { return x.displayName || x.emailAddress; }).join(", ") || "—"; };
    var att = v[4] === null ? "לא ידוע" : (v[4].length ? v[4].map(function (a) { return a.name; }).join(", ") : "אין");
    var info = ["נושא: " + (v[0] || "—"), "אל: " + who(v[1]), "עותק: " + who(v[2]), "קבצים מצורפים: " + att, "", "גוף המייל:", v[3]].join(NL);
    return callClaude(systemPrompt([
      "משימה: בדוק מייל לפני שליחה ומצא בעיות בלבד. אל תנסח מחדש.",
      "בדוק: קובץ שהוזכר בטקסט ולא צורף; מספר פרויקט שמופיע בגוף ולא בנושא (במשרד מספר פרויקט = מספר סידורי + שנה, למשל 15026); נושא חסר או לא ברור; פנייה בשם שלא תואם לנמען; טון חריף או לא מקצועי; שגיאות כתיב בולטות; התחייבות למועד או לערך שנראית לא מכוונת.",
      "החזר רשימה קצרה של בעיות, כל אחת בשורה שמתחילה ב-•. אם הכול תקין – כתוב רק: 'לא נמצאו בעיות.'"
    ]), info, 800);
  }).then(function (out) { $("checkOut").textContent = out; $("checkOut").classList.remove("hidden"); setStatus("הבדיקה הסתיימה. ההחלטה שלך."); })
    .catch(function (e) { setStatus("שגיאה: " + e.message, true); })
    .then(function () { busy(["btnCheck"], false); });
}

/* ---------- reply (read mode) ---------- */
function doReply() {
  busy(["btnReply"], true); setStatus("מכין טיוטת תשובה...");
  var item = Office.context.mailbox.item;
  getBody().then(function (body) {
    var from = item.from ? (item.from.displayName || item.from.emailAddress) : "";
    var info = ["מאת: " + from, "נושא: " + (item.subject || ""), "", "המייל (כולל שרשור):", body.slice(0, 20000), "", "הנחיה מבר לתשובה: " + ($("replyHint").value.trim() || "אין – הצע תשובה עניינית ומתאימה")].join(NL);
    return callClaude(systemPrompt([
      TONES[$("replyTone").value],
      "משימה: כתוב טיוטת תשובה בשם בר גפן, אפ.אמ. הנדסה. ענה רק על מה שנשאל ולפי ההנחיה. אל תמציא נתונים, ערכים או מועדים – אם חסר מידע, השאר [להשלמה].",
      "החזר את גוף התשובה בלבד."
    ]), info, 1500);
  }).then(function (out) { $("replyText").value = out; $("replyBox").classList.remove("hidden"); setStatus("הטיוטה מוכנה. ערוך ולחץ 'פתח כתשובה' – שום דבר לא נשלח."); })
    .catch(function (e) { setStatus("שגיאה: " + e.message, true); })
    .then(function () { busy(["btnReply"], false); });
}
function openReply(all) {
  var html = toHtml($("replyText").value), item = Office.context.mailbox.item;
  if (all) item.displayReplyAllForm(html); else item.displayReplyForm(html);
  setStatus("נפתח חלון תשובה לעריכה. בדוק שהחתימה 'בר' נוספה לפני שליחה.");
}

/* ---------- summary (read mode) ---------- */
function doSummary() {
  busy(["btnSummary"], true); setStatus("מסכם...");
  getBody().then(function (body) {
    return callClaude(systemPrompt([
      "משימה: סכם את המייל/השרשור לבר. מבנה: 'בקצרה:' (2–3 שורות), 'החלטות:' ו'משימות פתוחות:' (מי, מה, עד מתי – רק אם נכתב). תבליטים עם •. בלי המצאות."
    ]), "נושא: " + (Office.context.mailbox.item.subject || "") + NL + NL + body.slice(0, 30000), 1000);
  }).then(function (out) { $("sumOut").textContent = out; $("sumOut").classList.remove("hidden"); setStatus(""); })
    .catch(function (e) { setStatus("שגיאה: " + e.message, true); })
    .then(function () { busy(["btnSummary"], false); });
}

/* ---------- settings + tabs ---------- */
function saveSettings() {
  var s = settings();
  s.set("fm_api_key", $("apiKey").value.trim());
  s.set("fm_model", $("model").value.trim() || DEFAULT_MODEL);
  s.set("fm_glossary", $("glossary").value);
  s.saveAsync(function (r) { setStatus(ok(r) ? "ההגדרות נשמרו." : "שמירה נכשלה.", !ok(r)); });
}
function buildTabs(list) {
  var tabs = $("tabs");
  list.forEach(function (t, k) {
    var b = document.createElement("button"); b.className = "tab"; b.textContent = t[1];
    b.onclick = function () {
      Array.prototype.forEach.call(tabs.children, function (x) { x.classList.remove("on"); });
      b.classList.add("on");
      ["rewrite", "check", "reply", "summary"].forEach(function (p) { $("p-" + p).classList.toggle("hidden", p !== t[0]); });
      setStatus("");
    };
    tabs.appendChild(b);
    if (k === 0) b.onclick();
  });
}

Office.onReady(function () {
  var item = Office.context.mailbox.item;
  isCompose = !!(item && item.body && item.body.setAsync && item.subject && item.subject.getAsync);
  var s = settings();
  $("apiKey").value = s.get("fm_api_key") || "";
  $("model").value = s.get("fm_model") || DEFAULT_MODEL;
  $("glossary").value = s.get("fm_glossary") || "";
  $("btnRewrite").onclick = doRewrite;
  $("btnReplace").onclick = doReplace;
  $("btnCopy").onclick = function () { navigator.clipboard.writeText($("result").value); setStatus("הועתק."); };
  $("vDiff").onclick = function () { showDiff(true); };
  $("vEdit").onclick = function () { showDiff(false); };
  $("btnCheck").onclick = doCheck;
  $("btnReply").onclick = doReply;
  $("btnOpenReply").onclick = function () { openReply(false); };
  $("btnOpenReplyAll").onclick = function () { openReply(true); };
  $("btnSummary").onclick = doSummary;
  $("btnSave").onclick = saveSettings;
  buildTabs(isCompose ? [["rewrite", "נסח"], ["check", "בדוק לפני שליחה"]] : [["reply", "ענה"], ["summary", "סכם"]]);
});

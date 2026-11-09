/* FM-ENG | Outlook add-in "FM - mail assistant" v1.2 | pilot (single user)
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

var DEFAULT_TEMPLATES = [
  "קבלת תכניות: אישור קבלת התכניות, נעבור עליהן ונחזור עם הערות",
  "השלמת מסמכים: בקשה להשלים מסמכים או תכניות חסרים כדי להמשיך בתכנון",
  "העברה לחברת החשמל: עדכון שהתכניות הועברו לבדיקה ואישור של חברת החשמל, ונעדכן כשתתקבל תשובה",
  "הצעת מחיר בהמשך: תודה על הפנייה, הצעת מחיר תישלח בהמשך",
  "תיאום פגישה: בקשה לתאם פגישה או שיחת תיאום",
  "הערות לתכניות: העברת הערות התכנון של אפ.אמ. הנדסה לתכניות שהתקבלו"
].join(NL);

var BASE_RULES = [
  "אתה עורך לשון במשרד תכנון חשמל ״אפ.אמ. הנדסה בע״מ״ (FM-ENG).",
  "כתוב בעברית תקינה, מקצועית וטבעית – שתישמע כמו מהנדס שכתב בעצמו, לא כמו מכונה. קצר וישיר, בלי ביטויים נבובים.",
  "שמור על כל המידע, המספרים, התאריכים, השמות, מספרי הפרויקטים, הערכים והיחידות בדיוק כפי שהם.",
  "אל תוסיף עובדות, ערכים טכניים, סעיפי תקן, מועדים או התחייבויות שלא נכתבו. אם חסר פרט – כתוב [להשלמה].",
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
function lines(t) { return (t || "").split(NL).map(function (s) { return s.trim(); }).filter(Boolean); }

function glossaryText() {
  return "מונחים מקצועיים – לכתוב בדיוק כך: " + GLOSSARY.concat(lines(settings().get("fm_glossary"))).join("; ") + ".";
}
function systemPrompt(extra) { return BASE_RULES.concat([glossaryText()], extra || []).join(NL); }

/* ---------- technical values: flag for engineer review ---------- */
var FLAG_PATTERNS = [
  /[0-9]+ ?[xX×] ?[0-9]+ ?A/g,
  /[0-9]+([.,][0-9]+)? ?(kVA|KVA|kW|KW|kV|KV|mA|MW|A|V|W|Hz|אמפר|וולט|קוו״ט|קוו"ט|ממ״ר|ממ"ר|מ״מ|מ"מ|%)(?![A-Za-zא-ת])/g,
  /(ת״י|ת"י|IEC|NFPA|סעיף|תקנה|תקנות) ?[0-9][0-9.-]*/g,
  /מקדם[^.,;:]{0,30}?[0-9]+([.,][0-9]+)?/g
];
function findValues(t) {
  var found = [];
  FLAG_PATTERNS.forEach(function (re) {
    (t.match(re) || []).forEach(function (m) { m = m.trim(); if (found.indexOf(m) < 0) found.push(m); });
  });
  return found.filter(function (v) { return !found.some(function (o) { return o !== v && o.indexOf(v) >= 0; }); });
}
function norm(s) { return s.split(" ").join("").split(",").join(".").toLowerCase(); }
function showFlags(boxId, out, src) {
  var box = $(boxId), vals = findValues(out);
  if (!vals.length) { box.classList.add("hidden"); return; }
  var srcN = norm(src || "");
  var fresh = vals.filter(function (v) { return srcN.indexOf(norm(v)) < 0; });
  box.innerHTML = "";
  var b = document.createElement("b"); b.textContent = "ערכים טכניים – לבדיקת מהנדס לפני שליחה: "; box.appendChild(b);
  box.appendChild(document.createTextNode(vals.join(" | ")));
  if (fresh.length && src !== null) {
    box.appendChild(document.createElement("br"));
    var b2 = document.createElement("b"); b2.textContent = "לא הופיעו בטקסט המקורי: "; box.appendChild(b2);
    box.appendChild(document.createTextNode(fresh.join(" | ")));
  }
  box.classList.remove("hidden");
}

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
function subjectText() { var s = Office.context.mailbox.item.subject; return typeof s === "string" ? s : ""; }
function esc(t) { return t.split("&").join("&amp;").split("<").join("&lt;").split(">").join("&gt;"); }
/* email drafts: RTL, Arial 11 */
function toHtml(t) {
  return '<div dir="rtl" style="font-family:Arial,sans-serif;font-size:11pt;text-align:right">' + esc(t).split(NL).join("<br>") + "</div>";
}
function insertAtCursor(text) {
  Office.context.mailbox.item.body.setSelectedDataAsync(toHtml(text), { coercionType: Office.CoercionType.Html }, function (r) {
    setStatus(ok(r) ? "הוכנס למייל. בדוק שהחתימה 'בר' קיימת לפני שליחה." : "ההכנסה נכשלה: " + r.error.message, !ok(r));
  });
}
function busy(ids, on) { ids.forEach(function (id) { $(id).disabled = on; }); }
function run(btn, statusText, work) {
  busy([btn], true); setStatus(statusText);
  return Promise.resolve().then(work)
    .catch(function (e) { setStatus("שגיאה: " + e.message, true); })
    .then(function () { busy([btn], false); });
}

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
  return run("btnRewrite", "מנסח...", function () {
    return getSelection().then(function (sel) {
      sel = sel.trim(); source.mode = "selection";
      return sel ? sel : getBody().then(function (b) { source.mode = "body"; return b.trim(); });
    }).then(function (text) {
      if (!text) throw new Error("לא נמצא טקסט לניסוח.");
      lastOriginal = text;
      return callClaude(systemPrompt([TONES[$("tone").value], "החזר את הטקסט המנוסח בלבד, בלי הקדמה ובלי הסברים."]), "נסח את הטקסט הבא:" + NL + NL + text);
    }).then(function (out) {
      $("result").value = out;
      $("rwOut").classList.remove("hidden"); showDiff(true);
      showFlags("rwFlags", out, lastOriginal);
      setStatus("מוכן. ירוק = נוסף, אדום = הוסר. אפשר לעבור ל'עריכה' לפני ההחלפה.");
    });
  });
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

/* ---------- points -> email ---------- */
function doPoints() {
  var pts = $("points").value.trim();
  if (!pts) { setStatus("כתוב לפחות נקודה אחת.", true); return; }
  return run("btnPoints", "כותב מייל...", function () {
    return callClaude(systemPrompt([
      TONES[$("ptTone").value],
      "משימה: כתוב גוף מייל מלא בשם בר גפן, אפ.אמ. הנדסה, על בסיס הנקודות בלבד. אל תוסיף נושאים שלא נכתבו.",
      "החזר את גוף המייל בלבד."
    ]), "נושא המייל: " + (subjectText() || "—") + NL + "נקודות:" + NL + pts, 1200).then(function (out) {
      $("ptText").value = out; $("ptOut").classList.remove("hidden");
      showFlags("ptFlags", out, pts);
      setStatus("מוכן. ערוך ולחץ 'הכנס למייל' – ייכנס במקום הסמן.");
    });
  });
}

/* ---------- templates ---------- */
function templates() { return lines(settings().get("fm_templates") || DEFAULT_TEMPLATES); }
function fillTemplates() {
  var sel = $("tplSel"); sel.innerHTML = "";
  templates().forEach(function (t, k) {
    var o = document.createElement("option"); o.value = k; o.textContent = t.split(":")[0].trim(); sel.appendChild(o);
  });
  $("templates").value = settings().get("fm_templates") || DEFAULT_TEMPLATES;
}
function doTemplate() {
  var t = templates()[+$("tplSel").value];
  if (!t) { setStatus("אין תבניות – הוסף בהגדרות.", true); return; }
  var details = $("tplDetails").value.trim();
  return run("btnTpl", "מכין...", function () {
    return getBody().then(function (body) {
      return callClaude(systemPrompt([
        TONES.external,
        "משימה: כתוב גוף מייל קצר בשם בר גפן, אפ.אמ. הנדסה, לפי התבנית: " + t + ".",
        "התאם לפרטים שניתנו ולנושא המייל. פרט שלא ניתן – [להשלמה]. החזר את גוף המייל בלבד."
      ]), "נושא: " + (subjectText() || "—") + NL + "פרטים: " + (details || "אין") + NL + NL + "תוכן קיים במייל (להקשר בלבד):" + NL + body.slice(0, 8000), 900);
    }).then(function (out) {
      $("tplText").value = out; $("tplOut").classList.remove("hidden");
      showFlags("tplFlags", out, details);
      setStatus("מוכן. ערוך ולחץ 'הכנס למייל'.");
    });
  });
}

/* ---------- pre-send check ---------- */
function doCheck() {
  return run("btnCheck", "בודק...", function () {
    return Promise.all([getComposeField("subject"), getComposeField("to"), getComposeField("cc"), getBody(), getAttachments()]).then(function (v) {
      var who = function (list) { return (list || []).map(function (x) { return x.displayName || x.emailAddress; }).join(", ") || "—"; };
      var att = v[4] === null ? "לא ידוע" : (v[4].length ? v[4].map(function (a) { return a.name; }).join(", ") : "אין");
      var info = ["נושא: " + (v[0] || "—"), "אל: " + who(v[1]), "עותק: " + who(v[2]), "קבצים מצורפים: " + att, "", "גוף המייל:", v[3]].join(NL);
      showFlags("ckFlags", v[3], null);
      return callClaude(systemPrompt([
        "משימה: בדוק מייל לפני שליחה ומצא בעיות בלבד. אל תנסח מחדש.",
        "בדוק: קובץ שהוזכר בטקסט ולא צורף; מספר פרויקט שמופיע בגוף ולא בנושא (במשרד מספר פרויקט = מספר סידורי + שנה, למשל 15026); נושא חסר או לא ברור; פנייה בשם שלא תואם לנמען; טון חריף או לא מקצועי; שגיאות כתיב בולטות; התחייבות למועד או לערך שנראית לא מכוונת.",
        "החזר רשימה קצרה של בעיות, כל אחת בשורה שמתחילה ב-•. אם הכול תקין – כתוב רק: 'לא נמצאו בעיות.'"
      ]), info, 800);
    }).then(function (out) { $("checkOut").textContent = out; $("checkOut").classList.remove("hidden"); setStatus("הבדיקה הסתיימה. ההחלטה שלך."); });
  });
}

/* ---------- reply (read mode) ---------- */
function doReply() {
  var item = Office.context.mailbox.item, hint = $("replyHint").value.trim();
  return run("btnReply", "מכין טיוטת תשובה...", function () {
    return getBody().then(function (body) {
      var from = item.from ? (item.from.displayName || item.from.emailAddress) : "";
      var info = ["מאת: " + from, "נושא: " + subjectText(), "", "המייל (כולל שרשור):", body.slice(0, 20000), "", "הנחיה מבר לתשובה: " + (hint || "אין – הצע תשובה עניינית ומתאימה")].join(NL);
      return callClaude(systemPrompt([
        TONES[$("replyTone").value],
        "משימה: כתוב טיוטת תשובה בשם בר גפן, אפ.אמ. הנדסה. ענה רק על מה שנשאל ולפי ההנחיה.",
        "החזר את גוף התשובה בלבד."
      ]), info, 1500).then(function (out) {
        $("replyText").value = out; $("replyBox").classList.remove("hidden");
        showFlags("rpFlags", out, body + " " + hint);
        setStatus("הטיוטה מוכנה. ערוך ולחץ 'פתח כתשובה' – שום דבר לא נשלח.");
      });
    });
  });
}
function openReply(all) {
  var html = toHtml($("replyText").value), item = Office.context.mailbox.item;
  if (all) item.displayReplyAllForm(html); else item.displayReplyForm(html);
  setStatus("נפתח חלון תשובה לעריכה. בדוק שהחתימה 'בר' נוספה לפני שליחה.");
}

/* ---------- summary (read mode) ---------- */
function doSummary() {
  return run("btnSummary", "מסכם...", function () {
    return getBody().then(function (body) {
      return callClaude(systemPrompt([
        "משימה: סכם את המייל/השרשור לבר. מבנה: 'בקצרה:' (2–3 שורות), 'החלטות:' ו'משימות פתוחות:' (מי, מה, עד מתי – רק אם נכתב). תבליטים עם •. בלי המצאות."
      ]), "נושא: " + subjectText() + NL + NL + body.slice(0, 30000), 1000);
    }).then(function (out) { $("sumOut").textContent = out; $("sumOut").classList.remove("hidden"); setStatus(""); });
  });
}

/* ---------- tasks (read mode) ---------- */
function doTasks() {
  return run("btnTasks", "מוציא משימות...", function () {
    return getBody().then(function (body) {
      return callClaude(systemPrompt([
        "משימה: הפוך את ההערות והבקשות במייל לרשימת משימות ממוספרת לביצוע בתכנון החשמל.",
        "כל משימה בשורה אחת: מספר, מה לעשות, ובסוגריים – מיקום/גיליון/סעיף אם צוין. שמור על ניסוח ההערה המקורית ככל האפשר.",
        "אל תוסיף משימות שלא נכתבו. בסוף – שורה 'לא ברור:' עם הערות שדורשות הבהרה, אם יש. החזר את הרשימה בלבד."
      ]), "נושא: " + subjectText() + NL + NL + body.slice(0, 30000), 1500);
    }).then(function (out) { $("tasksText").value = out; $("tasksBox").classList.remove("hidden"); setStatus("הרשימה מוכנה. אפשר לערוך, להעתיק או לפתוח מייל חדש עם הרשימה."); });
  });
}
function tasksNewMail() {
  Office.context.mailbox.displayNewMessageForm({ subject: "משימות: " + subjectText(), htmlBody: toHtml($("tasksText").value) });
  setStatus("נפתח מייל חדש לעריכה – בחר נמען. שום דבר לא נשלח.");
}

/* ---------- settings + tabs ---------- */
function saveSettings() {
  var s = settings();
  s.set("fm_api_key", $("apiKey").value.trim());
  s.set("fm_model", $("model").value.trim() || DEFAULT_MODEL);
  s.set("fm_glossary", $("glossary").value);
  s.set("fm_templates", $("templates").value.trim() || DEFAULT_TEMPLATES);
  s.saveAsync(function (r) { setStatus(ok(r) ? "ההגדרות נשמרו." : "שמירה נכשלה.", !ok(r)); fillTemplates(); });
}
function buildTabs(list) {
  var tabs = $("tabs");
  list.forEach(function (t, k) {
    var b = document.createElement("button"); b.className = "tab"; b.textContent = t[1];
    b.onclick = function () {
      Array.prototype.forEach.call(tabs.children, function (x) { x.classList.remove("on"); });
      b.classList.add("on");
      ["rewrite", "points", "templates", "check", "reply", "summary", "tasks"].forEach(function (p) { $("p-" + p).classList.toggle("hidden", p !== t[0]); });
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
  fillTemplates();
  $("btnRewrite").onclick = doRewrite;
  $("btnReplace").onclick = doReplace;
  $("btnCopy").onclick = function () { navigator.clipboard.writeText($("result").value); setStatus("הועתק."); };
  $("vDiff").onclick = function () { showDiff(true); };
  $("vEdit").onclick = function () { showDiff(false); };
  $("btnPoints").onclick = doPoints;
  $("btnPtInsert").onclick = function () { insertAtCursor($("ptText").value); };
  $("btnTpl").onclick = doTemplate;
  $("btnTplInsert").onclick = function () { insertAtCursor($("tplText").value); };
  $("btnCheck").onclick = doCheck;
  $("btnReply").onclick = doReply;
  $("btnOpenReply").onclick = function () { openReply(false); };
  $("btnOpenReplyAll").onclick = function () { openReply(true); };
  $("btnSummary").onclick = doSummary;
  $("btnTasks").onclick = doTasks;
  $("btnTasksCopy").onclick = function () { navigator.clipboard.writeText($("tasksText").value); setStatus("הועתק."); };
  $("btnTasksFwd").onclick = tasksNewMail;
  $("btnSave").onclick = saveSettings;
  $("btnTplReset").onclick = function () { $("templates").value = DEFAULT_TEMPLATES; setStatus("הוחזרו תבניות ברירת המחדל – לחץ 'שמור הגדרות'."); };
  buildTabs(isCompose
    ? [["rewrite", "נסח"], ["points", "מנקודות"], ["templates", "תבניות"], ["check", "בדוק לפני שליחה"]]
    : [["reply", "ענה"], ["summary", "סכם"], ["tasks", "משימות"]]);
});

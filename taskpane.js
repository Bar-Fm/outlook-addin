/* FM-ENG | Outlook add-in "נסח מקצועית" | pilot (single user)
   The API key is kept in Office roamingSettings (per mailbox). For an office-wide rollout,
   move the Claude call to a small server so the key is not in the client. */

const DEFAULT_MODEL = "claude-sonnet-4-5"; // לא מאומת שזה המודל המעודכן – לבדוק ב-docs.claude.com ולעדכן בהגדרות
let source = { mode: null }; // "selection" | "body"

const SYSTEM_PROMPT = `אתה עורך לשון במשרד תכנון חשמל "אפ.אמ. הנדסה בע"מ" (FM-ENG).
תפקידך: לנסח מחדש טקסט של מייל בעברית תקינה, מקצועית ומכובדת, בשפה של משרד תכנון חשמל.
כללים:
- שמור על כל המידע, המספרים, התאריכים, השמות, מספרי הפרויקטים, הערכים והיחידות בדיוק כפי שהם. אל תוסיף עובדות, ערכים, סעיפי תקן או התחייבויות שלא נכתבו.
- מונחים מקצועיים (מתח גבוה/נמוך, לוח חשמל, חברת החשמל, תב"ע, כתב כמויות, אומדן וכו') – השאר ותקן כתיב בלבד.
- תקן דקדוק, כתיב, פיסוק ומבנה משפטים. חלק לפסקאות קצרות; רשימות – בתבליטים כשזה מבהיר.
- אל תוסיף חתימה. אם יש כבר פתיחה/סיום – שפר אותם.
- החזר את הטקסט המנוסח בלבד, בלי הקדמה, בלי הסברים ובלי מרכאות.`;

const TONES = {
  external: "סגנון: רשמי ומכובד לגורם חיצוני. פתיחה מנומסת, תמצית ברורה, וסיום בשורה 'בכבוד רב,' (בלי שם אחריה).",
  internal: "סגנון: ענייני וברור לעמית במשרד, בלי רשמיות יתר.",
  short: "סגנון: קצר ותמציתי ככל האפשר, בלי לאבד מידע.",
  fix: "תקן שגיאות כתיב, דקדוק ופיסוק בלבד. אל תשנה סגנון, מבנה או ניסוח מעבר לנדרש."
};

const $ = (id) => document.getElementById(id);
const setStatus = (t, err) => { $("status").textContent = t; $("status").className = "status" + (err ? " err" : ""); };

Office.onReady(() => {
  const s = Office.context.roamingSettings;
  $("apiKey").value = s.get("fm_api_key") || "";
  $("model").value = s.get("fm_model") || DEFAULT_MODEL;
  $("btnRun").onclick = run;
  $("btnReplace").onclick = replaceInMail;
  $("btnCopy").onclick = () => { navigator.clipboard.writeText($("result").value); setStatus("הועתק."); };
  $("btnSave").onclick = saveSettings;
});

function saveSettings() {
  const s = Office.context.roamingSettings;
  s.set("fm_api_key", $("apiKey").value.trim());
  s.set("fm_model", $("model").value.trim() || DEFAULT_MODEL);
  s.saveAsync((r) => setStatus(r.status === Office.AsyncResultStatus.Succeeded ? "ההגדרות נשמרו." : "שמירה נכשלה.", r.status !== Office.AsyncResultStatus.Succeeded));
}

function getSelection() {
  return new Promise((res) => Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, (r) => res(r.status === Office.AsyncResultStatus.Succeeded ? (r.value.data || "") : "")));
}
function getBody() {
  return new Promise((res, rej) => Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (r) => r.status === Office.AsyncResultStatus.Succeeded ? res(r.value) : rej(r.error)));
}

async function run() {
  const key = Office.context.roamingSettings.get("fm_api_key") || $("apiKey").value.trim();
  if (!key) { setStatus("חסר מפתח API – הזן אותו ב'הגדרות' ושמור.", true); return; }

  let text = (await getSelection()).trim();
  source.mode = "selection";
  if (!text) { text = (await getBody()).trim(); source.mode = "body"; }
  if (!text) { setStatus("לא נמצא טקסט לניסוח.", true); return; }

  $("btnRun").disabled = true; $("btnReplace").disabled = true; $("btnCopy").disabled = true;
  setStatus(source.mode === "selection" ? "מנסח את הטקסט המסומן..." : "מנסח את כל גוף המייל...");

  try {
    const resp = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-api-key": key,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true"
      },
      body: JSON.stringify({
        model: Office.context.roamingSettings.get("fm_model") || $("model").value.trim() || DEFAULT_MODEL,
        max_tokens: 2000,
        system: SYSTEM_PROMPT + "\n" + TONES[$("tone").value],
        messages: [{ role: "user", content: "נסח את הטקסט הבא:\n\n" + text }]
      })
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data?.error?.message || resp.status);
    $("result").value = data.content.filter(c => c.type === "text").map(c => c.text).join("").trim();
    $("btnReplace").disabled = false; $("btnCopy").disabled = false;
    setStatus("מוכן. עבור על ההצעה, ערוך לפי הצורך ולחץ 'החלף במייל'.");
  } catch (e) {
    setStatus("שגיאה: " + e.message, true);
  } finally {
    $("btnRun").disabled = false;
  }
}

function toHtml(t) {
  const esc = t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  return '<div dir="rtl" style="font-family:Heebo,Arial,sans-serif;text-align:right">' + esc.replace(/\n/g, "<br>") + "</div>";
}

function replaceInMail() {
  const html = toHtml($("result").value);
  const cb = (r) => setStatus(r.status === Office.AsyncResultStatus.Succeeded ? "הטקסט הוחלף במייל." : "ההחלפה נכשלה: " + r.error.message, r.status !== Office.AsyncResultStatus.Succeeded);
  const item = Office.context.mailbox.item;
  if (source.mode === "selection") item.body.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, cb);
  else item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, cb);
}

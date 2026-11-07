/*
 * End-to-end test suite for the budget app.
 * Run:  node tests/run.js            (needs Playwright + Chromium)
 * Optional: LibreOffice (soffice) for recalculating the exported Excel formulas.
 * Writes tests/results.json and prints a summary.
 */
const { chromium } = require("playwright");
const http = require("http");
const fs = require("fs");
const path = require("path");
const os = require("os");
const { execSync } = require("child_process");

const ROOT = path.resolve(__dirname, "..");
const OUT = fs.mkdtempSync(path.join(os.tmpdir(), "budget-tests-"));
const MOBILE = { width: 390, height: 844 };

/* ---------- tiny static server ---------- */
const TYPES = { ".html": "text/html; charset=utf-8", ".js": "text/javascript", ".png": "image/png", ".webmanifest": "application/manifest+json" };
function serve() {
  return new Promise(res => {
    const srv = http.createServer((req, rsp) => {
      let f = decodeURIComponent(req.url.split("?")[0]);
      if (f.endsWith("/")) f += "index.html";
      const p = path.join(ROOT, f);
      if (!p.startsWith(ROOT) || !fs.existsSync(p)) { rsp.writeHead(404); rsp.end(); return; }
      rsp.writeHead(200, { "Content-Type": TYPES[path.extname(p)] || "application/octet-stream" });
      fs.createReadStream(p).pipe(rsp);
    }).listen(0, () => res(srv));
  });
}

/* ---------- helpers ---------- */
let BASE, browser;
async function open(opts = {}) {
  const ctx = await browser.newContext({ viewport: opts.viewport || MOBILE, acceptDownloads: true, hasTouch: !!opts.touch });
  // block external fonts so tests are deterministic offline
  await ctx.route(/fonts\.(googleapis|gstatic)\.com/, r => r.abort());
  const p = await ctx.newPage();
  p.errors = [];
  p.on("pageerror", e => p.errors.push(e.message));
  p.on("console", m => { if (m.type() === "error" && !/ERR_FAILED|net::/.test(m.text())) p.errors.push(m.text()); });
  await p.goto(BASE);
  return p;
}
async function onboard(p, o = {}) {
  await p.fill("[name=nameA]", o.a || "ישראלה");
  await p.fill("[name=nameB]", o.b || "דני");
  if (o.moveDate !== undefined) await p.fill("[name=moveDate]", o.moveDate);
  if (o.budget !== undefined) await p.fill("[name=moveBudget]", String(o.budget));
  if (o.variable !== undefined) await p.fill("[name=variable]", String(o.variable));
  await p.click("#welcome-form button[type=submit]");
  await p.waitForSelector("#tabs:not([hidden])");
}
const state = p => p.evaluate(() => JSON.parse(localStorage.getItem("budget.v1")));
async function setState(p, fn) {
  const s = await state(p);
  fn(s);
  await p.evaluate(s => localStorage.setItem("budget.v1", JSON.stringify(s)), s);
  await p.reload();
}
async function submit(p) { await p.click(".sheet button[type=submit]"); await p.waitForSelector(".sheet", { state: "detached", timeout: 1500 }).catch(() => {}); }
async function tab(p, v) { await p.click(`.tabs button[data-v=${v}]`); }
async function toast(p, sub) {
  if (sub) await p.waitForFunction(x => { const t = document.querySelector(".toast"); return t && t.textContent.includes(x); }, sub, { timeout: 2500 }).catch(() => {});
  else await p.waitForSelector(".toast", { timeout: 2500 }).catch(() => {});
  const t = await p.$(".toast"); return t ? (await t.textContent()).trim() : "";
}
async function sheetOpen(p) { return !!(await p.$(".sheet")); }
async function openItem(p, name) { await tab(p, "shop"); await p.click(`.li[data-act=edit-item] .t:text-is("${name}")`); }
async function buy(p, name, actual, o = {}) {
  await openItem(p, name); await p.fill(".sheet [name=actual]", String(actual));
  if (o.payer || o.inst || o.date || o.method) await p.click(".sheet summary").catch(() => {});
  if (o.date) await p.fill(".sheet [name=date]", o.date);
  if (o.payer) await p.selectOption(".sheet [name=payer]", o.payer);
  if (o.method) await p.selectOption(".sheet [name=method]", o.method);
  if (o.inst) await p.fill(".sheet [name=inst]", String(o.inst));
  if (o.note !== undefined) await p.fill(".sheet [name=note]", o.note);
  await submit(p);
}
async function richData(p) {
  await onboard(p, { budget: 40000, moveDate: iso(24) });
  await buy(p, "מקרר", 3200, { payer: "a", note: "LG, \"מבצע\"" });
  await buy(p, "מכונת כביסה", 2100);
  await buy(p, "מזרן זוגי", 1900, { payer: "b", inst: 10, date: ymShift(-1) + "-20" });
  await openItem(p, "מייבש כביסה"); await p.click(".sheet [data-act=remove-item]");
}
async function download(p, sel) {
  const [d] = await Promise.all([p.waitForEvent("download", { timeout: 4000 }), p.click(sel)]);
  const f = path.join(OUT, d.suggestedFilename());
  await d.saveAs(f);
  return f;
}
function iso(offsetDays) { const d = new Date(); d.setDate(d.getDate() + offsetDays); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); }
function ymShift(n) { const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() + n); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0"); }
function xlsx(file, recalc) {
  // returns {sheet: [[values]]} using python/openpyxl; recalc via LibreOffice when asked
  let f = file;
  if (recalc) {
    const dir = fs.mkdtempSync(path.join(OUT, "rc-"));
    execSync(`soffice --headless --calc --convert-to xlsx --outdir "${dir}" "${file}"`, { stdio: "ignore", timeout: 90000 });
    f = path.join(dir, path.basename(file));
  }
  const py = `import openpyxl,json,datetime,sys
wb=openpyxl.load_workbook(sys.argv[1],data_only=${recalc ? "True" : "False"})
o={}
for ws in wb:
  o[ws.title]={"rtl":bool(ws.sheet_view.rightToLeft),"rows":[[ (v.isoformat() if isinstance(v,(datetime.date,datetime.datetime)) else v) for v in r] for r in ws.iter_rows(values_only=True)]}
print(json.dumps(o,ensure_ascii=False))`;
  return JSON.parse(execSync(`python3 -c '${py.replace(/'/g, "'\\''")}' "${f}"`).toString());
}
function contrast(a, b) {
  const L = h => { const c = [1, 3, 5].map(i => parseInt(h.slice(i, i + 2), 16) / 255).map(v => v <= .03928 ? v / 12.92 : Math.pow((v + .055) / 1.055, 2.4)); return .2126 * c[0] + .7152 * c[1] + .0722 * c[2]; };
  const x = L(a), y = L(b); return (Math.max(x, y) + .05) / (Math.min(x, y) + .05);
}
function assert(cond, msg) { if (!cond) throw new Error(msg); }

/* ---------- test cases ---------- */
const T = [];
function test(id, module, name, steps, expected, fn) { T.push({ id, module, name, steps, expected, fn }); }

/* --- onboarding & settings --- */
test("ONB-01", "פתיחה", "הפעלה ראשונה", "פתיחת האפליקציה עם אחסון ריק", "מוצג מסך פתיחה, סרגל הלשוניות מוסתר", async p => {
  assert(await p.isVisible("#welcome-form"), "welcome form missing");
  assert(await p.isHidden("#tabs"), "tabs visible before setup");
});
test("ONB-02", "פתיחה", "שמות חובה", "לחיצה על 'מתחילים' בלי שמות", "נשארים במסך הפתיחה", async p => {
  await p.click("#welcome-form button[type=submit]");
  assert(await p.isVisible("#welcome-form"), "left welcome without names");
});
test("ONB-03", "פתיחה", "השלמת פתיחה", "מילוי שמות, תאריך מעבר ותקציב 45000", "מסך בית עם ברכה בשמות וספירה לאחור; 30 פריטים ברשימת הקניות; אין חשבונות", async p => {
  await onboard(p, { moveDate: iso(29), budget: 45000 });
  assert((await p.textContent("h1")).includes("ישראלה") && (await p.textContent("h1")).includes("דני"), "greeting");
  const s = await state(p);
  assert(s.items.length === 30 && s.bills.length === 0, `items ${s.items.length} bills ${s.bills.length}`);
  assert(s.settings.moveBudget === 45000, "settings saved");
  assert(!(await p.$("[name=variable]")), "variable field still exists");
});
test("ONB-04", "פתיחה", "תקציב עם אגורות", "הזנת תקציב מעבר 45000.50", "הערך נשמר כ־45000.5", async p => {
  await onboard(p, { budget: "45000.50" });
  const s = await state(p);
  assert(s.settings.moveBudget === 45000.5, "moveBudget=" + s.settings.moveBudget);
});
test("SET-01", "הגדרות", "עדכון שמות", "הגדרות ← שינוי שם א׳ ל'יעל' ← שמירה ← בית", "הברכה מציגה 'יעל'", async p => {
  await onboard(p);
  await p.click("[data-act=go][data-v=settings]");
  await p.fill("#settings-form [name=nameA]", "יעל");
  await p.click("#settings-form button[type=submit]");
  await tab(p, "home");
  assert((await p.textContent("h1")).includes("יעל"), "name not updated");
});
test("SET-04", "הגדרות", "איפוס", "הגדרות ← איפוס ← מחיקה והתחלה מחדש", "חוזרים למסך פתיחה, ההוצאות נמחקו", async p => {
  await onboard(p); await buy(p, "מקרר", 3000);
  await p.click("[data-act=go][data-v=settings]"); await p.click("[data-act=reset]");
  await p.click(".sheet .btn.danger");
  assert(await p.isVisible("#welcome-form"), "not back to welcome");
  const s = await state(p); assert(!s.items.some(i => i.actual != null), "purchases kept");
});

/* --- move --- */
test("MOV-01", "מעבר", "קניית פריט", "מקרר ← מחיר בפועל 5200 ← שמירה", "סטטוס 'התקבל', 'שולם' 5,200, מוצג +1,700 מהתכנון", async p => {
  await onboard(p, { budget: 45000 }); await openItem(p, "מקרר");
  await p.fill(".sheet [name=actual]", "5200"); await submit(p);
  const s = await state(p), i = s.items.find(x => x.name === "מקרר");
  assert(i.actual === 5200 && i.status === "got", JSON.stringify(i));
  assert((await p.textContent("#main")).includes("+₪1,700"), "diff not shown");
});
test("MOV-02", "מעבר", "מחיר עם אגורות", "מקרר ← מחיר בפועל 2399.90", "המחיר נשמר כ־2399.9", async p => {
  await onboard(p); await openItem(p, "מקרר");
  await p.fill(".sheet [name=actual]", "2399.90"); await submit(p);
  const i = (await state(p)).items.find(x => x.name === "מקרר");
  assert(i.actual === 2399.9, "actual=" + i.actual + (await sheetOpen(p) ? " (form blocked by browser validation)" : ""));
});
test("MOV-03", "מעבר", "'לא צריך' ושחזור", "מייבש כביסה ← לא צריך ← הצגת פריטים שהוסרו ← החזרה", "המייבש יוצא מהצפי וחוזר אחרי שחזור", async p => {
  await onboard(p);
  const before = await p.evaluate(() => 0);
  await openItem(p, "מייבש כביסה"); await p.click(".sheet [data-act=remove-item]");
  let s = await state(p); assert(s.items.find(x => x.name === "מייבש כביסה").removed, "not removed");
  await p.click("[data-act=toggle-removed]"); await p.click("[data-act=restore-item]");
  s = await state(p); assert(!s.items.find(x => x.name === "מייבש כביסה").removed, "not restored");
});
test("MOV-06", "מעבר", "מחיקה לצמיתות", "טלוויזיה ← מחיקה לצמיתות ← מחיקה", "הפריט נמחק מהנתונים", async p => {
  await onboard(p); await openItem(p, "טלוויזיה"); await p.click(".sheet [data-act=delete-item]"); await p.click(".sheet .btn.danger");
  assert(!(await state(p)).items.some(x => x.name === "טלוויזיה"), "still exists");
});
test("MOV-07", "מעבר", "הוספת פריט", "סלון ← + פריט ← 'כורסה' 1800 ← שמירה", "הפריט נוסף לסלון", async p => {
  await onboard(p); await tab(p, "shop");
  await p.click("[data-act=add-item][data-room=living]");
  await p.fill(".sheet [name=name]", "כורסה"); await p.fill(".sheet [name=planned]", "1800"); await submit(p);
  const i = (await state(p)).items.find(x => x.name === "כורסה"); assert(i && i.room === "living" && i.planned === 1800, JSON.stringify(i));
});
test("MOV-08", "מעבר", "סינון נקנה / עוד לקנות", "קניית מקרר ← סינון 'נקנה' ← סינון 'עוד לקנות'", "'נקנה' מציג רק את המקרר; 'עוד לקנות' לא מציג אותו", async p => {
  await onboard(p); await openItem(p, "מקרר"); await p.fill(".sheet [name=actual]", "5000"); await submit(p);
  await p.click("[data-act=move-filter][data-v=done]");
  let names = await p.$$eval(".li[data-act=edit-item] .t", x => x.map(e => e.textContent));
  assert(names.length === 1 && names[0] === "מקרר", "done filter: " + names);
  await p.click("[data-act=move-filter][data-v=open]");
  names = await p.$$eval(".li[data-act=edit-item] .t", x => x.map(e => e.textContent));
  assert(!names.includes("מקרר"), "open filter shows bought item");
});
test("MOV-12", "מעבר", "מחיר שלילי", "מקרר ← מחיר בפועל ‎-500", "הערך נדחה", async p => {
  await onboard(p); await openItem(p, "מקרר"); await p.fill(".sheet [name=actual]", "-500"); await submit(p);
  const i = (await state(p)).items.find(x => x.name === "מקרר"); assert(i.actual == null || i.actual >= 0, "negative saved: " + i.actual);
});

test("MOV-13", "מעבר", "החזרת פריט דרך 'פריט חדש'", "מייבש כביסה ← לא צריך ← חשמל ← + פריט ← בחירתו מהרשימה ← שמירה", "הרשימה מציעה את המייבש כ'הוסר'; אחרי השמירה הפריט המקורי חוזר (בלי כפילות) עם התקציב שלו", async p => {
  await onboard(p); await openItem(p, "מייבש כביסה"); await p.click(".sheet [data-act=remove-item]");
  await tab(p, "shop"); await p.click("[data-act=add-item][data-room=elec]");
  const opts = await p.$$eval("#item-suggest option", o => o.map(x => x.value + "|" + x.textContent));
  assert(opts[0] && opts[0].startsWith("מייבש כביסה|הוסר"), "removed item not offered first: " + opts.slice(0, 3));
  await p.fill(".sheet [name=name]", "מייבש כביסה");
  assert(await p.isVisible("#restore-hint"), "no restore hint");
  assert((await p.inputValue(".sheet [name=planned]")) === "2300", "planned not prefilled");
  await submit(p);
  const all = (await state(p)).items.filter(x => x.name === "מייבש כביסה");
  assert(all.length === 1 && !all[0].removed && all[0].planned === 2300, JSON.stringify(all));
});
test("MOV-14", "מעבר", "הקלדה חופשית לצד הרשימה", "+ פריט ← הקלדת 'מנורת לילה' שלא ברשימה ← שמירה", "נוצר פריט חדש בשם שהוקלד", async p => {
  await onboard(p); await tab(p, "shop"); await p.click("[data-act=add-item][data-room=bedroom]");
  await p.fill(".sheet [name=name]", "מנורת לילה"); await p.fill(".sheet [name=planned]", "150"); await submit(p);
  const it = (await state(p)).items.find(x => x.name === "מנורת לילה");
  assert(it && it.room === "bedroom" && !it.removed, JSON.stringify(it));
});
test("MOV-15", "מעבר", "פריט ברירת מחדל שנמחק לצמיתות", "טלוויזיה ← מחיקה לצמיתות ← + פריט", "'טלוויזיה' מוצעת שוב ברשימה עם הקבוצה והתקציב המקוריים", async p => {
  await onboard(p); await openItem(p, "טלוויזיה"); await p.click(".sheet [data-act=delete-item]"); await p.click(".sheet .btn.danger");
  await tab(p, "shop"); await p.click("[data-act=add-item][data-room=living]");
  assert((await p.$$eval("#item-suggest option", o => o.map(x => x.value))).includes("טלוויזיה"), "not suggested");
  await p.fill(".sheet [name=name]", "טלוויזיה");
  assert((await p.inputValue(".sheet [name=room]")) === "elec" && (await p.inputValue(".sheet [name=planned]")) === "2500", "not prefilled");
});

test("MOV-16", "מעבר", "הורדת תקציב לפריט", "מקרר: תקציב 3500 ← 3000 ← שמירה", "הצפי הכולל יורד ב־500", async p => {
  await onboard(p); await tab(p, "shop");
  const before = +(await p.$$eval(".stat .v", v => v[1].textContent.replace(/[^\d]/g, "")));
  await openItem(p, "מקרר"); await p.fill(".sheet [name=planned]", "3000"); await submit(p);
  const after = +(await p.$$eval(".stat .v", v => v[1].textContent.replace(/[^\d]/g, "")));
  assert(before - after === 500, `before ${before} after ${after}`);
});
test("MOV-17", "מעבר", "תקציב לפריט שכבר נקנה", "מקרר: מחיר בפועל 5200 ← פתיחה מחדש", "מוצג הסבר שהצפי מחושב לפי המחיר בפועל; ההסבר נעלם כשמוחקים את המחיר", async p => {
  await onboard(p); await openItem(p, "מקרר"); await p.fill(".sheet [name=actual]", "5200"); await submit(p);
  await openItem(p, "מקרר"); assert(await p.isVisible("#plan-hint"), "no hint");
  await p.fill(".sheet [name=actual]", ""); assert(await p.isHidden("#plan-hint"), "hint stays");
});

test("MIG-01", "מעבר", "עדכון לרשימה החדשה בנתונים קיימים", "נתונים בגרסה הקודמת: תנור נקנה, כורסה שנוספה ידנית ← מעבר ← עדכון הרשימה", "תנור וכורסה נשארים; שואב אבק יוצא; בר מים נוסף; מקרר 3,500; אין כפילויות והבאנר נעלם", async p => {
  await onboard(p);
  await setState(p, s => {
    const mk = (room, name, planned, extra) => Object.assign({ id: Math.random().toString(36).slice(2), room, name, prio: "now", planned, actual: null, status: "todo", payer: "j", method: "אשראי", inst: 1, date: "", warranty: "", note: "", store: "", removed: false }, extra || {});
    s.items = [mk("elec", "מקרר", 4500), mk("elec", "תנור / כיריים", 3000, { actual: 2800, status: "got", date: "2026-11-01" }), mk("elec", "שואב אבק", 1200),
      mk("elec", "מזגן + התקנה", 3000, { removed: true }), mk("living", "כורסה", 1800), mk("move", "הובלה", 2500)];
    delete s.meta.listVer;
  });
  await tab(p, "shop");
  assert(await p.isVisible("[data-act=upgrade-list]"), "no upgrade banner");
  await p.click(".banner [data-act=upgrade-list]"); await p.click(".sheet .btn.primary");
  const it = (await state(p)).items, names = it.map(i => i.name);
  assert(names.includes("תנור / כיריים") && names.includes("כורסה"), "kept items lost");
  assert(!names.includes("שואב אבק") && !names.includes("מזגן + התקנה"), "old defaults not removed");
  assert(names.includes("בר מים") && it.find(i => i.name === "מקרר").planned === 3500, "new list not applied");
  assert(new Set(names).size === names.length, "duplicates: " + names.filter((n, i) => names.indexOf(n) !== i));
  assert(!(await p.isVisible(".banner [data-act=upgrade-list]")), "banner stays");
});
test("MIG-02", "מעבר", "משתמש חדש", "פתיחה ראשונה ← מעבר", "אין באנר עדכון; הרשימה כבר החדשה (מזנון לטלוויזיה, חדר נוסף / עבודה)", async p => {
  await onboard(p); await tab(p, "shop");
  assert(!(await p.isVisible(".banner [data-act=upgrade-list]")), "banner for new user");
  const t = await p.textContent("#main"); assert(t.includes("מזנון לטלוויזיה") && t.includes("חדר נוסף / עבודה"), "new list missing");
});

test("HOM-04", "בית", "התראת תקציב מעבר", "תקציב 10000 ← קניית מקרר 8500 ← בית; ואז ספה 2000", "התראה כתומה ב־85%; אדומה כשהתקציב נחצה", async p => {
  await onboard(p, { budget: 10000 }); await openItem(p, "מקרר"); await p.fill(".sheet [name=actual]", "8500"); await submit(p);
  await tab(p, "home"); assert((await p.textContent(".banner.warn")).includes("85%"), "no 80% warning");
  await openItem(p, "ספה"); await p.fill(".sheet [name=actual]", "2000"); await submit(p);
  await tab(p, "home"); assert((await p.textContent(".banner.bad")).includes("חרגתם"), "no over-budget alert");
});
test("HOM-05", "בית", "ספירה לאחור", "תאריך מעבר בעוד 29 ימים", "תג 'עוד 29 ימים'", async p => {
  await onboard(p, { moveDate: iso(29) }); assert((await p.textContent("#main")).includes("עוד 29 ימים"), "countdown");
});

/* --- reports --- */
/* --- excel / csv / backup --- */
test("BAK-02", "גיבוי", "קובץ לא תקין", "שחזור מקובץ טקסט", "הודעת שגיאה, הנתונים לא משתנים", async p => {
  await onboard(p); await buy(p, "מקרר", 3000);
  const f = path.join(OUT, "bad.json"); fs.writeFileSync(f, "not json");
  await p.setInputFiles("#import-file", f);
  assert((await toast(p, "קובץ")).includes("לא"), "no error");
  assert((await state(p)).items.find(i => i.name === "מקרר").actual === 3000, "data changed");
});

/* --- UI / accessibility --- */
test("UI-01", "תצוגה", "מסך צר 320px", "מעבר על כל הלשוניות ברוחב 320", "אין גלילה אופקית", async p => {
  await p.setViewportSize({ width: 320, height: 568 });
  await richData(p);
  for (const v of ["home", "shop", "check", "reports"]) {
    await tab(p, v);
    const w = await p.evaluate(() => document.documentElement.scrollWidth);
    assert(w <= 320, `${v}: scrollWidth ${w}`);
  }
});
test("UI-04", "תצוגה", "כפתור 'חזור' באנדרואיד", "פתיחת טופס ← חזור בדפדפן", "הטופס נסגר והאפליקציה נשארת פתוחה", async p => {
  await onboard(p); await tab(p, "shop"); await openItem(p, "מקרר");
  await p.goBack().catch(() => {});
  await p.waitForTimeout(300);
  const stillApp = p.url().startsWith(BASE) && await p.isVisible("#tabs");
  assert(stillApp && !(await sheetOpen(p)), stillApp ? "sheet still open" : "back left the app");
});
test("UI-05", "תצוגה", "ניגודיות טקסט בכרטיס העליון", "חישוב ניגודיות בין צבע הטקסט לשני צבעי הרקע של הכרטיס", "לפחות 4.5:1 (תקן נגישות AA)", async p => {
  const css = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
  const m = css.match(/\.hero \{ background: linear-gradient\(135deg, (#[0-9A-Fa-f]{6}) 0%, (#[0-9A-Fa-f]{6}) 100%\); color: (#[0-9A-Fa-f]{6})/);
  for (const bg of [m[1], m[2]]) { const c = contrast(m[3], bg); assert(c >= 4.5, `${m[3]} on ${bg} = ${c.toFixed(2)}:1`); }
});
test("UI-07", "תצוגה", "שגיאות JavaScript", "זרימה מלאה עם נתונים לדוגמה בכל הלשוניות", "אין שגיאות בקונסול", async p => {
  await richData(p); for (const v of ["shop", "check", "reports", "home"]) await tab(p, v);
  assert(p.errors.length === 0, p.errors.join(" | "));
});

/* --- move-only focus --- */
test("NAV-01", "מבנה", "רק מה שקשור למעבר", "פתיחה ← סקירת הלשוניות וההגדרות", "לשוניות: בית, קניות, משימות, דוחות. אין חשבונות, הוצאות, שכר דירה, העברות, הוצאות משתנות או פיקדונות", async p => {
  await onboard(p);
  const tabs = await p.$$eval(".tabs button span", x => x.map(e => e.textContent));
  assert(JSON.stringify(tabs) === JSON.stringify(["בית", "קניות", "משימות", "דוחות"]), tabs.join(","));
  const all = [];
  for (const v of ["home", "shop", "check", "reports"]) { await tab(p, v); all.push(await p.textContent("#main")); }
  await p.click("[data-act=go][data-v=settings]"); all.push(await p.textContent("#main"));
  const t = all.join(" ");
  for (const w of ["שכר דירה", "חשבון קבוע", "חשבונות קבועים", "העברות לחשבון המשותף", "הוצאות משתנות", "קטגוריות", "פיקדון"]) assert(!t.includes(w), "still mentions: " + w);
});
test("NAV-02", "מבנה", "כפתור + מוסיף פריט", "בית ← כפתור +", "נפתח טופס 'פריט חדש'", async p => {
  await onboard(p); await p.click("#fab");
  assert((await p.textContent(".sheet h3")) === "פריט חדש", await p.textContent(".sheet h3"));
});
test("HOM-01", "בית", "כרטיס תקציב המעבר", "תקציב 40,000 ← קניית מקרר 3,200 ← בית", "'נשאר מתקציב המעבר' 36,800; תקציב 40,000; שולם 3,200", async p => {
  await onboard(p, { budget: 40000 }); await buy(p, "מקרר", 3200); await tab(p, "home");
  assert((await p.textContent(".hero .big")).replace(/[^\d]/g, "") === "36800", await p.textContent(".hero .big"));
  const cols = await p.$$eval(".hero .cols b", x => x.map(e => e.textContent.replace(/[^\d]/g, "")));
  assert(cols[0] === "40000" && cols[1] === "3200", JSON.stringify(cols));
});
test("HOM-02", "בית", "לפי שלבים והבא בתור", "קניית מקרר ← בית", "שלב ראשון 1/7; 'הבא בתור' מציג פריטי חובה מיד שעוד לא נקנו (בלי המקרר)", async p => {
  await onboard(p); await buy(p, "מקרר", 3200); await tab(p, "home");
  const ph = await p.textContent(".hbars"); assert(ph.includes("1/7"), ph.replace(/\s+/g, " "));
  const next = await p.$$eval("section.card", s => [...s.find(x => x.textContent.includes("הבא בתור")).querySelectorAll(".li .t")].map(e => e.textContent));
  assert(next.length && !next.includes("מקרר") && next.includes("ספה"), JSON.stringify(next));
});
test("HOM-03", "בית", "איזון 50/50 וסגירת חשבון", "א׳ קונה מקרר ב־1,000 מהחשבון הפרטי ← בית ← סגירת חשבון", "'דני צריך להעביר 500', ואחרי הרישום – 'מאוזנים'", async p => {
  await onboard(p); await buy(p, "מקרר", 1000, { payer: "a" }); await tab(p, "home");
  assert((await p.textContent("#main")).includes("₪500"), "balance wrong");
  await p.click("[data-act=settle]"); await submit(p);
  assert((await p.textContent("#main")).includes("מאוזנים"), "not balanced after settle");
});
test("MOV-04", "מעבר", "הסרת פריט שכבר שולם", "מכונת כביסה 2,300 ← לא צריך ← אישור", "הפריט מוסר והתשלום יורד מ'שולם' (0)", async p => {
  await onboard(p); await buy(p, "מכונת כביסה", 2300);
  await openItem(p, "מכונת כביסה"); await p.click(".sheet [data-act=remove-item]"); await p.click(".sheet .btn.danger");
  const i = (await state(p)).items.find(x => x.name === "מכונת כביסה");
  assert(i.removed && i.actual == null, JSON.stringify(i));
  await tab(p, "shop"); assert((await p.$$eval(".stat .v", v => v[2].textContent.replace(/[^\d]/g, ""))) === "0", "spent not 0");
});
test("MOV-09", "משימות", "צ׳קליסט נשמר", "משימות ← סימון 2 משימות ← טעינה מחדש", "הסימונים נשארים, המונה מציג 2/14", async p => {
  await onboard(p); await tab(p, "check");
  await p.click("[data-k=meters]"); await p.click("[data-k=arnona]");
  await p.reload(); await tab(p, "check");
  assert((await p.textContent(".chip.done")).includes("2/14"), await p.textContent(".chip.done"));
});
test("MOV-10", "משימות", "קריאת מונים נשמרת", "הזנת קריאת מים 1234 ← טעינה מחדש", "הערך נשמר", async p => {
  await onboard(p); await tab(p, "check");
  await p.fill("[data-meter=water]", "1234"); await p.press("[data-meter=water]", "Tab");
  await p.reload(); await tab(p, "check");
  assert((await p.inputValue("[data-meter=water]")) === "1234", "lost");
});
test("MIG-03", "מעבר", "הוצאות מעבר מגרסה קודמת", "נתונים ישנים עם הוצאה 'מעבר דירה' 600 (הובלה) והוצאת סופר ← טעינה", "ההוצאה הופכת לפריט שנקנה ב'עלויות המעבר' ונספרת בתקציב; הוצאת הסופר לא מוצגת", async p => {
  await onboard(p, { budget: 10000 });
  await setState(p, s => { s.expenses = [{ id: "x1", amount: 600, cat: "מעבר דירה", date: "2026-11-01", payer: "a", method: "מזומן", inst: 1, note: "הובלה נוספת · מהדירה הישנה" }, { id: "x2", amount: 300, cat: "סופר ובית", date: "2026-11-01", payer: "j", note: "סופר" }]; });
  const s = await state(p), it = s.items.find(i => i.name === "הובלה נוספת");
  assert(it && it.actual === 600 && it.room === "move" && it.note === "מהדירה הישנה", JSON.stringify(it));
  assert(s.expenses.length === 1, "move expense not removed");
  await tab(p, "home"); assert((await p.$$eval(".hero .cols b", x => x[1].textContent.replace(/[^\d]/g, ""))) === "600", "not counted");
  assert(!(await p.textContent("#main")).includes("סופר"), "general expense shown");
});
test("REP-01", "דוחות", "לפי חדר", "נתונים לדוגמה ← דוחות", "שורת הסה״כ שווה לסכום השורות בכל עמודה; שולם 7,200", async p => {
  await richData(p); await tab(p, "reports");
  const rows = await p.$$eval("section.card", s => [...s.find(x => x.textContent.includes("לפי חדר")).querySelectorAll("tbody tr")].map(tr => [...tr.children].slice(1).map(td => +td.textContent.replace(/[^\d]/g, ""))));
  const tot = rows.pop();
  for (let c = 0; c < 3; c++) assert(rows.reduce((a, r) => a + r[c], 0) === tot[c], `col ${c}: ${tot[c]}`);
  assert(tot[2] === 7200, "spent " + tot[2]);
});
test("REP-02", "דוחות", "תשלומי אשראי", "מזרן 1,900 ב־10 תשלומים באשראי (בחודש שעבר) ← דוחות", "190 ₪ בחודש הנוכחי ובכל אחד מ־5 החודשים הבאים", async p => {
  await richData(p); await tab(p, "reports");
  const vals = await p.$$eval("section.card", s => [...s.find(x => x.textContent.includes("תשלומי אשראי")).querySelectorAll("tbody td:nth-child(2)")].map(td => td.textContent.replace(/[^\d]/g, "")));
  assert(vals.length === 6 && vals.every(v => v === "190"), JSON.stringify(vals));
});
test("REP-03", "דוחות", "קניות לפי חודש", "נתונים לדוגמה ← דוחות", "החודש הקודם 1,900; החודש הנוכחי 5,300", async p => {
  await richData(p); await tab(p, "reports");
  const t = await p.$$eval("section.card", s => s.find(x => x.textContent.includes("קניות לפי חודש")).textContent);
  assert(t.includes("1,900") && t.includes("5,300"), t.replace(/\s+/g, " "));
});
test("XLS-01", "אקסל", "ייצוא מלא", "נתונים לדוגמה ← דוחות ← הורדה; חישוב מחדש ב־LibreOffice", "2 גיליונות מימין לשמאל (רשימת קניות, סיכום); הנוסחאות מחושבות לאותם ערכים שבאפליקציה", async p => {
  await richData(p); await tab(p, "reports");
  const f = await download(p, "[data-act=export-xlsx]");
  const x = xlsx(f, false);
  assert(JSON.stringify(Object.keys(x)) === JSON.stringify(["רשימת קניות", "סיכום"]), Object.keys(x).join(","));
  assert(Object.values(x).every(s => s.rtl), "not RTL");
  const rc = xlsx(f, true);
  const fridge = rc["רשימת קניות"].rows.find(r => r[1] === "מקרר"); assert(fridge[5] === -300, "diff " + fridge[5]);
  const tot = rc["רשימת קניות"].rows.find(r => r[0] === "סה״כ (בלי פריטים שהוסרו)"); assert(tot[4] === 7200, "list total " + tot[4]);
  const sum = rc["סיכום"].rows;
  assert(sum.find(r => r[0] === "נשאר מהתקציב")[1] === 32800, "left");
  const roomTot = sum.filter(r => r[0] === "סה״כ")[0]; assert(roomTot[2] === 7200, "room spent " + roomTot[2]);
  assert(sum.find(r => r[0] === "מוצרי חשמל")[2] === 5300, "elec spent");
});
test("XLS-02", "אקסל", "סיכום בלבד", "ביטול 'רשימת הקניות' ← הורדה", "רק גיליון 'סיכום', עם ערכים נכונים בלי נוסחאות שבורות", async p => {
  await richData(p); await tab(p, "reports"); await p.uncheck("#x-move");
  const x = xlsx(await download(p, "[data-act=export-xlsx]"), true);
  assert(JSON.stringify(Object.keys(x)) === '["סיכום"]', Object.keys(x).join(","));
  assert(x["סיכום"].rows.find(r => r[0] === "מוצרי חשמל")[2] === 5300, "values");
});
test("XLS-03", "אקסל", "ייצוא בלי תוכן", "ביטול כל הסימונים ← הורדה", "הודעה 'צריך לבחור', אין הורדה", async p => {
  await onboard(p); await tab(p, "reports");
  for (const id of ["#x-move", "#x-sum"]) await p.uncheck(id);
  let got = false; p.once("download", () => { got = true; });
  await p.click("[data-act=export-xlsx]");
  assert((await toast(p, "צריך לבחור")).includes("צריך לבחור") && !got, "no validation");
});
test("XLS-04", "אקסל", "ייצוא קניות ל־CSV", "נתונים לדוגמה ← 'ייצוא הקניות ל־CSV'", "קובץ עם BOM, כותרת ו־3 קניות, מרכאות מוברחות", async p => {
  await richData(p); await tab(p, "reports");
  const t = fs.readFileSync(await download(p, "[data-act=export-csv]"), "utf8");
  assert(t.charCodeAt(0) === 0xFEFF, "no BOM");
  assert(t.trim().split(/\r\n/).length === 4, "rows " + t.trim().split(/\r\n/).length);
  assert(t.includes('"LG, ""מבצע"""'), "quote escaping");
});
test("BAK-01", "גיבוי", "גיבוי ושחזור", "נתונים לדוגמה ← שמירת גיבוי ← דפדפן נקי ← שחזור", "כל הנתונים חוזרים (קניות, פריטים שהוסרו, משימות)", async p => {
  await richData(p); await p.click("[data-act=go][data-v=settings]");
  const f = await download(p, "[data-act=backup]");
  const s1 = await state(p);
  const p2 = await open();
  await p2.setInputFiles("#import-file", f); await p2.click(".sheet .btn.primary");
  const s2 = await state(p2);
  assert(JSON.stringify(s2.items) === JSON.stringify(s1.items), "mismatch");
  await p2.context().close();
});
test("UI-02", "תצוגה", "סגירת חלון", "פתיחת טופס פריט ← Escape; פתיחה ← לחיצה על הרקע", "החלון נסגר בשני המקרים", async p => {
  await onboard(p);
  await p.click("#fab"); await p.keyboard.press("Escape"); assert(!(await sheetOpen(p)), "esc");
  await p.click("#fab"); await p.mouse.click(10, 10); assert(!(await sheetOpen(p)), "backdrop");
});
test("UI-03", "תצוגה", "זום אוטומטי באייפון", "בדיקת גודל הגופן בשדות קלט", "לפחות 16px", async p => {
  await onboard(p); await p.click("#fab");
  const fs_ = await p.$eval(".sheet [name=planned]", e => parseFloat(getComputedStyle(e).fontSize));
  assert(fs_ >= 16, "input font-size " + fs_ + "px");
});
test("UI-06", "תצוגה", "מקלדת ברשימת הקניות", "Tab לשורת פריט ← Enter", "נפתח טופס עריכת הפריט", async p => {
  await onboard(p); await tab(p, "shop");
  await p.focus(".li[data-act=edit-item]"); await p.keyboard.press("Enter");
  assert(await sheetOpen(p), "Enter did nothing");
});
/* ---------- runner ---------- */
(async () => {
  const srv = await serve();
  BASE = `http://localhost:${srv.address().port}/`;
  browser = await chromium.launch();
  const only = process.argv[2] ? new RegExp(process.argv[2]) : null;
  const results = [];
  for (const t of T) {
    if (only && !only.test(t.id)) continue;
    const p = await open();
    const t0 = Date.now();
    let status = "PASS", actual = "כמצופה";
    try { await t.fn(p); if (p.errors.length && t.id !== "UI-07") { status = "FAIL"; actual = "JS error: " + p.errors[0]; } }
    catch (e) { status = "FAIL"; actual = e.message.split("\n")[0].slice(0, 220); }
    await p.context().close();
    results.push({ id: t.id, module: t.module, name: t.name, steps: t.steps, expected: t.expected, status, actual, ms: Date.now() - t0 });
    console.log(`${status === "PASS" ? "✓" : "✗"} ${t.id} ${t.name}${status === "PASS" ? "" : "  →  " + actual}`);
  }
  await browser.close(); srv.close();
  fs.writeFileSync(path.join(__dirname, "results.json"), JSON.stringify(results, null, 1));
  const f = results.filter(r => r.status === "FAIL").length;
  console.log(`\n${results.length - f}/${results.length} passed`);
  process.exit(f ? 1 : 0);
})();

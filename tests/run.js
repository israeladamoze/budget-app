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
async function addExpense(p, e) {
  await tab(p, "expenses");
  await p.click("#fab");
  if (e.amount !== undefined) await p.fill(".sheet [name=amount]", String(e.amount));
  if (e.cat) await p.selectOption(".sheet [name=cat]", e.cat);
  if (e.date) await p.fill(".sheet [name=date]", e.date);
  if (e.payer) await p.selectOption(".sheet [name=payer]", e.payer);
  if (e.method) await p.selectOption(".sheet [name=method]", e.method);
  if (e.inst !== undefined) await p.fill(".sheet [name=inst]", String(e.inst));
  if (e.note !== undefined) await p.fill(".sheet [name=note]", e.note);
  await submit(p);
}
async function openItem(p, name) { await tab(p, "move"); await p.click(`.li[data-act=edit-item] .t:text-is("${name}")`); }
async function openBill(p, name) { await tab(p, "bills"); await p.click(`.li[data-act=edit-bill] .t:text-is("${name}")`); }
async function setBill(p, name, f) {
  await openBill(p, name);
  for (const [k, v] of Object.entries(f)) {
    if (k === "freq" || k === "payer") await p.selectOption(`.sheet [name=${k}]`, String(v));
    else if (k === "active") await p.setChecked(".sheet [name=active]", v);
    else await p.fill(`.sheet [name=${k}]`, String(v));
  }
  await submit(p);
}
async function payBill(p, name, amount, note) {
  await tab(p, "bills");
  const row = p.locator(".li[data-act=edit-bill]", { has: p.locator(`.t:text-is("${name}")`) });
  await row.locator("[data-act=pay-bill]").click();
  if (amount !== undefined) await p.fill(".sheet [name=amount]", String(amount));
  if (note) await p.fill(".sheet [name=note]", note);
  await submit(p);
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
test("ONB-03", "פתיחה", "השלמת פתיחה", "מילוי שמות, תאריך, תקציב 45000 ומשתנות 3000", "מסך בית עם ברכה בשמות; 32 פריטי מעבר ו־8 חשבונות ברירת מחדל", async p => {
  await onboard(p, { moveDate: iso(29), budget: 45000, variable: 3000 });
  assert((await p.textContent("h1")).includes("ישראלה") && (await p.textContent("h1")).includes("דני"), "greeting");
  const s = await state(p);
  assert(s.items.length === 32 && s.bills.length === 8, `items ${s.items.length} bills ${s.bills.length}`);
  assert(s.settings.moveBudget === 45000 && s.settings.variable === 3000, "settings saved");
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
test("SET-02", "הגדרות", "הוספת קטגוריה", "הוספת 'חיות מחמד', ואז הוספה חוזרת של אותו שם", "הקטגוריה מופיעה בטופס ההוצאה; כפילות נדחית עם הודעה", async p => {
  await onboard(p);
  await p.click("[data-act=go][data-v=settings]");
  await p.fill("#new-cat", "חיות מחמד"); await p.click("[data-act=add-cat]");
  await p.fill("#new-cat", "חיות מחמד"); await p.click("[data-act=add-cat]");
  assert((await toast(p, "כבר קיימת")).includes("כבר קיימת"), "no duplicate warning");
  const s = await state(p);
  assert(s.settings.cats.filter(c => c === "חיות מחמד").length === 1, "duplicate saved");
  await tab(p, "expenses"); await p.click("#fab");
  assert((await p.$$eval(".sheet [name=cat] option", o => o.map(x => x.textContent))).includes("חיות מחמד"), "not in dropdown");
});
test("SET-03", "הגדרות", "מחיקת קטגוריה בשימוש", "הוספת קטגוריה, רישום הוצאה בה, מחיקת הקטגוריה, פתיחת ההוצאה ושמירה", "ההוצאה שומרת על הקטגוריה המקורית שלה", async p => {
  await onboard(p);
  await p.click("[data-act=go][data-v=settings]");
  await p.fill("#new-cat", "חיות מחמד"); await p.click("[data-act=add-cat]");
  await addExpense(p, { amount: 200, cat: "חיות מחמד", note: "וטרינר" });
  await p.click("[data-act=go][data-v=settings]");
  await p.click("[data-act=del-cat]");
  if (await sheetOpen(p)) await p.click(".sheet .btn.danger");
  await tab(p, "expenses");
  await p.click(".li[data-act=edit-expense]");
  await submit(p);
  const s = await state(p);
  assert(s.expenses[0].cat === "חיות מחמד", "category silently changed to " + s.expenses[0].cat);
});
test("SET-04", "הגדרות", "איפוס", "הגדרות ← איפוס ← מחיקה והתחלה מחדש", "חוזרים למסך פתיחה, ההוצאות נמחקו", async p => {
  await onboard(p); await addExpense(p, { amount: 50 });
  await p.click("[data-act=go][data-v=settings]"); await p.click("[data-act=reset]");
  await p.click(".sheet .btn.danger");
  assert(await p.isVisible("#welcome-form"), "not back to welcome");
  const s = await state(p); assert(s.expenses.length === 0, "expenses kept");
});

/* --- expenses --- */
test("EXP-01", "הוצאות", "הוספת הוצאה", "כפתור + ← 420 ₪, סופר ובית, הערה ← שמירה", "ההוצאה מופיעה ברשימה והסכום החודשי 420", async p => {
  await onboard(p);
  await addExpense(p, { amount: 420, cat: "סופר ובית", note: "קניות לבית החדש" });
  assert((await p.textContent("#exp-list")).includes("קניות לבית החדש"), "not listed");
  assert((await p.textContent("#exp-list .card")).includes("420"), "total");
});
test("EXP-02", "הוצאות", "סכום חסר או אפס", "שמירת הוצאה בלי סכום, ואז עם 0", "ההוצאה לא נשמרת", async p => {
  await onboard(p);
  await addExpense(p, {});
  await p.fill(".sheet [name=amount]", "0"); await submit(p);
  const s = await state(p); assert(s.expenses.length === 0, "saved invalid expense");
});
test("EXP-03", "הוצאות", "סכום עשרוני", "הוצאה של 12.90", "נשמר 12.9 ומוצג ₪12.9", async p => {
  await onboard(p); await addExpense(p, { amount: "12.90" });
  const s = await state(p); assert(s.expenses[0] && s.expenses[0].amount === 12.9, "amount " + JSON.stringify(s.expenses));
});
test("EXP-04", "הוצאות", "עריכת הוצאה", "פתיחת הוצאה ← שינוי סכום ל־500 והערה ← שמירה", "הסכום וההערה מתעדכנים, אין הוצאה כפולה", async p => {
  await onboard(p); await addExpense(p, { amount: 100, note: "א" });
  await p.click(".li[data-act=edit-expense]");
  await p.fill(".sheet [name=amount]", "500"); await p.fill(".sheet [name=note]", "ב");
  await submit(p);
  const s = await state(p); assert(s.expenses.length === 1 && s.expenses[0].amount === 500 && s.expenses[0].note === "ב", JSON.stringify(s.expenses));
});
test("EXP-05", "הוצאות", "מחיקת הוצאה", "פתיחת הוצאה ← מחיקה ← אישור", "ההוצאה נמחקת אחרי אישור בלבד", async p => {
  await onboard(p); await addExpense(p, { amount: 100 });
  await p.click(".li[data-act=edit-expense]"); await p.click(".sheet [data-act=delete-expense]");
  assert((await state(p)).expenses.length === 1, "deleted before confirm");
  await p.click(".sheet .btn.danger");
  assert((await state(p)).expenses.length === 0, "not deleted");
});
test("EXP-06", "הוצאות", "חיפוש בכל החודשים", "הוצאה בחודש קודם עם הערה 'וטרינר' ← חיפוש 'וטרינר'", "ההוצאה נמצאת למרות שהיא בחודש אחר", async p => {
  await onboard(p); await addExpense(p, { amount: 80, date: ymShift(-2) + "-10", note: "וטרינר" });
  await p.fill("#exp-q", "וטרינר");
  assert((await p.textContent("#exp-list")).includes("וטרינר"), "not found");
});
test("EXP-07", "הוצאות", "סינון קטגוריה", "שתי הוצאות בקטגוריות שונות ← סינון לחשמל", "מוצגת רק הוצאת החשמל", async p => {
  await onboard(p); await addExpense(p, { amount: 80, cat: "חשמל", note: "חשמל1" }); await addExpense(p, { amount: 90, cat: "מים", note: "מים1" });
  await p.selectOption("#exp-cat", "חשמל");
  const t = await p.textContent("#exp-list"); assert(t.includes("חשמל1") && !t.includes("מים1"), "filter failed: " + t.replace(/\s+/g, " ").slice(0, 160));
});
test("EXP-08", "הוצאות", "ניווט בין חודשים", "הוצאה בתאריך בחודש הקודם ← מעבר בין החודשים", "ההוצאה מופיעה רק בחודש שלה", async p => {
  await onboard(p); await addExpense(p, { amount: 70, date: ymShift(-1) + "-15", note: "קודם" });
  await p.click("[data-act=go][data-v=home]"); await tab(p, "expenses");
  await p.click("[data-act=exp-month][data-d='1']");
  assert(!(await p.textContent("#exp-list")).includes("קודם"), "visible in the following month");
  await p.click("[data-act=exp-month][data-d='-1']");
  assert((await p.textContent("#exp-list")).includes("קודם"), "not in its own month");
});
test("EXP-09", "הוצאות", "תקרת תשלומים", "הזנת 100 תשלומים (המקסימום בשדה 36)", "הערך נדחה או מוגבל ל־36", async p => {
  await onboard(p); await addExpense(p, { amount: 1000, inst: 100 });
  const s = await state(p); assert(!s.expenses[0] || s.expenses[0].inst <= 36, "inst=" + (s.expenses[0] && s.expenses[0].inst));
});
test("EXP-10", "הוצאות", "הגנה מהזרקת קוד", "הערה: <img src=x onerror=window.__x=1>", "הטקסט מוצג כמו שהוא ולא מורץ", async p => {
  await onboard(p); await addExpense(p, { amount: 5, note: "<img src=x onerror=window.__x=1>" });
  await tab(p, "home");
  assert(!(await p.evaluate(() => window.__x)), "script executed");
  assert((await p.textContent("#main")).includes("<img src=x"), "text not shown");
});
test("EXP-11", "הוצאות", "הוצאה בחודש אחר – משוב", "הוספת הוצאה עם תאריך בחודש קודם מתוך מסך ההוצאות", "המסך עובר לחודש של ההוצאה כדי שאפשר יהיה לראות אותה", async p => {
  await onboard(p); await addExpense(p, { amount: 70, date: ymShift(-1) + "-15", note: "קודם" });
  assert((await p.textContent("#exp-list")).includes("קודם"), "expense saved but not visible – user sees nothing changed");
});

/* --- move --- */
test("MOV-01", "מעבר", "קניית פריט", "מקרר ← מחיר בפועל 5200 ← שמירה", "סטטוס 'התקבל', 'שולם' 5,200, מוצג +700 מהתכנון", async p => {
  await onboard(p, { budget: 45000 }); await openItem(p, "מקרר");
  await p.fill(".sheet [name=actual]", "5200"); await submit(p);
  const s = await state(p), i = s.items.find(x => x.name === "מקרר");
  assert(i.actual === 5200 && i.status === "got", JSON.stringify(i));
  assert((await p.textContent("#main")).includes("+₪700"), "diff not shown");
});
test("MOV-02", "מעבר", "מחיר עם אגורות", "מקרר ← מחיר בפועל 2399.90", "המחיר נשמר כ־2399.9", async p => {
  await onboard(p); await openItem(p, "מקרר");
  await p.fill(".sheet [name=actual]", "2399.90"); await submit(p);
  const i = (await state(p)).items.find(x => x.name === "מקרר");
  assert(i.actual === 2399.9, "actual=" + i.actual + (await sheetOpen(p) ? " (form blocked by browser validation)" : ""));
});
test("MOV-03", "מעבר", "'לא צריך' ושחזור", "מזגן ← לא צריך ← הצגת פריטים שהוסרו ← החזרה", "המזגן יוצא מהצפי (−3000) וחוזר אחרי שחזור", async p => {
  await onboard(p);
  const before = await p.evaluate(() => 0);
  await openItem(p, "מזגן + התקנה"); await p.click(".sheet [data-act=remove-item]");
  let s = await state(p); assert(s.items.find(x => x.name === "מזגן + התקנה").removed, "not removed");
  await p.click("[data-act=toggle-removed]"); await p.click("[data-act=restore-item]");
  s = await state(p); assert(!s.items.find(x => x.name === "מזגן + התקנה").removed, "not restored");
});
test("MOV-04", "מעבר", "הסרת פריט ששולם – השארה כהוצאה", "מכונת כביסה ← 2300 ← לא צריך ← 'להשאיר כהוצאה כללית'", "נוצרת הוצאה 'מעבר דירה' 2300; 'שולם' ו'צפי כולל' כוללים אותה", async p => {
  await onboard(p, { budget: 45000 }); await openItem(p, "מכונת כביסה");
  await p.fill(".sheet [name=actual]", "2300"); await submit(p);
  const planned0 = await p.evaluate(() => 0);
  await openItem(p, "מכונת כביסה"); await p.click(".sheet [data-act=remove-item]"); await p.click(".sheet [data-cb='0']");
  const s = await state(p);
  assert(s.expenses.length === 1 && s.expenses[0].amount === 2300 && s.expenses[0].cat === "מעבר דירה", "expense not created");
  await tab(p, "move");
  const stats = await p.$$eval(".stat .v", v => v.map(x => x.textContent.replace(/[^\d.]/g, "")));
  const planned = +stats[1], spent = +stats[2];
  assert(spent === 2300, "spent " + spent);
  // default plan 52,500 − washing machine plan 2,500 + 2,300 actually paid = 52,300
  assert(planned === 52300, `planned ${planned}, expected 52300 – the 2,300 kept as an expense is missing from the forecast`);
});
test("MOV-05", "מעבר", "הסרת פריט ששולם – מחיקת ההוצאה", "מכונת כביסה ← 2300 ← לא צריך ← 'למחוק גם את ההוצאה'", "אין הוצאה, 'שולם' חוזר ל־0", async p => {
  await onboard(p); await openItem(p, "מכונת כביסה");
  await p.fill(".sheet [name=actual]", "2300"); await submit(p);
  await openItem(p, "מכונת כביסה"); await p.click(".sheet [data-act=remove-item]"); await p.click(".sheet [data-cb='1']");
  const s = await state(p); const i = s.items.find(x => x.name === "מכונת כביסה");
  assert(s.expenses.length === 0 && i.actual == null && i.removed, JSON.stringify(i));
});
test("MOV-06", "מעבר", "מחיקה לצמיתות", "טלוויזיה ← מחיקה לצמיתות ← מחיקה", "הפריט נמחק מהנתונים", async p => {
  await onboard(p); await openItem(p, "טלוויזיה"); await p.click(".sheet [data-act=delete-item]"); await p.click(".sheet .btn.danger");
  assert(!(await state(p)).items.some(x => x.name === "טלוויזיה"), "still exists");
});
test("MOV-07", "מעבר", "הוספת פריט", "סלון ← + פריט ← 'כורסה' 1800 ← שמירה", "הפריט נוסף לסלון", async p => {
  await onboard(p); await tab(p, "move");
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
test("MOV-09", "מעבר", "צ׳קליסט נשמר", "צ׳קליסט ← סימון 2 משימות ← טעינה מחדש", "הסימונים נשארים, המונה מציג 2/15", async p => {
  await onboard(p); await tab(p, "move"); await p.click("[data-act=move-tab][data-v=check]");
  await p.click("[data-k=meters]"); await p.click("[data-k=arnona]");
  await p.reload(); await tab(p, "move"); await p.click("[data-act=move-tab][data-v=check]");
  assert((await p.textContent(".chip.done")).includes("2/15"), await p.textContent(".chip.done"));
});
test("MOV-10", "מעבר", "קריאת מונים נשמרת", "הזנת קריאת מים 1234 ← טעינה מחדש", "הערך נשמר", async p => {
  await onboard(p); await tab(p, "move"); await p.click("[data-act=move-tab][data-v=check]");
  await p.fill("[data-meter=water]", "1234"); await p.press("[data-meter=water]", "Tab");
  await p.reload(); await tab(p, "move"); await p.click("[data-act=move-tab][data-v=check]");
  assert((await p.inputValue("[data-meter=water]")) === "1234", "lost");
});
test("MOV-11", "מעבר", "פיקדון לא נספר כהוצאה", "פיקדונות ← הוספת פיקדון 11000 ← הוצאות", "הפיקדון מוצג בהמתנה ולא מופיע בהוצאות", async p => {
  await onboard(p); await tab(p, "move"); await p.click("[data-act=move-tab][data-v=deposits]");
  await p.click("[data-act=add-deposit]"); await p.fill(".sheet [name=name]", "פיקדון לבעל הדירה"); await p.fill(".sheet [name=amount]", "11000"); await submit(p);
  assert((await p.textContent("#main")).includes("11,000"), "not shown");
  await tab(p, "expenses"); assert(!(await p.textContent("#exp-list")).includes("11,000"), "counted as expense");
});
test("MOV-12", "מעבר", "מחיר שלילי", "מקרר ← מחיר בפועל ‎-500", "הערך נדחה", async p => {
  await onboard(p); await openItem(p, "מקרר"); await p.fill(".sheet [name=actual]", "-500"); await submit(p);
  const i = (await state(p)).items.find(x => x.name === "מקרר"); assert(i.actual == null || i.actual >= 0, "negative saved: " + i.actual);
});

/* --- bills --- */
test("BIL-01", "חשבונות", "ממוצע חודשי", "שכר דירה 5500 חודשי, חשמל 450 דו־חודשי", "ממוצע לחודש 5,725", async p => {
  await onboard(p); await setBill(p, "שכר דירה", { amount: 5500 }); await setBill(p, "חשמל", { amount: 450 });
  assert((await p.textContent(".stat .v")).replace(/[^\d]/g, "") === "5725", await p.textContent(".stat .v"));
});
test("BIL-02", "חשבונות", "סימון 'שולם'", "חשמל עם חיוב הבא היום ← שולם 430 עם הערה", "נוצרת הוצאה עם ההערה; החיוב הבא זז חודשיים קדימה", async p => {
  await onboard(p); await setBill(p, "חשמל", { amount: 450, nextDue: iso(0) });
  await payBill(p, "חשמל", 430, "מזגן בקיץ");
  const s = await state(p), b = s.bills.find(x => x.name === "חשמל");
  assert(s.expenses.length === 1 && s.expenses[0].note.includes("מזגן בקיץ") && s.expenses[0].amount === 430, JSON.stringify(s.expenses));
  const d = new Date(); d.setMonth(d.getMonth() + 2);
  assert(b.nextDue.slice(0, 7) === d.toISOString().slice(0, 7), "nextDue " + b.nextDue);
});
test("BIL-03", "חשבונות", "התראה על חשבון חריג", "שני תשלומי מים של 200 ← תשלום שלישי 300", "הודעה שהחשבון גבוה ב־50% מהממוצע", async p => {
  await onboard(p); await setBill(p, "מים", { amount: 200 });
  await payBill(p, "מים", 200); await payBill(p, "מים", 200); await payBill(p, "מים", 300);
  assert((await toast(p, "50%")).includes("50%"), "no warning: " + (await toast(p)));
});
test("BIL-04", "חשבונות", "כיבוי חשבון", "שכר דירה 5500 ← כיבוי 'רלוונטי לנו'", "יוצא מהממוצע החודשי ומההעברה המוצעת", async p => {
  await onboard(p, { variable: 0 }); await setBill(p, "שכר דירה", { amount: 5500 }); await setBill(p, "שכר דירה", { active: false });
  assert((await p.textContent(".stat .v")).replace(/[^\d]/g, "") === "0", "still counted");
});
test("BIL-05", "חשבונות", "חשבון באיחור", "ארנונה 600 עם חיוב הבא לפני 3 ימים ← בית", "התראה אדומה 'עבר מועד החיוב' עם כפתור תשלום", async p => {
  await onboard(p); await setBill(p, "ארנונה", { amount: 600, nextDue: iso(-3) }); await tab(p, "home");
  const t = await p.textContent(".banner.bad"); assert(t.includes("ארנונה") && t.includes("עבר מועד"), t);
});
test("BIL-06", "חשבונות", "חשבון קרוב", "מים עם חיוב בעוד יומיים ← בית", "התראה כתומה 'בעוד 2 ימים'", async p => {
  await onboard(p); await setBill(p, "מים", { amount: 150, nextDue: iso(2) }); await tab(p, "home");
  assert((await p.textContent(".banner.warn")).includes("בעוד 2 ימים"), "no warning");
});
test("BIL-07", "חשבונות", "הוספה ומחיקה של חשבון", "הוספת 'נטפליקס' 55 ← מחיקה", "נוסף ואז נמחק", async p => {
  await onboard(p); await tab(p, "bills"); await p.click("[data-act=add-bill]");
  await p.fill(".sheet [name=name]", "נטפליקס"); await p.fill(".sheet [name=amount]", "55"); await submit(p);
  assert((await state(p)).bills.some(b => b.name === "נטפליקס"), "not added");
  await openBill(p, "נטפליקס"); await p.click(".sheet [data-act=delete-bill]"); await p.click(".sheet .btn.danger");
  assert(!(await state(p)).bills.some(b => b.name === "נטפליקס"), "not deleted");
});
test("BIL-08", "חשבונות", "חשבון מחשבון פרטי", "אינטרנט 100 ← 'מי משלם' = א׳ ← בית", "לא נכלל בסכום שכל אחד מעביר לחשבון המשותף", async p => {
  await onboard(p, { variable: 0 }); await setBill(p, "אינטרנט", { amount: 100, payer: "a" }); await tab(p, "home");
  assert(!(await p.textContent("#main")).includes("כל אחד מעביר החודש"), "personal bill counted in joint need");
});
test("BIL-09", "חשבונות", "סכום חשבון עם אגורות", "חשמל ← סכום 452.37", "הסכום נשמר", async p => {
  await onboard(p); await setBill(p, "חשמל", { amount: "452.37" });
  assert((await state(p)).bills.find(b => b.name === "חשמל").amount === 452.37, "blocked");
});

/* --- home / joint account --- */
test("HOM-01", "בית", "סכום העברה מוצע", "שכר דירה 5500, חשמל 450 דו־חודשי, משתנות 3000", "כל אחד מעביר 4,363 (8,725 חלקי 2)", async p => {
  await onboard(p, { variable: 3000 }); await setBill(p, "שכר דירה", { amount: 5500 }); await setBill(p, "חשמל", { amount: 450 }); await tab(p, "home");
  assert((await p.textContent("#main")).includes("4,363"), "suggestion wrong");
});
test("HOM-02", "בית", "רישום העברה", "רישום העברה של א׳ בסכום המוצע ← הוצאה משותפת 600", "'הועבר אליו' 4,363, 'יצא ממנו' 600, א׳ מסומן ✓", async p => {
  await onboard(p, { variable: 3000 }); await setBill(p, "שכר דירה", { amount: 5500 }); await setBill(p, "חשמל", { amount: 450 }); await tab(p, "home");
  await p.click("[data-act=add-transfer]"); await submit(p);
  await addExpense(p, { amount: 600, cat: "סופר ובית" }); await tab(p, "home");
  const cols = await p.$$eval(".hero .cols b", x => x.map(e => e.textContent.replace(/[^\d]/g, "")));
  assert(cols[0] === "4363" && cols[1] === "600", JSON.stringify(cols));
  assert(await p.$(".chip.done"), "no ✓");
});
test("HOM-03", "בית", "איזון 50/50 וסגירת חשבון", "א׳ משלם 1000 מהחשבון הפרטי ← סגירת חשבון", "'דני צריך להעביר 500', ואחרי הרישום – 'מאוזנים'", async p => {
  await onboard(p); await addExpense(p, { amount: 1000, payer: "a" }); await tab(p, "home");
  assert((await p.textContent("#main")).includes("₪500"), "balance wrong");
  await p.click("[data-act=settle]"); await submit(p);
  assert((await p.textContent("#main")).includes("מאוזנים"), "not balanced after settle");
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
test("REP-01", "דוחות", "פילוח לפי קטגוריה", "הוצאות 300 סופר + 100 סופר + 200 חשמל ← דוחות", "סה״כ 600; סופר 400 (67%)", async p => {
  await onboard(p); await addExpense(p, { amount: 300, cat: "סופר ובית" }); await addExpense(p, { amount: 100, cat: "סופר ובית" }); await addExpense(p, { amount: 200, cat: "חשמל" });
  await tab(p, "reports"); const t = await p.textContent(".hbars");
  assert(t.includes("400") && t.includes("67%"), t);
});
test("REP-02", "דוחות", "תחזית תשלומי אשראי", "קנייה 1200 באשראי ב־12 תשלומים ← דוחות", "100 ₪ בכל אחד מ־6 החודשים הבאים", async p => {
  await onboard(p); await addExpense(p, { amount: 1200, method: "אשראי", inst: 12 }); await tab(p, "reports");
  const cells = await p.$$eval("table.t:nth-of-type(1)", () => 0);
  const rows = await p.$$eval("section.card", s => s.find(x => x.textContent.includes("תחזית")).querySelectorAll("tbody tr").length);
  const vals = await p.$$eval("section.card", s => [...s.find(x => x.textContent.includes("תחזית")).querySelectorAll("tbody tr td:nth-child(3)")].map(td => td.textContent.replace(/[^\d]/g, "")));
  assert(rows === 6 && vals.every(v => v === "100"), JSON.stringify(vals) + " state=" + JSON.stringify((await state(p)).expenses.map(e => [e.amount, e.method, e.inst, e.date])));
});
test("REP-03", "דוחות", "טבלת מעבר – סכום השורות", "מכונת כביסה 2300 ← לא צריך ← להשאיר כהוצאה ← דוחות", "שורת הסה״כ 'בפועל' שווה לסכום השורות שמעליה", async p => {
  await onboard(p); await openItem(p, "מכונת כביסה"); await p.fill(".sheet [name=actual]", "2300"); await submit(p);
  await openItem(p, "מכונת כביסה"); await p.click(".sheet [data-act=remove-item]"); await p.click(".sheet [data-cb='0']");
  await tab(p, "reports");
  const r = await p.$$eval("section.card", s => { const c = s.find(x => x.textContent.includes("תכנון מול בפועל")); const rows = [...c.querySelectorAll("tbody tr")]; return rows.map(tr => tr.lastElementChild.textContent.replace(/[^\d]/g, "")); });
  const total = +r.pop(), sumRows = r.reduce((a, b) => a + +b, 0);
  assert(total === sumRows, `total ${total} ≠ rows ${sumRows}`);
});

/* --- excel / csv / backup --- */
async function richData(p) {
  await onboard(p, { budget: 45000, variable: 3000 });
  await setBill(p, "שכר דירה", { amount: 5500, nextDue: iso(5) });
  await setBill(p, "חשמל", { amount: 450 });
  await payBill(p, "שכר דירה", 5500, "נובמבר");
  await openItem(p, "מקרר"); await p.fill(".sheet [name=actual]", "5200"); await submit(p);
  await addExpense(p, { amount: 420, cat: "סופר ובית", payer: "a", note: "שופרסל, \"מבצע\"" });
  await addExpense(p, { amount: 90, cat: "סופר ובית", date: ymShift(-1) + "-20", note: "חודש קודם" });
  await tab(p, "home"); await p.click("[data-act=add-transfer]"); await submit(p);
}
test("XLS-01", "אקסל", "ייצוא מלא", "נתונים לדוגמה ← דוחות ← הכול ← הורדה; חישוב מחדש ב־LibreOffice", "4 גיליונות מימין לשמאל; הנוסחאות מחושבות לאותם ערכים שבאפליקציה", async p => {
  await richData(p); await tab(p, "reports");
  const f = await download(p, "[data-act=export-xlsx]");
  const x = xlsx(f, false);
  assert(JSON.stringify(Object.keys(x)) === JSON.stringify(["הוצאות", "מעבר דירה", "חשבונות קבועים", "סיכום"]), Object.keys(x).join(","));
  assert(Object.values(x).every(s => s.rtl), "not RTL");
  const exp = x["הוצאות"].rows.filter(r => typeof r[1] === "number");
  assert(exp.length === 4, "expense rows " + exp.length);
  const rc = xlsx(f, true);
  const tot = rc["הוצאות"].rows.find(r => r[0] === "סה״כ");
  assert(Math.abs(tot[1] - (5500 + 5200 + 420 + 90)) < .01, "sum formula " + tot[1]);
  const sumTot = rc["סיכום"].rows.find(r => r[0] === "סה״כ הוצאות");
  assert(Math.abs(sumTot[1] - 11210) < .01, "summary total " + sumTot[1]);
  const fridge = rc["מעבר דירה"].rows.find(r => r[1] === "מקרר");
  assert(fridge[5] === 700, "diff formula " + fridge[5]);
});
test("XLS-02", "אקסל", "ייצוא חודש אחד", "טווח 'החודש שנבחר' ← הורדה", "רק הוצאות החודש הנוכחי (בלי ה־90 מהחודש הקודם)", async p => {
  await richData(p); await tab(p, "reports"); await p.selectOption("#x-range", "month");
  const x = xlsx(await download(p, "[data-act=export-xlsx]"), false);
  const notes = x["הוצאות"].rows.slice(4).map(r => r[6]).filter(Boolean);
  assert(!notes.includes("חודש קודם") && notes.length === 3, JSON.stringify(notes));
});
test("XLS-03", "אקסל", "ייצוא סיכום בלבד", "ביטול כל הסימונים חוץ מ'סיכום' ← הורדה", "בקובץ רק גיליון 'סיכום', עם ערכים נכונים", async p => {
  await richData(p); await tab(p, "reports");
  for (const id of ["#x-exp", "#x-move", "#x-bills"]) await p.uncheck(id);
  const x = xlsx(await download(p, "[data-act=export-xlsx]"), true);
  assert(JSON.stringify(Object.keys(x)) === '["סיכום"]', "sheets: " + Object.keys(x).join(","));
  assert(Math.abs(x["סיכום"].rows.find(r => r[0] === "סה״כ הוצאות")[1] - 11210) < .01, "summary values");
});
test("XLS-04", "אקסל", "ייצוא בלי תוכן", "ביטול כל הסימונים ← הורדה", "הודעה 'צריך לבחור', אין הורדה", async p => {
  await onboard(p); await tab(p, "reports");
  for (const id of ["#x-exp", "#x-move", "#x-bills", "#x-sum"]) await p.uncheck(id);
  let got = false; p.once("download", () => { got = true; });
  await p.click("[data-act=export-xlsx]");
  assert((await toast(p, "צריך לבחור")).includes("צריך לבחור") && !got, "no validation");
});
test("XLS-05", "אקסל", "ייצוא CSV", "נתונים לדוגמה ← 'ייצוא הוצאות ל־CSV'", "קובץ עם BOM (עברית תקינה באקסל), כותרת ו־4 שורות, מרכאות מוברחות", async p => {
  await richData(p); await tab(p, "reports");
  const t = fs.readFileSync(await download(p, "[data-act=export-csv]"), "utf8");
  assert(t.charCodeAt(0) === 0xFEFF, "no BOM");
  assert(t.trim().split(/\r\n/).length === 5, "rows " + t.trim().split(/\r\n/).length);
  assert(t.includes('"שופרסל, ""מבצע"""'), "quote escaping");
});
test("BAK-01", "גיבוי", "גיבוי ושחזור", "נתונים לדוגמה ← שמירת גיבוי ← דפדפן נקי ← שחזור מהקובץ", "כל הנתונים חוזרים (הוצאות, פריטים, העברות)", async p => {
  await richData(p); await p.click("[data-act=go][data-v=settings]");
  const f = await download(p, "[data-act=backup]");
  const s1 = await state(p);
  const p2 = await open();
  await p2.evaluate(() => { localStorage.clear(); }); await p2.reload();
  await p2.setInputFiles("#import-file", f); await p2.click(".sheet .btn.primary");
  const s2 = await state(p2);
  assert(s2.expenses.length === s1.expenses.length && s2.transfers.length === 1 && s2.items.length === s1.items.length, "mismatch");
  await p2.context().close();
});
test("BAK-02", "גיבוי", "קובץ לא תקין", "שחזור מקובץ טקסט", "הודעת שגיאה, הנתונים לא משתנים", async p => {
  await onboard(p); await addExpense(p, { amount: 10 });
  const f = path.join(OUT, "bad.json"); fs.writeFileSync(f, "not json");
  await p.setInputFiles("#import-file", f);
  assert((await toast(p, "קובץ")).includes("לא"), "no error");
  assert((await state(p)).expenses.length === 1, "data changed");
});

/* --- UI / accessibility --- */
test("UI-01", "תצוגה", "מסך צר 320px", "מעבר על כל הלשוניות ברוחב 320", "אין גלילה אופקית", async p => {
  await p.setViewportSize({ width: 320, height: 568 });
  await richData(p);
  for (const v of ["home", "move", "bills", "expenses", "reports"]) {
    await tab(p, v);
    const w = await p.evaluate(() => document.documentElement.scrollWidth);
    assert(w <= 320, `${v}: scrollWidth ${w}`);
  }
});
test("UI-02", "תצוגה", "סגירת חלון", "פתיחת טופס הוצאה ← Escape; פתיחה ← לחיצה על הרקע", "החלון נסגר בשני המקרים", async p => {
  await onboard(p); await tab(p, "expenses");
  await p.click("#fab"); await p.keyboard.press("Escape"); assert(!(await sheetOpen(p)), "esc");
  await p.click("#fab"); await p.mouse.click(10, 10); assert(!(await sheetOpen(p)), "backdrop");
});
test("UI-03", "תצוגה", "זום אוטומטי באייפון", "בדיקת גודל הגופן בשדות קלט", "לפחות 16px (מתחת לזה ספארי מגדיל את המסך בכל לחיצה על שדה)", async p => {
  await onboard(p); await tab(p, "expenses"); await p.click("#fab");
  const fs_ = await p.$eval(".sheet [name=amount]", e => parseFloat(getComputedStyle(e).fontSize));
  assert(fs_ >= 16, "input font-size " + fs_ + "px");
});
test("UI-04", "תצוגה", "כפתור 'חזור' באנדרואיד", "פתיחת טופס ← חזור בדפדפן", "הטופס נסגר והאפליקציה נשארת פתוחה", async p => {
  await onboard(p); await tab(p, "move"); await openItem(p, "מקרר");
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
test("UI-06", "תצוגה", "מקלדת בחשבונות", "Tab לשורת חשבון ← Enter", "נפתח טופס עריכת החשבון", async p => {
  await onboard(p); await tab(p, "bills");
  await p.focus(".li[data-act=edit-bill]"); await p.keyboard.press("Enter");
  assert(await sheetOpen(p), "Enter did nothing");
});
test("UI-07", "תצוגה", "שגיאות JavaScript", "זרימה מלאה עם נתונים לדוגמה בכל הלשוניות", "אין שגיאות בקונסול", async p => {
  await richData(p); for (const v of ["move", "bills", "expenses", "reports", "home"]) await tab(p, v);
  assert(p.errors.length === 0, p.errors.join(" | "));
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

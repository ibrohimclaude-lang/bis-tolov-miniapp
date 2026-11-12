// Reconciliation act (Акт сверки) per client as an Excel file, built from the Google Sheets month cells:
// a month with an amount or «qarz» is invoiced at the line's monthly price (one invoice per contract,
// dated the last day of the month), and the amounts written in the month cells are the payments.
(function () {
  const US = { name: "\"Bclouds\" Mchj", director: "Ergashev Axmadjon Kaxramon", short: "Ergashev A.K." };
  const RU_MONTH = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"];
  // ExcelJS 4.4.0: the copy stored with the site, then the public CDN
  const EXCELJS = ["/_blob/8ae8be88696da1e8bb1be569a851c7eb", "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"];

  const n = (v) => Number(v) || 0;
  const pad = (v) => String(v).padStart(2, "0");
  const norm = (s) => String(s || "").toLowerCase().replace(/[«»"“”'’]/g, "").replace(/[^0-9a-zа-яёўқғҳ]+/gi, " ").trim();
  const sameCo = (a, b) => { const x = norm(a), y = norm(b); return !!x && !!y && (x === y || x.startsWith(y + " ") || y.startsWith(x + " ")); };
  const lastDay = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
  const dmy = (y, m, d) => pad(d) + "." + pad(m) + "." + y;

  // ---------- amount in Russian words ----------
  const W1 = ["", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"];
  const W1F = ["", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"];
  const W10 = ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"];
  const W20 = ["", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто"];
  const W100 = ["", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот"];
  const plural = (v, one, few, many) => { const a = v % 100, b = v % 10; return a >= 11 && a <= 19 ? many : b === 1 ? one : b >= 2 && b <= 4 ? few : many; };
  const triad = (v, fem) => {
    const out = [W100[Math.floor(v / 100)]], t = v % 100;
    if (t >= 10 && t < 20) out.push(W10[t - 10]);
    else out.push(W20[Math.floor(t / 10)], (fem ? W1F : W1)[t % 10]);
    return out.filter(Boolean).join(" ");
  };
  function words(amount) {
    const v = Math.round(Math.abs(amount) * 100), sum = Math.floor(v / 100), tiyin = v % 100;
    const parts = [];
    [[1e9, false, ["миллиард", "миллиарда", "миллиардов"]], [1e6, false, ["миллион", "миллиона", "миллионов"]], [1e3, true, ["тысяча", "тысячи", "тысяч"]]].forEach(([k, fem, f]) => {
      const t = Math.floor(sum / k) % 1000;
      if (t) parts.push(triad(t, fem) + " " + plural(t, f[0], f[1], f[2]));
    });
    const rest = sum % 1000;
    if (rest) parts.push(triad(rest, false));
    const text = (parts.join(" ") || "ноль") + " сум " + pad(tiyin) + " тийин";
    return text.charAt(0).toUpperCase() + text.slice(1);
  }

  // ---------- act data ----------
  // Billed months follow the sheet: in the sheet's own year a line is due for J («Bu yilgidan necha oyga
  // qarzligi») months from its first month; in other years every month marked with an amount or «qarz».
  // Every amount written in a month cell is a payment, so months paid ahead show up as prepayment.
  function billedMonths(r, sheetYear) {
    const qarz = new Set(String(r.qarz_oylar || "").split(";").filter(Boolean).map((k) => k.replace("-", "_")));
    const keys = Array.from(new Set(Object.keys(r).filter((k) => /^m\d{4}_\d{2}$/.test(k)).map((k) => k.slice(1)).concat(Array.from(qarz)))).sort();
    const marked = (k) => n(r["m" + k]) > 0 || qarz.has(k);
    const out = new Set(keys.filter((k) => k.slice(0, 4) !== sheetYear && marked(k)));
    const j = Math.max(0, Math.round(n(r.months_debt)));
    const own = keys.filter((k) => k.slice(0, 4) === sheetYear);
    let start = own.findIndex(marked);
    if (start < 0 && j) {
      const d = String(r.date || "");
      start = d.slice(0, 4) === sheetYear ? own.indexOf(sheetYear + "_" + d.slice(5, 7)) : 0;
    }
    if (start >= 0) own.slice(start, start + j).forEach((k) => out.add(k));
    return out;
  }

  function buildAct(lines, requisites, company, year, today) {
    const rows = lines.filter((r) => String(r.company) === String(company));
    const all = lines.reduce((ks, r) => ks.concat(Object.keys(r).filter((k) => /^m\d{4}_\d{2}$/.test(k))), []).sort();
    const sheetYear = all.length ? all[0].slice(1, 5) : String(year);
    const ys = String(year);
    const t = today || new Date();
    const endISO = Number(year) === t.getFullYear() ? t.getFullYear() + "-" + pad(t.getMonth() + 1) + "-" + pad(t.getDate()) : ys + "-12-31";
    const to = endISO.split("-").reverse().join(".");
    let opening = 0;
    const docs = [];
    const inv = new Map(), pay = new Map();
    rows.forEach((r) => {
      const billed = billedMonths(r, sheetYear);
      const keys = new Set(Array.from(billed).concat(Object.keys(r).filter((k) => /^m\d{4}_\d{2}$/.test(k) && n(r[k]) > 0).map((k) => k.slice(1))));
      keys.forEach((k) => {
        const b = billed.has(k) ? n(r.monthly) : 0, p = n(r["m" + k]) > 0 ? n(r["m" + k]) : 0;
        if (k.slice(0, 4) < ys) { opening += b - p; return; }
        if (k.slice(0, 4) !== ys) return;
        if (b) { const ik = k + "|" + (r.date || ""); inv.set(ik, (inv.get(ik) || 0) + b); }
        if (p) pay.set(k, (pay.get(k) || 0) + p);
      });
    });
    // a document is dated the month's last day, or the act's end date when that comes first
    const when = (k) => {
      const y = Number(k.slice(0, 4)), m = Number(k.slice(5, 7)), iso = y + "-" + pad(m) + "-" + pad(lastDay(y, m));
      return iso <= endISO ? { iso, text: dmy(y, m, lastDay(y, m)), ahead: false } : { iso: endISO, text: to, ahead: true };
    };
    inv.forEach((v, ik) => {
      const [k, cdate] = ik.split("|"), w = when(k);
      docs.push({ key: w.iso + "|0|" + k + "|" + cdate, date: w.text, text: "Счет–фактура №   от " + w.text, debit: v, credit: 0 });
    });
    pay.forEach((v, k) => {
      const w = when(k), m = Number(k.slice(5, 7));
      docs.push({ key: w.iso + "|1|" + k, date: w.text, text: (w.ahead ? "Предоплата за " : "Оплата за ") + RU_MONTH[m - 1] + " " + k.slice(0, 4) + " г.", debit: 0, credit: v });
    });
    docs.sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : 0));
    const debit = docs.reduce((t, x) => t + x.debit, 0), credit = docs.reduce((t, x) => t + x.credit, 0);
    const req = requisites.filter((r) => sameCo(r.company, company))
      .sort((a, b) => String(b.contract_date || "").localeCompare(String(a.contract_date || "")))[0] || {};
    return {
      company, year: Number(year),
      client: String(req.legal_name || company), director: String(req.director || ""),
      from: "01.01." + year, to,
      opening, docs, debit, credit, closing: opening + debit - credit
    };
  }

  // ---------- one worksheet ----------
  function fillSheet(ws, a) {
    const FONT = { name: "Arial", size: 9 };
    const NUM = "#,##0.00";
    const thin = { style: "thin" };
    const box = { top: thin, left: thin, bottom: thin, right: thin };
    ws.columns = [{ width: 9 }, { width: 31 }, { width: 15 }, { width: 15 }, { width: 9 }, { width: 31 }, { width: 15 }, { width: 15 }];
    ws.pageSetup = { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0,
      margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } };
    let r = 1;
    const line = (text, opts) => {
      ws.mergeCells(r, 1, r, 8);
      const c = ws.getCell(r, 1);
      c.value = text;
      c.font = Object.assign({}, FONT, opts && opts.font);
      c.alignment = { horizontal: (opts && opts.align) || "center", vertical: "middle", wrapText: true };
      if (opts && opts.height) ws.getRow(r).height = opts.height;
      r += 1;
    };
    const half = (left, right, opts) => {
      [[1, left], [5, right]].forEach(([col, text]) => {
        ws.mergeCells(r, col, r, col + 3);
        const c = ws.getCell(r, col);
        c.value = text;
        c.font = Object.assign({}, FONT, opts && opts.font);
        c.alignment = { horizontal: "left", vertical: "top", wrapText: true };
      });
      if (opts && opts.height) ws.getRow(r).height = opts.height;
      r += 1;
    };
    line("Акт сверки", { font: { size: 14, bold: true }, height: 22 });
    line("взаимных расчетов за период: " + a.from + " - " + a.to);
    line("между " + US.name);
    line("и " + a.client);
    r += 1;
    line("Мы, нижеподписавшиеся, Директор " + US.name + " " + US.director + ", с одной стороны, и Директор " + a.client + " " +
      (a.director || "____________________") + ", с другой стороны, составили настоящий акт сверки в том, что состояние взаимных расчетов по данным учета следующее:",
      { align: "left", height: 30, font: { size: 10 } });
    r += 1;

    const t0 = r;
    const put = (col, value, o) => {
      const c = ws.getCell(r, col);
      c.value = value === "" || value === 0 && o && o.blankZero ? null : value;
      c.font = Object.assign({}, FONT, o && o.bold ? { bold: true } : null);
      c.border = box;
      if (value instanceof Date) { c.numFmt = "dd.mm.yy"; c.alignment = { horizontal: "left" }; }
      else if (typeof value === "number") { c.numFmt = NUM; c.alignment = { horizontal: "right" }; }
      else if (o && o.center) c.alignment = { horizontal: "center" };
    };
    const label = (col, text) => {
      ws.mergeCells(r, col, r, col + 1);
      put(col, text, { bold: true }); ws.getCell(r, col + 1).border = box;
    };
    ws.mergeCells(r, 1, r, 4); put(1, "По данным " + US.name + ", сум.");
    [2, 3, 4].forEach((c) => (ws.getCell(r, c).border = box));
    ws.mergeCells(r, 5, r, 8); put(5, "По данным " + a.client + ", сум.");
    [6, 7, 8].forEach((c) => (ws.getCell(r, c).border = box));
    r += 1;
    ["Дата", "Документ", "Дебет", "Кредит", "Дата", "Документ", "Дебет", "Кредит"].forEach((h, i) => put(i + 1, h, { center: true }));
    r += 1;
    const sides = (labelText, d, c, bold) => {
      label(1, labelText); put(3, d, { bold, blankZero: true }); put(4, c, { bold, blankZero: true });
      label(5, labelText); put(7, c, { bold, blankZero: true }); put(8, d, { bold, blankZero: true });
      r += 1;
    };
    sides("Сальдо начальное", a.opening > 0 ? a.opening : 0, a.opening < 0 ? -a.opening : 0, true);
    a.docs.forEach((x) => {
      const [d, m, y] = x.date.split(".").map(Number), day = new Date(Date.UTC(y, m - 1, d));
      put(1, day); put(2, x.text); put(3, x.debit, { blankZero: true }); put(4, x.credit, { blankZero: true });
      put(5, day); put(6, x.text); put(7, x.credit, { blankZero: true }); put(8, x.debit, { blankZero: true });
      r += 1;
    });
    sides("Обороты за период", a.debit, a.credit, true);
    sides("Сальдо конечное", a.closing > 0 ? a.closing : 0, a.closing < 0 ? -a.closing : 0, true);
    ws.views = [{ state: "frozen", ySplit: t0 + 1, showGridLines: false }];
    r += 1;

    half("По данным " + US.name, "По данным " + a.client);
    const amt = Math.abs(a.closing).toLocaleString("ru-RU", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).replace(/\s/g, " ");
    const state = Math.round(a.closing) === 0
      ? "на " + a.to + " задолженность отсутствует."
      : "на " + a.to + " задолженность в пользу " + (a.closing > 0 ? US.name : a.client) + " " + amt + " сум.\n(" + words(a.closing) + ").";
    half(state, state, { height: 42, font: { bold: true } });
    r += 1;
    half("От " + US.name, "От " + a.client);
    half("Директор", "Директор");
    r += 1;
    half("_________________ (" + US.short + ")", "_________________ (" + (a.director || "____________") + ")");
    half("М.П.", "М.П.");
  }

  const safeSheetName = (s, used) => {
    let base = String(s).replace(/[\\/*?:\[\]]/g, " ").replace(/\s+/g, " ").trim().slice(0, 28) || "Mijoz";
    let name = base, i = 2;
    while (used.has(name.toLowerCase())) name = base.slice(0, 26) + " " + i++;
    used.add(name.toLowerCase());
    return name;
  };
  const safeFile = (s) => String(s).replace(/[\\/:*?"<>|«»“”]/g, "").replace(/\s+/g, " ").trim();

  async function workbook(ExcelJS, acts) {
    const wb = new ExcelJS.Workbook();
    wb.creator = "BIS";
    const used = new Set();
    acts.forEach((a) => fillSheet(wb.addWorksheet(acts.length === 1 ? "Акт сверки" : safeSheetName(a.company, used)), a));
    return wb.xlsx.writeBuffer();
  }

  const core = { buildAct, fillSheet, workbook, words, safeFile };
  if (typeof module !== "undefined" && module.exports) { module.exports = core; return; }

  // ---------- page ----------
  let libP = null;
  const load = (src) => new Promise((ok, bad) => {
    const s = document.createElement("script");
    s.src = src;
    s.onload = () => (window.ExcelJS ? ok(window.ExcelJS) : bad(new Error("ExcelJS")));
    s.onerror = () => { s.remove(); bad(new Error("ExcelJS")); };
    document.head.appendChild(s);
  });
  const lib = () => (libP = libP || (window.ExcelJS ? Promise.resolve(window.ExcelJS) : load(EXCELJS[0]).catch(() => load(EXCELJS[1])))
    .catch((e) => { libP = null; throw e; }));
  const years = (lines) => {
    const ys = new Set();
    lines.forEach((r) => Object.keys(r).forEach((k) => { if (/^m\d{4}_\d{2}$/.test(k) && (n(r[k]) > 0 || String(r.qarz_oylar || "").indexOf(k.slice(1, 5) + "-") >= 0)) ys.add(Number(k.slice(1, 5))); }));
    return Array.from(ys).sort();
  };
  const el = (tag, attrs, text) => { const e = document.createElement(tag); Object.entries(attrs || {}).forEach(([k, v]) => e.setAttribute(k, v)); if (text != null) e.textContent = text; return e; };

  function yearSelect(id) {
    const s = el("select", { id, class: "bk-btn", "aria-label": "Akt yili" });
    const fill = () => {
      const d = dash.data("lines_eff");
      const list = d.status === "ok" ? years(d.data) : [];
      const now = new Date().getFullYear();
      if (list.indexOf(now) < 0) list.push(now);
      const keep = s.value || String(now);
      s.textContent = "";
      list.sort().forEach((y) => s.appendChild(el("option", { value: String(y) }, y + "-yil")));
      s.value = list.indexOf(Number(keep)) >= 0 ? keep : String(now);
    };
    fill();
    s.addEventListener("focus", fill);
    return s;
  }

  async function download(companies, year, say) {
    const dl = dash.data("lines_eff"), dr = dash.data("requisites");
    if (dl.status !== "ok") { say("Jadval hali yuklanmoqda, birozdan so’ng qayta bosing."); return; }
    const reqs = dr.status === "ok" ? dr.data : [];
    const acts = companies.map((co) => buildAct(dl.data, reqs, co, year)).filter((a) => companies.length === 1 || a.docs.length || Math.round(a.opening));
    if (!acts.length) { say(year + "-yil uchun hech kimda yozuv yo’q."); return; }
    let downloads = null;
    try { downloads = await window.claude.use("downloads"); } catch (e) { downloads = null; }
    if (!downloads) { say("Bu ko’rinishda yuklab olish ishlamaydi."); return; }
    say("Tayyorlanmoqda…");
    let buf;
    try { buf = await workbook(await lib(), acts); }
    catch (e) { say("Excel kutubxonasi yuklanmadi — internetni tekshirib, qayta bosing."); return; }
    const name = companies.length === 1
      ? "Akt sverka - " + safeFile(acts[0].company) + " - " + year + ".xlsx"
      : "Akt sverka - barcha mijozlar - " + year + ".xlsx";
    try {
      await downloads.save({ filename: name, data: new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }) });
      say("");
    } catch (e) {
      const c = e && e.code;
      say(c === "declined" ? "" : c === "rate_limited" ? "Oldingi yuklab olish oynasi hali ochiq." : "Faylni yuklab bo’lmadi.");
    }
  }

  // client card: one client's act
  const cardBar = document.querySelector("#bk-cmodal .bk-dl");
  if (cardBar) {
    const ysel = yearSelect("bk-akt-year");
    const btn = el("button", { class: "bk-btn", type: "button", id: "bk-akt-one" }, "Akt sverka (Excel)");
    const msg = el("span", { class: "bk-note", role: "status", id: "bk-akt-msg" });
    cardBar.insertBefore(btn, cardBar.firstChild);
    cardBar.insertBefore(ysel, btn);
    document.getElementById("bk-c-sub").after(msg);
    btn.addEventListener("click", () => {
      const co = document.getElementById("bk-c-title").textContent.trim();
      if (co) download([co], Number(ysel.value), (t) => (msg.textContent = t));
    });
  }

  // reminder card: every client, one sheet each
  const exp = document.getElementById("bk-export");
  if (exp) {
    const ysel = yearSelect("bk-akt-all-year");
    const btn = el("button", { class: "bk-btn", type: "button", id: "bk-akt-all" }, "Akt sverka: barcha mijozlar");
    exp.after(ysel, btn);
    btn.addEventListener("click", () => {
      const d = dash.data("lines_eff");
      if (d.status !== "ok") { document.getElementById("bk-exmsg").textContent = "Jadval hali yuklanmoqda."; return; }
      const cos = Array.from(new Set(d.data.map((r) => String(r.company))));
      download(cos, Number(ysel.value), (t) => (document.getElementById("bk-exmsg").textContent = t));
    });
  }
  window.bkAkt = core;
})();

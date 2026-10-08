(() => {
"use strict";

// Presentation and filters ported from the supplied dashboard; credentials stay on the server.
const API = "/mingkh/collective/api";
const owner = document.querySelector(".collective-dashboard").dataset.owner;
const PREFIX = "neurona.collective." + owner + ".";

const THEMES = {
  ii: { name: "Инженерная инфраструктура", color: "var(--ii)" },
  kr: { name: "Капитальный ремонт", color: "var(--kr)" },
  pr: { name: "Прочее", color: "var(--pr)" }
};
const NC = "не учитывается";

/* Правила тем: сначала явная пометка в аннотации, потом ключевые слова. */
// В JS \w и \b не понимают кириллицу — границы слов задаём явно
const L = "а-яёa-z";
const rx = s => new RegExp(s.replace(/\\W/g, `[${L}]*`).replace(/\\B/g, `(?<![${L}])`).replace(/\\E/g, `(?![${L}])`), "i");
const RE_KR = rx(String.raw`капитальн\W\s+ремонт|капремонт|\/кр\E|фонд\W\s+капитальн`);
const RE_II = rx(String.raw`инженерн\W\s+(инфраструктур|систем|сет)|водоснаб|водоотвед|водопровод|канализ|\Bвод[аыуе]\E|питьев|колод|скважин|отоплен|теплоснаб|теплос[еи]т|теплотрасс|котельн|\Bгвс\E|\Bхвс\E|газоснаб|газифик|газопровод|\Bгаз(а|ом|ов\W)?\E|электроснаб|электроэнерг|\Bсвет(а|ом)?\E|очистн\W\s+сооруж|\Bсток|сточн|\/вс\E|водопонижен`);

function classify(r) {
  const ann = r["Аннотация/краткое содержание"] || "";
  if (/"?инженерная инфраструктура"?/i.test(ann)) return "ii";
  if (/"?капитальный ремонт"?/i.test(ann) && /\/кр|коллективное/i.test(ann)) return "kr";
  if (RE_KR.test(ann)) return "kr";
  if (RE_II.test(ann)) return "ii";
  const fact = r["Факт"] || "";
  if (RE_KR.test(fact)) return "kr";
  if (RE_II.test(fact)) return "ii";
  return "pr";
}

const store = {
  get(k) { try { return localStorage.getItem(PREFIX + k); } catch (e) { return null; } },
  set(k, v) { try { v == null ? localStorage.removeItem(PREFIX + k) : localStorage.setItem(PREFIX + k, v); } catch (e) {} }
};

let HEAD = [], ROWS = [], META = "";
const BIG = 100;  // «крупные» жалобы — более 100 подписей
const NO_STATUS = "Без статуса";
// status — набор выбранных статусов; пустой = все
const state = { theme: "all", omsu: new Set(), status: new Set(), q: "", sort: "sig", hideNc: false, big: false, rep: false, omsuSort: "sig" };
const isRep = r => /^да$/i.test(String(r["Повтор"] ?? "").trim());

function num(v) { const n = parseInt(String(v ?? "").replace(/\s/g, ""), 10); return Number.isSafeInteger(n) ? Math.max(0, n) : 0; }
function parseDate(s) { const m = String(s || "").match(/(\d{2})\.(\d{2})\.(\d{4})/); return m ? m[3] + m[2] + m[1] : ""; }
function fmt(n) { return n.toLocaleString("ru-RU"); }
function esc(s) { return String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }
function keyOf(r) {
  const ids = [r["Номер обращения ЕЦУР, МСЭД"], r["Номер обращения исходной системы"]];
  if (ids.some(Boolean)) return "theme." + JSON.stringify(ids);
  // Rows without IDs must not overwrite every other anonymous row's correction.
  const text = [r["ОМСУ"], r["Дата поступления обращения"], r["Аннотация/краткое содержание"], r["Факт"]].join("|");
  let h = 2166136261; for (const c of text) h = Math.imul(h ^ c.charCodeAt(0), 16777619);
  return "theme.fallback." + (h >>> 0).toString(16);
}

function load(head, rows, meta) {
  HEAD = head.map(h => String(h ?? "").trim());
  META = meta || "";
  ROWS = rows.filter(r => r && r.some(v => v != null && String(v).trim() !== "")).map((arr, i) => {
    const r = {};
    HEAD.forEach((h, j) => { r[h] = arr[j] == null ? "" : String(arr[j]); });
    r._id = i;
    r._sig = num(r["Подписи (фактическое число подписей)"]);
    r._date = parseDate(r["Дата поступления обращения"]);
    r._auto = classify(r);
    const manual = store.get(keyOf(r));
    r._theme = Object.hasOwn(THEMES, manual) ? manual : r._auto;
    r._nc = (r["Статус"] || "").toLowerCase().includes(NC);
    r._status = (r["Статус"] || "").trim() || NO_STATUS;
    r._text = Object.values(r).join(" ").toLowerCase();
    return r;
  });
  msBuild("omsu", countBy(r => r["ОМСУ"] || "—").sort((a, b) => a[0].localeCompare(b[0], "ru")));
  msBuild("status", countBy(r => r._status).sort((a, b) => b[1] - a[1]));
  document.getElementById("meta").textContent = META + " · строк: " + ROWS.length;
  render();
}

function baseFiltered() {
  const q = state.q.trim().toLowerCase();
  return ROWS.filter(r =>
    (!state.hideNc || !r._nc) &&
    (!state.big || r._sig > BIG) &&
    (!state.rep || isRep(r)) &&
    (!state.status.size || state.status.has(r._status)) &&
    (!state.omsu.size || state.omsu.has(r["ОМСУ"] || "—")) &&
    (!q || r._text.includes(q)));
}

function render() {
  document.querySelectorAll(".toggle input").forEach(i => i.closest(".toggle").classList.toggle("on", i.checked));
  const base = baseFiltered();
  const rows = state.theme === "all" ? base : base.filter(r => r._theme === state.theme);

  // Вкладки тем
  const cnt = { all: base.length, ii: 0, kr: 0, pr: 0 };
  base.forEach(r => cnt[r._theme]++);
  document.getElementById("tabs").innerHTML = [["all", "Все"], ["ii", THEMES.ii.name], ["kr", THEMES.kr.name], ["pr", THEMES.pr.name]]
    .map(([k, n]) => `<button class="tab" data-t="${k}" aria-pressed="${state.theme === k}">${n}<b>${cnt[k]}</b></button>`).join("");

  // KPI
  const sig = rows.reduce((s, r) => s + r._sig, 0);
  const om = new Set(rows.map(r => r["ОМСУ"])).size;
  const rep = rows.filter(isRep).length;
  const big = rows.filter(r => r._sig > BIG).length;
  const kp = [["Обращений", fmt(rows.length)], ["Подписей", fmt(sig)], ["ОМСУ", fmt(om)],
    ["Подписей на обращение", rows.length ? (sig / rows.length).toFixed(1).replace(".", ",") : "0"],
    ["Более 100 подписей", fmt(big), "big"], ["Повторных", fmt(rep), "rep"]];
  // Карточки с ключом — фильтры: клик включает/выключает
  document.getElementById("kpis").innerHTML = kp.map(([l, v, f]) => f
    ? `<button class="kpi kpi-filter" data-f="${f}" aria-pressed="${state[f]}" title="${state[f] ? "Снять фильтр" : "Показать только их"}"><div class="l">${l}</div><div class="v">${v}</div></button>`
    : `<div class="kpi"><div class="l">${l}</div><div class="v">${v}</div></div>`).join("");

  renderOmsu(state.theme === "all" ? base : rows);
  renderList(rows);
}

function omsuSummary(rows) {
  const m = new Map();
  rows.forEach(r => {
    const k = r["ОМСУ"] || "—";
    if (!m.has(k)) m.set(k, { omsu: k, n: 0, ii: 0, kr: 0, pr: 0, sig: 0, sii: 0, skr: 0, spr: 0 });
    const o = m.get(k); o.n++; o[r._theme]++; o.sig += r._sig; o["s" + r._theme] += r._sig;
  });
  return [...m.values()];
}

function renderOmsu(rows) {
  // Таблица ОМСУ считается без фильтра по ОМСУ, чтобы видеть всех
  const q = state.q.trim().toLowerCase();
  const src = ROWS.filter(r => (!state.hideNc || !r._nc) && (!state.big || r._sig > BIG) && (!state.rep || isRep(r)) &&
    (!state.status.size || state.status.has(r._status)) && (!q || r._text.includes(q)) &&
    (state.theme === "all" || r._theme === state.theme));
  const list = omsuSummary(src);
  const k = state.omsuSort;
  list.sort((a, b) => k === "omsu" ? a.omsu.localeCompare(b.omsu, "ru") : (b[k] - a[k]) || (b.sig - a.sig));
  const tb = document.querySelector("#omsuTable tbody");
  tb.innerHTML = list.length ? list.map(o => `<tr data-o="${esc(o.omsu)}" tabindex="0" role="button" aria-pressed="${state.omsu.has(o.omsu)}" class="${state.omsu.has(o.omsu) ? "sel" : ""}">
      <td>${esc(o.omsu)}</td><td class="n">${o.n}</td>
      <td class="n hide-sm">${o.ii || ""}</td><td class="n hide-sm">${o.kr || ""}</td><td class="n hide-sm">${o.pr || ""}</td>
      <td class="n"><b>${fmt(o.sig)}</b></td></tr>`).join("")
    : `<tr><td colspan="6" class="empty">Нет данных</td></tr>`;
  const t = omsuSummary(src).reduce((a, o) => { ["n", "ii", "kr", "pr", "sig"].forEach(x => a[x] += o[x]); return a; }, { n: 0, ii: 0, kr: 0, pr: 0, sig: 0 });
  document.querySelector("#omsuTable tfoot").innerHTML =
    `<tr><td>Итого</td><td class="n">${t.n}</td><td class="n hide-sm">${t.ii}</td><td class="n hide-sm">${t.kr}</td><td class="n hide-sm">${t.pr}</td><td class="n">${fmt(t.sig)}</td></tr>`;
}

function sorted(rows) {
  const s = state.sort;
  return [...rows].sort((a, b) =>
    s === "date" ? b._date.localeCompare(a._date) || b._sig - a._sig :
    s === "omsu" ? (a["ОМСУ"] || "").localeCompare(b["ОМСУ"] || "", "ru") || b._sig - a._sig :
    b._sig - a._sig || b._date.localeCompare(a._date));
}

function chip(t) { return `<span class="chip ${t}">${THEMES[t].name}</span>`; }

function renderList(rows) {
  const title = (state.theme === "all" ? "Все обращения" : THEMES[state.theme].name) + (state.omsu.size ? " — " + (state.omsu.size > 2 ? state.omsu.size + " ОМСУ" : [...state.omsu].join(", ")) : "");
  document.getElementById("listTitle").textContent = title + " (" + rows.length + ")";
  const el = document.getElementById("list");
  if (!rows.length) { el.innerHTML = '<div class="empty">Ничего не найдено</div>'; return; }
  el.innerHTML = sorted(rows).map(r => `<div class="item" data-id="${r._id}" tabindex="0" role="button" aria-label="Открыть обращение ${esc(r["Номер обращения ЕЦУР, МСЭД"] || r["Номер обращения исходной системы"] || String(r._id + 1))}, ${esc(r["ОМСУ"] || "ОМСУ не указан")}">
      <div class="top"><span>${esc(r["Дата поступления обращения"])}</span><span class="omsu">${esc(r["ОМСУ"])}</span>
        ${chip(r._theme)}${r._nc ? '<span class="chip warn">Не учитывается</span>' : ""}
        ${isRep(r) ? '<span class="chip warn">Повтор</span>' : ""}
        ${r._sig > BIG ? '<span class="chip big">100+</span>' : ""}
        <span>${esc(r["Источник"])}${r["Источник обращения ЕЦУР"] ? " · " + esc(r["Источник обращения ЕЦУР"]) : ""}</span>
        <span class="sig">${r._sig} подп.</span></div>
      <div class="txt">${esc(r["Аннотация/краткое содержание"] || r["Факт"])}</div></div>`).join("");
}

/* Карточка */
let current = null;
function openCard(id) {
  const r = ROWS[id]; current = r;
  document.getElementById("dlgTitle").textContent = (r["ОМСУ"] || "—") + " · " + r._sig + " подписей";
  document.getElementById("dlgSub").textContent = (r["Дата поступления обращения"] || "") + " · " + (r["Номер обращения ЕЦУР, МСЭД"] || "");
  const sel = document.getElementById("dlgTheme");
  sel.innerHTML = Object.entries(THEMES).map(([k, t]) =>
    `<option value="${k}" ${k === r._theme ? "selected" : ""}>${t.name}${k === r._auto ? " (авто)" : ""}</option>`).join("");
  document.getElementById("dlgBody").innerHTML =
    `<dt>Тема</dt><dd>${chip(r._theme)}</dd>` +
    HEAD.filter(h => h).map(h => `<dt>${esc(h)}</dt><dd>${esc(r[h]) || "—"}</dd>`).join("");
  document.getElementById("dlg").showModal();
}

/* Excel */
function exportRows(rows, name) {
  const head = ["Тема", ...HEAD.filter(h => h)];
  const data = [head, ...sorted(rows).map(r => [THEMES[r._theme].name, ...HEAD.filter(h => h).map(h =>
    h.startsWith("Подписи") ? r._sig : r[h])])];
  return { name, data };
}
function save(sheets, file) {
  if (!window.XLSX) {   // нет интернета — отдаём CSV
    const csv = sheets[0].data.map(row => row.map(v => '"' + String(typeof v === 'string' && /^[=+@\-]/.test(v) ? "'" + v : v ?? "").replace(/"/g, '""') + '"').join(";")).join("\r\n");
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob(["﻿" + csv], { type: "text/csv" }));
    a.download = file.replace(/\.xlsx$/, ".csv"); a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 5000); return;
  }
  const wb = XLSX.utils.book_new();
  sheets.forEach(s => {
    const ws = XLSX.utils.aoa_to_sheet(s.data);
    ws["!cols"] = s.data[0].map((h, i) => ({ wch: Math.min(60, Math.max(10, ...s.data.slice(0, 50).map(r => String(r[i] ?? "").length))) }));
    XLSX.utils.book_append_sheet(wb, ws, s.name);
  });
  XLSX.writeFile(wb, file);
}
function exportAll() {
  const base = baseFiltered();
  const rows = state.theme === "all" ? base : base.filter(r => r._theme === state.theme);
  const om = omsuSummary(rows).sort((a, b) => b.sig - a.sig);
  const omSheet = { name: "По ОМСУ", data: [["ОМСУ", "Обращений", "Инженерная инфраструктура", "Капитальный ремонт", "Прочее", "Подписей",
    "Подписи: ИИ", "Подписи: КР", "Подписи: прочее"], ...om.map(o => [o.omsu, o.n, o.ii, o.kr, o.pr, o.sig, o.sii, o.skr, o.spr])] };
  const sheets = [exportRows(rows, "Обращения"), omSheet];
  if (state.theme === "all") ["ii", "kr", "pr"].forEach(t => sheets.push(exportRows(rows.filter(r => r._theme === t), THEMES[t].name.slice(0, 31))));
  const d = new Date();
  save(sheets, `Коллективки_дашборд_${d.toLocaleDateString("ru-RU")}.xlsx`);
}

/* Выпадающие списки с множественным выбором (ОМСУ, статусы) */
function countBy(fn) {
  const m = new Map();
  ROWS.forEach(r => { const k = fn(r); m.set(k, (m.get(k) || 0) + 1); });
  return [...m];
}
function msRoot(key) { return document.querySelector(`.ms[data-key="${key}"]`); }
function msBuild(key, items) {
  const root = msRoot(key), set = state[key];
  [...set].forEach(v => { if (!items.some(([k]) => k === v)) set.delete(v); });
  const wasOpen = root.querySelector(".ms-panel") && !root.querySelector(".ms-panel").hidden;
  root.innerHTML = `<button type="button" class="ms-btn" aria-haspopup="true" aria-expanded="false"><span class="ms-lbl"></span><span class="ms-caret">▾</span></button>
    <div class="ms-panel" hidden>
      ${items.length > 8 ? '<input type="search" class="ms-search" placeholder="Найти…">' : ""}
      <div class="ms-actions"><button type="button" data-act="all">Выбрать все</button><button type="button" data-act="none">Очистить</button></div>
      <div class="ms-list">${items.map(([v, n]) =>
        `<label><input type="checkbox" value="${esc(v)}"><span>${esc(v)}</span><span class="cnt">${n}</span></label>`).join("")}</div>
    </div>`;
  msSync(key);
  if (wasOpen) msOpen(root, true);
}
function msSync(key) {
  const root = msRoot(key), set = state[key], all = root.dataset.all;
  root.querySelectorAll(".ms-list input").forEach(i => { i.checked = set.has(i.value); });
  const v = [...set];
  root.querySelector(".ms-lbl").textContent = !v.length ? all : v.length <= 2 ? v.join(", ") : `Выбрано ${v.length} ${root.dataset.many}`;
  root.querySelector(".ms-btn").classList.toggle("active", v.length > 0);
  root.querySelector(".ms-btn").title = v.join("\n");
}
function msOpen(root, open) {
  const p = root.querySelector(".ms-panel"); if (!p) return;
  p.hidden = !open;
  root.querySelector(".ms-btn").setAttribute("aria-expanded", open);
  if (open) { const s = p.querySelector(".ms-search"); if (s) s.focus(); }
}

/* События */
document.getElementById("tabs").addEventListener("click", e => { const b = e.target.closest(".tab"); if (b) { state.theme = b.dataset.t; render(); } });
document.addEventListener("click", e => {
  const ms = e.target.closest(".ms");
  document.querySelectorAll(".ms").forEach(m => { if (m !== ms) msOpen(m, false); });
  if (!ms) return;
  const key = ms.dataset.key, set = state[key];
  if (e.target.closest(".ms-btn")) { msOpen(ms, ms.querySelector(".ms-panel").hidden); return; }
  const act = e.target.closest("[data-act]");
  if (act) {
    if (act.dataset.act === "none") set.clear();
    else ms.querySelectorAll(".ms-list label:not([hidden]) input").forEach(i => set.add(i.value));  // все видимые (с учётом поиска)
    msSync(key); render();
  }
});
document.addEventListener("change", e => {
  const ms = e.target.closest(".ms"); if (!ms || e.target.type !== "checkbox") return;
  const set = state[ms.dataset.key];
  e.target.checked ? set.add(e.target.value) : set.delete(e.target.value);
  msSync(ms.dataset.key); render();
});
document.addEventListener("input", e => {
  if (!e.target.classList.contains("ms-search")) return;
  const q = e.target.value.trim().toLowerCase();
  e.target.closest(".ms").querySelectorAll(".ms-list label").forEach(l => { l.hidden = q && !l.textContent.toLowerCase().includes(q); });
});
document.addEventListener("keydown", e => { if (e.key === "Escape") document.querySelectorAll(".ms").forEach(m => msOpen(m, false)); });
document.getElementById("resetFilters").addEventListener("click", () => {
  Object.assign(state, { theme: "all", q: "", sort: "sig", hideNc: false, big: false, rep: false });
  state.omsu.clear(); state.status.clear(); msSync("omsu"); msSync("status");
  document.getElementById("q").value = "";
  document.getElementById("sortSel").value = "sig";
  document.getElementById("hideNc").checked = false;
  document.getElementById("bigOnly").checked = false;
  document.getElementById("repOnly").checked = false;
  render();
});
document.getElementById("q").addEventListener("input", e => { state.q = e.target.value; render(); });
document.getElementById("sortSel").addEventListener("change", e => { state.sort = e.target.value; render(); });
document.getElementById("repOnly").addEventListener("change", e => { state.rep = e.target.checked; render(); });
document.getElementById("hideNc").addEventListener("change", e => { state.hideNc = e.target.checked; render(); });
document.getElementById("bigOnly").addEventListener("change", e => { state.big = e.target.checked; render(); });
document.getElementById("kpis").addEventListener("click", e => {
  const b = e.target.closest(".kpi-filter"); if (!b) return;
  state[b.dataset.f] = !state[b.dataset.f];
  document.getElementById("bigOnly").checked = state.big;
  document.getElementById("repOnly").checked = state.rep;
  render();
});
document.querySelector("#omsuTable thead").addEventListener("click", e => { const th = e.target.closest("th"); if (th && th.dataset.k) { state.omsuSort = th.dataset.k; render(); } });
document.querySelector("#omsuTable tbody").addEventListener("click", e => {
  const tr = e.target.closest("tr[data-o]"); if (!tr) return;
  const o = tr.dataset.o;
  state.omsu.has(o) ? state.omsu.delete(o) : state.omsu.add(o);
  msSync("omsu"); render();
});
document.querySelector("#omsuTable tbody").addEventListener("keydown", e => {
  const row = e.target.closest("tr[data-o]");
  if (row && ["Enter", " "].includes(e.key)) { e.preventDefault(); row.click(); }
});
const listEl = document.getElementById("list");
listEl.addEventListener("click", e => { const it = e.target.closest(".item"); if (it) openCard(+it.dataset.id); });
listEl.addEventListener("keydown", e => { const it = e.target.closest(".item"); if (it && ["Enter", " "].includes(e.key)) { e.preventDefault(); openCard(+it.dataset.id); } });
document.getElementById("dlgClose").addEventListener("click", () => document.getElementById("dlg").close());
document.getElementById("dlgTheme").addEventListener("change", e => {
  if (!current || !Object.hasOwn(THEMES, e.target.value)) return;
  current._theme = e.target.value;
  store.set(keyOf(current), current._theme === current._auto ? null : current._theme);
  openCard(current._id); render();
});
document.getElementById("dlgExport").addEventListener("click", () => {
  const r = current;
  save([{ name: "Карточка", data: [["Поле", "Значение"], ["Тема", THEMES[r._theme].name], ...HEAD.filter(h => h).map(h => [h, r[h]])] }],
    `Карточка_${(r["ОМСУ"] || "").replace(/[\\/:*?"<>|]/g, "")}_${r._id + 1}.xlsx`);
});
document.getElementById("exportAll").addEventListener("click", exportAll);

/* The API uses the shared administrator integration and application session. */
const iso = d => d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
function periodRange() {
  const p = document.getElementById("periodSel").value, t = new Date(), s = new Date(t);
  if (p === "custom") return [document.getElementById("dStart").value, document.getElementById("dEnd").value];
  if (p === "d7") s.setDate(t.getDate() - 6);
  if (p === "d30") s.setDate(t.getDate() - 29);
  if (p === "month") s.setDate(1);
  return [iso(s), iso(t)];
}
let loadedPeriod = null;

/* Презентация по образцу — из того, что сейчас на экране */
async function exportPptx() {
  const btn = document.getElementById("exportPptx");
  if (!loadedPeriod) { alert("Сначала загрузи данные кнопкой «Обновить»."); return; }
  const base = baseFiltered();
  const rows = sorted(state.theme === "all" ? base : base.filter(r => r._theme === state.theme)).map(r => {
    const o = { _theme: r._theme, _sig: r._sig };
    HEAD.forEach(h => { if (h) o[h] = r[h]; });
    return o;
  });
  if (!rows.length) { alert("На экране нет обращений — нечего класть в презентацию."); return; }
  btn.disabled = true; btn.textContent = "Собираю…";
  try {
    const resp = await fetch(API + "/pptx", { method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ start: loadedPeriod[0], end: loadedPeriod[1], rows }) });
    if (!resp.ok) { const d = await resp.json().catch(() => ({})); throw new Error(d.error || resp.statusText); }
    const cd = resp.headers.get("Content-Disposition") || "";
    const m = cd.match(/filename\*=UTF-8''([^;]+)/);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(await resp.blob());
    a.download = m ? decodeURIComponent(m[1]) : "Коллективные обращения.pptx";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  } catch (e) {
    alert("Не удалось собрать презентацию: " + (e.message || e));
  } finally { btn.disabled = false; btn.textContent = "Презентация"; }
}
document.getElementById("exportPptx").addEventListener("click", exportPptx);

let requestNumber = 0;
let activeRequest = null;
function setExportState() {
  const disabled = !loadedPeriod || !ROWS.length;
  document.getElementById("exportAll").disabled = disabled;
  document.getElementById("exportPptx").disabled = disabled;
}
async function refresh() {
  const [start, end] = periodRange();
  const err = document.getElementById("err"), btn = document.getElementById("refresh");
  if (!start || !end || start > end || end > iso(new Date())) {
    err.textContent = "Укажите корректный период без будущих дат."; err.hidden = false; return;
  }
  const request = ++requestNumber;
  if (activeRequest) activeRequest.abort();
  activeRequest = new AbortController();
  err.hidden = true; btn.disabled = true; btn.textContent = "Загружаю…";
  document.getElementById("list").setAttribute("aria-busy", "true");
  if (!ROWS.length) document.getElementById("list").innerHTML = '<div class="loading">Забираем данные с портала…</div>';
  try {
    const resp = await fetch(`${API}/data?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`, { cache: "no-store", signal: activeRequest.signal });
    const data = await resp.json().catch(() => ({}));
    if (request !== requestNumber) return;
    if (!resp.ok) throw new Error(data.error || data.detail || "Не удалось получить данные.");
    if (!Array.isArray(data.head) || !Array.isArray(data.rows)) throw new Error("Сервер вернул неполные данные.");
    load(data.head, data.rows, data.meta + " · обновлено " + data.updated);
    loadedPeriod = [data.start || start, data.end || end];
    store.set("period", document.getElementById("periodSel").value);
  } catch (e) {
    if (e.name === "AbortError" || request !== requestNumber) return;
    const retained = loadedPeriod ? ` Показаны ранее загруженные данные за ${loadedPeriod[0]} — ${loadedPeriod[1]}.` : "";
    err.textContent = "Не удалось обновить: " + (e.message || e) + retained; err.hidden = false;
    if (!loadedPeriod) document.getElementById("list").innerHTML = '<div class="empty">Данные пока не загружены</div>';
  } finally {
    if (request === requestNumber) {
      btn.disabled = false; btn.textContent = "⟳ Обновить";
      document.getElementById("list").setAttribute("aria-busy", "false"); setExportState();
    }
  }
}
const ps = document.getElementById("periodSel"), saved = store.get("period");
if (["today", "d7", "month", "d30"].includes(saved)) ps.value = saved;
const today = new Date(), week = new Date(today); week.setDate(today.getDate() - 6);
document.getElementById("dStart").value = iso(week); document.getElementById("dEnd").value = iso(today);
document.getElementById("dStart").max = iso(today); document.getElementById("dEnd").max = iso(today);
ps.addEventListener("change", () => {
  document.getElementById("customDates").hidden = ps.value !== "custom";
  if (ps.value !== "custom") refresh();
});
document.getElementById("refresh").addEventListener("click", refresh);
load([], [], "Выберите период и загрузите обращения");
setExportState();
refresh();

})();

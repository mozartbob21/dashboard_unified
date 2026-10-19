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

let HEAD = [], ROWS = [], META = "", ASSIGNEES = [], SINCE = "";
const DRAFTS = new Map();
let canManage = false;
const BIG = 100;  // «крупные» жалобы — более 100 подписей
const NO_STATUS = "Без статуса";
// status — набор выбранных статусов; пустой = все
const state = { theme: "all", omsu: new Set(), status: new Set(), q: "", sort: "sig", hideNc: false, big: false, rep: false, omsuSort: "sig", assignee: "", unassigned: false };
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

function load(head, rows, meta, extra = {}) {
  if(current && !document.getElementById("dlg").open)current=null;
  if (Array.isArray(extra.assignees)) ASSIGNEES = extra.assignees;
  SINCE = extra.since || SINCE;
  canManage = document.querySelector(".collective-dashboard").dataset.canManage === "true";
  document.getElementById("manageAssignees").hidden = !canManage;
  HEAD = head.map(h => String(h ?? "").trim());
  META = meta || "";
  ROWS = rows.map((arr,sourceIndex)=>({arr,sourceIndex})).filter(({arr})=>Array.isArray(arr) && arr.some(v=>v!=null && String(v).trim()!=="")).map(({arr,sourceIndex}, i) => {
    const r = {};
    HEAD.forEach((h, j) => { r[h] = arr[j] == null ? "" : String(arr[j]); });
    r._id = i;
    r._sig = num(r["Подписи (фактическое число подписей)"]);
    r._date = parseDate(r["Дата поступления обращения"]);
    r._auto = classify(r);
    r._key = String(extra.keys?.[sourceIndex] || "");
    r._workflow = normalizeWorkflow(extra.workflows?.[r._key]);
    r._theme = Object.hasOwn(THEMES, r._workflow.theme) ? r._workflow.theme : r._auto;
    r._nc = (r["Статус"] || "").toLowerCase().includes(NC);
    r._status = (r["Статус"] || "").trim() || NO_STATUS;
    r._text = HEAD.map(h=>r[h] || "").concat(r._workflow.assignee_name,r._workflow.report).join(" ").toLowerCase();
    return r;
  });
  msBuild("omsu", countBy(r => r["ОМСУ"] || "—").sort((a, b) => a[0].localeCompare(b[0], "ru")));
  msBuild("status", countBy(r => r._status).sort((a, b) => b[1] - a[1]));
  document.getElementById("meta").textContent = META + " · строк: " + ROWS.length;
  syncAssigneeFilter();
  showLegacyThemes();
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
    assignmentMatches(r) &&
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
  renderAssigneeSummary();
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
    (!state.status.size || state.status.has(r._status)) && assignmentMatches(r) && (!q || r._text.includes(q)) &&
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
      <div class="txt">${esc(r["Аннотация/краткое содержание"] || r["Факт"])}</div>
      <div class="assignment-line"><span class="assignee-tag ${!r._workflow.assignee_id && r._workflow.can_assign ? "unassigned" : ""}">${esc(r._workflow.assignee_name || (r._workflow.can_assign ? "Не распределено" : "Ответственный не назначен"))}${r._workflow.assignee_archived ? " · архив" : ""}</span>${r._workflow.report ? '<span class="report-mark">Доклад сохранён</span>' : ""}</div></div>`).join("");
}

/* Server workflow: only filters remain device-local. */
function normalizeWorkflow(w = {}) {
  return {version:0,assignee_id:null,assignee_name:"",assignee_archived:false,report:"",theme:null,can_assign:false,...w};
}
function assignmentMatches(r) {
  return (!state.assignee || String(r._workflow.assignee_id || "")===state.assignee) && (!state.unassigned || (r._workflow.can_assign && !r._workflow.assignee_id));
}
function assigneeSummary(rows) {
  const groups=new Map();
  for(const r of rows) {
    const w=r._workflow;
    if(!w.assignee_id && !w.can_assign)continue;
    const key=w.assignee_id?String(w.assignee_id):"__unassigned";
    if(!groups.has(key))groups.set(key,{key,name:w.assignee_name || "Не распределены",archived:w.assignee_archived,count:0,sig:0,reports:0});
    const item=groups.get(key);item.count++;item.sig+=r._sig;if(w.report.trim())item.reports++;
  }
  return [...groups.values()].sort((a,b)=>b.count-a.count || b.sig-a.sig || a.name.localeCompare(b.name,"ru"));
}
function renderAssigneeSummary() {
  const previous=state.assignee, unassigned=state.unassigned;
  state.assignee="";state.unassigned=false;
  const rows=baseFiltered().filter(r=>state.theme==="all" || r._theme===state.theme);
  state.assignee=previous;state.unassigned=unassigned;
  const list=assigneeSummary(rows);
  document.querySelector("#assigneeTable tbody").innerHTML=list.length?list.map(o=>`<tr data-assignee="${esc(o.key)}" tabindex="0" role="button" aria-pressed="${o.key==="__unassigned"?state.unassigned:state.assignee===o.key}" class="${(o.key==="__unassigned"?state.unassigned:state.assignee===o.key)?"sel":""}"><td>${esc(o.name)}${o.archived?'<span class="field-note"> · архив</span>':""}</td><td class="n">${fmt(o.count)}</td><td class="n">${fmt(o.sig)}</td><td class="n">${fmt(o.reports)}</td></tr>`).join(""):'<tr><td colspan="4" class="empty">Нет обращений для распределения</td></tr>';
  document.getElementById("assignmentNote").textContent="В распределении — открытые обращения, уже назначенные и новые"+(SINCE?" с "+SINCE.split("-").reverse().join("."):"")+". Таблица учитывает выбранный период и остальные фильтры.";
}
function syncAssigneeFilter() {
  const sel=document.getElementById("assigneeFilter");
  sel.innerHTML='<option value="">Все ответственные</option>'+ASSIGNEES.map(a=>`<option value="${esc(a.id)}">${esc(a.name)}${a.archived?" · архив":""}</option>`).join("");
  if(!ASSIGNEES.some(a=>String(a.id)===state.assignee))state.assignee="";
  sel.value=state.assignee;
}
async function requestJSON(path, options={}) {
  const controller=new AbortController();
  const abort=()=>controller.abort();
  if(options.signal?.aborted)controller.abort();else options.signal?.addEventListener("abort",abort,{once:true});
  const timer=setTimeout(()=>controller.abort(),45000);
  let response,data;
  try {
    response=await fetch(API+path,{credentials:"same-origin",cache:"no-store",...options,signal:controller.signal});
    data=await response.json().catch(()=>({}));
  }catch(error){throw new Error(error.name==="AbortError"?"Сервер отвечает дольше обычного. Повторите действие: введённый текст сохранён в карточке.":"Связь с сервером прервалась. Проверьте подключение и повторите действие.");}
  finally{clearTimeout(timer);options.signal?.removeEventListener("abort",abort);}
  if(!response.ok){const error=new Error(data.error || (typeof data.detail==="string"?data.detail:"") || (response.status===401?"Войдите в платформу повторно.":"Не удалось сохранить изменения."));error.status=response.status;error.workflow=data.workflow;throw error;}
  return data;
}
function status(id,text,kind="") {const el=document.getElementById(id);el.textContent=text;el.className="save-state"+(kind?" "+kind:"");}
let current=null, fieldBusy=false, reportBusy=false;
function draftFor(r) {
  if(!DRAFTS.has(r._key))DRAFTS.set(r._key,{text:r._workflow.report,baseReport:r._workflow.report,dirty:false,conflict:false});
  const d=DRAFTS.get(r._key);
  if(!d.dirty){d.text=r._workflow.report;d.baseReport=r._workflow.report;d.conflict=false;}
  else if(d.baseReport!==r._workflow.report){d.conflict=d.text!==r._workflow.report;if(!d.conflict){d.dirty=false;d.baseReport=r._workflow.report;}}
  return d;
}
function applyWorkflow(r,w) {
  const before=r._workflow, updated=normalizeWorkflow(w);
  const matching=new Set([r,...ROWS.filter(row=>row._key && row._key===r._key)]);
  for(const row of matching){
    row._workflow=updated;
    row._theme=Object.hasOwn(THEMES,updated.theme)?updated.theme:row._auto;
    row._text=HEAD.map(h=>row[h] || "").concat(updated.assignee_name,updated.report).join(" ").toLowerCase();
  }
  const d=DRAFTS.get(r._key);
  if(d && !d.dirty){d.text=r._workflow.report;d.baseReport=r._workflow.report;d.conflict=false;}
  else if(d && before.report!==r._workflow.report && d.baseReport!==r._workflow.report){d.conflict=true;}
  r._text=HEAD.map(h=>r[h] || "").concat(r._workflow.assignee_name,r._workflow.report).join(" ").toLowerCase();
  render();
}
function updateCardFields() {
  if(!current)return;const w=current._workflow;
  const choices=ASSIGNEES.filter(a=>!a.archived || a.id===w.assignee_id);
  document.getElementById("dlgAssignee").innerHTML='<option value="">Не назначен</option>'+choices.map(a=>`<option value="${esc(a.id)}"${a.archived?' disabled':""}>${esc(a.name)}${a.archived?" · архив":""}</option>`).join("");
  document.getElementById("dlgAssignee").value=w.assignee_id?String(w.assignee_id):"";
  document.getElementById("dlgAssignee").disabled=fieldBusy || reportBusy || !w.can_assign || !current._key;
  document.getElementById("dlgTheme").innerHTML=`<option value="">Автоматически — ${THEMES[current._auto].name}</option>`+Object.entries(THEMES).map(([k,t])=>`<option value="${k}">${t.name}</option>`).join("");
  document.getElementById("dlgTheme").value=w.theme || "";
  document.getElementById("dlgTheme").disabled=fieldBusy || reportBusy || !current._key;
  document.getElementById("dlgReport").disabled=!current._key;
  document.getElementById("closeWithDraft").disabled=fieldBusy || reportBusy;
  const savedDate=w.report_at?new Date(w.report_at):null;
  document.getElementById("dlgReportSavedAt").textContent=savedDate && !Number.isNaN(savedDate.getTime()) ? "Последнее сохранение доклада: "+savedDate.toLocaleString("ru-RU",{timeZone:"Europe/Moscow",day:"2-digit",month:"2-digit",year:"numeric",hour:"2-digit",minute:"2-digit"})+" МСК" : "";
  document.getElementById("dlgAssignNote").textContent=!current._key?"Сначала обновите данные обращения.":w.can_assign?"Ответственный и тема сохраняются сразу. Доклад сохраняется отдельной кнопкой.":"Это обращение не входит в распределение по статусу и дате. Доклад и тему можно обновить.";
  updateDraftState();
}
function updateDraftState() {
  if(!current)return;const d=draftFor(current);
  document.getElementById("dlgReportCount").textContent=fmt(d.text.length)+" / 20 000";
  document.getElementById("dlgReportDraft").hidden=!d.dirty;
  document.getElementById("saveReport").disabled=reportBusy || fieldBusy || !current._key || !d.dirty || d.conflict;
  document.getElementById("keepOwnReport").disabled=reportBusy || fieldBusy;
  document.getElementById("useServerReport").disabled=reportBusy;
  document.getElementById("reportConflict").hidden=!d.conflict;
  document.getElementById("conflictServerReport").textContent=current._workflow.report || "Доклад пуст.";
}
function openCard(id) {
  if(current)captureDraft();
  const r=ROWS[id];if(!r)return;current=r;
  const d=draftFor(r);
  document.getElementById("dlgTitle").textContent=(r["ОМСУ"] || "—")+" · "+r._sig+" подписей";
  document.getElementById("dlgSub").textContent=(r["Дата поступления обращения"] || "")+" · "+(r["Номер обращения ЕЦУР, МСЭД"] || "");
  document.getElementById("dlgBody").innerHTML=HEAD.filter(h=>h).map(h=>`<dt>${esc(h)}</dt><dd>${esc(r[h]) || "—"}</dd>`).join("");
  document.getElementById("dlgReport").value=d.text;
  document.getElementById("closeDraftWarning").hidden=true;
  status("dlgFieldStatus","");status("dlgReportStatus","");updateCardFields();
  const dialog=document.getElementById("dlg");if(!dialog.open)dialog.showModal();
}
function captureDraft() {
  if(!current)return;const d=draftFor(current);
  d.text=document.getElementById("dlgReport").value;d.dirty=d.text!==current._workflow.report;
  if(!d.dirty){d.baseReport=current._workflow.report;d.conflict=false;}
  updateDraftState();
}
function requestCloseCard() {
  captureDraft();if(current && (draftFor(current).dirty || reportBusy || fieldBusy)){
    document.getElementById("closeDraftWarning").hidden=false;return;
  }document.getElementById("dlg").close();
}
async function patchWorkflow(r,changes) {
  const data=await requestJSON("/appeals/"+encodeURIComponent(r._key),{method:"PATCH",headers:{"Content-Type":"application/json"},body:JSON.stringify({version:r._workflow.version,...changes})});
  if(!data.workflow)throw new Error("Сервер не подтвердил сохранение. Обновите данные и проверьте результат.");
  applyWorkflow(r,data.workflow);return data.workflow;
}
async function changeField(field,value) {
  if(!current || fieldBusy || reportBusy)return;captureDraft();const r=current;
  fieldBusy=true;updateCardFields();status("dlgFieldStatus","Сохраняем…");
  try{await patchWorkflow(r,{[field]:field==="assignee_id" && value?Number(value):value});status("dlgFieldStatus","Сохранено на сервере", "success");}
  catch(error){if(error.workflow)applyWorkflow(r,error.workflow);status("dlgFieldStatus",error.status===409?"Коллега уже изменил карточку. Показана актуальная версия. Повторите выбор при необходимости.":error.message,"error");}
  finally{fieldBusy=false;if(current===r)updateCardFields();}
}
async function saveReport(overwrite=false) {
  if(!current || reportBusy || fieldBusy)return;captureDraft();const r=current,d=draftFor(r);
  if(d.conflict && !overwrite)return;
  if(d.text.length>20000){status("dlgReportStatus","Сократите доклад до 20 000 символов.","error");return;}
  if(!d.dirty)return;const sent=d.text;
  reportBusy=true;updateCardFields();status("dlgReportStatus","Сохраняем доклад…");
  try{
    await patchWorkflow(r,{report:sent});
    if(d.text===sent){d.text=r._workflow.report;if(current===r)document.getElementById("dlgReport").value=d.text;}
    d.baseReport=r._workflow.report;d.dirty=d.text!==r._workflow.report;d.conflict=false;
    status("dlgReportStatus",d.dirty?"Предыдущая версия сохранена. Новые изменения в поле ещё не отправлены.":"Доклад сохранён на сервере","success");
    document.getElementById("closeDraftWarning").hidden=true;
  }catch(error){
    if(error.workflow){applyWorkflow(r,error.workflow);if(d.text===sent && sent.trim()===r._workflow.report){d.text=r._workflow.report;if(current===r)document.getElementById("dlgReport").value=d.text;}d.dirty=d.text!==r._workflow.report;d.conflict=d.dirty;if(!d.dirty)d.baseReport=r._workflow.report;}
    status("dlgReportStatus",error.status===409 && !d.dirty?"Сервер подтвердил: этот доклад уже сохранён.":error.status===409?"Карточка была изменена другим пользователем. Сравните доклады ниже.":error.message+" Доклад остался в поле; его можно отправить повторно.",error.status===409 && !d.dirty?"success":"error");
  }finally{reportBusy=false;if(current)updateCardFields();}
}
function useServerReport(){if(!current || reportBusy)return;const d=draftFor(current);d.text=current._workflow.report;d.baseReport=d.text;d.dirty=false;d.conflict=false;document.getElementById("dlgReport").value=d.text;status("dlgReportStatus","Загружен доклад с сервера.");updateDraftState();}
function renderDirectory(){
  for(const [id,archived]of[["activeAssignees",false],["archivedAssignees",true]]){
    const list=ASSIGNEES.filter(a=>!!a.archived===archived);
    document.getElementById(id).innerHTML=list.length?list.map(a=>`<div class="assignee-row"><span class="person-name">${esc(a.name)}</span><button class="btn" type="button" data-person-id="${esc(a.id)}" data-active="${archived?"true":"false"}">${archived?"Вернуть":"В архив"}</button></div>`).join(""):'<p class="field-note">'+(archived?"Архив пуст.":"Ответственные ещё не добавлены.")+"</p>";
  }
  document.getElementById("archivedCount").textContent="("+ASSIGNEES.filter(a=>a.archived).length+")";
  syncAssigneeFilter();
}
async function reloadDirectory(){const data=await requestJSON("/assignees");if(!Array.isArray(data.assignees))throw new Error("Не удалось получить справочник ответственных.");ASSIGNEES=data.assignees;renderDirectory();if(current)updateCardFields();}
async function openAssignees(){if(!canManage)return;renderDirectory();document.getElementById("assigneeDialog").showModal();status("assigneeStatus","Обновляем справочник…");try{await reloadDirectory();status("assigneeStatus","");}catch(error){status("assigneeStatus",error.message,"error");}}
async function addAssignee(event){
  event.preventDefault();if(!canManage)return;const input=document.getElementById("newAssigneeName"),name=input.value.trim(),btn=document.getElementById("addAssignee");
  if(!name || name.length>120){status("assigneeStatus","Введите имя длиной от 1 до 120 символов.","error");return;}
  if(btn.disabled)return;btn.disabled=true;status("assigneeStatus","Добавляем…");
  try{const data=await requestJSON("/assignees",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({name})});if(data.assignee){ASSIGNEES.push(data.assignee);renderDirectory();}else await reloadDirectory();input.value="";status("assigneeStatus","Ответственный добавлен.","success");}
  catch(error){status("assigneeStatus",error.message,"error");}finally{btn.disabled=false;}
}
async function changeAssigneeActive(event){
  const button=event.target.closest("[data-person-id]");if(!button || button.disabled || !canManage)return;
  const person=ASSIGNEES.find(a=>String(a.id)===button.dataset.personId);if(!person)return;
  button.disabled=true;status("assigneeStatus","Сохраняем…");
  try{const data=await requestJSON("/assignees/"+person.id,{method:"PATCH",headers:{"Content-Type":"application/json"},body:JSON.stringify({version:person.version,active:button.dataset.active==="true"})});if(data.assignee){ASSIGNEES=ASSIGNEES.map(a=>a.id===data.assignee.id?data.assignee:a);renderDirectory();}else await reloadDirectory();
    for(const r of ROWS){const a=ASSIGNEES.find(a=>a.id===r._workflow.assignee_id);if(a)r._workflow.assignee_archived=a.archived;}
    render();status("assigneeStatus","Справочник обновлён. Прежние назначения сохранены.","success");
  }catch(error){status("assigneeStatus",error.message,"error");if(error.status===409){try{await reloadDirectory();}catch(_){}}}finally{button.disabled=false;}
}

function legacyThemeCandidates(){return ROWS.filter(r=>r._key && !r._workflow.theme && Object.hasOwn(THEMES,store.get(keyOf(r))) && store.get(keyOf(r))!==r._auto);}
function showLegacyThemes(){const count=legacyThemeCandidates().length;document.getElementById("legacyThemes").hidden=!count;document.getElementById("legacyThemeNote").textContent=`На этом устройстве есть ручные темы для ${count} обращений выбранного периода. Можно перенести их на сервер: после этого они будут видны коллегам. Существующие серверные темы останутся без изменений.`;}
async function migrateLegacyThemes(){
  const btn=document.getElementById("migrateThemes");if(btn.disabled)return;btn.disabled=true;const candidates=legacyThemeCandidates();let saved=0,failed=0;
  for(const r of candidates){try{await patchWorkflow(r,{theme:store.get(keyOf(r))});store.set(keyOf(r),null);saved++;}catch(error){if(error.workflow)applyWorkflow(r,error.workflow);failed++;}status("migrationStatus",`Перенесено: ${saved}. Не удалось перенести: ${failed}.`);}
  btn.disabled=false;showLegacyThemes();if(failed)status("migrationStatus","Часть тем не перенесена: карточки изменились или прервалась связь. Уже сохранённые темы доступны коллегам.","error");
}
document.getElementById("migrateThemes").addEventListener("click",migrateLegacyThemes);

const importForm=document.getElementById("importDistributionForm");
if(importForm){
  document.getElementById("distributionFile").addEventListener("change",event=>{
    const file=event.target.files?.[0];document.getElementById("distributionFileName").textContent=file?.name || "Файл не выбран";
    document.getElementById("importDistribution").disabled=!file;status("distributionImportStatus","");
  });
  importForm.addEventListener("submit",async event=>{
    event.preventDefault();if(!canManage)return;
    const file=document.getElementById("distributionFile").files?.[0],btn=document.getElementById("importDistribution");
    if(!file || btn.disabled)return;
    if(!/\.json$/i.test(file.name) || file.size>20*1024*1024){status("distributionImportStatus","Выберите файл JSON размером до 20 МБ.","error");return;}
    btn.disabled=true;status("distributionImportStatus","Переносим сохранённое распределение…");
    try{
      const result=await requestJSON("/import-distribution",{method:"POST",headers:{"Content-Type":"application/json"},body:file});
      if(Array.isArray(result.assignees)){ASSIGNEES=result.assignees;syncAssigneeFilter();}
      const [start,end]=loadedPeriod || periodRange();
      try{const snapshot=await requestJSON("/data?start="+encodeURIComponent(start)+"&end="+encodeURIComponent(end));acceptSnapshot(snapshot,start,end);setExportState();}
      catch(error){if(error.status!==404){status("distributionImportStatus","Данные перенесены, но список не удалось обновить. Обновите его позднее.","error");return;}}
      status("distributionImportStatus","Данные распределения перенесены. Существующие правки сохранены.","success");
    }catch(error){status("distributionImportStatus",error.message,"error");}
    finally{btn.disabled=false;}
  });
}

/* Excel */
function exportRows(rows, name) {
  const head = ["Тема", "Ответственный", "Доклад", ...HEAD.filter(h => h)];
  const data = [head, ...sorted(rows).map(r => [THEMES[r._theme].name, r._workflow.assignee_name || "", r._workflow.report || "", ...HEAD.filter(h => h).map(h =>
    h.startsWith("Подписи") ? r._sig : r[h])])];
  return { name, data };
}
function save(sheets, file) {
  if (!window.XLSX) {   // нет интернета — отдаём CSV
    const csv = sheets[0].data.map(row => row.map(v => '"' + String(typeof v === 'string' && /^(?:[\t\r\n]|\s*[=+@\-])/.test(v) ? "'" + v : v ?? "").replace(/"/g, '""') + '"').join(";")).join("\r\n");
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
  const sheets = [exportRows(rows, "Обращения"), omSheet, {name:"По ответственным",data:[["Ответственный","Обращений","Подписей","С докладом"],...assigneeSummary(rows).map(o=>[o.name,o.count,o.sig,o.reports])]}];
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
  Object.assign(state, { theme: "all", q: "", sort: "sig", hideNc: false, big: false, rep: false, assignee: "", unassigned: false });
  state.omsu.clear(); state.status.clear(); msSync("omsu"); msSync("status");
  document.getElementById("q").value = "";
  document.getElementById("sortSel").value = "sig";
  document.getElementById("hideNc").checked = false;
  document.getElementById("bigOnly").checked = false;
  document.getElementById("repOnly").checked = false;
  document.getElementById("unassignedOnly").checked = false;
  document.getElementById("assigneeFilter").value = "";
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
document.getElementById("dlgClose").addEventListener("click", requestCloseCard);
document.getElementById("dlg").addEventListener("cancel", e => {e.preventDefault();requestCloseCard();});
document.getElementById("dlgTheme").addEventListener("change", e => changeField("theme", e.target.value || null));
document.getElementById("dlgAssignee").addEventListener("change", e => changeField("assignee_id", e.target.value || null));
document.getElementById("dlgReport").addEventListener("input", captureDraft);
document.getElementById("saveReport").addEventListener("click", () => saveReport(false));
document.getElementById("keepOwnReport").addEventListener("click", () => saveReport(true));
document.getElementById("useServerReport").addEventListener("click", useServerReport);
document.getElementById("closeWithDraft").addEventListener("click", () => {if(!fieldBusy&&!reportBusy)document.getElementById("dlg").close();});
window.addEventListener("beforeunload", e => { if ([...DRAFTS.values()].some(d=>d.dirty)) {e.preventDefault(); e.returnValue="";} });
document.getElementById("dlgExport").addEventListener("click", () => {
  const r=current; if(!r)return;
  save([{ name: "Карточка", data: [["Поле", "Значение"], ["Тема", THEMES[r._theme].name], ["Ответственный",r._workflow.assignee_name || ""], ["Доклад",r._workflow.report || ""], ...HEAD.filter(h => h).map(h => [h, r[h]])] }],
    `Карточка_${(r["ОМСУ"] || "").replace(/[\\/:*?"<>|]/g, "")}_${r._id + 1}.xlsx`);
});
document.getElementById("assigneeFilter").addEventListener("change",e=>{state.assignee=e.target.value;state.unassigned=false;document.getElementById("unassignedOnly").checked=false;render();});
document.getElementById("unassignedOnly").addEventListener("change",e=>{state.unassigned=e.target.checked;if(state.unassigned){state.assignee="";document.getElementById("assigneeFilter").value="";}render();});
document.querySelector("#assigneeTable tbody").addEventListener("click",e=>{const row=e.target.closest("tr[data-assignee]");if(!row)return;const value=row.dataset.assignee;if(value==="__unassigned"){state.unassigned=!state.unassigned;state.assignee="";}else{state.assignee=state.assignee===value?"":value;state.unassigned=false;}document.getElementById("assigneeFilter").value=state.assignee;document.getElementById("unassignedOnly").checked=state.unassigned;render();});
document.querySelector("#assigneeTable tbody").addEventListener("keydown",e=>{const row=e.target.closest("tr[data-assignee]");if(row&&["Enter"," "].includes(e.key)){e.preventDefault();row.click();}});
document.getElementById("manageAssignees").addEventListener("click",openAssignees);
document.getElementById("closeAssignees").addEventListener("click",()=>document.getElementById("assigneeDialog").close());
document.getElementById("addAssigneeForm").addEventListener("submit",addAssignee);
for(const id of ["activeAssignees","archivedAssignees"])document.getElementById(id).addEventListener("click",changeAssigneeActive);
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
    const o = { _theme: r._theme, _sig: r._sig, _key:r._key, _assignee:r._workflow.assignee_name || "", _report:r._workflow.report || "" };
    HEAD.forEach(h => { if (h) o[h] = r[h]; });
    return o;
  });
  if (!rows.length) { alert("На экране нет обращений — нечего класть в презентацию."); return; }
  btn.disabled = true; btn.textContent = "Собираю…";
  try {
    const resp = await fetch(API + "/pptx", { method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ start: loadedPeriod[0], end: loadedPeriod[1], rows, keys: rows.map(r=>r._key) }) });
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

let requestNumber=0,activeRequest=null;
const jobTickets=new Map();
function readTicket(key){try{return JSON.parse(sessionStorage.getItem(PREFIX+"refresh."+key) || "null");}catch(_){return jobTickets.get(key) || null;}}
function writeTicket(key,value){if(value)jobTickets.set(key,value);else jobTickets.delete(key);try{value?sessionStorage.setItem(PREFIX+"refresh."+key,JSON.stringify(value)):sessionStorage.removeItem(PREFIX+"refresh."+key);}catch(_){}}
const pause=ms=>new Promise(resolve=>setTimeout(resolve,ms));
function setExportState() {
  const disabled=!loadedPeriod || !ROWS.length;
  document.getElementById("exportAll").disabled=disabled;
  document.getElementById("exportPptx").disabled=disabled;
}
function acceptSnapshot(data,start,end){
  if(!Array.isArray(data.head)||!Array.isArray(data.rows))throw new Error("Сервер вернул неполные данные.");
  load(data.head,data.rows,(data.meta || "")+" · обновлено "+(data.updated || data.fetched_at || "—"),data);
  loadedPeriod=[data.start || start,data.end || end];
  store.set("period",document.getElementById("periodSel").value);
}
async function refresh(force=true) {
  const [start,end]=periodRange(),err=document.getElementById("err"),btn=document.getElementById("refresh");
  if(!start||!end||start>end||end>iso(new Date())){err.textContent="Укажите корректный период без будущих дат.";err.hidden=false;return;}
  const request=++requestNumber;if(activeRequest)activeRequest.abort();activeRequest=new AbortController();
  const signal=activeRequest.signal;err.hidden=true;btn.disabled=true;btn.textContent="Загружаем…";
  document.getElementById("list").setAttribute("aria-busy","true");
  if(!ROWS.length)document.getElementById("list").innerHTML='<div class="loading">Загружаем сохранённые обращения…</div>';
  const query="?start="+encodeURIComponent(start)+"&end="+encodeURIComponent(end),ticketKey=start+"_"+end;
  try{
    if(!force){
      try{const data=await requestJSON("/data"+query,{signal});if(request!==requestNumber)return;acceptSnapshot(data,start,end);if(!readTicket(ticketKey))return;}
      catch(error){if(error.status!==404)throw error;}
    }
    let ticket=readTicket(ticketKey);
    if(!ticket){ticket={idempotency_key:globalThis.crypto?.randomUUID?.() || Date.now()+"-"+Math.random().toString(36).slice(2)};writeTicket(ticketKey,ticket);}
    if(!ticket.status_url){
      let data;
      for(let attempt=0;attempt<3;attempt++){
        try{data=await requestJSON("/refresh",{method:"POST",headers:{"Content-Type":"application/json","Idempotency-Key":ticket.idempotency_key},body:JSON.stringify({start,end}),signal});break;}
        catch(error){if(error.status || signal.aborted || attempt===2)throw error;await pause(1200*(attempt+1));}
      }
      if(!data || typeof data.status_url!=="string" || !data.status_url.startsWith(API+"/jobs/") || data.status_url.includes("..") || data.status_url.includes("?") || data.status_url.includes("#"))throw new Error("Сервер не подтвердил запуск обновления.");
      ticket={...ticket,job_id:data.job_id,status_url:data.status_url};writeTicket(ticketKey,ticket);
    }
    let failures=0;
    while(true){
      if(request!==requestNumber || signal.aborted)return;
      let job;
      try{job=await requestJSON(ticket.status_url.slice(API.length),{signal});failures=0;}
      catch(error){if([401,403,404].includes(error.status)){writeTicket(ticketKey,null);throw error;}if(signal.aborted || ++failures>3)throw error;btn.textContent="Восстанавливаем связь…";await pause(1500*failures);continue;}
      if(job.status==="success"){writeTicket(ticketKey,null);break;}
      if(job.status==="error"){writeTicket(ticketKey,null);throw new Error(job.error || "Источник не завершил обновление. Повторите позже.");}
      if(!["queued","running"].includes(job.status))throw new Error("Сервер вернул неизвестное состояние обновления.");
      btn.textContent=job.status==="queued"?"Ожидаем очередь…":"Получаем обращения…";
      await pause(1800);
    }
    const data=await requestJSON("/data"+query,{signal});if(request!==requestNumber)return;acceptSnapshot(data,start,end);
  }catch(error){
    if(signal.aborted || request!==requestNumber)return;
    const retained=loadedPeriod?` Показаны ранее загруженные данные за ${loadedPeriod[0]} — ${loadedPeriod[1]}.`:"";
    err.textContent="Не удалось обновить: "+error.message+retained;err.hidden=false;
    if(!loadedPeriod)document.getElementById("list").innerHTML='<div class="empty">Данные пока не загружены. Повторите обновление.</div>';
  }finally{if(request===requestNumber){activeRequest=null;btn.disabled=false;btn.textContent="⟳ Обновить";document.getElementById("list").setAttribute("aria-busy","false");setExportState();}}
}
let snapshotPolling=false;
function canSyncSnapshot(){return !!loadedPeriod && !activeRequest && !snapshotPolling && document.visibilityState!=="hidden" && !document.getElementById("dlg").open && !document.getElementById("assigneeDialog").open && ![...DRAFTS.values()].some(d=>d.dirty);}
async function syncSavedSnapshot(){
  if(!canSyncSnapshot())return;const period=loadedPeriod.slice();snapshotPolling=true;
  try{
    const data=await requestJSON("/data?start="+encodeURIComponent(period[0])+"&end="+encodeURIComponent(period[1]));
    snapshotPolling=false;
    if(!canSyncSnapshot() || period.join()!==loadedPeriod.join())return;
    acceptSnapshot(data,...period);setExportState();
    document.getElementById("workflowSyncStatus").textContent="Изменения коллег проверены в "+new Date().toLocaleTimeString("ru-RU",{hour:"2-digit",minute:"2-digit"})+". Синхронизация — раз в минуту.";
  }catch(_){document.getElementById("workflowSyncStatus").textContent="Не удалось проверить изменения коллег. Повторим автоматически; текущие данные сохранены.";}
  finally{snapshotPolling=false;}
}
setInterval(syncSavedSnapshot,60000);
document.addEventListener("visibilitychange",()=>{if(document.visibilityState!=="hidden")syncSavedSnapshot();});
const ps=document.getElementById("periodSel"),saved=store.get("period");
if(["today","d7","month","d30"].includes(saved))ps.value=saved;
const today=new Date(),week=new Date(today);week.setDate(today.getDate()-6);
document.getElementById("dStart").value=iso(week);document.getElementById("dEnd").value=iso(today);
document.getElementById("dStart").max=iso(today);document.getElementById("dEnd").max=iso(today);
ps.addEventListener("change",()=>{document.getElementById("customDates").hidden=ps.value!=="custom";if(ps.value!=="custom")refresh(false);});
document.getElementById("refresh").addEventListener("click",()=>refresh(true));
load([],[],"Выберите период и загрузите обращения");
setExportState();refresh(false);

})();

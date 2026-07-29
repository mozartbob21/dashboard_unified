const COLS = {
    id: 0, addr: 1, dist: 2, cat: 3, sub: 4, fact: 5,
    exec: 6, cur: 7, synth: 8, text: 9, created: 10, deadline: 11, status: 12
};

const STATE = {
    rows: window.PRELOADED_ROWS || [],
    meta: window.PRELOADED_META || null,
    filters: { status: "", category: "", curator: "", synth: "", day: null, bucket: null },
    calMonth: null,
    view: "curator",
};

function parseRuDate(s) {
    if (!s) return null;
    const p = String(s).trim().split(" ")[0].split(".");
    if (p.length !== 3) return null;
    const [d, m, y] = p.map(Number);
    if (!d || !m || !y) return null;
    return new Date(y, m - 1, d);
}
function fmtRuDate(dt) {
    if (!dt) return "";
    return String(dt.getDate()).padStart(2,"0") + "." + String(dt.getMonth()+1).padStart(2,"0") + "." + dt.getFullYear();
}
function startOfDay(dt) { const d = new Date(dt); d.setHours(0,0,0,0); return d; }
function endOfWeek(dt) { const d = startOfDay(dt); const day = d.getDay() || 7; d.setDate(d.getDate() + (7 - day)); return d; }
function endOfMonth(dt) { return new Date(dt.getFullYear(), dt.getMonth()+1, 0); }
function bucketFor(deadline) {
    if (!deadline) return null;
    const today = startOfDay(new Date());
    const dl = startOfDay(deadline);
    if (dl < today) return "overdue";
    if (dl.getTime() === today.getTime()) return "today";
    if (dl <= endOfWeek(today)) return "week";
    if (dl <= endOfMonth(today)) return "month";
    return null;
}
function toast(msg) {
    const el = document.getElementById("toast");
    if (!el) return;
    el.textContent = msg;
    el.classList.add("show");
    setTimeout(() => el.classList.remove("show"), 2400);
}
function uniqueSorted(arr) {
    return Array.from(new Set(arr.filter(Boolean))).sort((a,b) => a.localeCompare(b, "ru"));
}
function dataRows() { return STATE.rows.length > 1 ? STATE.rows.slice(1) : []; }

function applyFilters(rows) {
    const f = STATE.filters;
    return rows.filter(r => {
        if (f.status   && r[COLS.status] !== f.status) return false;
        if (f.category && r[COLS.cat]    !== f.category) return false;
        if (f.curator  && r[COLS.cur]    !== f.curator) return false;
        if (f.synth    && r[COLS.synth]  !== f.synth) return false;
        if (f.day) {
            const dl = parseRuDate(r[COLS.deadline]);
            if (!dl || fmtRuDate(dl) !== f.day) return false;
        }
        if (f.bucket) {
            const dl = parseRuDate(r[COLS.deadline]);
            if (bucketFor(dl) !== f.bucket) return false;
        }
        return true;
    });
}

function fillFilterSelects() {
    const rows = dataRows();
    const fill = (id, values) => {
        const sel = document.getElementById(id);
        if (!sel) return;
        const cur = sel.value;
        sel.innerHTML = '<option value="">все</option>' +
            values.map(v => `<option value="${v.replace(/"/g,"&quot;")}">${v}</option>`).join("");
        if (values.includes(cur)) sel.value = cur;
    };
    fill("statusSel", uniqueSorted(rows.map(r => r[COLS.status])));
    fill("catSel",    uniqueSorted(rows.map(r => r[COLS.cat])));
    fill("curSel",    uniqueSorted(rows.map(r => r[COLS.cur])));
    fill("synthSel",  uniqueSorted(rows.map(r => r[COLS.synth])));
}

function updateTopKPI() {
    const filtered = applyFilters(dataRows());
    let o=0,t=0,w=0,mo=0;
    filtered.forEach(r => {
        const b = bucketFor(parseRuDate(r[COLS.deadline]));
        if (b==="overdue") o++; else if (b==="today") t++;
        else if (b==="week") w++; else if (b==="month") mo++;
    });
    document.getElementById("kTotal").textContent   = filtered.length;
    document.getElementById("kOverdue").textContent = o;
    document.getElementById("kToday").textContent   = t;
    document.getElementById("kWeek").textContent    = w;
    document.getElementById("kMonth").textContent   = mo;
    document.getElementById("ecurFilterCount").textContent = "оказано: " + filtered.length;
}

function renderCategoryKPI() {
    const cont = document.getElementById("kpiCats");
    if (!cont) return;
    const filtered = applyFilters(dataRows());
    const byCat = new Map();
    filtered.forEach(r => {
        const c = r[COLS.cat] || "—";
        if (!byCat.has(c)) byCat.set(c, {total:0, overdue:0, today:0});
        const rec = byCat.get(c);
        rec.total++;
        const b = bucketFor(parseRuDate(r[COLS.deadline]));
        if (b==="overdue") rec.overdue++;
        if (b==="today") rec.today++;
    });
    const arr = Array.from(byCat.entries()).sort((a,b) => b[1].total - a[1].total);
    cont.innerHTML = arr.map(([name,r]) => `
        <div class="ecur-kpi-cat">
            <div class="ecur-kpi-cat-name">${name}</div>
            <div class="ecur-kpi-cat-stats">
                <span class="ecur-kpi-cat-stat pill pill-blue">${r.total}</span>
                ${r.overdue ? `<span class="ecur-kpi-cat-stat pill pill-critical">просроч. ${r.overdue}</span>` : ""}
                ${r.today   ? `<span class="ecur-kpi-cat-stat pill pill-risk">сег. ${r.today}</span>` : ""}
            </div>
        </div>
    `).join("");
}

function initCalendar() {
    if (!STATE.calMonth) {
        const now = new Date();
        STATE.calMonth = { y: now.getFullYear(), m: now.getMonth() };
    }
}
function renderCalendar() {
    initCalendar();
    const grid = document.getElementById("calGrid");
    const label = document.getElementById("calLabel");
    if (!grid || !label) return;
    const { y, m } = STATE.calMonth;
    label.textContent = new Date(y,m,1).toLocaleString("ru", { month:"long", year:"numeric" });

    const filtered = applyFilters(dataRows());
    const byDay = new Map();
    filtered.forEach(r => {
        const dl = parseRuDate(r[COLS.deadline]);
        if (!dl) return;
        if (dl.getFullYear() !== y || dl.getMonth() !== m) return;
        const key = fmtRuDate(dl);
        byDay.set(key, (byDay.get(key) || 0) + 1);
    });

    const startOffset = (new Date(y,m,1).getDay() || 7) - 1;
    const daysInMonth = new Date(y,m+1,0).getDate();
    const todayStr = fmtRuDate(startOfDay(new Date()));

    let html = "";
    for (let i=0; i<startOffset; i++) html += `<div class="ecur-cal-cell empty"></div>`;
    for (let d=1; d<=daysInMonth; d++) {
        const dt = new Date(y,m,d);
        const key = fmtRuDate(dt);
        const cnt = byDay.get(key) || 0;
        const b = bucketFor(dt);
        const cls = ["ecur-cal-cell"];
        if (!cnt) cls.push("void");
        if (key === todayStr) cls.push("today");
        if (STATE.filters.day === key) cls.push("selected");
        html += `<div class="${cls.join(" ")}" data-day="${key}">
            <span class="d">${d}</span>
            ${cnt ? `<span class="cnt ${b || ""}">${cnt}</span>` : ""}
        </div>`;
    }
    grid.innerHTML = html;
    grid.querySelectorAll(".ecur-cal-cell:not(.void):not(.empty)").forEach(cell => {
        cell.addEventListener("click", () => {
            const day = cell.dataset.day;
            STATE.filters.day = (STATE.filters.day === day) ? null : day;
            STATE.filters.bucket = null;
            rerenderAll();
        });
    });
}

function renderFactList() {
    const cont = document.getElementById("factList");
    if (!cont) return;
    const filtered = applyFilters(dataRows());
    const key = STATE.view === "curator" ? COLS.cur : COLS.synth;
    document.getElementById("grpTitle").textContent =
        STATE.view === "curator" ? "о куратору" : "о синтетической группе";

    const groups = new Map();
    filtered.forEach(r => {
        const g = r[key] || "— не указан —";
        if (!groups.has(g)) groups.set(g, []);
        groups.get(g).push(r);
    });
    const arr = Array.from(groups.entries()).sort((a,b) => b[1].length - a[1].length);
    if (!arr.length) {
        cont.innerHTML = `<div class="utnkr-no-filter-results">ет данных под текущие фильтры.</div>`;
        return;
    }
    cont.innerHTML = arr.map(([name, items]) => {
        let o=0, t=0;
        items.forEach(r => {
            const b = bucketFor(parseRuDate(r[COLS.deadline]));
            if (b==="overdue") o++;
            if (b==="today") t++;
        });
        return `<div class="ecur-group">
            <div class="ecur-group-head">
                <div class="ecur-group-name">${name}</div>
                <div class="ecur-group-stats">
                    <span class="pill pill-blue">${items.length}</span>
                    ${o ? `<span class="pill pill-critical">просроч. ${o}</span>` : ""}
                    ${t ? `<span class="pill pill-risk">сег. ${t}</span>` : ""}
                    <span class="ecur-group-arrow">▾</span>
                </div>
            </div>
            <div class="ecur-group-body">
                ${items.slice(0,200).map(r => {
                    const b = bucketFor(parseRuDate(r[COLS.deadline]));
                    const dlCls = b==="overdue" ? "overdue" : (b==="today" ? "today" : "ok");
                    return `<div class="ecur-item" data-id="${r[COLS.id]}">
                        <span class="id">#${r[COLS.id]}</span>
                        <span class="addr">${(r[COLS.addr]||"—").slice(0,120)}</span>
                        <span class="dl ${dlCls}">${r[COLS.deadline]||"—"}</span>
                    </div>`;
                }).join("")}
                ${items.length > 200 ? `<div class="subtext" style="margin-top:8px">оказаны первые 200 из ${items.length}</div>` : ""}
            </div>
        </div>`;
    }).join("");

    cont.querySelectorAll(".ecur-group-head").forEach(h =>
        h.addEventListener("click", () => h.parentElement.classList.toggle("on")));
    cont.querySelectorAll(".ecur-item").forEach(it =>
        it.addEventListener("click", () => openComplaintModal(it.dataset.id)));
}

function renderChips() {
    const cont = document.getElementById("activeChips");
    if (!cont) return;
    const f = STATE.filters;
    const chips = [];
    const push = (label, key, val) =>
        chips.push(`<span class="ecur-chip">${label}: ${val}<span class="x" data-key="${key}">✕</span></span>`);
    if (f.status)   push("Статус",    "status",   f.status);
    if (f.category) push("атегория", "category", f.category);
    if (f.curator)  push("уратор",   "curator",  f.curator);
    if (f.synth)    push("Синт.гр.",  "synth",    f.synth);
    if (f.day)      push("ень",      "day",      f.day);
    if (f.bucket) {
        const map = { overdue:"росрочено", today:"Сегодня", week:"та неделя", month:"тот месяц" };
        push("Срок", "bucket", map[f.bucket]);
    }
    if (chips.length) chips.push(`<span class="ecur-chip reset" id="chipsResetAll">Сбросить все ✕</span>`);
    cont.innerHTML = chips.join("");
    cont.querySelectorAll(".ecur-chip .x").forEach(x => {
        x.addEventListener("click", () => {
            const key = x.dataset.key;
            STATE.filters[key] = (key==="day" || key==="bucket") ? null : "";
            syncSelects();
            rerenderAll();
        });
    });
    document.getElementById("chipsResetAll")?.addEventListener("click", resetFilters);
}

function syncSelects() {
    const map = { statusSel:"status", catSel:"category", curSel:"curator", synthSel:"synth" };
    Object.entries(map).forEach(([id, key]) => {
        const sel = document.getElementById(id);
        if (sel) sel.value = STATE.filters[key] || "";
    });
}
function resetFilters() {
    STATE.filters = { status:"", category:"", curator:"", synth:"", day:null, bucket:null };
    syncSelects();
    rerenderAll();
}

function openComplaintModal(id) {
    const row = dataRows().find(r => String(r[COLS.id]) === String(id));
    if (!row) return;
    document.getElementById("mTitle").textContent = "#" + row[COLS.id] + " · " + (row[COLS.cat] || "—");
    document.getElementById("mSub").textContent = row[COLS.addr] || "—";
    const b = bucketFor(parseRuDate(row[COLS.deadline]));
    const pill = b==="overdue" ? `<span class="pill pill-critical">росрочено</span>`
        : b==="today" ? `<span class="pill pill-risk">Срок сегодня</span>`
        : b==="week" ? `<span class="pill pill-blue">о конца недели</span>`
        : b==="month" ? `<span class="pill pill-ok">о конца месяца</span>`
        : `<span class="pill pill-outline">ез спец. срока</span>`;

    document.getElementById("mBody").innerHTML = `
        <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px;">
            ${pill}<span class="pill pill-outline">${row[COLS.status]||"—"}</span>
        </div>
        <table>
            <tr><th style="width:180px">дрес</th><td>${row[COLS.addr]||"—"}</td></tr>
            <tr><th>айон</th><td>${row[COLS.dist]||"—"}</td></tr>
            <tr><th>атегория </th><td>${row[COLS.cat]||"—"}</td></tr>
            <tr><th>одкатегория</th><td>${row[COLS.sub]||"—"}</td></tr>
            <tr><th> факт</th><td>${row[COLS.fact]||"—"}</td></tr>
            <tr><th>сполнитель</th><td>${row[COLS.exec]||"—"}</td></tr>
            <tr><th>уратор</th><td>${row[COLS.cur]||"—"}</td></tr>
            <tr><th>Синт. группа</th><td>${row[COLS.synth]||"—"}</td></tr>
            <tr><th>ата создания</th><td>${row[COLS.created]||"—"}</td></tr>
            <tr><th>Срок</th><td>${row[COLS.deadline]||"—"}</td></tr>
            <tr><th>Статус</th><td>${row[COLS.status]||"—"}</td></tr>
        </table>
        <div style="margin-top:14px;">
            <div class="subtext" style="margin-bottom:6px;font-weight:800;">Текст обращения</div>
            <pre>${(row[COLS.text]||"—").replace(/</g,"&lt;")}</pre>
        </div>`;
    document.getElementById("complaintModal").classList.remove("hidden");
    document.body.classList.add("modal-open");
}
function closeComplaintModal() {
    document.getElementById("complaintModal").classList.add("hidden");
    document.body.classList.remove("modal-open");
}

function exportExcel() {
    const filtered = applyFilters(dataRows());
    if (!filtered.length) { toast("ет данных для экспорта"); return; }
    const header = STATE.rows[0];
    const ws = XLSX.utils.aoa_to_sheet([header, ...filtered]);
    ws["!cols"] = header.map((h,i) => {
        const maxLen = Math.max(String(h).length, ...filtered.slice(0,200).map(r => String(r[i]||"").length));
        return { wch: Math.min(60, Math.max(10, maxLen+2)) };
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "алобы ");
    const now = new Date();
    const stamp = now.getFullYear() + "-" + String(now.getMonth()+1).padStart(2,"0") + "-" + String(now.getDate()).padStart(2,"0");
    XLSX.writeFile(wb, "ecur_complaints_" + stamp + ".xlsx");
    toast("Excel сформирован");
}

async function refreshData() {
    const btn = document.getElementById("btnRefresh");
    btn.disabled = true;
    btn.textContent = "бновление...";
    document.getElementById("loaderBlock").classList.remove("hidden");
    const badge = document.getElementById("systemStateBadge");
    badge.className = "system-badge running";
    badge.textContent = "ыполняется выгрузка";
    try {
        const r = await fetch("/ecur/refresh", { method:"POST" });
        const data = await r.json();
        if (!data.ok) {
            toast(data.message || "е удалось запустить обновление");
            btn.disabled = false; btn.textContent = "бновить свод";
            return;
        }
        pollStatus();
    } catch (e) {
        toast("шибка сети");
        btn.disabled = false; btn.textContent = "бновить свод";
    }
}
async function pollStatus() {
    try {
        const r = await fetch("/ecur/run-status");
        const s = await r.json();
        document.getElementById("statusText").textContent = s.stage || "—";
        document.getElementById("currentStageMessage").textContent = s.message || "";
        if (s.running) {
            setTimeout(pollStatus, 1500);
        } else {
            document.getElementById("loaderBlock").classList.add("hidden");
            const badge = document.getElementById("systemStateBadge");
            badge.className = "system-badge ready";
            badge.textContent = "отова";
            const btn = document.getElementById("btnRefresh");
            btn.disabled = false; btn.textContent = "бновить свод";
            if (s.last_error) toast("шибка: " + s.last_error);
            else { toast("Свод обновлён — перезагружаю…"); setTimeout(() => location.reload(), 800); }
        }
    } catch (e) {
        setTimeout(pollStatus, 2500);
    }
}

function rerenderAll() {
    updateTopKPI();
    renderCategoryKPI();
    renderCalendar();
    renderFactList();
    renderChips();
}

document.addEventListener("DOMContentLoaded", () => {
    document.getElementById("btnRefresh")?.addEventListener("click", refreshData);
    document.getElementById("btnExport")?.addEventListener("click", exportExcel);

    document.getElementById("mClose")?.addEventListener("click", closeComplaintModal);
    document.getElementById("mCloseFooter")?.addEventListener("click", closeComplaintModal);
    document.getElementById("complaintModal")?.addEventListener("click", (e) => {
        if (e.target.id === "complaintModal") closeComplaintModal();
    });
    document.addEventListener("keydown", (e) => {
        if (e.key === "Escape") closeComplaintModal();
    });

    document.getElementById("calPrev")?.addEventListener("click", () => {
        initCalendar();
        STATE.calMonth.m--;
        if (STATE.calMonth.m < 0) { STATE.calMonth.m = 11; STATE.calMonth.y--; }
        renderCalendar();
    });
    document.getElementById("calNext")?.addEventListener("click", () => {
        initCalendar();
        STATE.calMonth.m++;
        if (STATE.calMonth.m > 11) { STATE.calMonth.m = 0; STATE.calMonth.y++; }
        renderCalendar();
    });

    document.getElementById("grpSeg")?.querySelectorAll("button").forEach(btn => {
        btn.addEventListener("click", () => {
            document.querySelectorAll("#grpSeg button").forEach(b => b.classList.remove("on"));
            btn.classList.add("on");
            STATE.view = btn.dataset.view;
            renderFactList();
        });
    });

    ["statusSel","catSel","curSel","synthSel"].forEach(id => {
        const sel = document.getElementById(id);
        if (!sel) return;
        sel.addEventListener("change", () => {
            const map = { statusSel:"status", catSel:"category", curSel:"curator", synthSel:"synth" };
            STATE.filters[map[id]] = sel.value;
            rerenderAll();
        });
    });

    document.getElementById("resetEcurFiltersBtn")?.addEventListener("click", resetFilters);

    if (STATE.rows && STATE.rows.length > 1) {
        fillFilterSelects();
        rerenderAll();
    }
});
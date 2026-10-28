// Reference: original archive JavaScript, calculations only; no EXE, network or secrets.
const PERIODS = ['1 квартал', '1 полугодие', '9 месяцев', '11 месяцев'];
// для заголовка «Динамика к …»: периоды нарастающим итогом, в скобках — какой это квартал
const PERIODS_TO = ['1 кв.', '1 пол. (2 кв.)', '9 мес. (3 кв.)', '11 мес. (4 кв.)'];

// Места — по «Методике расчёта … Повышение эффективности РСО. Эффективность ФХД» на 2026 год
// (сверено с официальной таблицей «Рейтинг 3 кв» — по 4-му показателю совпадает целиком).

// ЗАТО по методике рейтингуются отдельно, с 1 по 5 место
const ZATO = new Set(['ВЛАСИХА', 'ВОСХОД', 'ЗВЕЗДНЫЙ ГОРОДОК', 'КРАСНОЗНАМЕНСК', 'МОЛОДЕЖНЫЙ']);
// Все ОМСУ МО (51 + 5 ЗАТО) — как в официальном своде; кто не участвует в критерии, показывается без места
const ALL_OMSU = ['БАЛАШИХА', 'БОГОРОДСКИЙ', 'БРОННИЦЫ', 'ВОЛОКОЛАМСКИЙ', 'ВОСКРЕСЕНСК', 'ДМИТРОВСКИЙ', 'ДОЛГОПРУДНЫЙ',
  'ДОМОДЕДОВО', 'ДУБНА', 'ЕГОРЬЕВСК', 'ЖУКОВСКИЙ', 'ЗАРАЙСК', 'ИСТРА', 'КАШИРА', 'КЛИН', 'КОЛОМНА', 'КОРОЛЕВ', 'КОТЕЛЬНИКИ',
  'КРАСНОГОРСК', 'ЛЕНИНСКИЙ', 'ЛОБНЯ', 'ЛОСИНО-ПЕТРОВСКИЙ', 'ЛОТОШИНО', 'ЛУХОВИЦЫ', 'ЛЫТКАРИНО', 'ЛЮБЕРЦЫ', 'МОЖАЙСКИЙ',
  'МЫТИЩИ', 'НАРО-ФОМИНСКИЙ', 'ОДИНЦОВСКИЙ', 'ОРЕХОВО-ЗУЕВСКИЙ', 'ПАВЛОВО-ПОСАДСКИЙ', 'ПОДОЛЬСК', 'ПУШКИНСКИЙ', 'РАМЕНСКИЙ',
  'РЕУТОВ', 'РУЗСКИЙ', 'СЕРГИЕВО-ПОСАДСКИЙ', 'СЕРЕБРЯНЫЕ ПРУДЫ', 'СЕРПУХОВ', 'СОЛНЕЧНОГОРСК', 'СТУПИНО', 'ТАЛДОМСКИЙ',
  'ФРЯЗИНО', 'ХИМКИ', 'ЧЕРНОГОЛОВКА', 'ЧЕХОВ', 'ШАТУРА', 'ШАХОВСКАЯ', 'ЩЕЛКОВО', 'ЭЛЕКТРОСТАЛЬ', ...ZATO];
const normOmsu = o => o.toUpperCase().replace(/Ё/g, 'Е').trim();
const isZato = o => ZATO.has(o.toUpperCase().replace(/Ё/g, 'Е'));
const LAST = {main: 51, zato: 5};

// 4-й показатель: категория РСО по НВВ (тыс. руб.) и план по периодам, % от НВВ (методика, стр. 20–21)
const cat4 = nvv => nvv < 300000 ? 1 : nvv < 600000 ? 2 : 3;
const PLAN4 = {1: [1, 2, 4, 6], 2: [1, 3, 6, 8], 3: [1, 4, 7, 10]};
const CAT_TXT = {1: 'НВВ < 300 млн руб.', 2: 'НВВ от 300 до 600 млн руб.', 3: 'НВВ ≥ 600 млн руб.'};

// 7-й показатель берётся не из RM, а из Google-таблицы дорожных карт (через /gsheet в server.py)
const GSHEET_ID = '<long-id-redacted>';
const GSHEET_URL = `<url>`;

const IND = {
  4: {
    query: 53, roman: 'IV', name: 'Увеличение платы за негативное воздействие', short: 'Негативное воздействие',
    valName: 'Выполнение плана, %',
    better: 'desc', fmt: 2, scale: null,
    c: [165, 166, 167, 168],               // С1нв…С4нв, тыс. руб. — в RM уже нарастающим итогом
    c2: [209, 210, 211, 212], nvv: 169, nvv2: 213, rso2: 208,   // вторая РСО ОМСУ (название — поле «РСО2»)
    rso: i => cfv(i, 70),
    about: '<b>Критерий отражает эффективность работы РСО по взиманию платы за негативное воздействие на централизованные системы водоотведения.</b> ' +
           'Характеризуется долей сбора платы за негативное воздействие в НВВ РСО на 2026 год, утверждённой Комитетом по ценам и тарифам МО.',
    rule: 'Выполнил план — 1 место, остальные — по % выполнения. Факт = С<sub>нв</sub> / НВВ, план — по категории РСО.',
  },
  5: {
    query: 52, roman: 'V', name: 'Снижение сверхнормативных расходов', short: 'Сверхнормативные расходы',
    valName: 'Балл', better: 'desc', fmt: 2, scale: 10,
    val: [158, 159, 160, 161],             // «Пок. 5 Балл» за период
    zero: [null, 259, 260, 261],           // «Обнулить балл …?»
    zeroYes: [null, '337', '339', '341'],  // значение «Да» в этих полях
    rso: i => i.subject,
    about: '<b>Критерий отражает работу РСО по снижению сверхнормативных расходов воды</b>: снижение потерь при транспортировке и выявление бездоговорного водопотребления. ' +
           'Балл от 0 до 10; если у ОМСУ несколько РСО/задач — берётся средний балл.',
    rule: 'Больше балл — выше место; несколько РСО — среднее. Балл обнулён — 51 место (ЗАТО — 5).',
  },
  6: {
    query: 54, roman: 'VI', name: 'Увеличение платы за отсутствие ИПУ', short: 'Отсутствие ИПУ',
    valName: 'Балл (доля ЛС без ИПУ), %', better: 'asc', fmt: 2, scale: null,
    val: [189, 195, 201, 207],             // «Балл N кв»
    extra: {cf: [[184, 185, 186, 187, 188], [190, 191, 192, 193, 194], [196, 197, 198, 199, 200], [202, 203, 204, 205, 206]]},
    rso: () => '',
    about: '<b>Цель — стимулировать РСО к полному учёту ресурсов и установке ИПУ</b> холодной/горячей воды либо применению повышающего коэффициента 3 при расчётах по нормативу.',
    rule: 'Всем ОМСУ — 1 место, как в официальном своде: ЛС<sub>факт</sub> из методики в RM нет. Балл — для справки.',
  },
  7: {
    source: 'gsheet', roman: 'VII', name: 'Снижение затрат на оплату электроэнергии', short: 'Электроэнергия',
    valName: 'Балл К7', better: 'desc', fmt: 2, scale: 100,
    about: '<b>Снижение затрат на оплату электроэнергии РСО</b> за счёт выполнения мероприятий «дорожной карты» по переводу точек учёта на оптимальную ценовую категорию. ' +
           'Данные — из <a href="' + GSHEET_URL + '" target="_blank">Google-таблицы</a> (лист «Исполнение дорожной карты»).',
    rule: 'Балл РСО = выполнено мероприятий дорожной карты / план × 100, ОМСУ — среднее. По ответам РСО в таблице, официально ещё не оценивался.',
  },
};

// ---------- 7-й показатель: дорожные карты по электроэнергии (Google-таблица, не RM) ----------
const K7_SHEET = 'Исполнение дорожной';
// Мероприятия «дорожной карты» (методика, стр. 38–40): колонка листа, срок; pct — колонка в %, выполнено только при 100
const K7_ITEMS = [
  {n: 1, col: 4, due: '27.02', name: 'Идентификация всех точек учёта'},
  {n: 2, col: 5, due: '06.03', name: 'Идентификация точек учёта ≥ 10 000 кВт·ч/мес.'},
  {n: 3, col: 6, due: '15.03', name: 'Перечень точек ТП с потреблением ≥ 10 000 кВт·ч/мес.'},
  {n: 4, col: 7, due: '15.03', name: 'Запросы ГП и СО при отсутствии/неполноте данных'},
  {n: 5, col: 9, due: '30.03', name: 'Сбор актов ТП и актов разграничения балансовой принадлежности', pct: true},
  {n: 6, col: 10, due: '10.04', name: 'Запросы на восстановление документов о ТП'},
  {n: 7, col: 11, due: '15.04', name: 'Сбор актов допуска ПУ в эксплуатацию', pct: true},
  {n: 8, col: 12, due: '30.04', name: 'Анализ ПУ на соответствие 3–4 ЦК', pct: true},
  {n: 9, col: 13, due: '15.05', name: 'Запросы в СО/ГП при отсутствии актов допуска ПУ'},
  {n: 10, col: 14, due: '22.05', name: 'Запрос архивных данных о потреблении, ЦК и МПИ ПУ'},
  {n: 11, col: 15, due: '31.05', name: 'Предварительная оценка потенциала снижения затрат', pct: true},
  {n: 12, col: 16, due: '30.06', name: 'Расчёт экономического эффекта смены ЦК', pct: true},
  {n: 13, col: 17, due: '15.07', name: 'Перечень точек с подтверждённой целесообразностью перехода'},
  {n: 14, col: 18, due: '31.07', name: 'Обращения к ГП о смене ЦК'},
  {n: 15, col: 19, due: '15.08', name: 'Выявление точек с несоответствующими ПУ'},
  {n: 16, col: 20, due: '31.08', name: 'Заявления в СО о замене (модернизации) ПУ'},
  {n: 17, col: 21, due: '30.09', name: 'Контроль замены (модернизации) ПУ'},
  {n: 18, col: null, due: '10.10', name: 'Перечень точек с эффектом после модернизации ПУ'},
  {n: 19, col: null, due: '16.11', name: 'Обращения к ГП по модернизированным ПУ'},
  {n: 20, col: null, due: '30.11', name: 'Подведение итогов, отчёт об экономическом эффекте'},
];
// План П по периодам — мероприятия со сроком до конца периода. В таблице только текущее состояние,
// поэтому считаем один период — 9 месяцев (мероприятия 1–17; 18–20 в таблице ещё нет).
const K7_PERIODS = [2];
const K7_PLAN = p => K7_ITEMS.filter(it => it.n <= [5, 12, 17, 20][p]);

// РСО — участники критерия по методике (стр. 40–41) и как их найти в таблице: [ОМСУ, РСО, шаблон названия, ОМСУ в таблице]
const K7_RSO = [
  ['БАЛАШИХА', 'МУП «БКС»', /БКС/],
  ['БОГОРОДСКИЙ', 'АО «Ногинсктрастинвест»', /Ногинсктраст|[«"]НТИ[»"]/i],   // в таблице — АО «НТИ»
  ['БОГОРОДСКИЙ', 'АО «НПТО ЖКХ»', /НПТО/],
  ['БРОННИЦЫ', 'АО «Бронницкий ТВК»', /Бронниц/i],
  ['ВЛАСИХА', 'МУП «Благоустройство и развитие»', /Благоустройство/i],
  ['ВОЛОКОЛАМСКИЙ', 'АО «Волоколамское ПТП РЖКХ»', /Волоколам/i],
  ['ДОЛГОПРУДНЫЙ', 'МУП «Инженерные сети г. Долгопрудного»', /Долгопрудн/i],
  ['ДОМОДЕДОВО', 'МУП «Теплосеть»', /^МУП\s*[«"]Теплосеть[»"]\s*(г\.?\s*о\.?\s*Домодедово)?\s*$/i, 'ДОМОДЕДОВО'],
  ['ДОМОДЕДОВО', 'МУП «Домодедовский водоканал»', /Домодедовский/i],
  ['ДУБНА', 'АО «ПТО ГХ»', /^АО\s*[«"]ПТО ГХ/i],
  ['ЕГОРЬЕВСК', 'МУП КХ «Егорьевские инженерные сети»', /Егорьевск/i],
  ['ЖУКОВСКИЙ', 'МП «Теплоцентраль»', /Теплоцентраль/i],
  ['ЗАРАЙСК', 'МУП «ЕСКХ Зарайского района»', /Зарайск/i],
  ['ИСТРА', 'МУП «Истринский водоканал»', /^МУП\s*[«"]Истрин/i],
  ['ИСТРА', 'АО «Истринская теплосеть»', /^АО\s*[«"]Истрин/i],
  ['КАШИРА', 'МУП «КГХ»', /КГХ|Кашир/i],
  ['КЛИН', 'МУП «Клинские тепловые сети»', /Клинск/i],
  ['КОЛОМНА', 'МУП «Тепло Коломны»', /Коломн/i],
  ['КОРОЛЕВ', 'АО «Водоканал»', /Водоканал.*Корол[её]в/i],
  ['КОРОЛЕВ', 'АО «Теплосеть»', /Теплосеть.*Корол[её]в/i],
  ['КОТЕЛЬНИКИ', 'МУЖКП «Котельники»', /Котельники/i],
  ['КРАСНОГОРСК', 'АО «Красногорская теплосеть»', /Красногорская/i],
  ['КРАСНОГОРСК', 'АО «Водоканал»', /^АО\s*[«"]Водоканал[»"]\s*г\.?\s*о\.?\s*Красногорск/i],
  ['КРАСНОГОРСК', 'АО «Водоканал «Павшино»', /Павшино/i],
  ['КРАСНОЗНАМЕНСК', 'МУП ВКХ «Водоканал»', /ВКХ/],
  ['КРАСНОЗНАМЕНСК', 'МУП ТХ «Теплосервис»', /Теплосервис/i],
  ['ЛЕНИНСКИЙ', 'МУП «Видновское ПТО ГХ»', /Видновское/i],
  ['ЛОСИНО-ПЕТРОВСКИЙ', 'МП «ЛП КТВС»', /КТВС/],
  ['ЛОТОШИНО', 'МКП «Лотошинское ЖКХ»', /Лотошин/i],
  ['ЛУХОВИЦЫ', 'ООО «ЛСК»', /ЛСК/],
  ['ЛЫТКАРИНО', 'МП «Водоканал»', /ВОДОКАНАЛ.*Лыткарино/i],
  ['ЛЫТКАРИНО', 'МП «Лыткаринская теплосеть»', /Лыткаринская/i],
  ['ЛЮБЕРЦЫ', 'АО «Люберецкий водоканал»', /Люберецкий/i],
  ['ЛЮБЕРЦЫ', 'АО «Люберецкая теплосеть»', /Люберецкая/i],
  ['ЛЮБЕРЦЫ', 'МУП «Водоканал»', /^МУП\s*[«"]Водоканал[»"]\s*$/i, 'ЛЮБЕРЦЫ'],
  ['НАРО-ФОМИНСКИЙ', 'МУП «Водоканал Наро-Фоминского ГО»', /Водоканал Наро/i],
  ['НАРО-ФОМИНСКИЙ', 'МУП «Теплосеть Наро-Фоминского ГО»', /Теплосеть Наро/i],
  ['ОДИНЦОВСКИЙ', 'МП «ЗИС»', /ЗИС/],
  ['ОДИНЦОВСКИЙ', 'АО «Одинцовская теплосеть»', /Одинцовская/i],
  ['ОДИНЦОВСКИЙ', 'МУП «ЖКХ Назарьево»', /Назарьево/i],
  ['ПОДОЛЬСК', 'МУП «Подольская теплосеть»', /Подольская/i],
  ['ПОДОЛЬСК', 'МУП «Водоканал» г. Подольск', /Водоканал/i, 'ПОДОЛЬСК'],
  ['ПУШКИНСКИЙ', 'МУП «ПКС»', /ПКС/],
  ['ПУШКИНСКИЙ', 'АО «Ивантеевская теплосеть»', /Ивантеев/i],
  ['РАМЕНСКИЙ', 'АО «Раменская теплосеть»', /Раменская/i],
  ['РАМЕНСКИЙ', 'АО «Раменский водоканал»', /Раменский водоканал/i],
  ['РЕУТОВ', 'МУП «Реутовский водоканал»', /Реутов/i],
  ['СЕРГИЕВО-ПОСАДСКИЙ', 'МУП «РКС»', /РКС/],
  ['СЕРЕБРЯНЫЕ ПРУДЫ', 'МУП «РСО МО Серебряные Пруды»', /Серебряные/i],
  ['СЕРПУХОВ', 'МУП «ТВК г. Пущино»', /Пущино/i],
  ['СЕРПУХОВ', 'МУП «Водоканал-Сервис»', /Водоканал-Сервис/i],
  ['СТУПИНО', 'МУП «ПТО ЖКХ»', /ПТО ЖКХ/],
  ['ФРЯЗИНО', 'АО «Теплосеть Фрязино»', /Теплосеть Фрязино/i],
  ['ЧЕРНОГОЛОВКА', 'МУП «УЭ»', /[«"]УЭ[»"]/],
  ['ШАТУРА', 'МУП «ШПТО ГХ»', /ШПТО/],
  ['ШАХОВСКАЯ', 'МКП «РСО Шаховская»', /Шаховская/i],
  ['ЩЕЛКОВО', 'МУП «Межрайонный Щёлковский водоканал»', /МЩВ|Щ[её]лковск/i],
  ['ЭЛЕКТРОСТАЛЬ', 'МУП «ЭЦУ»', /ЭЦУ/],
].map(([omsu, rso, re, where]) => ({omsu, rso, re, where}));

// Ответы в таблице — свободный текст. Правило: «выполнено» — только мероприятие, сделанное полностью
// (методика: частичное выполнение = 0 баллов); «не требуется / нецелесообразно / не выявлено» — выполнено.
// prevOk — выполнено ли предыдущее мероприятие (нужно, чтобы понять смысл «0»).
function k7cell(v, it, prevOk) {
  if (v === null || v === undefined || String(v).trim() === '') return {ok: false, why: 'не заполнено'};
  const s = String(v).toLowerCase().replace(/\s+/g, ' ').trim();
  if (/^\d+([.,]\d+)?$/.test(s)) {
    const x = Number(s.replace(',', '.'));
    if (it.pct) return x >= 100 ? {ok: true} : {ok: false, why: `выполнено на ${fmtN(x, 0)} %`};
    if (x > 0) return {ok: true};
    // «0» в колонках-количествах: 15–16 — несоответствующих ПУ нет, менять нечего;
    // 13–14 — целесообразных точек нет, но только если предыдущий шаг (расчёт / перечень) сделан
    if (it.n >= 15) return {ok: true, why: '«0» — несоответствующих ПУ нет'};
    return prevOk ? {ok: true, why: '«0» — целесообразных точек нет'} : {ok: false, why: '«0», а предыдущий шаг не выполнен'};
  }
  if (/\d+\s*%/.test(s)) return {ok: false, why: 'выполнено частично'};
  // 15: «осталось / выявлено N точек» — точки выявлены, мероприятие сделано
  if (it.n === 15 && /\d+\s*(точ|при|пу\b)/.test(s)) return {ok: true};
  // 16: «заявки / письма направлены, ведётся работа по замене» — заявления поданы, замена — это уже п. 17
  if (it.n === 16 && /направлен|заявк|исх|№|и-\d\d-/.test(s) && !/не\s*направл|будет/.test(s)) return {ok: true};
  if (/не\s*выполн|в работе|буд(ет|ут)|не\s*направл|проводится|вед[её]тся|выполняется|в стадии|на стадии|не\s*получ|на этапе|ожида|осталось|осталась|в ноябре|в сентябре|некоррект/.test(s))
    return {ok: false, why: 'не завершено'};
  if (/выполн|получен|направ|не\s*треб|не\s*целесообраз|не\s*выявл|в наличии|данные полные|данные есть|сформирован|проведен|выявлен|заменен|устан|запрос|исх|вх\.|заявк|письм|№|и-\d\d-|ответ|экономи|истек мпи|^да$|точ(ек|ки|ка) уч/.test(s))
    return {ok: true};
  // 17 — контроль замены ПУ: указан только исполнитель или канал («АО "Теплосеть"», «через личный кабинет Россети»,
  // «на контроле сетевых организаций») — подтверждения, что замена выполнена, нет
  if (it.n === 17) return {ok: false, why: 'нет подтверждения, что замена ПУ выполнена'};
  return {ok: false, unclear: true, why: 'неясная запись'};
}

function cfv(issue, id) {
  const c = (issue.custom_fields || []).find(f => f.id === id);
  return c ? c.value : null;
}
function num(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim().replace(',', '.');
  if (s === '') return null;
  const n = Number(s);
  return isFinite(n) ? n : null;
}
const round2 = n => Math.round(n * 100) / 100;
const omsuOf = i => (i.assigned_to && i.assigned_to.name || '').trim();
const fmtN = (n, d) => n === null || n === undefined ? '—' : n.toLocaleString('ru-RU', {minimumFractionDigits: d, maximumFractionDigits: d});

// ---------- расчёт рейтинга ----------
function isZeroed(cfg, issue, p) {
  if (!cfg.zero || !cfg.zero[p]) return false;
  const z = cfv(issue, cfg.zero[p]);
  return z !== null && String(z) === cfg.zeroYes[p];
}
function valueOf(cfg, issue, p) {
  if (isZeroed(cfg, issue, p)) return 0;
  return num(cfv(issue, cfg.val[p]));
}

// 4-й показатель по методике: по каждой РСО ОМСУ — факт, категория, план, % выполнения
function rso4(cfg, issues, p) {
  const out = [];
  const add = (name, c, nvv) => {
    if (c === null || !nvv) return;
    const fact = round2(c / nvv * 100), cat = cat4(nvv), plan = PLAN4[cat][p];
    out.push({name, c, nvv, fact, cat, plan, dost: round2(fact / plan * 100)});
  };
  for (const i of issues) {
    add(cfg.rso(i) || i.subject, num(cfv(i, cfg.c[p])), num(cfv(i, cfg.nvv)));
    if (num(cfv(i, cfg.nvv2))) add(cfv(i, cfg.rso2) || 'РСО 2', num(cfv(i, cfg.c2[p])), num(cfv(i, cfg.nvv2)));
  }
  return out;
}

// конкурентное ранжирование: одинаковые значения делят место, следующее — по порядковому номеру
function rankBy(list, better, offset = 0) {
  list.sort((a, b) => better === 'desc' ? b.v - a.v : a.v - b.v);
  list.forEach((r, k) => { r.place = (k && Math.abs(list[k - 1].v - r.v) < 1e-9) ? list[k - 1].place : offset + k + 1; });
}

// 7-й показатель: по каждой РСО из списка методики — строка таблицы и выполненные мероприятия
function k7rso(p) {
  const plan = K7_PLAN(p), sheet = DATA[7] || [];
  const norm = s => String(s || '').toUpperCase().replace(/Ё/g, 'Е').trim();
  const rowsBySlot = new Map(), unmatched = [];
  for (const row of sheet) {
    const name = String(row[2] || '').trim();
    if (!name) continue;
    const slot = K7_RSO.find(x => x.re.test(name) && (!x.where || norm(row[3]) === x.where));
    if (!slot) { unmatched.push(`${name} (${row[3] || '—'})`); continue; }
    if (!rowsBySlot.has(slot)) rowsBySlot.set(slot, []);
    rowsBySlot.get(slot).push(row);
  }
  const res = K7_RSO.map(slot => {
    const cand = (rowsBySlot.get(slot) || []).map(row => {
      const items = [];
      for (const it of plan) {
        const prevOk = items.length ? items[items.length - 1].ok : true;
        items.push({...it, raw: it.col === null ? null : row[it.col], ...(it.col === null ? {ok: false, why: 'нет в таблице'} : k7cell(row[it.col], it, prevOk))});
      }
      return {row, items, done: items.filter(x => x.ok).length};
    });
    // у РСО несколько строк (заполняли по разным кварталам) — берём самую полную
    cand.sort((a, b) => b.done - a.done);
    const best = cand[0];
    const done = best ? best.done : 0;
    return {...slot, found: !!best, dup: cand.length, items: best ? best.items : plan.map(it => ({...it, ok: false, why: 'РСО нет в таблице'})),
            sheetName: best ? String(best.row[2]).trim() : null, sheetOmsu: best ? String(best.row[3] || '').trim() : null,
            done, plan: plan.length, v: round2(done / plan.length * 100)};
  });
  return {rso: res, unmatched};
}

function rating7(p) {
  if (!DATA[7] || !K7_PERIODS.includes(p)) return [];
  const {rso} = k7rso(p);
  const by = new Map();
  for (const r of rso) { if (!by.has(r.omsu)) by.set(r.omsu, []); by.get(r.omsu).push(r); }
  const rows = [];
  for (const [o, list] of by) {
    rows.push({omsu: o, issues: [], k7: list, rso: list.map(x => x.rso),
               v: round2(list.reduce((a, x) => a + x.v, 0) / list.length),
               extra: {done: list.reduce((a, x) => a + x.done, 0), plan: list.reduce((a, x) => a + x.plan, 0)},
               group: isZato(o) ? 'zato' : 'main', noDocs: false});
  }
  for (const g of ['main', 'zato']) rankBy(rows.filter(r => r.group === g), 'desc');
  rows.sort((a, b) => (a.group === 'zato') - (b.group === 'zato') || a.place - b.place || a.omsu.localeCompare(b.omsu, 'ru'));
  return rows;
}

function rating(ind, p) {
  if (ind === 7) return rating7(p);
  const cfg = IND[ind], groups = new Map();
  for (const i of DATA[ind] || []) {
    const o = omsuOf(i);
    if (!o || num(cfv(i, 17)) === null) continue;   // задачи без ОМСУ (напр. Мособлводоканал) в рейтинг не идут
    if (!groups.has(o)) groups.set(o, []);
    groups.get(o).push(i);
  }
  const rows = [];
  for (const [o, issues] of groups) {
    const row = {omsu: o, issues, rso: [...new Set(issues.map(cfg.rso).filter(Boolean))], extra: null, perRso: null,
                 group: isZato(o) ? 'zato' : 'main', noDocs: false};
    if (ind === 4) {
      const rs = rso4(cfg, issues, p);
      if (!rs.length) continue;
      row.m4 = rs;
      row.rso = rs.map(x => x.name);
      row.v = round2(rs.reduce((a, x) => a + x.dost, 0) / rs.length);
      row.extra = rs.length === 1 ? rs[0].fact : round2(rs.reduce((a, x) => a + x.c, 0) / rs.reduce((a, x) => a + x.nvv, 0) * 100);
      row.plan = [...new Set(rs.map(x => x.plan))];
    } else {
      const vals = issues.map(i => valueOf(cfg, i, p)).filter(v => v !== null);
      if (!vals.length) continue;
      // значение не округляем: 6,6667 и 6,67 — разные места
      row.v = vals.reduce((a, b) => a + b, 0) / vals.length;
      if (ind === 6) {
        let tot = 0, no = 0;
        for (const i of issues) {
          const f = cfg.extra.cf[p];
          tot += num(cfv(i, f[0])) || 0;
          for (const k of f.slice(1)) no += num(cfv(i, k)) || 0;
        }
        row.extra = tot ? {no, tot} : null;
      }
      if (issues.length > 1) row.perRso = issues.map(i => ({rso: cfg.rso(i), v: valueOf(cfg, i, p), id: i.id}));
      // документы не представлены → балл обнулён по всем задачам ОМСУ
      if (cfg.zero) row.noDocs = issues.every(i => isZeroed(cfg, i, p));
    }
    rows.push(row);
  }

  for (const g of ['main', 'zato']) {
    const gr = rows.filter(r => r.group === g);
    const ok = gr.filter(r => !r.noDocs);
    gr.filter(r => r.noDocs).forEach(r => r.place = LAST[g]);
    if (ind === 4) {
      const ach = ok.filter(r => r.v >= 100), rest = ok.filter(r => r.v < 100);
      ach.forEach(r => r.place = 1);
      rankBy(rest, 'desc', ach.length);
    } else if (ind === 6) {
      ok.forEach(r => r.place = 1);
    } else {
      rankBy(ok, cfg.better);
    }
  }
  rows.sort((a, b) => (a.group === 'zato') - (b.group === 'zato') || a.place - b.place ||
    (cfg.better === 'desc' ? b.v - a.v : a.v - b.v));
  return rows;
}

function periodsWithData(ind) {
  return PERIODS.map((_, p) => rating(ind, p).length > 0);
}

// период по умолчанию — последний, где данные есть хотя бы у половины ОМСУ
// (в начале следующего периода сдают единицы, и рейтинг из 5 строк вводит в заблуждение)
function defaultPeriod(ind) {
  const n = PERIODS.map((_, p) => rating(ind, p).length), max = Math.max(...n);
  for (let p = n.length - 1; p >= 0; p--) if (n[p] && n[p] >= max / 2) return p;
  return -1;
}
// делят место — показываем диапазон: пять ОМСУ на 1-м месте → «1–5», следующий — 6
function placeLabel(r, rows) {
  if (r.noDocs) return String(r.place);   // обнулённые — все на последнем месте (51 / у ЗАТО 5), диапазона нет
  const n = rows.filter(x => x.group === r.group && x.place === r.place).length;
  return n > 1 ? `${r.place}–${r.place + n - 1}` : String(r.place);
}
let DATA = {};
module.exports = {
  evaluate(input) {
    DATA=input.data || {};
    return {periods:[0,1,2,3].map(p=>rating(input.ind,p).map(r=>({omsu:r.omsu,group:r.group,value:r.v,place:r.place,no_docs:r.noDocs,place_label:placeLabel(r,rating(input.ind,p))}))),default_period:defaultPeriod(input.ind)};
  },
  cell(input) {const r=k7cell(input.value,input.item,input.previous);return {ok:r.ok,reason:r.why,unclear:r.unclear};},
  numbers(input) {return input.map(num);},
  presentation(input) {
    DATA=input.data || {};
    const cfg=IND[input.ind],rows=rating(input.ind,input.period),prevRows=input.period>0?rating(input.ind,input.period-1):[];
    const prev=prevRows.length?new Map(prevRows.map(r=>[r.omsu,r.place])):null;
    const maxV=Math.max(...rows.map(r=>r.v),cfg.scale||0)||1;
    return rows.map(r=>{
      const gn=rows.filter(x=>x.group===r.group).length;
      const zone=r.place<=gn/3?'z1':r.place<=2*gn/3?'z2':'z3';
      const barW=cfg.better==='desc'?r.v/maxV:1-r.v/maxV;
      const pp=prev?prev.get(r.omsu):undefined;
      return {omsu:r.omsu,zone,bar_pct:Number(Math.max(2,Math.min(100,barW*100)).toFixed(1)),delta:pp===undefined?null:pp-r.place};
    });
  },
  constants() {
    return {items:K7_ITEMS.map(i=>({...i,pct:!!i.pct})),rso:K7_RSO.map(s=>({omsu:s.omsu,rso:s.rso,pattern:s.re.source,ignorecase:s.re.ignoreCase,where:s.where||null})),municipalities:ALL_OMSU,zato:[...ZATO],plans:PLAN4};
  }
};

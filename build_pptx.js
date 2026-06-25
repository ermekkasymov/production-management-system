const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.author = "Аналитический обзор";
pres.title = "Рынок акций KASE — обзор за последний год";

const W = 13.33, H = 7.5;

// ---- Палитра (финансы / KASE) ----
const NAVY = "0E2A47", BLUE = "1C6DD0", BLUE2 = "13507F";
const GOLD = "E0A82E", TEAL = "178A8A";
const GREEN = "1E9E5A", RED = "C0392B";
const INK = "1C2533", MUTED = "5B6B7B";
const PANEL = "F2F6FA", LINEC = "DCE4EC", WHITE = "FFFFFF";
const HEAD = "Cambria", BODY = "Calibri";

const shadow = () => ({ type: "outer", color: "9AABBC", blur: 8, offset: 3, angle: 135, opacity: 0.28 });

// ---- Хелперы ----
function motif(slide, x, y) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.18, h: 0.18, fill: { color: GOLD } });
  slide.addShape(pres.shapes.RECTANGLE, { x: x + 0.24, y, w: 0.18, h: 0.18, fill: { color: BLUE } });
}
function contentHeader(slide, title, kicker) {
  motif(slide, 0.6, 0.55);
  slide.addText((kicker || "РЫНОК АКЦИЙ KASE").toUpperCase(), {
    x: 0.6, y: 0.78, w: 11, h: 0.3, fontFace: BODY, fontSize: 11, color: BLUE, bold: true, charSpacing: 2, margin: 0
  });
  slide.addText(title, {
    x: 0.6, y: 1.05, w: 12.1, h: 0.8, fontFace: HEAD, fontSize: 28, color: NAVY, bold: true, margin: 0
  });
}
function footer(slide, n) {
  slide.addText("Обзор рынка акций KASE · 2025 – середина 2026", {
    x: 0.6, y: 7.06, w: 9, h: 0.3, fontFace: BODY, fontSize: 9, color: MUTED, margin: 0
  });
  slide.addText(String(n), {
    x: 12.5, y: 7.06, w: 0.5, h: 0.3, fontFace: BODY, fontSize: 9, color: MUTED, align: "right", margin: 0
  });
}
function statCard(slide, x, y, w, h, value, label, accent) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: WHITE }, line: { color: LINEC, width: 1 }, shadow: shadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.1, h, fill: { color: accent || GOLD } });
  slide.addText(value, { x: x + 0.22, y: y + 0.16, w: w - 0.35, h: h * 0.5, fontFace: HEAD, fontSize: 30, bold: true, color: NAVY, margin: 0, valign: "middle" });
  slide.addText(label, { x: x + 0.22, y: y + h * 0.58, w: w - 0.35, h: h * 0.38, fontFace: BODY, fontSize: 12, color: MUTED, margin: 0, valign: "top" });
}

// ============ 1. ТИТУЛ ============
(() => {
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.22, fill: { color: GOLD } });
  // motif column
  for (let i = 0; i < 6; i++) {
    s.addShape(pres.shapes.RECTANGLE, { x: 11.9, y: 1.4 + i * 0.55, w: 0.9, h: 0.32, fill: { color: i % 2 ? BLUE : "16365C" } });
  }
  s.addText("КАЗАХСТАНСКАЯ ФОНДОВАЯ БИРЖА", { x: 0.9, y: 1.7, w: 10, h: 0.4, fontFace: BODY, fontSize: 15, color: GOLD, bold: true, charSpacing: 3, margin: 0 });
  s.addText("Рынок акций KASE", { x: 0.9, y: 2.25, w: 11, h: 1.2, fontFace: HEAD, fontSize: 54, bold: true, color: WHITE, margin: 0 });
  s.addText("Обзор торгуемых акций за последний год", { x: 0.9, y: 3.55, w: 11, h: 0.7, fontFace: HEAD, fontSize: 26, color: "CFE0F2", italic: true, margin: 0 });
  s.addText([
    { text: "Индекс KASE · ", options: { color: "9DB6D2" } },
    { text: "+36,2% за 52 недели", options: { color: GREEN, bold: true } },
    { text: "   ·   капитализация рынка ", options: { color: "9DB6D2" } },
    { text: "≈ $90 млрд", options: { color: GOLD, bold: true } },
  ], { x: 0.9, y: 4.55, w: 11.5, h: 0.5, fontFace: BODY, fontSize: 16, margin: 0 });
  s.addText("Период: 2025 год – середина 2026 года   ·   Данные: KASE, AFK, открытые источники", {
    x: 0.9, y: 6.5, w: 11.5, h: 0.4, fontFace: BODY, fontSize: 12, color: "8FA8C6", margin: 0
  });
})();

// ============ 2. СОДЕРЖАНИЕ ============
(() => {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  contentHeader(s, "Содержание", "Навигация");
  const items = [
    ["01", "Ключевые итоги года", "Индекс, капитализация, число эмитентов"],
    ["02", "Динамика индекса KASE", "Траектория за 12 месяцев и максимумы"],
    ["03", "Представительский список", "10 голубых фишек индекса"],
    ["04", "Доходность акций за год", "Лидеры и аутсайдеры роста"],
    ["05", "Дивиденды и доходность", "Кто и сколько платит инвестору"],
    ["06", "Ликвидность и объёмы", "Самые торгуемые бумаги"],
    ["07", "Профили эмитентов", "Финансы · сырьё · телеком"],
    ["08", "Расширенный рынок и риски", "За пределами индекса, выводы"],
  ];
  const colW = 5.9, x0 = 0.6, x1 = 6.9, y0 = 2.0, rh = 1.18;
  items.forEach((it, i) => {
    const col = i < 4 ? x0 : x1;
    const row = i % 4;
    const y = y0 + row * rh;
    s.addShape(pres.shapes.RECTANGLE, { x: col, y, w: colW, h: 1.0, fill: { color: PANEL }, line: { color: LINEC, width: 1 } });
    s.addText(it[0], { x: col + 0.15, y: y + 0.12, w: 0.95, h: 0.76, fontFace: HEAD, fontSize: 30, bold: true, color: BLUE, align: "center", valign: "middle", margin: 0 });
    s.addText(it[1], { x: col + 1.2, y: y + 0.15, w: colW - 1.35, h: 0.42, fontFace: BODY, fontSize: 15, bold: true, color: NAVY, margin: 0, valign: "middle" });
    s.addText(it[2], { x: col + 1.2, y: y + 0.55, w: colW - 1.35, h: 0.38, fontFace: BODY, fontSize: 11, color: MUTED, margin: 0, valign: "middle" });
  });
  footer(s, 2);
})();

// ============ 3. КЛЮЧЕВЫЕ ИТОГИ ============
(() => {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  contentHeader(s, "Ключевые итоги года", "Раздел 01");
  const y = 2.05, h = 1.55, w = 2.85, gap = 0.22, x0 = 0.6;
  statCard(s, x0 + 0*(w+gap), y, w, h, "+36,2%", "рост индекса KASE за 52 недели", GREEN);
  statCard(s, x0 + 1*(w+gap), y, w, h, "7 728", "пунктов — значение индекса", BLUE);
  statCard(s, x0 + 2*(w+gap), y, w, h, "8 039", "исторический максимум (март 2026)", GOLD);
  statCard(s, x0 + 3*(w+gap), y, w, h, "≈ $90", "млрд — капитализация рынка", TEAL);
  // вторая строка
  const y2 = 3.85;
  statCard(s, x0 + 0*(w+gap), y2, w, h, "+26,1%", "доходность индекса за 2025 год", GREEN);
  statCard(s, x0 + 1*(w+gap), y2, w, h, "≈ 33%", "полная доходность с дивидендами", GREEN);
  statCard(s, x0 + 2*(w+gap), y2, w, h, "89 / 76", "видов акций / эмитентов в списках", BLUE);
  statCard(s, x0 + 3*(w+gap), y2, w, h, "62%", "доля физлиц в обороте по акциям", GOLD);

  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 5.7, w: 12.13, h: 1.0, fill: { color: NAVY } });
  s.addText([
    { text: "Главное:  ", options: { bold: true, color: GOLD } },
    { text: "рынок акций KASE завершил год уверенным ростом на фоне сильных финансовых результатов эмитентов, высоких дивидендов и рекордной капитализации. Основной вклад внесли банки и сырьевые компании.", options: { color: "E8EFF7" } },
  ], { x: 0.85, y: 5.78, w: 11.7, h: 0.84, fontFace: BODY, fontSize: 14, valign: "middle", margin: 0 });
  footer(s, 3);
})();

// ============ 4. ДИНАМИКА ИНДЕКСА ============
(() => {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  contentHeader(s, "Динамика индекса KASE", "Раздел 02");
  s.addChart(pres.charts.LINE, [{
    name: "Индекс KASE, пунктов",
    labels: ["май '25", "сен '25", "окт '25", "дек '25", "янв '26", "мар '26", "июн '26"],
    values: [5542, 7149, 7212, 7031, 7100, 8039, 7728],
  }], {
    x: 0.6, y: 2.0, w: 7.7, h: 4.6,
    chartColors: [BLUE], lineSize: 3, lineSmooth: true,
    showLegend: false, showTitle: false,
    chartArea: { fill: { color: WHITE } },
    catAxisLabelColor: MUTED, valAxisLabelColor: MUTED,
    catAxisLabelFontSize: 11, valAxisLabelFontSize: 11,
    valAxisMinVal: 5000, valAxisMaxVal: 8500, valGridLine: { color: "EBF0F5", size: 0.5 },
    showValue: true, dataLabelColor: NAVY, dataLabelFontSize: 9, dataLabelPosition: "t",
  });
  // правая колонка — вехи
  const bx = 8.7, bw = 4.0;
  const milestones = [
    ["Минимум 52 недель", "5 542 пункта · май 2025", GREEN],
    ["Итог 2025 года", "7 031 пункт · +26,1% за год", BLUE],
    ["Исторический максимум", "8 039 пунктов · март 2026", GOLD],
    ["Текущий уровень", "7 728 пунктов · −3,9% от пика", TEAL],
  ];
  s.addText("Ключевые уровни", { x: bx, y: 1.95, w: bw, h: 0.4, fontFace: HEAD, fontSize: 16, bold: true, color: NAVY, margin: 0 });
  milestones.forEach((m, i) => {
    const y = 2.5 + i * 1.05;
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y, w: bw, h: 0.9, fill: { color: PANEL }, line: { color: LINEC, width: 1 } });
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y, w: 0.09, h: 0.9, fill: { color: m[2] } });
    s.addText(m[0], { x: bx + 0.25, y: y + 0.12, w: bw - 0.4, h: 0.36, fontFace: BODY, fontSize: 13, bold: true, color: NAVY, margin: 0 });
    s.addText(m[1], { x: bx + 0.25, y: y + 0.48, w: bw - 0.4, h: 0.34, fontFace: BODY, fontSize: 11, color: MUTED, margin: 0 });
  });
  s.addText("Динамика приведена по отчётным ключевым уровням; промежуточные точки — ориентировочно.", {
    x: 0.6, y: 6.65, w: 7.7, h: 0.3, fontFace: BODY, fontSize: 9, italic: true, color: MUTED, margin: 0
  });
  footer(s, 4);
})();

// ============ 5. ПРЕДСТАВИТЕЛЬСКИЙ СПИСОК (таблица) ============
(() => {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  contentHeader(s, "Представительский список индекса", "Раздел 03");
  const head = ["Тикер", "Компания", "Сектор", "Цена, ₸", "За день", "Дох. 2025"];
  const rows = [
    ["HSBK", "Народный банк (Halyk)", "Финансы", "378,85", "+0,23%", "≈ +16%"],
    ["KSPI", "Kaspi.kz", "Финтех", "40 720", "−0,51%", "≈ +15%"],
    ["CCBN", "Банк ЦентрКредит", "Финансы", "4 630", "+0,11%", "≈ +15%"],
    ["KMGZ", "КазМунайГаз", "Нефть и газ", "33 065", "+0,20%", "≈ +15%"],
    ["KZAP", "Казатомпром", "Уран", "35 295", "+0,13%", "≈ +14%"],
    ["KCEL", "Kcell", "Телеком", "3 210", "−1,08%", "≈ +11%"],
    ["AIRA", "Air Astana", "Авиа", "659,99", "+0,33%", "≈ +6%"],
    ["KZTK", "Казахтелеком", "Телеком", "40 500", "−0,25%", "≈ +4%"],
    ["KEGC", "KEGOC", "Энергетика", "1 440,40", "−0,11%", "≈ +3%"],
    ["KZTO", "КазТрансОйл", "Транспорт", "1 115", "−0,98%", "≈ +2%"],
  ];
  const headRow = head.map(h => ({ text: h, options: { fill: { color: NAVY }, color: WHITE, bold: true, fontSize: 13, align: "center", valign: "middle" } }));
  const body = rows.map((r, i) => r.map((c, j) => {
    const opt = { fill: { color: i % 2 ? "EEF3F8" : WHITE }, color: INK, fontSize: 12, valign: "middle",
      align: (j === 0 || j >= 3) ? "center" : "left" };
    if (j === 0) { opt.bold = true; opt.color = BLUE2; }
    if (j === 4) opt.color = r[4].startsWith("+") ? GREEN : RED;
    if (j === 5) { opt.color = GREEN; opt.bold = true; }
    return { text: c, options: opt };
  }));
  s.addTable([headRow, ...body], {
    x: 0.6, y: 2.0, w: 12.13, colW: [1.3, 3.6, 2.0, 2.0, 1.6, 1.63],
    rowH: 0.42, border: { pt: 0.5, color: LINEC }, fontFace: BODY, valign: "middle",
  });
  s.addText("Индекс взвешен по капитализации с учётом акций в свободном обращении (free-float). Цены — на дату обзора.", {
    x: 0.6, y: 6.7, w: 12, h: 0.3, fontFace: BODY, fontSize: 9, italic: true, color: MUTED, margin: 0
  });
  footer(s, 5);
})();

// ============ 6. ДОХОДНОСТЬ ЗА ГОД (bar) ============
(() => {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  contentHeader(s, "Доходность акций за 2025 год", "Раздел 04");
  s.addChart(pres.charts.BAR, [{
    name: "Доходность, %",
    labels: ["HSBK", "CCBN", "KSPI", "KMGZ", "KZAP", "KCEL", "AIRA", "KZTK", "KEGC", "KZTO"],
    values: [16, 15, 15, 15, 14, 11, 6, 4, 3, 2],
  }], {
    x: 0.6, y: 2.0, w: 8.2, h: 4.7, barDir: "col",
    chartColors: [BLUE],
    chartArea: { fill: { color: WHITE } },
    catAxisLabelColor: MUTED, valAxisLabelColor: MUTED,
    catAxisLabelFontSize: 11, valAxisLabelFontSize: 11,
    valGridLine: { color: "EBF0F5", size: 0.5 }, catGridLine: { style: "none" },
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: NAVY, dataLabelFontSize: 11,
    valAxisMaxVal: 20, showLegend: false,
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 9.1, y: 2.0, w: 3.63, h: 2.2, fill: { color: PANEL }, line: { color: LINEC, width: 1 } });
  s.addText("Лидеры роста", { x: 9.3, y: 2.15, w: 3.3, h: 0.4, fontFace: HEAD, fontSize: 15, bold: true, color: GREEN, margin: 0 });
  s.addText([
    { text: "Halyk Bank  ", options: { bold: true } }, { text: "≈ +16%", options: { color: GREEN, bold: true, breakLine: true } },
    { text: "Банки и финтех  ", options: { bold: true } }, { text: "+15%", options: { color: GREEN, bold: true, breakLine: true } },
    { text: "Казатомпром  ", options: { bold: true } }, { text: "+14% (рост цен на уран)", options: { color: GREEN } },
  ], { x: 9.3, y: 2.6, w: 3.3, h: 1.5, fontFace: BODY, fontSize: 13, color: INK, margin: 0, lineSpacingMultiple: 1.3 });

  s.addShape(pres.shapes.RECTANGLE, { x: 9.1, y: 4.45, w: 3.63, h: 2.25, fill: { color: PANEL }, line: { color: LINEC, width: 1 } });
  s.addText("Отстающие", { x: 9.3, y: 4.6, w: 3.3, h: 0.4, fontFace: HEAD, fontSize: 15, bold: true, color: RED, margin: 0 });
  s.addText([
    { text: "KEGOC  ", options: { bold: true } }, { text: "≈ +3%", options: { color: MUTED, bold: true, breakLine: true } },
    { text: "Казахтелеком  ", options: { bold: true } }, { text: "+4%", options: { color: MUTED, bold: true, breakLine: true } },
    { text: "КазТрансОйл  ", options: { bold: true } }, { text: "+2% (снижение дивидендов)", options: { color: MUTED } },
  ], { x: 9.3, y: 5.05, w: 3.3, h: 1.5, fontFace: BODY, fontSize: 13, color: INK, margin: 0, lineSpacingMultiple: 1.3 });
  footer(s, 6);
})();

// ============ 7. ДИВИДЕНДЫ ============
(() => {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  contentHeader(s, "Дивиденды и доходность", "Раздел 05");
  statCard(s, 0.6, 2.05, 3.8, 1.5, "8 из 10", "акций индекса платят дивиденды", GOLD);
  statCard(s, 4.6, 2.05, 3.8, 1.5, "12,4%", "див. доходность Halyk — лидер", GREEN);
  statCard(s, 8.6, 2.05, 4.13, 1.5, "6–10%", "типичная доходность в тенге", TEAL);

  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 3.8, w: 12.13, h: 2.9, fill: { color: PANEL }, line: { color: LINEC, width: 1 } });
  s.addText("Что важно знать о выплатах", { x: 0.85, y: 3.95, w: 11.6, h: 0.4, fontFace: HEAD, fontSize: 16, bold: true, color: NAVY, margin: 0 });
  s.addText([
    { text: "Halyk Bank (HSBK) ", options: { bold: true, color: BLUE2 } },
    { text: "— самая дивидендная бумага индекса: ~12,4% доходности, среднегодовой рост дивидендов 26% за 5 лет, перешёл на выплаты дважды в год.", options: {} },
    { text: "\nKEGOC ", options: { bold: true, color: BLUE2 } },
    { text: "— второй по дивидендной доходности, также платит дважды в год; стабильный денежный поток.", options: {} },
    { text: "\nКазатомпром (KZAP) ", options: { bold: true, color: BLUE2 } },
    { text: "— на фоне роста цен на уран ожидается увеличение дивидендов (прогноз ≈ 1 500 ₸ на акцию по итогам 2025).", options: {} },
    { text: "\nКазТрансОйл (KZTO) ", options: { bold: true, color: BLUE2 } },
    { text: "— ранее дивидендный лидер рынка, но резко снизил выплаты — отношение инвесторов ухудшилось.", options: {} },
    { text: "\nKazakhtelecom, KEGOC, Halyk ", options: { bold: true, color: BLUE2 } },
    { text: "— обеспечивают регулярную доходность 6–10% в тенге.", options: {} },
  ], { x: 0.85, y: 4.4, w: 11.6, h: 2.2, fontFace: BODY, fontSize: 13.5, color: INK, margin: 0, lineSpacingMultiple: 1.18 });
  footer(s, 7);
})();

// ============ 8. ЛИКВИДНОСТЬ ============
(() => {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  contentHeader(s, "Ликвидность и объёмы торгов", "Раздел 06");
  s.addChart(pres.charts.BAR, [{
    name: "Объём торгов, млрд ₸",
    labels: ["HSBK", "CCBN", "AIRA", "KZAP", "KCEL"],
    values: [7.1, 6.2, 3.3, 1.9, 1.3],
  }], {
    x: 0.6, y: 2.0, w: 7.6, h: 4.6, barDir: "bar",
    chartColors: [TEAL],
    chartArea: { fill: { color: WHITE } },
    catAxisLabelColor: MUTED, valAxisLabelColor: MUTED,
    catAxisLabelFontSize: 12, valAxisLabelFontSize: 11,
    valGridLine: { color: "EBF0F5", size: 0.5 }, catGridLine: { style: "none" },
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: NAVY, dataLabelFontSize: 12,
    showLegend: false,
  });
  const bx = 8.55, bw = 4.18;
  statCard(s, bx, 2.0, bw, 1.5, "30,5%", "доля Halyk в обороте по акциям", BLUE);
  statCard(s, bx, 3.65, bw, 1.5, "62%", "оборота формируют физлица", GOLD);
  s.addShape(pres.shapes.RECTANGLE, { x: bx, y: 5.3, w: bw, h: 1.4, fill: { color: NAVY } });
  s.addText("Ликвидность сконцентрирована в нескольких голубых фишках: на топ-5 бумаг приходится большая часть оборота рынка.", {
    x: bx + 0.2, y: 5.42, w: bw - 0.4, h: 1.16, fontFace: BODY, fontSize: 13, color: "E8EFF7", valign: "middle", margin: 0
  });
  s.addText("Данные по объёмам — сентябрь 2025 (репрезентативный месяц).", {
    x: 0.6, y: 6.65, w: 7.6, h: 0.3, fontFace: BODY, fontSize: 9, italic: true, color: MUTED, margin: 0
  });
  footer(s, 8);
})();

// ============ 9. ПРОФИЛИ: ФИНАНСЫ ============
function profileSlide(n, title, kicker, cards) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  contentHeader(s, title, kicker);
  const cols = cards.length;
  const gap = 0.3, x0 = 0.6;
  const w = (12.13 - gap * (cols - 1)) / cols;
  cards.forEach((c, i) => {
    const x = x0 + i * (w + gap);
    const y = 2.05, h = 4.5;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: WHITE }, line: { color: LINEC, width: 1 }, shadow: shadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.85, fill: { color: NAVY } });
    s.addText(c.ticker, { x: x + 0.2, y: y + 0.12, w: w - 0.4, h: 0.34, fontFace: HEAD, fontSize: 18, bold: true, color: GOLD, margin: 0 });
    s.addText(c.name, { x: x + 0.2, y: y + 0.48, w: w - 0.4, h: 0.32, fontFace: BODY, fontSize: 12, color: "DCE8F5", margin: 0 });
    s.addText(c.ret, { x: x + 0.2, y: y + 1.0, w: w - 0.4, h: 0.55, fontFace: HEAD, fontSize: 26, bold: true, color: GREEN, margin: 0 });
    s.addText("доходность за 2025", { x: x + 0.2, y: y + 1.55, w: w - 0.4, h: 0.3, fontFace: BODY, fontSize: 10, color: MUTED, margin: 0 });
    s.addText(c.bullets.map((b, j) => ({ text: b, options: { bullet: true, breakLine: true, paraSpaceAfter: 6 } })),
      { x: x + 0.2, y: y + 2.0, w: w - 0.4, h: h - 2.15, fontFace: BODY, fontSize: 12, color: INK, margin: 0, valign: "top" });
  });
  footer(s, n);
}

profileSlide(9, "Профили: финансовый сектор", "Раздел 07 · Финансы", [
  { ticker: "HSBK", name: "Народный банк (Halyk)", ret: "≈ +16%", bullets: ["Крупнейший банк страны", "Самая ликвидная акция (30,5% оборота)", "Див. доходность ~12,4%", "Выплаты дважды в год"] },
  { ticker: "KSPI", name: "Kaspi.kz", ret: "≈ +15%", bullets: ["Финтех-экосистема (платежи, e-com)", "Высокая капитализация", "Листинг также на Nasdaq", "Драйвер цифровой экономики"] },
  { ticker: "CCBN", name: "Банк ЦентрКредит", ret: "≈ +15%", bullets: ["2-я по обороту бумага рынка", "Сильный рост прибыли", "Платит дивиденды", "Активный розничный спрос"] },
]);

// ============ 10. ПРОФИЛИ: СЫРЬЁ И ЭНЕРГЕТИКА ============
profileSlide(10, "Профили: сырьё и энергетика", "Раздел 07 · Сырьё", [
  { ticker: "KZAP", name: "Казатомпром", ret: "≈ +14%", bullets: ["Мировой лидер по урану", "Рост на фоне цен на уран", "Прогноз дивидендов ≈1500 ₸/акц.", "Экспортная выручка"] },
  { ticker: "KMGZ", name: "КазМунайГаз", ret: "≈ +15%", bullets: ["Нефтегазовый гигант", "Чувствителен к ценам на нефть", "Дивидендные выплаты", "Высокая цена акции (33 065 ₸)"] },
  { ticker: "KEGC", name: "KEGOC", ret: "≈ +3%", bullets: ["Оператор электросетей", "Защитный «дивидендный» актив", "Выплаты дважды в год", "Стабильный денежный поток"] },
  { ticker: "KZTO", name: "КазТрансОйл", ret: "≈ +2%", bullets: ["Транспортировка нефти", "Резко снизил дивиденды", "Слабая динамика за год", "Под давлением настроений"] },
]);

// ============ 11. ПРОФИЛИ: ТЕЛЕКОМ И ТРАНСПОРТ ============
profileSlide(11, "Профили: телеком и транспорт", "Раздел 07 · Телеком", [
  { ticker: "KZTK", name: "Казахтелеком", ret: "≈ +4%", bullets: ["Крупнейший телеком-оператор", "Регулярные дивиденды 6–10%", "Стабильный бизнес", "Умеренный рост котировок"] },
  { ticker: "KCEL", name: "Kcell", ret: "≈ +11%", bullets: ["Мобильный оператор", "Входит в топ-5 по ликвидности", "Дивидендные выплаты", "Хорошая годовая динамика"] },
  { ticker: "AIRA", name: "Air Astana", ret: "≈ +6%", bullets: ["Национальный авиаперевозчик", "3-я по обороту бумага", "IPO 2024 года", "Чувствительность к спросу на перелёты"] },
]);

// ============ 12. РАСШИРЕННЫЙ РЫНОК ============
(() => {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  contentHeader(s, "За пределами индекса: рынок в целом", "Раздел 08");
  const cards = [
    ["76 эмитентов", "В торговых списках KASE — 89 видов акций от 76 эмитентов. Помимо 10 голубых фишек индекса торгуются десятки бумаг второго эшелона.", BLUE],
    ["Низкая ликвидность", "По большинству бумаг вне индекса объёмы малы и сделки нерегулярны — котировки могут быть неактуальны, а спред — широким.", GOLD],
    ["ETF на индексы KASE", "Запущены первые биржевые фонды на индексы KASE — простой способ вложиться сразу в корзину ликвидных акций.", TEAL],
    ["KASE Global", "Отдельный сегмент с акциями зарубежных компаний (Apple, Tesla и др.) — расширяет выбор для инвесторов в Казахстане.", GREEN],
  ];
  const w = 5.95, h = 2.05, gap = 0.23, x0 = 0.6, y0 = 2.05;
  cards.forEach((c, i) => {
    const x = x0 + (i % 2) * (w + gap);
    const y = y0 + Math.floor(i / 2) * (h + gap);
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: PANEL }, line: { color: LINEC, width: 1 } });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.1, h, fill: { color: c[2] } });
    s.addText(c[0], { x: x + 0.28, y: y + 0.15, w: w - 0.5, h: 0.45, fontFace: HEAD, fontSize: 18, bold: true, color: NAVY, margin: 0 });
    s.addText(c[1], { x: x + 0.28, y: y + 0.62, w: w - 0.5, h: h - 0.75, fontFace: BODY, fontSize: 12.5, color: INK, margin: 0, valign: "top", lineSpacingMultiple: 1.12 });
  });
  footer(s, 12);
})();

// ============ 13. РИСКИ И ФАКТОРЫ ============
(() => {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  contentHeader(s, "Факторы роста и риски", "Раздел 08");
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 2.05, w: 5.95, h: 4.6, fill: { color: "EAF6EF" }, line: { color: "BFE3CD", width: 1 } });
  s.addText("Драйверы роста", { x: 0.85, y: 2.2, w: 5.5, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true, color: GREEN, margin: 0 });
  s.addText(["Сильные финансовые результаты эмитентов", "Высокие и растущие дивиденды", "Рост цен на сырьё (уран, нефть)", "Ускорение роста ВВП Казахстана", "Приток розничных инвесторов (62% оборота)", "Запуск ETF и развитие KASE Global"].map(b => ({ text: b, options: { bullet: true, breakLine: true, paraSpaceAfter: 9 } })),
    { x: 0.85, y: 2.7, w: 5.5, h: 3.8, fontFace: BODY, fontSize: 13.5, color: INK, margin: 0 });

  s.addShape(pres.shapes.RECTANGLE, { x: 6.78, y: 2.05, w: 5.95, h: 4.6, fill: { color: "FBECEC" }, line: { color: "F0CBC8", width: 1 } });
  s.addText("Риски и ограничения", { x: 7.03, y: 2.2, w: 5.5, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true, color: RED, margin: 0 });
  s.addText(["Концентрация ликвидности в 5–10 бумагах", "Низкая ликвидность второго эшелона", "Зависимость от цен на сырьё и курса тенге", "Снижение дивидендов отдельными эмитентами", "Близость индекса к историческим максимумам", "Геополитические и внешние факторы"].map(b => ({ text: b, options: { bullet: true, breakLine: true, paraSpaceAfter: 9 } })),
    { x: 7.03, y: 2.7, w: 5.5, h: 3.8, fontFace: BODY, fontSize: 13.5, color: INK, margin: 0 });
  footer(s, 13);
})();

// ============ 14. ВЫВОДЫ ============
(() => {
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.22, fill: { color: GOLD } });
  motif(s, 0.9, 0.85);
  s.addText("Выводы", { x: 0.9, y: 1.15, w: 11, h: 0.9, fontFace: HEAD, fontSize: 40, bold: true, color: WHITE, margin: 0 });
  s.addText([
    { text: "Рынок акций KASE показал сильный год: ", options: { bold: true, color: GOLD } },
    { text: "индекс прибавил +36% за 52 недели и обновил исторический максимум.", options: { color: "E8EFF7", breakLine: true } },
    { text: "Лидеры — банки, финтех и Казатомпром; ", options: { color: "E8EFF7" } },
    { text: "отстают телеком-инфраструктура и КазТрансОйл.", options: { color: "E8EFF7", breakLine: true } },
    { text: "Дивиденды остаются ключевым преимуществом: ", options: { color: "E8EFF7" } },
    { text: "Halyk и KEGOC обеспечивают 10–12% доходности в тенге.", options: { color: "E8EFF7", breakLine: true } },
    { text: "Главные ограничения — ", options: { color: "E8EFF7" } },
    { text: "концентрация ликвидности и близость к максимумам.", options: { color: "E8EFF7" } },
  ], { x: 0.9, y: 2.3, w: 11.3, h: 2.6, fontFace: BODY, fontSize: 18, margin: 0, lineSpacingMultiple: 1.3 });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.9, y: 5.2, w: 11.5, h: 1.5, fill: { color: "16365C" }, line: { color: BLUE, width: 1 } });
  s.addText([
    { text: "Дисклеймер: ", options: { bold: true, color: GOLD } },
    { text: "материал носит информационно-аналитический характер и не является индивидуальной инвестиционной рекомендацией. Цифры приведены на дату обзора по открытым источникам (KASE, AFK, отраслевые СМИ) и могут отличаться от актуальных. Источники: kase.kz, afk.kz, kapital.kz.", options: { color: "CFE0F2" } },
  ], { x: 1.15, y: 5.32, w: 11.0, h: 1.26, fontFace: BODY, fontSize: 11.5, valign: "middle", margin: 0, lineSpacingMultiple: 1.1 });
})();

pres.writeFile({ fileName: "KASE_акции_обзор.pptx" }).then(f => console.log("Saved:", f));

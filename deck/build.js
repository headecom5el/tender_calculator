// Generates the tender calculator proposal deck.
// Run: node deck/build.js
const pptxgen = require("pptxgenjs");
const path = require("path");

const IMG = path.join(__dirname, "..", "docs", "img");
const OUT = path.join(__dirname, "..", "docs", "tendernyy-kalkulyator.pptx");

const C = {
  dark: "0D3B3D",
  darkSoft: "17514F",
  petrol: "0E5C5E",
  mint: "5FBDB6",
  ink: "12191A",
  muted: "566A6B",
  light: "FFFFFF",
  tint: "EDF2F2",
  line: "CBD8D8",
  ok: "2E6E3F",
  warn: "9A6510",
  bad: "A5372C",
  white: "FFFFFF",
};
const F = { head: "Cambria", body: "Calibri" };
const M = 0.62;              // left margin
const W = 13.33 - M * 2;     // content width

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Управление e-commerce, ЗАО ПАТИО";
pres.title = "Тендерный калькулятор";

let pageNo = 0;

function slide(opts = {}) {
  const s = pres.addSlide();
  s.background = { color: opts.dark ? C.dark : C.light };
  pageNo += 1;
  if (!opts.noNum) {
    s.addText(String(pageNo), {
      isTextBox: true, x: 13.33 - M - 0.6, y: 6.95, w: 0.6, h: 0.3,
      align: "right", fontFace: F.body, fontSize: 10,
      color: opts.dark ? C.mint : C.line, margin: 0,
    });
  }
  return s;
}

function heading(s, eyebrow, title, opts = {}) {
  const onDark = !!opts.dark;
  s.addText(eyebrow, {
    isTextBox: true, x: M, y: 0.42, w: W, h: 0.26,
    fontFace: F.body, fontSize: 10.5, bold: true, charSpacing: 2,
    color: onDark ? C.mint : C.muted, margin: 0,
  });
  s.addText(title, {
    isTextBox: true, x: M, y: 0.72, w: opts.titleW || W, h: 0.62,
    fontFace: F.head, fontSize: opts.size || 32, bold: true,
    color: onDark ? C.white : C.ink, margin: 0,
  });
}

function card(s, x, y, w, h, fill) {
  s.addShape(pres.ShapeType.rect, {
    x, y, w, h, fill: { color: fill || C.tint }, line: { color: fill || C.tint, width: 0 },
  });
}

// ─────────────────────────────────────────── 1. титул
{
  const s = slide({ dark: true, noNum: true });
  s.addText("ЗАО ПАТИО   ·   УПРАВЛЕНИЕ E-COMMERCE", {
    isTextBox: true, x: M, y: 1.55, w: W, h: 0.3,
    fontFace: F.body, fontSize: 12, bold: true, charSpacing: 2.5, color: C.mint, margin: 0,
  });
  s.addText("Тендерный калькулятор", {
    isTextBox: true, x: M, y: 2.0, w: W, h: 1.0,
    fontFace: F.head, fontSize: 46, bold: true, color: C.white, margin: 0,
  });
  s.addText("Предложение по автоматизации закупки Apple у трейдеров:\nчто болит, что предлагаю сделать, сколько это займёт и кто это делает", {
    isTextBox: true, x: M, y: 3.05, w: 9.2, h: 1.0,
    fontFace: F.body, fontSize: 17, color: "C7DEDD", lineSpacing: 26, margin: 0,
  });
  s.addText("Артём Боровский, начальник управления e-commerce   ·   7 сентября 2026", {
    isTextBox: true, x: M, y: 6.35, w: W, h: 0.35,
    fontFace: F.body, fontSize: 12, color: C.mint, margin: 0,
  });
  s.addNotes("Задача поставлена 17 августа. Это предложение по её реализации.");
}

// ─────────────────────────────────────────── 2. что я разобрал
{
  const s = slide();
  heading(s, "ОСНОВАНИЕ", "С чего я начал");
  s.addText(
    "Задачу поставили 17 августа: собрать в одном месте условия работы с трейдерами и сделать " +
    "тендерный калькулятор.\n\nЯ разобрал пять рабочих файлов и восстановил по ним, как процесс " +
    "устроен на самом деле. Всё, что дальше в презентации, посчитано на этих данных, а не на " +
    "предположениях.",
    { isTextBox: true, x: M, y: 1.75, w: 5.1, h: 2.6,
      fontFace: F.body, fontSize: 15, color: C.ink, lineSpacing: 24, margin: 0 });

  const files = [
    ["Техническое задание", "требования к автоматизации, 6 страниц"],
    ["Свод условий по трейдерам, 18.08", "базис, оплата, курсы, сроки, инспекция"],
    ["Форма заявки, 04.08", "заполнена трейдером Inonics"],
    ["Форма заявки, оптимизированная", "заполнена трейдером Noviva"],
    ["Сводная сравнения, 25.06", "реальный расчёт по девяти трейдерам"],
    ["Прототип калькулятора, 20.08", "10 листов, коллегам не подошёл"],
  ];
  let y = 1.75;
  files.forEach(([t, d]) => {
    card(s, 6.2, y, 6.5, 0.72);
    s.addText(t, { isTextBox: true, x: 6.45, y: y + 0.08, w: 6.0, h: 0.28,
      fontFace: F.body, fontSize: 13.5, bold: true, color: C.ink, margin: 0 });
    s.addText(d, { isTextBox: true, x: 6.45, y: y + 0.37, w: 6.0, h: 0.28,
      fontFace: F.body, fontSize: 11.5, color: C.muted, margin: 0 });
    y += 0.84;
  });
}

// ─────────────────────────────────────────── 3. как сейчас
{
  const s = slide();
  heading(s, "КАК СЕЙЧАС", "Четыре ручных этапа");
  const steps = [
    ["1", "Рассылка", "Менеджер заполняет Excel и вручную отправляет его в девять отдельных чатов в Telegram. Общая таблица запрещена: трейдеры не видят друг друга."],
    ["2", "Сбор", "Девять файлов приходят в девять чатов. Менеджер открывает каждый и переносит данные в общую сводную вручную."],
    ["3", "Нотификация", "По каждой модели менеджер вручную ищет запись в реестре ЕАЭС, пробуя латинскую и кириллическую «А». Без нотификации товар нельзя продавать в РБ."],
    ["4", "Расчёт", "Себестоимость считается девятью разными формулами. У каждого трейдера свой источник курса, своя надбавка и своя логистика."],
  ];
  const cw = 2.92, gap = 0.16;
  steps.forEach(([n, t, d], i) => {
    const x = M + i * (cw + gap);
    card(s, x, 1.8, cw, 3.5);
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.28, y: 2.05, w: 0.5, h: 0.5,
      fill: { color: C.petrol }, line: { color: C.petrol, width: 0 } });
    s.addText(n, { isTextBox: true, x: x + 0.28, y: 2.13, w: 0.5, h: 0.34,
      align: "center", fontFace: F.body, fontSize: 15, bold: true, color: C.white, margin: 0 });
    s.addText(t, { isTextBox: true, x: x + 0.28, y: 2.72, w: cw - 0.56, h: 0.35,
      fontFace: F.head, fontSize: 17, bold: true, color: C.ink, margin: 0 });
    s.addText(d, { isTextBox: true, x: x + 0.28, y: 3.15, w: cw - 0.56, h: 1.95,
      fontFace: F.body, fontSize: 12.5, color: C.muted, lineSpacing: 18, margin: 0, valign: "top" });
  });
  s.addText("На всё это уходит рабочий день менеджера, и каждый шаг можно сделать неправильно.", {
    isTextBox: true, x: M, y: 5.55, w: W, h: 0.4,
    fontFace: F.body, fontSize: 14, italic: true, color: C.petrol, margin: 0 });
}

// ─────────────────────────────────────────── 4. масштаб
{
  const s = slide();
  heading(s, "МАСШТАБ", "О каких деньгах идёт речь");
  const tiles = [
    ["17", "позиций в одном тендере", "заявка от 04.08.2026"],
    ["1 730", "штук техники", "суммарная потребность"],
    ["1,35", "млн долларов", "объём по таргетам закупки"],
    ["9", "трейдеров", "и у каждого своя формула"],
  ];
  const cw = 2.92, gap = 0.16;
  tiles.forEach(([n, t, d], i) => {
    const x = M + i * (cw + gap);
    card(s, x, 1.85, cw, 2.35);
    s.addText(n, { isTextBox: true, x: x + 0.3, y: 2.05, w: cw - 0.6, h: 0.9,
      fontFace: F.head, fontSize: 44, bold: true, color: C.petrol, margin: 0 });
    s.addText(t, { isTextBox: true, x: x + 0.3, y: 3.0, w: cw - 0.6, h: 0.6,
      fontFace: F.body, fontSize: 14, bold: true, color: C.ink, margin: 0 });
    s.addText(d, { isTextBox: true, x: x + 0.3, y: 3.55, w: cw - 0.6, h: 0.45,
      fontFace: F.body, fontSize: 11.5, color: C.muted, margin: 0 });
  });
  s.addText("Это один тендер. Они идут регулярно, и по каждому решение принимается вручную.", {
    isTextBox: true, x: M, y: 4.5, w: W, h: 0.4,
    fontFace: F.body, fontSize: 15, color: C.ink, margin: 0 });
  s.addNotes("Цифры взяты из реальной заявки от 04.08.2026 и таргетов от 10.08.");
}

// ─────────────────────────────────────────── 5. почему риск (тёмный)
{
  const s = slide({ dark: true });
  heading(s, "ГЛАВНОЕ", "Почему я считаю это риском, а не неудобством", { dark: true, size: 30 });
  s.addText("0,07 %", {
    isTextBox: true, x: M, y: 1.95, w: 4.4, h: 1.1,
    fontFace: F.head, fontSize: 58, bold: true, color: C.mint, margin: 0 });
  s.addText("разрыв между первым и вторым предложением в реальном расчёте по позиции iPhone 17 256GB", {
    isTextBox: true, x: M, y: 3.05, w: 4.6, h: 1.2,
    fontFace: F.body, fontSize: 14.5, color: "C7DEDD", lineSpacing: 21, margin: 0 });

  card(s, 5.9, 1.9, 6.8, 3.5, C.darkSoft);
  s.addText("Что я нашёл в файле от 25.06", {
    isTextBox: true, x: 6.25, y: 2.15, w: 6.1, h: 0.35,
    fontFace: F.head, fontSize: 19, bold: true, color: C.white, margin: 0 });
  s.addText([
    { text: "Надбавка 0,5 % по одному из трейдеров учтена дважды: и в курсе, и коэффициентом к цене.", options: { breakLine: true } },
    { text: "В переписке 19 августа трейдер прямо ответил, что надбавка одна.", options: { breakLine: true } },
    { text: "Это 15 рублей на единице товара при разрыве между первым и вторым местом в 2 рубля. На позиции в 320 штук — 4 800 рублей и другой победитель тендера.", options: { breakLine: true } },
    { text: "Файл делали грамотные люди. Ошибка прожила в нём несколько месяцев, потому что в таблице её не видно." },
  ], { isTextBox: true, x: 6.25, y: 2.65, w: 6.1, h: 2.5,
       fontFace: F.body, fontSize: 14, color: "C7DEDD", lineSpacing: 21, paraSpaceAfter: 8, margin: 0 });
  s.addText("Решение о закупке на 1,35 млн долларов принимается по цифре, которую никто не может проверить.", {
    isTextBox: true, x: M, y: 5.75, w: W, h: 0.5,
    fontFace: F.body, fontSize: 15, bold: true, color: C.white, margin: 0 });
}

// ─────────────────────────────────────────── 6. девять формул
{
  const s = slide();
  heading(s, "ПОЧЕМУ ЭТО СЛОЖНО", "Девять трейдеров, девять разных формул");
  const rows = [
    [{ text: "Трейдер" }, { text: "Валюта и источник курса" }, { text: "Надбавка" }, { text: "Природа надбавки" }, { text: "Логистика" }],
    ["Inonics trading", "RUB, НБ РБ", "1,0 %", "добавляют наши финансы при оплате", "в цене, CIP Минск"],
    ["Tech Planet FZCO", "RUB, инвестинг + 10 коп.", "1,5 %", "комиссия платёжного агента", "7 $ за единицу"],
    ["Al Shiva General", "RUB, инвестинг + 10 коп.", "1,5 %", "комиссия платёжного агента", "7 $ за единицу"],
    ["Noviva DWC", "RUB, ЦБ РФ", "1,25 %", "конвертация рублей в дирхамы", "7 $ за единицу"],
    ["KVN Group, Бигап", "RUB, ЦБ РФ", "1,0 %", "добавляют наши финансы при оплате", "7 $ за единицу"],
    ["Angel Trade", "CNY, инвестинг + 0,5 %", "0,5 %", "курс банка трейдера", "в цене, DAP Минск"],
    ["HuaXun Group", "CNY", "1,0 %", "курс в формуле вписан числом", "в цене, CIP Минск"],
  ];
  s.addTable(rows, {
    x: M, y: 1.8, w: W, colW: [2.4, 2.9, 1.3, 4.0, 1.53],
    fontFace: F.body, fontSize: 12, color: C.ink,
    border: { type: "solid", pt: 0.75, color: C.line },
    fill: { color: C.white },
    rowH: 0.42, valign: "middle",
    margin: [4, 8, 4, 8],
  });
  s.addText("Формулы живут в ячейках Excel и меняются вручную. Проверить их может только тот, кто их писал.", {
    isTextBox: true, x: M, y: 5.75, w: W, h: 0.4,
    fontFace: F.body, fontSize: 14, italic: true, color: C.petrol, margin: 0 });
  s.addNotes("Ещё два трейдера, ETK и VVP, есть в файле сравнения, но отсутствуют в своде условий. Списки не согласованы между собой.");
}

// ─────────────────────────────────────────── 7. что ещё мешает
{
  const s = slide();
  heading(s, "ЧТО ЕЩЁ МЕШАЕТ", "Четыре проблемы, которые я вижу в данных");
  const items = [
    ["Нотификации никто не заполняет", "Ни в одном полученном файле колонки «Нотификация» и «Модель» не заполнены. Проверять по факту нечего, а без нотификации партию нельзя продавать в РБ."],
    ["Файлы приходят грязными", "Цена приходит текстом: «$825». В поле итоговой цены встречается «$1425 orange, $1450 silver». Числами это не посчитать."],
    ["В шаблоне остались старые формулы", "В новой форме колонка названа «Статус позиции», а внутри осталась формула сложения от прежней версии. Трейдер видит в ней мусор."],
    ["Тендер идёт в несколько раундов", "В форме от 04.08 два столбца таргета, с разными датами. Торг многораундовый, а в техническом задании описан как разовый сбор предложений."],
  ];
  const cw = 6.05, ch = 1.85, gx = 0.22, gy = 0.25;
  items.forEach(([t, d], i) => {
    const x = M + (i % 2) * (cw + gx);
    const y = 1.85 + Math.floor(i / 2) * (ch + gy);
    card(s, x, y, cw, ch);
    s.addText(t, { isTextBox: true, x: x + 0.3, y: y + 0.2, w: cw - 0.6, h: 0.35,
      fontFace: F.head, fontSize: 17, bold: true, color: C.ink, margin: 0 });
    s.addText(d, { isTextBox: true, x: x + 0.3, y: y + 0.63, w: cw - 0.6, h: 1.05,
      fontFace: F.body, fontSize: 13, color: C.muted, lineSpacing: 19, margin: 0, valign: "top" });
  });
  s.addText("Ни одна из этих проблем не решается аккуратностью. Их решает система, которая не принимает неправильные данные.", {
    isTextBox: true, x: M, y: 6.1, w: W, h: 0.4,
    fontFace: F.body, fontSize: 14, italic: true, color: C.petrol, margin: 0 });
}

// ─────────────────────────────────────────── 8. что предлагаю (тёмный)
{
  const s = slide({ dark: true });
  heading(s, "ПРЕДЛОЖЕНИЕ", "Сервис вместо таблицы", { dark: true });
  s.addText(
    "Трейдер остаётся в своём чате в Telegram и ничего нового не осваивает. Менеджер работает " +
    "в кабинете и получает готовую сводную.",
    { isTextBox: true, x: M, y: 1.75, w: 8.6, h: 0.8,
      fontFace: F.body, fontSize: 17, color: "C7DEDD", lineSpacing: 26, margin: 0 });
  const pts = [
    ["Рассылка и сбор", "Бот сам отправляет запрос каждому трейдеру и сам забирает ответ. Данные других участников недоступны никому из них."],
    ["Расчёт", "Себестоимость в рублях по единой методике, с разложением по шагам. Маржа считается к целевой РРЦ."],
    ["Нотификации", "Проверяются автоматически в реестре ЕАЭС, в двух раскладках. Позиция без нотификации блокируется к выбору."],
    ["История", "Каждый расчёт сохраняется вместе с курсами на дату и воспроизводится задним числом."],
  ];
  const cw = 2.92, gap = 0.16;
  pts.forEach(([t, d], i) => {
    const x = M + i * (cw + gap);
    card(s, x, 2.85, cw, 2.6, C.darkSoft);
    s.addText(t, { isTextBox: true, x: x + 0.28, y: 3.08, w: cw - 0.56, h: 0.35,
      fontFace: F.head, fontSize: 17, bold: true, color: C.white, margin: 0 });
    s.addText(d, { isTextBox: true, x: x + 0.28, y: 3.55, w: cw - 0.56, h: 1.75,
      fontFace: F.body, fontSize: 12.5, color: "C7DEDD", lineSpacing: 18, margin: 0, valign: "top" });
  });
  s.addText("Excel остаётся запасным каналом. Кто не хочет форму, отправляет файл в чат как раньше.", {
    isTextBox: true, x: M, y: 5.85, w: W, h: 0.4,
    fontFace: F.body, fontSize: 14, color: C.mint, margin: 0 });
}

// ─────────────────────────────────────────── 9. схема
{
  const s = slide();
  heading(s, "АРХИТЕКТУРА", "Схема реализации");
  s.addImage({ path: path.join(IMG, "05-diagram.png"), x: 1.55, y: 1.6, w: 10.2, h: 4.8 });
  s.addText("Пунктиром показана вторая очередь. До неё потребность и целевая РРЦ вводятся вручную.", {
    isTextBox: true, x: M, y: 6.55, w: W, h: 0.35,
    fontFace: F.body, fontSize: 12, color: C.muted, margin: 0 });
}

// ─────────────────────────────────────────── 10. кабинет менеджера
{
  const s = slide();
  heading(s, "ИНТЕРФЕЙС, МАКЕТ", "Что видит менеджер");
  s.addImage({ path: path.join(IMG, "06-kabinet-win.png"), x: 0.9, y: 1.55, w: 8.4, h: 5.28 });
  const notes = [
    ["Одна позиция, все предложения", "Себестоимость сразу в рублях, маржа к целевой РРЦ, покрытие потребности и срок прихода на ЦС."],
    ["Разложение цифры по шагам", "Та же цепочка, которую менеджер сегодня считает вручную. Поэтому ей можно верить."],
    ["Блокировки видны сразу", "Позиция без нотификации подсвечена и к выбору недоступна."],
  ];
  let y = 1.75;
  notes.forEach(([t, d]) => {
    s.addText(t, { isTextBox: true, x: 9.6, y, w: 3.1, h: 0.5,
      fontFace: F.head, fontSize: 15, bold: true, color: C.ink, margin: 0 });
    s.addText(d, { isTextBox: true, x: 9.6, y: y + 0.45, w: 3.1, h: 1.2,
      fontFace: F.body, fontSize: 12, color: C.muted, lineSpacing: 17, margin: 0 });
    y += 1.7;
  });
  s.addNotes("Цифры на макете посчитаны настоящими формулами на курсах из файла 25.06. Целевая РРЦ и артикул условные.");
}

// ─────────────────────────────────────────── 11. методика расчёта
{
  const s = slide();
  heading(s, "МЕТОДИКА", "Как считается себестоимость");
  const steps = [
    ["Цена предложения трейдера", "в валюте оффера"],
    ["Надбавка трейдера", "процент и его природа берутся из справочника, а не из формулы в ячейке"],
    ["Логистика до склада", "только если базис поставки этого требует"],
    ["Конвертация по траншам", "каждый платёж по курсу своей даты и своего источника: НБ РБ, ЦБ РФ или инвестинг"],
    ["Пошлина, НДС, утильсбор, гарантийный сбор", "ставки задаются финансовой службой и хранятся с датой действия"],
    ["Стоимость денег", "предоплата стоит денег, отсрочка их экономит. Считается по обеим сторонам"],
    ["Маржа к целевой РРЦ", "то, в чём менеджер реально принимает решение"],
  ];
  let y = 1.78;
  steps.forEach(([t, d], i) => {
    card(s, M, y, W, 0.62, i % 2 ? C.white : C.tint);
    s.addText(String(i + 1), { isTextBox: true, x: M + 0.22, y: y + 0.15, w: 0.4, h: 0.32,
      fontFace: F.body, fontSize: 14, bold: true, color: C.petrol, margin: 0 });
    s.addText(t, { isTextBox: true, x: M + 0.72, y: y + 0.14, w: 4.7, h: 0.34,
      fontFace: F.body, fontSize: 14, bold: true, color: C.ink, margin: 0 });
    s.addText(d, { isTextBox: true, x: M + 5.5, y: y + 0.16, w: W - 5.7, h: 0.34,
      fontFace: F.body, fontSize: 12.5, color: C.muted, margin: 0 });
    y += 0.66;
  });
  s.addText("Ни одного числа в коде. Курсы, ставки и надбавки лежат в справочниках, каждый расчёт сохраняется целиком.", {
    isTextBox: true, x: M, y: 6.5, w: W, h: 0.4,
    fontFace: F.body, fontSize: 13.5, italic: true, color: C.petrol, margin: 0 });
}

// ─────────────────────────────────────────── 12. вид трейдера
{
  const s = slide();
  heading(s, "ИНТЕРФЕЙС, МАКЕТ", "Что видит трейдер");
  s.addImage({ path: path.join(IMG, "04-phone.png"), x: 9.55, y: 1.35, w: 2.2, h: 5.27 });
  const items = [
    ["Открывается прямо в Telegram", "По кнопке из чата. Логина и пароля нет, трейдер опознан своим аккаунтом."],
    ["Поля с типами", "Цена принимает только число. Запись «1425 orange» отклоняется в момент ввода, а не при разборе файла у нас."],
    ["Обязательные поля", "Модель и номер нотификации пропустить нельзя. Сегодня их не заполняет ни один трейдер."],
    ["Изоляция", "Трейдер видит только свои строки. Ни цен, ни имён других участников в интерфейсе нет."],
    ["Excel остаётся", "Кто не хочет форму, отправляет файл в чат как раньше. Бот его примет и разберёт."],
  ];
  let y = 1.72;
  items.forEach(([t, d]) => {
    s.addText(t, { isTextBox: true, x: M, y, w: 8.5, h: 0.32,
      fontFace: F.head, fontSize: 16, bold: true, color: C.ink, margin: 0 });
    s.addText(d, { isTextBox: true, x: M, y: y + 0.32, w: 8.5, h: 0.55,
      fontFace: F.body, fontSize: 13, color: C.muted, lineSpacing: 18, margin: 0 });
    y += 0.98;
  });
}

// ─────────────────────────────────────────── 13. стек
{
  const s = slide();
  heading(s, "ТЕХНИЧЕСКАЯ ЧАСТЬ", "Из чего собирается сервис");
  const rows = [
    [{ text: "Слой" }, { text: "Чем делается" }, { text: "Почему так" }],
    ["Расчёт и домен", "Python, FastAPI", "лучший инструментарий для разбора Excel и расчётов"],
    ["Бот в Telegram", "aiogram", "рассылка, приём файлов, статусы, напоминания"],
    ["Форма трейдера", "Telegram Mini App", "открывается в чате, без логинов и паролей"],
    ["Кабинет менеджера", "Vue 3", "та же технология, что на сайте компании"],
    ["Хранение", "PostgreSQL, S3", "история тендеров, курсы на дату, оригиналы файлов"],
    ["Развёртывание", "Docker", "переносится в контур IT без переписывания"],
  ];
  s.addTable(rows, {
    x: M, y: 1.8, w: 7.7, colW: [2.3, 2.2, 3.2],
    fontFace: F.body, fontSize: 12.5, color: C.ink,
    border: { type: "solid", pt: 0.75, color: C.line },
    fill: { color: C.white }, rowH: 0.46, valign: "middle", margin: [4, 8, 4, 8],
  });
  card(s, 8.6, 1.8, 4.1, 3.22);
  s.addText("Что берём снаружи", { isTextBox: true, x: 8.9, y: 2.0, w: 3.5, h: 0.3,
    fontFace: F.head, fontSize: 16, bold: true, color: C.ink, margin: 0 });
  s.addText([
    { text: "НБ РБ и ЦБ РФ отдают курсы машине, забираем автоматически", options: { bullet: true, breakLine: true } },
    { text: "Курс инвестинга вводит менеджер один раз в день, публичного источника нет", options: { bullet: true, breakLine: true } },
    { text: "Реестр ЕАЭС проверяем по каждой модели, ответ сохраняем как подтверждение", options: { bullet: true } },
  ], { isTextBox: true, x: 8.9, y: 2.4, w: 3.5, h: 2.4,
       fontFace: F.body, fontSize: 12.5, color: C.muted, lineSpacing: 18, paraSpaceAfter: 8, margin: 0 });
  s.addText("Вторая очередь: подтягивать потребность и остатки из 1С:УТ, а целевую РРЦ и продажи из Qlik. Без них сервис работает, эти данные вводятся вручную.", {
    isTextBox: true, x: M, y: 5.45, w: W, h: 0.5,
    fontFace: F.body, fontSize: 13.5, color: C.ink, lineSpacing: 20, margin: 0 });
}

// ─────────────────────────────────────────── 14. как делаем
{
  const s = slide({ dark: true });
  heading(s, "МОДЕЛЬ РАБОТЫ", "Кто делает и кто поддерживает", { dark: true });
  const stages = [
    ["Этап 1", "Разработка у нас", "Управление e-commerce делает сервис целиком: справочники, расчёт, бот, форму трейдера, кабинет и проверку нотификаций. Обкатываем на живых тендерах.", "25–35 чел-дней, 6–8 недель"],
    ["Этап 2", "Передача в IT", "Отдаём исходный код, схему данных, автотесты расчёта на исторических тендерах, инструкцию по развёртыванию и описание методики.", "5–10 чел-дней IT"],
    ["Этап 3", "Эксплуатация в IT", "IT переносит сервис в управляемую среду, отвечает за доступность, резервные копии и доступы, заводит на help desk.", "постоянно"],
  ];
  const cw = 3.96, gap = 0.19;
  stages.forEach(([lbl, t, d, term], i) => {
    const x = M + i * (cw + gap);
    card(s, x, 1.85, cw, 3.9, i === 0 ? C.petrol : C.darkSoft);
    s.addText(lbl, { isTextBox: true, x: x + 0.3, y: 2.08, w: cw - 0.6, h: 0.3,
      fontFace: F.body, fontSize: 11.5, bold: true, charSpacing: 2, color: C.mint, margin: 0 });
    s.addText(t, { isTextBox: true, x: x + 0.3, y: 2.42, w: cw - 0.6, h: 0.4,
      fontFace: F.head, fontSize: 20, bold: true, color: C.white, margin: 0 });
    s.addText(d, { isTextBox: true, x: x + 0.3, y: 2.95, w: cw - 0.6, h: 2.0,
      fontFace: F.body, fontSize: 13, color: "C7DEDD", lineSpacing: 19, margin: 0, valign: "top" });
    s.addText(term, { isTextBox: true, x: x + 0.3, y: 5.1, w: cw - 0.6, h: 0.4,
      fontFace: F.body, fontSize: 13, bold: true, color: C.white, margin: 0 });
  });
  s.addText("Мы отвечаем за то, что сервис правильно считает и закрывает потребность бизнеса. IT отвечает за то, что он работает и поддерживается. Инфраструктуры и дежурства в управлении e-commerce нет, и держать их у себя я не предлагаю.", {
    isTextBox: true, x: M, y: 6.0, w: W, h: 0.7,
    fontFace: F.body, fontSize: 14, color: C.mint, lineSpacing: 21, margin: 0 });
}

// ─────────────────────────────────────────── 15. передача
{
  const s = slide();
  heading(s, "ПЕРЕДАЧА", "Что именно получает IT");
  const arts = [
    ["Исходный код", "в репозиторий IT, без внешних зависимостей за пределами согласованного стека"],
    ["Схема данных", "структура базы и миграции"],
    ["Автотесты расчёта", "прогон на прошлых тендерах: если формула поедет, тест это покажет"],
    ["Инструкция по развёртыванию", "как поднять сервис с нуля"],
    ["Описание методики", "откуда берётся каждая цифра в себестоимости"],
  ];
  let y = 1.8;
  arts.forEach(([t, d]) => {
    card(s, M, y, 7.2, 0.78);
    s.addText(t, { isTextBox: true, x: M + 0.28, y: y + 0.11, w: 6.6, h: 0.3,
      fontFace: F.body, fontSize: 14, bold: true, color: C.ink, margin: 0 });
    s.addText(d, { isTextBox: true, x: M + 0.28, y: y + 0.42, w: 6.6, h: 0.3,
      fontFace: F.body, fontSize: 12, color: C.muted, margin: 0 });
    y += 0.9;
  });
  card(s, 8.3, 1.8, 4.4, 3.4, C.petrol);
  s.addText("Условие, без которого схема не работает", {
    isTextBox: true, x: 8.6, y: 2.05, w: 3.8, h: 0.7,
    fontFace: F.head, fontSize: 19, bold: true, color: C.white, margin: 0 });
  s.addText(
    "Технологический стек согласуем с IT до начала разработки.\n\n" +
    "Тогда на приёмке обсуждается качество кода, а не выбор языка и базы. " +
    "Если IT предложит свой стек, мы его примем.",
    { isTextBox: true, x: 8.6, y: 2.85, w: 3.8, h: 2.1,
      fontFace: F.body, fontSize: 13.5, color: "D9EAE9", lineSpacing: 20, margin: 0 });
  s.addText("Развитие функциональности остаётся за управлением e-commerce: новый трейдер, изменение методики, новые справочники. Поддержка в IT — это доступность и инциденты.", {
    isTextBox: true, x: M, y: 6.35, w: W, h: 0.5,
    fontFace: F.body, fontSize: 13.5, color: C.ink, lineSpacing: 20, margin: 0 });
}

// ─────────────────────────────────────────── 16. оценка
{
  const s = slide();
  heading(s, "ОЦЕНКА", "Сроки по этапам");
  const rows = [
    [{ text: "Этап работ" }, { text: "Что входит" }, { text: "Чел-дни" }],
    ["Справочники и расчётное ядро", "условия девяти трейдеров, схемы оплаты, ставки сборов, курсы, сверка на прошлых тендерах", "4–5"],
    ["Приём и разбор ответов", "разбор файлов, нормализация значений, отчёт об ошибках трейдеру", "3–4"],
    ["Telegram", "бот и форма трейдера с проверкой полей", "4–5"],
    ["Кабинет менеджера", "потребность, мониторинг ответов, сводная, распределение объёма, архив", "4–5"],
    ["Проверка нотификаций ЕАЭС", "разведка реестра, запрос в двух раскладках, блокировка позиции", "2–6"],
    ["Развёртывание и обкатка", "запуск, два живых тендера, доработки, инструкции", "5–8"],
    [{ text: "Разработка у нас", options: { bold: true } }, { text: "итого по этапам выше", options: { bold: true } }, { text: "22–33", options: { bold: true } }],
    ["Передача и приёмка в IT", "перенос в контур, документация, help desk", "5–10"],
  ];
  s.addTable(rows, {
    x: M, y: 1.75, w: W, colW: [3.5, 6.6, 1.99],
    fontFace: F.body, fontSize: 12.5, color: C.ink,
    border: { type: "solid", pt: 0.75, color: C.line },
    fill: { color: C.white }, rowH: 0.42, valign: "middle", margin: [4, 8, 4, 8],
  });
  card(s, M, 5.95, 5.9, 0.95, C.tint);
  s.addText("До промышленной эксплуатации: 2,5–3 месяца", {
    isTextBox: true, x: M + 0.3, y: 6.12, w: 5.3, h: 0.6,
    fontFace: F.head, fontSize: 18, bold: true, color: C.petrol, margin: 0 });
  s.addText("Календарь упирается не в скорость разработки, а в обкатку на живых тендерах и в согласования. Для сравнения: передача задачи в IT целиком — 70–85 чел-дней и 2–2,5 месяца только на разработку, без учёта погружения в предметную область.", {
    isTextBox: true, x: 6.9, y: 5.9, w: 5.4, h: 0.9,
    fontFace: F.body, fontSize: 12, color: C.muted, lineSpacing: 17, margin: 0, valign: "top" });
}

// ─────────────────────────────────────────── 17. риски
{
  const s = slide();
  heading(s, "РИСКИ", "Где оценка может поехать");
  const rows = [
    [{ text: "Риск" }, { text: "В чём он" }, { text: "Что делаем" }],
    ["Реестр ЕАЭС", "неизвестно, отдаёт ли портал данные машине", "день разведки в самом начале. Если не отдаёт, плюс 4 дня и поддержка при изменениях портала"],
    ["Дисциплина трейдеров", "если модель и нотификацию не заполняют, проверять нечего", "делаем поля обязательными в форме и фиксируем это в договорённостях"],
    ["Ставки сборов", "множитель 1,24 в текущем файле это компромисс, а не выведенная величина", "запрашиваем ставки у финансовой службы, иначе расчёт нечем принимать"],
    ["Курс инвестинга", "публичного источника нет, трейдеры оперируют курсом на 11:00", "менеджер вводит его раз в день, бот напоминает. Это решение, а не недоработка"],
    ["Инфраструктура", "нужны сервер, токен бота и права администратора в девяти чатах", "запрашиваем сразу, без этого не стартует ни один вариант"],
  ];
  s.addTable(rows, {
    x: M, y: 1.8, w: W, colW: [2.4, 4.6, 5.09],
    fontFace: F.body, fontSize: 12.5, color: C.ink,
    border: { type: "solid", pt: 0.75, color: C.line },
    fill: { color: C.white }, rowH: 0.72, valign: "middle", margin: [6, 8, 6, 8],
  });
}

// ─────────────────────────────────────────── 18. решения
{
  const s = slide({ dark: true });
  heading(s, "РЕШЕНИЕ", "Что я прошу решить", { dark: true });
  const asks = [
    ["1", "Схему работы", "Разработка в управлении e-commerce, поддержка в IT."],
    ["2", "Согласование стека", "Поручить IT согласовать с нами технологический стек до старта работ."],
    ["3", "Ставки сборов", "Запросить у финансовой службы пошлину по категориям, НДС, утильсбор и гарантийный сбор."],
    ["4", "Доступы", "Сервер на время разработки и токен Telegram-бота."],
  ];
  let y = 1.9;
  asks.forEach(([n, t, d]) => {
    card(s, M, y, 8.9, 0.95, C.darkSoft);
    s.addShape(pres.ShapeType.ellipse, { x: M + 0.3, y: y + 0.24, w: 0.47, h: 0.47,
      fill: { color: C.mint }, line: { color: C.mint, width: 0 } });
    s.addText(n, { isTextBox: true, x: M + 0.3, y: y + 0.32, w: 0.47, h: 0.32,
      align: "center", fontFace: F.body, fontSize: 14, bold: true, color: C.dark, margin: 0 });
    s.addText(t, { isTextBox: true, x: M + 1.0, y: y + 0.16, w: 7.6, h: 0.32,
      fontFace: F.head, fontSize: 17, bold: true, color: C.white, margin: 0 });
    s.addText(d, { isTextBox: true, x: M + 1.0, y: y + 0.52, w: 7.6, h: 0.32,
      fontFace: F.body, fontSize: 13, color: "C7DEDD", margin: 0 });
    y += 1.06;
  });
  s.addText("Пункты 3 и 4 нужны к концу первого этапа и старт не блокируют.", {
    isTextBox: true, x: M, y: 6.3, w: 8.9, h: 0.35,
    fontFace: F.body, fontSize: 14, color: C.mint, margin: 0 });

  card(s, 9.9, 1.9, 2.81, 4.2, C.petrol);
  s.addText("Что будет\nдальше", { isTextBox: true, x: 10.2, y: 2.12, w: 2.3, h: 0.9,
    fontFace: F.head, fontSize: 20, bold: true, color: C.white, margin: 0 });
  s.addText([
    { text: "Собираю справочник трейдеров и расчётное ядро", options: { bullet: true, breakLine: true } },
    { text: "Прогоняю его на тендере от 25.06 и показываю расхождения", options: { bullet: true, breakLine: true } },
    { text: "После этого начинаем сборку сервиса", options: { bullet: true } },
  ], { isTextBox: true, x: 10.2, y: 3.15, w: 2.3, h: 2.7,
       fontFace: F.body, fontSize: 12.5, color: "D9EAE9", lineSpacing: 17, paraSpaceAfter: 8, margin: 0 });
  s.addNotes("Первый этап можно начинать сразу: сверка расчёта делается на тех ставках, которые фактически применялись в прошлых тендерах.");
}

pres.writeFile({ fileName: OUT }).then(() => console.log("written", OUT));

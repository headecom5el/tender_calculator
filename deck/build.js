// Generates the tender calculator proposal deck. 4 slides, plain style.
// Run: node deck/build.js
const pptxgen = require("pptxgenjs");
const path = require("path");

const IMG = path.join(__dirname, "..", "docs", "img");
const OUT = path.join(__dirname, "..", "docs", "tendernyy-kalkulyator.pptx");

const C = { ink: "1A1A1A", muted: "5A5A5A", accent: "14595B", line: "BFBFBF", tint: "F0F3F3", white: "FFFFFF" };
const F = { head: "Cambria", body: "Calibri" };
const M = 0.7;
const W = 13.33 - M * 2;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Управление e-commerce, ЗАО ПАТИО";
pres.title = "Тендерный калькулятор";

let n = 0;
function slide(title, sub) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  n += 1;
  s.addText(title, { isTextBox: true, x: M, y: 0.42, w: W, h: 0.5,
    fontFace: F.head, fontSize: 26, bold: true, color: C.ink, margin: 0 });
  if (sub) {
    s.addText(sub, { isTextBox: true, x: M, y: 0.95, w: W, h: 0.3,
      fontFace: F.body, fontSize: 13, color: C.muted, margin: 0 });
  }
  s.addText(`${n} / 4`, { isTextBox: true, x: 11.9, y: 6.95, w: 0.73, h: 0.3, align: "right",
    fontFace: F.body, fontSize: 10, color: C.line, margin: 0 });
  return s;
}
function txt(s, t, o) {
  s.addText(t, Object.assign({ isTextBox: true, fontFace: F.body, fontSize: 14,
    color: C.ink, lineSpacing: 20, margin: 0, valign: "top" }, o));
}
function tbl(s, rows, o) {
  s.addTable(rows, Object.assign({ x: M, fontFace: F.body, fontSize: 12, color: C.ink,
    border: { type: "solid", pt: 0.75, color: C.line }, fill: { color: C.white },
    valign: "middle", margin: [4, 7, 4, 7] }, o));
}

// 1 ─ проблема
{
  const s = slide("Тендерный калькулятор", "Артём Боровский, управление e-commerce, 7 сентября 2026");

  txt(s, "Сейчас", { x: M, y: 1.5, w: 5.5, h: 0.3, fontSize: 15, bold: true });
  txt(s, "Заявку на тендер КМ вручную рассылает в девять чатов в Telegram, вручную сводит девять " +
        "ответов в одну таблицу, вручную ищет нотификации в реестре ЕАЭС и считает себестоимость " +
        "девятью разными формулами. У каждого трейдера свой источник курса, своя надбавка, своя логистика.",
    { x: M, y: 1.85, w: 6.0, h: 1.5 });
  txt(s, "Один тендер: 17 позиций, 1 730 штук, 1,35 млн долларов. Заявка от 04.08.2026.",
    { x: M, y: 3.45, w: 6.0, h: 0.4, bold: true });

  txt(s, "Почему это риск", { x: 7.3, y: 1.5, w: 5.3, h: 0.3, fontSize: 15, bold: true });
  txt(s, "Разрыв между первым и вторым предложением в реальном расчёте — 0,07 %. " +
        "Победителя определяют доли процента.",
    { x: 7.3, y: 1.85, w: 5.3, h: 0.8 });
  s.addShape(pres.ShapeType.rect, { x: 7.3, y: 2.75, w: 5.3, h: 1.55,
    fill: { color: C.tint }, line: { color: C.tint, width: 0 } });
  txt(s, "В файле от 25.06 надбавка 0,5 % по одному из трейдеров учтена дважды: и в курсе, и в цене. " +
        "Это 15 рублей на единице при разрыве в 2 рубля. На позиции в 320 штук — 4 800 рублей " +
        "и другой победитель тендера.",
    { x: 7.55, y: 2.95, w: 4.8, h: 1.2, fontSize: 13 });

  txt(s, "Плюс к этому", { x: M, y: 4.35, w: 5.5, h: 0.3, fontSize: 15, bold: true });
  const more = [
    ["Нотификации не заполняются", "ни в одном полученном файле, проверять по факту нечего"],
    ["Файлы приходят грязными", "цена текстом «$825», в итоговой цене «$1425 orange, $1450 silver»"],
    ["В шаблоне старая формула", "колонка «Статус позиции», внутри сложение от прежней версии"],
    ["Торг многораундовый", "в форме от 04.08 два столбца таргета, в ТЗ описан разовый сбор"],
  ];
  let y = 4.75;
  more.forEach(([t, d]) => {
    txt(s, t, { x: M, y, w: 3.3, h: 0.3, fontSize: 13, bold: true });
    txt(s, d, { x: M + 3.4, y, w: W - 3.4, h: 0.3, fontSize: 13, color: C.muted });
    y += 0.42;
  });
  txt(s, "Аккуратностью это не лечится. Лечится системой, которая не принимает неправильные данные.",
    { x: M, y: 6.55, w: W, h: 0.35, fontSize: 13.5, italic: true, color: C.accent });
  s.addNotes("Задача поставлена 17 августа. Цифры из заявки 04.08.2026 и сводной 25.06.2026.");
}

// 2 ─ решение
{
  const s = slide("Что предлагаю сделать");
  txt(s, "Трейдер остаётся в своём чате в Telegram, бот сам отправляет запрос и забирает ответ. " +
        "КМ получает готовую сводную: себестоимость в рублях, маржа к целевой РРЦ, покрытие, сроки. " +
        "Позиция без нотификации блокируется, каждый расчёт сохраняется с курсами на дату.",
    { x: M, y: 1.05, w: W, h: 0.7, fontSize: 14 });
  s.addImage({ path: path.join(IMG, "05-diagram.png"), x: 2.0, y: 1.95, w: 9.3, h: 4.38 });
  txt(s, "Пунктиром — вторая очередь. До неё потребность и РРЦ вводятся вручную. " +
        "Стек: Python, Telegram, Vue, PostgreSQL, Docker.",
    { x: M, y: 6.5, w: W, h: 0.35, fontSize: 13, color: C.muted });
}

// 3 ─ макеты
{
  const s = slide("Как это выглядит", "Макеты. Цифры посчитаны настоящей формулой на курсах из файла 25.06");
  s.addImage({ path: path.join(IMG, "06-kabinet-win.png"), x: M, y: 1.4, w: 8.3, h: 5.21 });
  s.addImage({ path: path.join(IMG, "04-phone.png"), x: 9.3, y: 1.4, w: 1.95, h: 4.67 });
  txt(s, "Кабинет КМ: все предложения по позиции, снизу разложение себестоимости по шагам. " +
        "HuaXun заблокирован, нотификация не найдена.",
    { x: M, y: 6.68, w: 8.3, h: 0.35, fontSize: 12.5, color: C.muted });
  txt(s, "Форма трейдера в Telegram. Поля с типами: «1425 orange» отклоняется сразу.",
    { x: 9.3, y: 6.2, w: 3.3, h: 0.6, fontSize: 12.5, color: C.muted });
}

// 4 ─ трудозатраты и решения
{
  const s = slide("Трудозатраты, кто делает, что прошу решить");

  const rows = [
    [{ text: "Работа" }, { text: "Кто" }, { text: "Чел-дни" }],
    ["Справочники и расчёт себестоимости, сверка на прошлых тендерах", "мы", "4–5"],
    ["Приём и разбор ответов трейдеров", "мы", "3–4"],
    ["Бот и форма в Telegram", "мы", "4–5"],
    ["Кабинет менеджера", "мы", "4–5"],
    ["Проверка нотификаций в реестре ЕАЭС", "мы", "2–6"],
    ["Запуск и обкатка на двух живых тендерах", "мы", "5–8"],
    [{ text: "Итого разработка", options: { bold: true } }, { text: "мы", options: { bold: true } }, { text: "22–33", options: { bold: true } }],
    ["Передача, приёмка, постановка на поддержку", "IT", "5–10"],
    [{ text: "Календарь до промышленной эксплуатации", options: { bold: true } }, { text: "", options: {} }, { text: "2,5–3 мес", options: { bold: true } }],
  ];
  tbl(s, rows, { y: 1.15, w: 7.5, colW: [5.0, 1.1, 1.4], rowH: 0.4 });
  txt(s, "Если отдавать задачу в IT целиком — 70–85 чел-дней и 2–2,5 месяца только на разработку.",
    { x: M, y: 5.3, w: 7.5, h: 0.5, fontSize: 12.5, color: C.muted });

  txt(s, "Модель работы", { x: 8.5, y: 1.15, w: 4.1, h: 0.3, fontSize: 15, bold: true });
  txt(s, "Разработку делает управление e-commerce. Готовый сервис вместе с исходным кодом, " +
        "автотестами и документацией передаём в IT. IT переносит его в управляемую среду, " +
        "отвечает за доступность и заводит на help desk.\n\n" +
        "Модель себестоимости — знание коммерции. Что 1,5 % это комиссия агента, а 1,25 % конвертация " +
        "рублей в дирхамы, выяснялось в переписке с трейдерами, в ТЗ этого нет.",
    { x: 8.5, y: 1.5, w: 4.1, h: 3.1, fontSize: 13 });

  s.addShape(pres.ShapeType.rect, { x: 8.5, y: 4.95, w: 4.1, h: 0.95,
    fill: { color: C.tint }, line: { color: C.tint, width: 0 } });
  txt(s, "Условие: стек согласуем с IT до начала разработки. Если предложат свой, примем его.",
    { x: 8.7, y: 5.12, w: 3.7, h: 0.7, fontSize: 13, bold: true });

  txt(s, "Прошу решить", { x: M, y: 5.95, w: 5.5, h: 0.3, fontSize: 15, bold: true });
  txt(s, "1.  Схема работы: разработка у нас, поддержка в IT\n" +
        "2.  Поручить IT согласовать с нами стек до старта\n" +
        "3.  Ставки сборов от финансов: пошлина, НДС, утиль, гарантия\n" +
        "4.  Сервер на время разработки и токен Telegram-бота",
    { x: M, y: 6.3, w: 7.5, h: 0.8, fontSize: 12.5, lineSpacing: 16 });
  txt(s, "Пункты 3 и 4 нужны к концу первого этапа, старт не блокируют.",
    { x: 8.5, y: 6.2, w: 4.1, h: 0.5, fontSize: 12.5, color: C.muted });
}

pres.writeFile({ fileName: OUT }).then(() => console.log("written", OUT));

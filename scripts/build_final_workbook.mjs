#!/usr/bin/env node

import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const inputFile = path.resolve(process.argv[2] || "seo-result.json");
const data = JSON.parse(await fs.readFile(inputFile, "utf8"));
const productCard = data.productCard || {};
const characteristicAnalysis = data.characteristicAnalysis || {};
const characteristicItems = characteristicAnalysis.items || [];

function safeFolderName(value) {
  return String(value || "SEO без поискового запроса")
    .trim()
    .replace(/\s+/g, " ")
    .replace(/[\\/:*?"<>|]+/g, "-")
    .replace(/[. ]+$/g, "") || "SEO без поискового запроса";
}

function timestamp(date = new Date()) {
  const pad = (value) => String(value).padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}-${pad(date.getMinutes())}`;
}

async function defaultOutputFile() {
  const queryFolder = safeFolderName(data.brief?.searchQuery);
  const baseDir = path.join(os.homedir(), "Desktop", "Codex", "SEO");
  let resultDir = path.join(baseDir, queryFolder);
  try {
    await fs.access(path.join(resultDir, "Итоговое SEO.xlsx"));
    resultDir = path.join(baseDir, `${queryFolder} — ${timestamp()}`);
  } catch {
    // Первый результат по этому запросу: использовать папку без даты.
  }
  return path.join(resultDir, "Итоговое SEO.xlsx");
}

const outputFile = process.argv[3] ? path.resolve(process.argv[3]) : await defaultOutputFile();

if (!String(productCard.title || "").trim()) {
  throw new Error("Не заполнено финальное название товара");
}
if ([...String(productCard.title)].length > 60) {
  throw new Error(`Название длиннее 60 символов: ${productCard.title}`);
}
if (!String(productCard.description || "").trim()) {
  throw new Error("Описание не заполнено");
}
const descriptionLength = [...String(productCard.description)].length;
const descriptionLimit = Number(productCard.descriptionLimit || 0);
if (descriptionLimit > 0 && descriptionLength > descriptionLimit) {
  throw new Error(`Описание длиннее лимита категории: ${descriptionLength} из ${descriptionLimit} символов`);
}

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Итоговое SEO");
sheet.showGridLines = false;
sheet.freezePanes.freezeRows(2);

const purple = "#7C3AED";
const dark = "#312E81";
const light = "#EDE9FE";
const pale = "#F5F3FF";
const gray = "#6B7280";
const red = "#FEE2E2";
const amber = "#FEF3C7";

let row = 1;

function title(text) {
  const range = sheet.getRange(`A${row}:F${row}`);
  range.merge();
  range.values = [[text]];
  range.format = { fill: dark, font: { bold: true, color: "#FFFFFF", size: 18 }, rowHeight: 34 };
  row += 2;
}

function section(text) {
  const range = sheet.getRange(`A${row}:F${row}`);
  range.merge();
  range.values = [[text]];
  range.format = { fill: purple, font: { bold: true, color: "#FFFFFF", size: 12 }, rowHeight: 24 };
  row += 1;
}

function table(headers, rows, formats = {}) {
  const headerRange = sheet.getRangeByIndexes(row - 1, 0, 1, headers.length);
  headerRange.values = [headers];
  headerRange.format = {
    fill: light,
    font: { bold: true, color: dark },
    wrapText: true,
    borders: { preset: "outside", style: "thin", color: "#C4B5FD" },
  };
  row += 1;
  if (rows.length) {
    const body = sheet.getRangeByIndexes(row - 1, 0, rows.length, headers.length);
    body.values = rows;
    body.format = {
      fill: pale,
      font: { color: "#111827" },
      wrapText: true,
      verticalAlignment: "top",
      borders: { preset: "inside", style: "thin", color: "#E5E7EB" },
    };
    for (const [column, numberFormat] of Object.entries(formats)) {
      sheet.getRangeByIndexes(row - 1, Number(column), rows.length, 1).format.numberFormat = numberFormat;
    }
    row += rows.length;
  } else {
    sheet.getRange(`A${row}`).values = [["Нет данных"]];
    sheet.getRange(`A${row}`).format.font = { italic: true, color: gray };
    row += 1;
  }
  row += 1;
}

function textBlock(label, value) {
  sheet.getRange(`A${row}`).values = [[label]];
  sheet.getRange(`A${row}`).format = { fill: light, font: { bold: true, color: dark } };
  const valueRange = sheet.getRange(`B${row}:F${row}`);
  valueRange.merge();
  valueRange.values = [[String(value || "")]];
  valueRange.format = { fill: pale, wrapText: true, verticalAlignment: "top" };
  valueRange.format.rowHeight = Math.max(28, Math.min(160, 20 + String(value || "").length / 7));
  row += 1;
}

title("Итоговое SEO для Wildberries");

section("1. Техническое задание");
const briefLabels = {
  product: "Товар",
  category: "Категория",
  gender: "Пол",
  color: "Цвет",
  material: "Материал",
  silhouette: "Фасон",
  fit: "Посадка",
  season: "Сезон",
  sizes: "Размеры",
  package: "Комплектация",
  features: "Особенности",
  searchQuery: "Поисковый запрос",
  minRevenue: "Минимальная выручка, ₽",
};
for (const [key, label] of Object.entries(briefLabels)) textBlock(label, data.brief?.[key] ?? "");
row += 1;

section("2. Проверенные артикулы");
table(
  ["Артикул", "Ссылка", "Выручка 30 дней, ₽", "Оценочная", "Статус", "Причина / результат"],
  (data.checkedArticles || []).map((item) => [
    String(item.article || ""),
    item.url || "",
    Number(item.revenue30d || 0),
    item.estimated ? "Да" : "Нет",
    item.status || "",
    `${item.reason || ""}${item.queriesCollected ? " | запросы получены" : " | запросы не получены"}`,
  ]),
  { 2: "#,##0" },
);

section("3. Принятые конкуренты");
table(
  ["Артикул", "Ссылка", "Выручка 30 дней, ₽", "Оценочная", "Почему похож"],
  (data.competitors || []).map((item) => [
    String(item.article || ""),
    item.url || "",
    Number(item.revenue30d || 0),
    item.estimated ? "Да" : "Нет",
    item.similarityReason || "",
  ]),
  { 2: "#,##0" },
);

section("4. Анализ характеристик конкурентов");
textBlock("Полнота перечня", characteristicAnalysis.sourceScope || "Наблюдаемые характеристики конкурентов");
textBlock("Источник схемы категории", characteristicAnalysis.categorySchemaSource || "Полная схема категории не предоставлена");
table(
  ["Характеристика", "У конкурентов", "Значения конкурентов", "Наше значение", "Решение", "Причина / вопрос"],
  characteristicItems.map((item) => [
    item.field || "",
    Number(item.competitorCount || 0),
    Array.isArray(item.competitorValues) ? item.competitorValues.filter(Boolean).join(", ") : String(item.competitorValues || ""),
    item.ourValue || "",
    item.status || "",
    item.status === "Уточнить" ? (item.question || item.reason || "") : (item.reason || ""),
  ]),
  { 1: "#,##0" },
);

section("5. Очищенное SEO-ядро");
table(
  ["Запрос", "Частотность за 30 дней", "У конкурентов", "Использование"],
  (data.semanticCore || []).map((item) => [
    item.query || "",
    Number(item.frequency30d || 0),
    Number(item.competitorCount || 0),
    item.usage || "",
  ]),
  { 1: "#,##0", 2: "#,##0" },
);

section("6. Сомнительные запросы");
const questionableStart = row;
table(
  ["Запрос", "Частотность за 30 дней", "Причина проверки"],
  (data.questionable || []).map((item) => [item.query || "", Number(item.frequency30d || 0), item.reason || ""]),
  { 1: "#,##0" },
);
if ((data.questionable || []).length) sheet.getRange(`A${questionableStart}:C${row - 2}`).format.fill = amber;

section("7. Исключенные запросы");
const excludedStart = row;
table(
  ["Запрос", "Частотность за 30 дней", "Причина исключения"],
  (data.excluded || []).map((item) => [item.query || "", Number(item.frequency30d || 0), item.reason || ""]),
  { 1: "#,##0" },
);
if ((data.excluded || []).length) sheet.getRange(`A${excludedStart}:C${row - 2}`).format.fill = red;

if ((data.notes || []).length) {
  section("8. Примечания");
  for (const note of data.notes) textBlock("Примечание", note);
}

sheet.getRange(`A1:F${row}`).format.font = { name: "Arial" };
sheet.getRange("A:A").format.columnWidth = 24;
sheet.getRange("B:B").format.columnWidth = 38;
sheet.getRange("C:C").format.columnWidth = 24;
sheet.getRange("D:D").format.columnWidth = 20;
sheet.getRange("E:E").format.columnWidth = 35;
sheet.getRange("F:F").format.columnWidth = 18;

const cardSheet = workbook.worksheets.add("Карточка товара");
cardSheet.showGridLines = false;
cardSheet.freezePanes.freezeRows(2);

const cardTitle = cardSheet.getRange("A1:F1");
cardTitle.merge();
cardTitle.values = [["Готовая карточка товара Wildberries"]];
cardTitle.format = { fill: dark, font: { bold: true, color: "#FFFFFF", size: 18 }, rowHeight: 36 };

const nameHeader = cardSheet.getRange("A3:F3");
nameHeader.merge();
nameHeader.values = [["Название товара"]];
nameHeader.format = { fill: purple, font: { bold: true, color: "#FFFFFF", size: 12 }, rowHeight: 24 };

const nameValue = cardSheet.getRange("A4:F5");
nameValue.merge();
nameValue.values = [[String(productCard.title)]];
nameValue.format = { fill: pale, font: { bold: true, color: dark, size: 16 }, wrapText: true, verticalAlignment: "center" };

cardSheet.getRange("A6:B6").values = [["Количество символов", [...String(productCard.title)].length]];
cardSheet.getRange("A6:B6").format = { fill: light, font: { color: dark } };

const charHeader = cardSheet.getRange("A8:F8");
charHeader.merge();
charHeader.values = [["Характеристики"]];
charHeader.format = { fill: purple, font: { bold: true, color: "#FFFFFF", size: 12 }, rowHeight: 24 };

cardSheet.getRange("A9:C9").values = [["Характеристика", "Значение", "Ключи-основания"]];
cardSheet.getRange("A9:C9").format = { fill: light, font: { bold: true, color: dark }, wrapText: true };
const characteristicRows = (productCard.characteristics || []).map((item) => [
  item.field || "",
  item.value || "",
  item.supportingQueries || "",
]);
let cardRow = 10;
if (characteristicRows.length) {
  const range = cardSheet.getRangeByIndexes(cardRow - 1, 0, characteristicRows.length, 3);
  range.values = characteristicRows;
  range.format = { fill: pale, wrapText: true, verticalAlignment: "top", borders: { preset: "inside", style: "thin", color: "#E5E7EB" } };
  cardRow += characteristicRows.length;
} else {
  cardSheet.getRange(`A${cardRow}`).values = [["Нет данных"]];
  cardRow += 1;
}

cardRow += 1;
const clarificationItems = characteristicItems.filter((item) => item.status === "Уточнить");
const clarificationHeader = cardSheet.getRange(`A${cardRow}:F${cardRow}`);
clarificationHeader.merge();
clarificationHeader.values = [["Нужно уточнить"]];
clarificationHeader.format = { fill: "#D97706", font: { bold: true, color: "#FFFFFF", size: 12 }, rowHeight: 24 };
cardRow += 1;
cardSheet.getRange(`A${cardRow}:C${cardRow}`).values = [["Характеристика", "Почему важно", "Вопрос"]];
cardSheet.getRange(`A${cardRow}:C${cardRow}`).format = { fill: amber, font: { bold: true, color: "#92400E" }, wrapText: true };
cardRow += 1;
if (clarificationItems.length) {
  const clarificationRows = clarificationItems.map((item) => [item.field || "", item.reason || "", item.question || "Уточните значение характеристики"]);
  const range = cardSheet.getRangeByIndexes(cardRow - 1, 0, clarificationRows.length, 3);
  range.values = clarificationRows;
  range.format = { fill: "#FFFBEB", wrapText: true, verticalAlignment: "top", borders: { preset: "inside", style: "thin", color: "#FDE68A" } };
  cardRow += clarificationRows.length;
} else {
  cardSheet.getRange(`A${cardRow}:C${cardRow}`).merge();
  cardSheet.getRange(`A${cardRow}:C${cardRow}`).values = [["Все важные характеристики подтверждены"]];
  cardSheet.getRange(`A${cardRow}:C${cardRow}`).format = { fill: "#ECFDF5", font: { color: "#065F46" } };
  cardRow += 1;
}

cardRow += 1;
const descriptionHeader = cardSheet.getRange(`A${cardRow}:F${cardRow}`);
descriptionHeader.merge();
descriptionHeader.values = [["Описание товара"]];
descriptionHeader.format = { fill: purple, font: { bold: true, color: "#FFFFFF", size: 12 }, rowHeight: 24 };
cardRow += 1;

const cardDescription = cardSheet.getRange(`A${cardRow}:F${cardRow + 11}`);
cardDescription.merge();
cardDescription.values = [[String(productCard.description)]];
cardDescription.format = { fill: pale, wrapText: true, verticalAlignment: "top", font: { color: "#111827", size: 11 } };
cardRow += 12;
cardSheet.getRange(`A${cardRow}:C${cardRow}`).values = [[
  "Количество символов в описании",
  descriptionLength,
  descriptionLimit > 0 ? `Лимит категории: ${descriptionLimit}` : "Лимит категории не указан",
]];
cardSheet.getRange(`A${cardRow}:C${cardRow}`).format = { fill: light, font: { color: dark } };
cardRow += 2;

const usedHeader = cardSheet.getRange(`A${cardRow}:F${cardRow}`);
usedHeader.merge();
usedHeader.values = [["Использованные ключевые запросы"]];
usedHeader.format = { fill: light, font: { bold: true, color: dark }, rowHeight: 24 };
cardRow += 1;
const usedQueries = (productCard.usedQueries || []).join(", ");
const usedValue = cardSheet.getRange(`A${cardRow}:F${cardRow + 2}`);
usedValue.merge();
usedValue.values = [[usedQueries]];
usedValue.format = { fill: pale, wrapText: true, verticalAlignment: "top", font: { color: gray } };

cardSheet.getRange(`A1:F${cardRow + 2}`).format.font = { name: "Arial" };
cardSheet.getRange("A:A").format.columnWidth = 28;
cardSheet.getRange("B:B").format.columnWidth = 38;
cardSheet.getRange("C:C").format.columnWidth = 42;
cardSheet.getRange(`D1:F${cardRow + 2}`).format.columnWidth = 18;

await fs.mkdir(path.dirname(outputFile), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputFile);

const preview = await workbook.render({ sheetName: "Итоговое SEO", range: `A1:F${Math.min(row, 100)}`, scale: 1 });
await fs.writeFile(outputFile.replace(/\.xlsx$/i, "-analysis-preview.png"), new Uint8Array(await preview.arrayBuffer()));
const cardPreview = await workbook.render({ sheetName: "Карточка товара", range: `A1:F${cardRow + 2}`, scale: 1 });
await fs.writeFile(outputFile.replace(/\.xlsx$/i, "-card-preview.png"), new Uint8Array(await cardPreview.arrayBuffer()));

console.log(outputFile);

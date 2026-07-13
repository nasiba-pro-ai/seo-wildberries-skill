#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputDir = path.resolve(process.argv[2] || ".");
const outputFile = path.resolve(process.argv[3] || path.join(inputDir, "merged-semantics.json"));

const QUERY_NAMES = ["запрос", "запросы", "поисковый запрос", "ключевая фраза"];
const FREQ_NAMES = [
  "частотность за 30 дней",
  "частота за 30 дней",
  "частотность 30 дней",
  "частотность",
];

function normalizeText(value) {
  return String(value ?? "").trim().replace(/\s+/g, " ").toLowerCase();
}

function parseNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const cleaned = String(value ?? "")
    .replace(/[\u00a0\u202f\s]/g, "")
    .replace(/,/g, ".")
    .replace(/[^\d.-]/g, "");
  const number = Number(cleaned);
  return Number.isFinite(number) ? number : 0;
}

function findColumn(headers, accepted) {
  return headers.findIndex((header) => accepted.some((name) => header === name || header.includes(name)));
}

function findHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 80); i += 1) {
    const headers = rows[i].map(normalizeText);
    if (findColumn(headers, QUERY_NAMES) >= 0 && findColumn(headers, FREQ_NAMES) >= 0) return i;
  }
  return -1;
}

async function readWorkbook(filePath) {
  const input = await FileBlob.load(filePath);
  const workbook = await SpreadsheetFile.importXlsx(input);
  const sheet = workbook.worksheets.getItemAt(0);
  const used = sheet.getUsedRange(true);
  const rows = used?.values || [];
  const headerRow = findHeaderRow(rows);
  if (headerRow < 0) throw new Error("не найдены колонки запроса и частотности за 30 дней");
  const headers = rows[headerRow].map(normalizeText);
  const queryColumn = findColumn(headers, QUERY_NAMES);
  const frequencyColumn = findColumn(headers, FREQ_NAMES);
  const records = [];
  for (const row of rows.slice(headerRow + 1)) {
    const query = normalizeText(row[queryColumn]);
    if (!query) continue;
    records.push({ query, frequency30d: parseNumber(row[frequencyColumn]) });
  }
  return records;
}

const names = (await fs.readdir(inputDir))
  .filter((name) => /\.xlsx$/i.test(name) && !name.startsWith("~$"))
  .sort((a, b) => a.localeCompare(b, "ru"));

if (!names.length) throw new Error(`В папке нет файлов .xlsx: ${inputDir}`);

const merged = new Map();
const files = [];

for (const name of names) {
  const filePath = path.join(inputDir, name);
  try {
    const rows = await readWorkbook(filePath);
    files.push({ file: name, status: "ok", queries: rows.length });
    for (const row of rows) {
      const current = merged.get(row.query) || {
        query: row.query,
        frequency30d: 0,
        competitorCount: 0,
        sources: [],
      };
      current.frequency30d = Math.max(current.frequency30d, row.frequency30d);
      current.competitorCount += 1;
      current.sources.push(name);
      merged.set(row.query, current);
    }
  } catch (error) {
    files.push({ file: name, status: "error", error: String(error.message || error) });
  }
}

const queries = [...merged.values()].sort(
  (a, b) => b.frequency30d - a.frequency30d || b.competitorCount - a.competitorCount || a.query.localeCompare(b.query, "ru"),
);

await fs.writeFile(
  outputFile,
  JSON.stringify(
    {
      generatedAt: new Date().toISOString(),
      inputDirectory: inputDir,
      files,
      successfulFiles: files.filter((file) => file.status === "ok").length,
      uniqueQueries: queries.length,
      queries,
    },
    null,
    2,
  ),
  "utf8",
);

console.log(`Готово: ${queries.length} уникальных запросов из ${files.filter((file) => file.status === "ok").length} файлов`);
console.log(outputFile);

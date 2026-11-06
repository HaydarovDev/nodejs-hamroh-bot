import ExcelJS from "exceljs";
import fs from "node:fs";
import path from "node:path";
import { allApplications, getShifts } from "./db";
import type { Application, AppStatus } from "./db";
import { CATEGORIES } from "./menu";
import { monthLabel, monthOf, monthlyRows, monthsWithData, summarize } from "./stats";
import { tzParts } from "./time";

// Arizalar avtomatik ravishda data/arizalar.xlsx fayliga yoziladi.
//  • "Arizalar" varag'i — har bir ariza bitta qator. Yangi ariza qo'shiladi; status yoki baho o'zgarsa,
//    faqat shu arizaning "Status" va "Baho" kataklari yangilanadi (qolganini, jumladan "Eslatma"ni,
//    qo'lda tahrirlasangiz, o'chib ketmaydi).
//  • "Statistika" varag'i — umumiy va operatorlar kesimi (formulalar bilan).
//  • "Oylik" varag'i — har oy uchun har bir operator: jami, hal qilindi, hal qilinmadi, jarayonda, o'rtacha ball.
//  • Qatorni topish uchun oxirgi ustun — "Ariza ID". Uni yashirishingiz mumkin, lekin o'chirmang.

export const XLSX_FILE = path.join(process.cwd(), "data", "arizalar.xlsx");

const SHEET = "Arizalar";
const STATS = "Statistika";
const MONTHLY = "Oylik";
const HEADERS = [
  "Murojaat mavzusi",
  "Module",
  "Batafsil",
  "Status",
  "Sana",
  "Soat",
  "Eslatma",
  "Murojaat manbasi",
  "Aloqa ma'lumotlari",
  "Operator",
  "Smena",
  "Baho",
  "Ariza ID",
];
const WIDTHS = [26, 22, 48, 16, 12, 13, 30, 18, 20, 14, 24, 8, 10];
const COL = { status: 4, operator: 10, rating: 12, id: 13 };
const MAXROW = 100000;

// Status: Hal qilindi / Hal qilinmadi — foydalanuvchining javobi; Jarayonda — hali tasdiqlanmagan
const STATUS_TEXT: Record<AppStatus, string> = {
  new: "Jarayonda",
  done: "Jarayonda",
  resolved: "Hal qilindi",
  unresolved: "Hal qilinmadi",
};

const OBJECT_NAME = { car: "Avtomobil", home: "Hisob raqami", other: "Mavzu" } as const;

/** Baho ranglari: 1 — qizil, 2 — to'q sariq, 3 — sariq, 4 — och yashil, 5 — yashil */
const RATING_FILL: Record<number, string> = {
  1: "FFF8696B",
  2: "FFFBAA77",
  3: "FFFFEB84",
  4: "FFB1D580",
  5: "FF63BE7B",
};

function paintRating(cell: ExcelJS.Cell, rating: number | undefined) {
  if (!rating) {
    cell.value = null;
    cell.fill = { type: "pattern", pattern: "none" };
    return;
  }
  cell.value = rating;
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: RATING_FILL[Math.round(rating)] } };
  cell.alignment = { horizontal: "center" };
  cell.font = { bold: true };
}

/** "🚗 Mening avtomobilim" → "Mening avtomobilim" */
const stripEmoji = (s: string) => s.replace(/^[^\p{L}\p{N}]+/u, "").trim();

function rowValues(a: Application): (string | number | Date | null)[] {
  const cat = CATEGORIES[a.category];
  const t = tzParts(new Date(a.createdAt));
  const details = [a.object ? `${OBJECT_NAME[cat.objectKind]}: ${a.object}` : "", a.comment]
    .filter(Boolean)
    .join(" | ");
  return [
    a.service,
    stripEmoji(cat.title),
    details,
    STATUS_TEXT[a.status],
    new Date(Date.UTC(t.year, t.month - 1, t.day)), // Sana (O'zbekiston vaqti bo'yicha)
    (t.hour * 3600 + t.minute * 60 + t.second) / 86400, // Soat (kunning ulushi)
    "", // Eslatma — operator o'zi yozadi
    "Telegram",
    a.phone.replace(/\D/g, ""),
    a.operator ?? "",
    a.shift ?? "",
    null, // Baho — pastda bo'yaladi
    a.id,
  ];
}

// ───────────────────────── Statistika va Oylik varaqlari ─────────────────────────

const headStyle = (row: ExcelJS.Row) => {
  row.font = { bold: true };
  row.eachCell((c) => {
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE7EEF7" } };
  });
};

function buildStats(wb: ExcelJS.Workbook, apps: Application[]) {
  const old = wb.getWorksheet(STATS);
  if (old) wb.removeWorksheet(old.id);
  const ws = wb.addWorksheet(STATS);
  ws.columns = [{ width: 28 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 }];

  const R = (col: string) => `${SHEET}!${col}2:${col}${MAXROW}`;
  const all = summarize(apps);

  // Umumiy
  headStyle(ws.addRow(["Umumiy", "Soni"]));
  ws.addRow(["Jami arizalar", { formula: `COUNT(${R("M")})`, result: all.total }]);
  ws.addRow(["✅ Hal qilindi", { formula: `COUNTIF(${R("D")},"Hal qilindi")`, result: all.resolved }]);
  ws.addRow(["❌ Hal qilinmadi", { formula: `COUNTIF(${R("D")},"Hal qilinmadi")`, result: all.unresolved }]);
  ws.addRow(["⏳ Jarayonda", { formula: `COUNTIF(${R("D")},"Jarayonda")`, result: all.inProgress }]);
  const avgRow = ws.addRow([
    "⭐ O'rtacha ball",
    { formula: `IFERROR(AVERAGE(${R("L")}),"—")`, result: all.avg === null ? "—" : Number(all.avg.toFixed(2)) },
  ]);
  avgRow.getCell(2).numFmt = "0.00";
  if (all.avg !== null) paintRating(avgRow.getCell(2), all.avg);
  ws.addRow([]);

  // Operatorlar bo'yicha (barcha vaqt)
  headStyle(
    ws.addRow(["Operator", "Jami arizalar", "✅ Hal qilindi", "❌ Hal qilinmadi", "⏳ Jarayonda", "⭐ O'rtacha ball"])
  );
  const names = [...new Set([...getShifts().map((s) => s.operator), ...apps.map((a) => a.operator ?? "")])].filter(
    (n) => n && n !== "—"
  );
  for (const name of names) {
    const mine = summarize(apps.filter((a) => a.operator === name));
    const r = ws.rowCount + 1;
    const J = R("J");
    const D = R("D");
    const row = ws.addRow([
      name,
      { formula: `COUNTIF(${J},A${r})`, result: mine.total },
      { formula: `COUNTIFS(${J},A${r},${D},"Hal qilindi")`, result: mine.resolved },
      { formula: `COUNTIFS(${J},A${r},${D},"Hal qilinmadi")`, result: mine.unresolved },
      { formula: `COUNTIFS(${J},A${r},${D},"Jarayonda")`, result: mine.inProgress },
      {
        formula: `IFERROR(AVERAGEIF(${J},A${r},${R("L")}),"—")`,
        result: mine.avg === null ? "—" : Number(mine.avg.toFixed(2)),
      },
    ]);
    const c = row.getCell(6);
    c.numFmt = "0.00";
    if (mine.avg !== null) paintRating(c, mine.avg);
  }
}

/** Oylik: har oy uchun har bir operator bo'yicha qator (qiymatlar bot tomonidan har safar qayta hisoblanadi) */
function buildMonthly(wb: ExcelJS.Workbook, apps: Application[]) {
  const old = wb.getWorksheet(MONTHLY);
  if (old) wb.removeWorksheet(old.id);
  const ws = wb.addWorksheet(MONTHLY, { views: [{ state: "frozen", ySplit: 1 }] });
  ws.columns = [{ width: 18 }, { width: 18 }, { width: 14 }, { width: 16 }, { width: 18 }, { width: 14 }, { width: 16 }];
  headStyle(ws.addRow(["Oy", "Operator", "Jami arizalar", "✅ Hal qilindi", "❌ Hal qilinmadi", "⏳ Jarayonda", "⭐ O'rtacha ball"]));

  const addRow = (label: string, name: string, r: ReturnType<typeof summarize>, bold: boolean) => {
    const row = ws.addRow([label, name, r.total, r.resolved, r.unresolved, r.inProgress, r.avg === null ? "—" : r.avg]);
    const c = row.getCell(7);
    c.numFmt = "0.00";
    c.alignment = { horizontal: "center" };
    if (r.avg !== null) paintRating(c, r.avg);
    if (bold) {
      row.eachCell((cell, colNumber) => {
        cell.font = { bold: true };
        if (colNumber !== 7) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF3F3F3" } };
      });
    }
  };

  for (const ym of monthsWithData(apps)) {
    const label = monthLabel(ym);
    for (const r of monthlyRows(apps, ym)) addRow(label, r.operator, r, false);
    addRow(label, "Jami", summarize(apps.filter((a) => monthOf(a) === ym)), true);
  }
}

// ───────────────────────── Yozish ─────────────────────────

async function writeOnce(apps: Application[]) {
  fs.mkdirSync(path.dirname(XLSX_FILE), { recursive: true });

  let wb = new ExcelJS.Workbook();
  let batch = apps;
  if (fs.existsSync(XLSX_FILE)) {
    await wb.xlsx.readFile(XLSX_FILE);
    const old = wb.getWorksheet(SHEET);
    // eski tuzilmadagi fayl ("Baho" ustuni yo'q) — zaxiraga olib, yangidan quramiz
    if (old && String(old.getRow(1).getCell(12).value) === "Ariza ID") {
      const backup = XLSX_FILE.replace(/\.xlsx$/, `-eski-${Date.now()}.xlsx`);
      fs.renameSync(XLSX_FILE, backup);
      wb = new ExcelJS.Workbook();
      batch = allApplications();
    }
  }

  let ws = wb.getWorksheet(SHEET);
  if (!ws) {
    ws = wb.addWorksheet(SHEET, { views: [{ state: "frozen", ySplit: 1 }] });
    ws.columns = HEADERS.map((header, i) => ({ header, key: `c${i}`, width: WIDTHS[i] }));
    const head = ws.getRow(1);
    head.font = { bold: true };
    head.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE7EEF7" } };
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: HEADERS.length } };
  }

  // Ariza ID → qator
  const rowById = new Map<number, ExcelJS.Row>();
  ws.eachRow((row, n) => {
    if (n === 1) return;
    const v = Number(row.getCell(COL.id).value);
    if (v) rowById.set(v, row);
  });

  for (const a of batch) {
    const existing = rowById.get(a.id);
    if (existing) {
      existing.getCell(COL.status).value = STATUS_TEXT[a.status];
      paintRating(existing.getCell(COL.rating), a.rating);
      existing.commit();
      continue;
    }
    const row = ws.addRow(rowValues(a));
    row.getCell(5).numFmt = "m/d/yyyy";
    row.getCell(6).numFmt = "h:mm:ss AM/PM";
    row.getCell(9).numFmt = "@";
    paintRating(row.getCell(COL.rating), a.rating);
    row.commit();
  }

  buildStats(wb, allApplications());
  buildMonthly(wb, allApplications());
  wb.calcProperties.fullCalcOnLoad = true; // formulalar fayl ochilganda qayta hisoblansin

  const tmp = XLSX_FILE + ".tmp";
  await wb.xlsx.writeFile(tmp);
  fs.renameSync(tmp, XLSX_FILE);
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

async function withRetry(fn: () => Promise<void>) {
  for (let attempt = 1; ; attempt++) {
    try {
      await fn();
      return;
    } catch (e) {
      const code = (e as NodeJS.ErrnoException).code;
      const locked = code === "EBUSY" || code === "EPERM" || code === "EACCES";
      if (!locked || attempt >= 3) throw e;
      console.warn("⚠️ arizalar.xlsx band (Excelda ochiq bo'lishi mumkin). Qayta urinilmoqda...");
      await sleep(3000);
    }
  }
}

// Yozuvlar navbat bilan bajariladi — fayl buzilib qolmasligi uchun
let queue: Promise<void> = Promise.resolve();

function enqueue(apps: Application[]): Promise<void> {
  queue = queue
    .then(() => withRetry(() => writeOnce(apps)))
    .catch((e) =>
      console.error(
        "❌ Excelga yozib bo'lmadi (faylni Excelda yopib qo'ying, keyingi o'zgarishda yoki qayta ishga tushganda yoziladi):",
        (e as Error).message
      )
    );
  return queue;
}

export const syncApplication = (a: Application) => enqueue([a]);
export const syncAll = (apps: Application[]) => enqueue(apps);

import type { Application } from "./db";
import { tzParts } from "./time";

// Oylik statistika uchun umumiy hisob-kitob (Excel va bot ikkalasi shuni ishlatadi).
//   Hal qilindi  — foydalanuvchi "Muammo hal qilindi" deb tasdiqlagan
//   Hal qilinmadi — foydalanuvchi "Muammo hal qilinmadi" degan (yozishma davom etadi)
//   Jarayonda    — hali ko'rib chiqilmoqda yoki operator yopgan, lekin foydalanuvchi tasdiqlamagan

export const MONTH_UZ = [
  "yanvar", "fevral", "mart", "aprel", "may", "iyun",
  "iyul", "avgust", "sentabr", "oktabr", "noyabr", "dekabr",
];

export interface StatRow {
  operator: string;
  total: number;
  resolved: number;
  unresolved: number;
  inProgress: number;
  /** O'rtacha ball (faqat baho qo'yilgan arizalar bo'yicha), baho yo'q bo'lsa — null */
  avg: number | null;
}

/** 202611 ko'rinishidagi oy kaliti (O'zbekiston vaqti bo'yicha) */
export function monthOf(a: Application): number {
  const p = tzParts(new Date(a.createdAt));
  return p.year * 100 + p.month;
}

export const currentMonth = (): number => {
  const p = tzParts(new Date());
  return p.year * 100 + p.month;
};

export const monthLabel = (ym: number) => `${Math.floor(ym / 100)} ${MONTH_UZ[(ym % 100) - 1]}`;

export function shiftMonth(ym: number, delta: number): number {
  let y = Math.floor(ym / 100);
  let m = (ym % 100) + delta;
  while (m < 1) { m += 12; y--; }
  while (m > 12) { m -= 12; y++; }
  return y * 100 + m;
}

export function summarize(apps: Application[]): Omit<StatRow, "operator"> {
  const rated = apps.filter((a) => a.rating);
  return {
    total: apps.length,
    resolved: apps.filter((a) => a.status === "resolved").length,
    unresolved: apps.filter((a) => a.status === "unresolved").length,
    inProgress: apps.filter((a) => a.status === "new" || a.status === "done").length,
    avg: rated.length ? rated.reduce((x, a) => x + (a.rating ?? 0), 0) / rated.length : null,
  };
}

/** Berilgan oy uchun har bir operator bo'yicha qator (arizalari ko'pi birinchi) */
export function monthlyRows(apps: Application[], ym: number): StatRow[] {
  const byOp = new Map<string, Application[]>();
  for (const a of apps) {
    if (monthOf(a) !== ym) continue;
    const op = a.operator && a.operator !== "" ? a.operator : "—";
    byOp.set(op, [...(byOp.get(op) ?? []), a]);
  }
  return [...byOp.entries()]
    .map(([operator, list]) => ({ operator, ...summarize(list) }))
    .sort((x, y) => y.total - x.total);
}

/** Ma'lumot bor oylar, yangisi birinchi */
export function monthsWithData(apps: Application[]): number[] {
  return [...new Set(apps.map(monthOf))].sort((a, b) => b - a);
}

// Vaqt bilan ishlash: server qaysi soat mintaqasida bo'lmasin, hisob O'zbekiston vaqti bo'yicha yuritiladi.
// Boshqa mintaqa kerak bo'lsa, .env da BOT_TIMEZONE=Asia/Samarkand kabi yozing.

export const TIMEZONE = process.env.BOT_TIMEZONE?.trim() || "Asia/Tashkent";

export interface TzParts {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
}

const fmt = new Intl.DateTimeFormat("en-GB", {
  timeZone: TIMEZONE,
  year: "numeric",
  month: "2-digit",
  day: "2-digit",
  hour: "2-digit",
  minute: "2-digit",
  second: "2-digit",
  hourCycle: "h23",
});

export function tzParts(date: Date): TzParts {
  const p: Record<string, string> = {};
  for (const part of fmt.formatToParts(date)) p[part.type] = part.value;
  return {
    year: Number(p.year),
    month: Number(p.month),
    day: Number(p.day),
    hour: Number(p.hour),
    minute: Number(p.minute),
    second: Number(p.second),
  };
}

import fs from "node:fs";
import path from "node:path";
import type { CategoryKey } from "./menu";
import { tzParts } from "./time";

// Oddiy JSON fayl asosidagi saqlash. Production uchun PostgreSQL/MongoDB ga almashtirish mumkin.

export interface User {
  id: number;
  username?: string;
  name?: string;
  phone?: string;
  cars: string[];
  homes: string[];
  registeredAt?: string;
  /** Guruhdagi (forum) shu foydalanuvchi uchun ochilgan topic ID'si */
  topicId?: number;
}

/**
 * new        — ariza qabul qilindi
 * done       — operator yopdi, foydalanuvchi bahosi kutilmoqda
 * resolved   — "Muammo hal qilindi"
 * unresolved — "Muammo hal qilinmadi"
 */
export type AppStatus = "new" | "done" | "resolved" | "unresolved";

export interface Application {
  id: number;
  userId: number;
  category: CategoryKey;
  service: string;
  object: string;
  phone: string;
  comment: string;
  status: AppStatus;
  createdAt: string;
  /** Ariza qoldirilgan paytdagi smena operatori (keyin smena o'zgarsa ham o'zgarmaydi) */
  operator?: string;
  shift?: string;
  /** Foydalanuvchining operator ishiga bergan bahosi (1–5) */
  rating?: number;
}

export interface Shift {
  operator: string;
  /** "9:00" ko'rinishida */
  start: string;
  end: string;
}

interface DbShape {
  lastAppId: number;
  users: Record<string, User>;
  applications: Application[];
  /** Hammaga xabar yuboriladigan topic ID'si */
  broadcastTopicId?: number;
  shifts?: Shift[];
}

const DATA_DIR = path.join(process.cwd(), "data");
const DB_FILE = path.join(DATA_DIR, "db.json");

function load(): DbShape {
  try {
    return JSON.parse(fs.readFileSync(DB_FILE, "utf8")) as DbShape;
  } catch {
    return { lastAppId: 0, users: {}, applications: [] };
  }
}

const db = load();

function save() {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  const tmp = DB_FILE + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(db, null, 2));
  fs.renameSync(tmp, DB_FILE);
}

export function getUser(id: number, username?: string): User {
  const key = String(id);
  let user = db.users[key];
  if (!user) {
    user = { id, username, cars: [], homes: [] };
    db.users[key] = user;
    save();
  } else if (username && user.username !== username) {
    user.username = username;
    save();
  }
  return user;
}

export function updateUser(id: number, patch: Partial<User>): User {
  const user = getUser(id);
  Object.assign(user, patch);
  save();
  return user;
}

/** Avtomobil raqami / hisob raqamini eslab qoladi — faqat eng oxirgisini (oldingisi almashadi) */
export function addUserItem(id: number, kind: "car" | "home", value: string) {
  const user = getUser(id);
  if (kind === "car") user.cars = [value];
  else user.homes = [value];
  save();
}

export function createApplication(
  data: Omit<Application, "id" | "status" | "createdAt" | "operator" | "shift" | "rating">
): Application {
  const now = new Date();
  const current = shiftFor(now);
  const application: Application = {
    ...data,
    id: ++db.lastAppId,
    status: "new",
    createdAt: now.toISOString(),
    operator: current.operator,
    shift: current.label,
  };
  db.applications.push(application);
  save();
  return application;
}

export function getApplication(id: number): Application | undefined {
  return db.applications.find((a) => a.id === id);
}

export function setStatus(id: number, status: AppStatus): Application | undefined {
  const a = getApplication(id);
  if (a) {
    a.status = status;
    save();
  }
  return a;
}

export function userApplications(userId: number): Application[] {
  return db.applications.filter((a) => a.userId === userId);
}

/** Hali yopilmagan (resolved bo'lmagan) eng oxirgi ariza — jonli yozishma shu ariza bo'yicha ketadi */
export function activeApplication(userId: number): Application | undefined {
  const list = userApplications(userId).filter((a) => a.status !== "resolved");
  return list[list.length - 1];
}

export function findUserByTopic(topicId: number): User | undefined {
  return Object.values(db.users).find((u) => u.topicId === topicId);
}

export function getBroadcastTopic(): number | undefined {
  return db.broadcastTopicId;
}

export function setBroadcastTopic(id: number) {
  db.broadcastTopicId = id;
  save();
}

export function allUserIds(): number[] {
  return Object.values(db.users).map((u) => u.id);
}

export function allApplications(): Application[] {
  return db.applications;
}

export function setRating(id: number, rating: number): Application | undefined {
  const a = getApplication(id);
  if (a) {
    a.rating = rating;
    save();
  }
  return a;
}

// ───────────────────────── Smenalar (admin tahrirlaydi) ─────────────────────────

const DEFAULT_SHIFTS: Shift[] = [
  { operator: "Asadbek", start: "9:00", end: "18:00" },
  { operator: "Abdug'ani", start: "18:00", end: "2:00" },
  { operator: "Bonu", start: "2:00", end: "9:00" },
];

export function getShifts(): Shift[] {
  if (!db.shifts) {
    db.shifts = DEFAULT_SHIFTS.map((x) => ({ ...x }));
    save();
  }
  return db.shifts;
}

export function updateShift(index: number, patch: Partial<Shift>) {
  const shifts = getShifts();
  if (!shifts[index]) return;
  Object.assign(shifts[index], patch);
  save();
}

export const shiftLabel = (s: Shift, index: number) => `${index + 1}-smena ${s.start}-${s.end}`;

const toMin = (t: string) => {
  const [h, m] = t.split(":").map(Number);
  return h * 60 + m;
};

/** Berilgan vaqtga to'g'ri keladigan smena (yarim tundan o'tadigan smenalar ham ishlaydi) */
export function shiftFor(date: Date): { operator: string; label: string } {
  const p = tzParts(date);
  const now = p.hour * 60 + p.minute;
  const shifts = getShifts();
  for (let i = 0; i < shifts.length; i++) {
    const a = toMin(shifts[i].start);
    const b = toMin(shifts[i].end);
    const inside = a < b ? now >= a && now < b : a > b ? now >= a || now < b : false;
    if (inside) return { operator: shifts[i].operator, label: shiftLabel(shifts[i], i) };
  }
  return { operator: "—", label: "" };
}

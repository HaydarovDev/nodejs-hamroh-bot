// Struktura: Start -> bo'lim -> xizmat -> ariza qadamlari -> operator ishini baholash

export type CategoryKey = "car" | "home" | "package" | "navigator";

/**
 * Ariza qadamlari. Tartib `flow` massivida:
 *   object  — avtomobil raqami yoki hisob raqami
 *   phone   — telefon raqam (har doim ixtiyoriy: "O'tkazib yuborish" tugmasi bor)
 *   comment — muammo / savol matni
 */
export type Field = "object" | "phone" | "comment";

export interface Service {
  label: string;
  /** Shu xizmat uchun alohida qadamlar (bo'lmasa — bo'limning qadamlari) */
  flow?: Field[];
  /** Izoh so'ralganda chiqadigan matn */
  commentPrompt?: string;
}

export interface Category {
  /** Bosh menyudagi tugma matni (ikonka kerak bo'lmasa — olib tashlang) */
  title: string;
  /** "object" qadamida nima so'raladi: avtomobil raqami yoki hisob raqami */
  objectKind: "car" | "home" | "other";
  /** Bo'limdagi xizmatlar uchun standart qadamlar */
  flow: Field[];
  /** true bo'lsa, ichki menyu yo'q — tugma bosilishi bilan to'g'ridan-to'g'ri ariza boshlanadi */
  direct?: boolean;
  services: Service[];
}

const svc = (label: string, extra: Partial<Service> = {}): Service => ({ label, ...extra });

/** Har bir bo'limdagi "Boshqa" tugmasi: faqat savol yoziladi */
const other = () => svc("Boshqa", { flow: ["comment"], commentPrompt: "Savolingizni qisqacha yozing:" });

export const CATEGORIES: Record<CategoryKey, Category> = {
  car: {
    title: "🚗 Mening avtomobilim",
    objectKind: "car",
    flow: ["object", "phone", "comment"],
    services: [
      svc("Avtomobil qo'shish"),
      svc("Jarimalar"),
      svc("Yaqinlarni jarimalarini to'lash"),
      svc("Sug'urta"),
      svc("Tonirovka"),
      svc("Ishonchnoma"),
      svc("Car sharing"),
      other(),
    ],
  },
  home: {
    title: "🏠 Mening uyim",
    objectKind: "home",
    flow: ["object"], // faqat hisob raqami
    services: [
      svc("Internet"),
      svc("Ichimlik suvi"),
      svc("Tabiiy gaz"),
      svc("Elektr energiya"),
      svc("Issiq suv"),
      svc("Chiqindi"),
      svc("Telefon"),
      other(),
    ],
  },
  package: {
    title: "📦 Paket",
    objectKind: "other",
    flow: ["comment", "phone"], // muammo, keyin ixtiyoriy telefon
    direct: true,
    services: [svc("Paket")],
  },
  navigator: {
    title: "🧭 Navigator",
    objectKind: "other",
    flow: ["comment"], // faqat muammo
    direct: true,
    services: [svc("Navigator")],
  },
};

import "./silence";
import "dotenv/config";
import fs from "node:fs";
import {
  Bot,
  Context,
  GrammyError,
  InlineKeyboard,
  InputFile,
  Keyboard,
  SessionFlavor,
  session,
} from "grammy";
import { CATEGORIES, CategoryKey, Field, Service } from "./menu";
import {
  Application,
  activeApplication,
  allApplications,
  allUserIds,
  getShifts,
  setRating,
  shiftLabel,
  updateShift,
  addUserItem,
  createApplication,
  findUserByTopic,
  getBroadcastTopic,
  setBroadcastTopic,
  getApplication,
  getUser,
  setStatus,
  updateUser,
  userApplications,
} from "./db";
import { XLSX_FILE, syncAll, syncApplication } from "./excel";
import { currentMonth, monthLabel, monthOf, monthlyRows, shiftMonth, summarize } from "./stats";

// ───────────────────────── Sozlamalar ─────────────────────────

const BOT_TOKEN = process.env.BOT_TOKEN?.trim();
if (!BOT_TOKEN) {
  console.error("❌ BOT_TOKEN topilmadi. .env faylga BOT_TOKEN=... deb yozing.");
  process.exit(1);
}
const ADMIN_ID = process.env.ADMIN_ID?.trim() ? Number(process.env.ADMIN_ID.trim()) : undefined;
if (ADMIN_ID !== undefined && Number.isNaN(ADMIN_ID)) {
  console.error("❌ ADMIN_ID raqam bo'lishi kerak (masalan: 123456789 yoki -1001234567890).");
  process.exit(1);
}

// Adminlar: guruh adminlari avtomatik admin hisoblanadi. Qo'shimcha odamlar uchun .env da
// ADMIN_USER_IDS=111111,222222 deb yozish mumkin.
const ADMIN_USER_IDS = (process.env.ADMIN_USER_IDS ?? "")
  .split(",")
  .map((x) => Number(x.trim()))
  .filter((x) => x > 0);

// ───────────────────────── Session (qadamlar holati) ─────────────────────────

type Step =
  | "app_object"
  | "app_phone"
  | "app_comment"
  | "adm_name"
  | "adm_time"
  | "adm_names";

interface Draft {
  cat?: CategoryKey;
  service?: string;
  object?: string;
  phone?: string;
  comment?: string;
  /** Hali so'ralmagan qadamlar (navbat), umumiy qadamlar soni */
  flow?: Field[];
  total?: number;
  commentOptional?: boolean;
  commentPrompt?: string;
  /** Admin paneldagi tahrirlanayotgan smena raqami */
  shiftIdx?: number;
}

interface SessionData {
  step: Step | null;
  draft: Draft;
  /** Hozir qaysi bo'lim menyusi ochiq (reply tugmalar matn bo'lib kelgani uchun kerak) */
  menuCat?: CategoryKey;
  /** "Ma'lumotlarni tekshiring" xabari — tasdiqlangach o'chiriladi */
  checkMsgId?: number;
}

type MyContext = Context & SessionFlavor<SessionData>;

const bot = new Bot<MyContext>(BOT_TOKEN);

bot.use(session({ initial: (): SessionData => ({ step: null, draft: {} }) }));

/** Foydalanuvchi bilan shaxsiy chat. Guruh (operatorlar) xabarlari alohida ishlanadi. */
const priv = bot.chatType("private");

// ───────────────────────── Yordamchi funksiyalar ─────────────────────────

const CANCEL = "❌ Bekor qilish";
const BACK = "⬅️ Orqaga";
const SKIP = "⏭ O'tkazib yuborish";

const OBJECT_LABEL = { car: "🚘 Avtomobil", home: "🔢 Hisob raqami", other: "📝 Mavzu" } as const;

const esc = (s: string) =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/** Faqat joriy qadam va qoralamani tozalaydi (menyu holati saqlanadi) */
function reset(ctx: MyContext) {
  ctx.session.step = null;
  ctx.session.draft = {};
}

function normalizePhone(raw: string): string | null {
  const digits = raw.replace(/[^\d]/g, "");
  if (digits.length === 9) return "+998" + digits;
  if (digits.length === 12 && digits.startsWith("998")) return "+" + digits;
  return null;
}


/** Callback (inline tugma) bo'lsa xabarni tahrirlaydi, aks holda yangi xabar yuboradi */
async function show(ctx: MyContext, text: string, kb?: InlineKeyboard) {
  const opts = { parse_mode: "HTML" as const, reply_markup: kb };
  if (ctx.callbackQuery) {
    try {
      await ctx.editMessageText(text, opts);
      return;
    } catch {
      /* xabar o'zgarmagan bo'lishi mumkin — pastda yangi yuboramiz */
    }
  }
  await ctx.reply(text, opts);
}

// ───────────────────────── Klaviaturalar (Reply Keyboard — klaviatura o'rnida chiqadi) ─────────────────────────

/** Tugmalarni 2 tadan qatorga joylaydi */
function twoPerRow(kb: Keyboard, labels: string[]): Keyboard {
  labels.forEach((label, i) => {
    kb.text(label);
    if (i % 2 === 1 && i !== labels.length - 1) kb.row();
  });
  return kb;
}

function mainMenu(): Keyboard {
  const titles = (Object.keys(CATEGORIES) as CategoryKey[]).map((k) => CATEGORIES[k].title);
  return twoPerRow(new Keyboard(), titles).resized();
}

function categoryMenu(key: CategoryKey): Keyboard {
  const labels = CATEGORIES[key].services.map((s) => s.label);
  return twoPerRow(new Keyboard(), labels).row().text(BACK).resized();
}

function cancelKb(extra?: string[]) {
  const kb = new Keyboard();
  extra?.forEach((t) => kb.text(t).row());
  return kb.text(CANCEL).resized();
}

function phoneKb() {
  return new Keyboard().requestContact("📱 Raqamimni yuborish").row().text(SKIP).text(CANCEL).resized();
}

async function sendMainMenu(ctx: MyContext, text = "Quyidagi bo'limlardan birini tanlang 👇") {
  ctx.session.menuCat = undefined;
  await ctx.reply(text, { parse_mode: "HTML", reply_markup: mainMenu() });
}

async function sendCategoryMenu(ctx: MyContext, cat: CategoryKey) {
  ctx.session.menuCat = cat;
  await ctx.reply(`<b>${CATEGORIES[cat].title}</b>\nXizmatni tanlang 👇`, {
    parse_mode: "HTML",
    reply_markup: categoryMenu(cat),
  });
}

// ───────────────────────── Buyruqlar ─────────────────────────

priv.command("start", async (ctx) => {
  reset(ctx);
  getUser(ctx.from!.id, ctx.from!.username);
  void isOwner(ctx.from!.id).then((ok) => {
    if (ok) {
      bot.api
        .setMyCommands(OWNER_COMMANDS, { scope: { type: "chat", chat_id: ctx.from!.id } })
        .catch(() => {});
    }
  });
  await sendMainMenu(
    ctx,
    `Assalomu alaykum, <b>${esc(ctx.from!.first_name)}</b>! 👋\n\n` +
      `Men orqali xizmatlar bo'yicha ariza qoldirishingiz mumkin.\n\n` +
      `Quyidagi bo'limlardan birini tanlang 👇`
  );
});

priv.command("menu", async (ctx) => {
  reset(ctx);
  await sendMainMenu(ctx, "🏠 Bosh menyu");
});

/** Chat va foydalanuvchi ID'sini ko'rsatadi (guruh ID'sini bilish uchun ham qulay) */
bot.command("id", async (ctx) => {
  await ctx.reply(
    `🆔 Chat ID: <code>${ctx.chat.id}</code>\n👤 Sizning ID: <code>${ctx.from?.id}</code>`,
    { parse_mode: "HTML" }
  );
});

// ───────────────────────── Xizmatni ochish ─────────────────────────

async function openService(ctx: MyContext, cat: CategoryKey, service: Service) {
  reset(ctx);
  const flow = [...(service.flow ?? CATEGORIES[cat].flow)];
  ctx.session.draft = {
    cat,
    service: service.label,
    object: "",
    phone: "",
    comment: "",
    flow,
    total: flow.length,
    commentOptional: flow.includes("object"), // izoh faqat avtomobil/hisob raqami bilan birga bo'lsa ixtiyoriy
    commentPrompt: service.commentPrompt,
  };
  await askNext(ctx);
}

// ───────────────────────── Qadamlar ─────────────────────────

/** Navbatdagi qadamni so'raydi; qadamlar tugagan bo'lsa — tasdiqlashga o'tadi */
async function askNext(ctx: MyContext) {
  const d = ctx.session.draft;
  const field = d.flow?.shift();
  if (!field) {
    await askConfirm(ctx);
    return;
  }

  const n = d.total! - d.flow!.length;
  const header = n === 1 ? `📝 <b>${esc(d.service!)}</b> bo'yicha ariza\n\n` : "";
  const num = d.total! > 1 ? `<b>${n}/${d.total}.</b> ` : "";
  const kind = CATEGORIES[d.cat!].objectKind;

  if (field === "object") {
    ctx.session.step = "app_object";
    // faqat eng oxirgi kiritilgani tugma bo'lib chiqadi
    const user = getUser(ctx.from!.id);
    const list = kind === "car" ? user.cars : user.homes;
    const last = list.length ? [list[list.length - 1]] : [];
    await ctx.reply(
      header +
        num +
        (kind === "car"
          ? "Avtomobil davlat raqamini kiriting (masalan: <code>01 A 123 BC</code>):"
          : "Hisob raqamini kiriting (gaz, elektr, suv yoki chiqindi hisob raqami):"),
      { parse_mode: "HTML", reply_markup: cancelKb(last) }
    );
  } else if (field === "phone") {
    ctx.session.step = "app_phone";
    await ctx.reply(header + num + "Bog'lanish uchun telefon raqamingizni yuboring (ixtiyoriy):", {
      parse_mode: "HTML",
      reply_markup: phoneKb(),
    });
  } else {
    ctx.session.step = "app_comment";
    await ctx.reply(header + num + (d.commentPrompt ?? "Muammo yoki so'rovingizni qisqacha yozing:"), {
      parse_mode: "HTML",
      reply_markup: d.commentOptional ? cancelKb([SKIP]) : cancelKb(),
    });
  }
}

async function askConfirm(ctx: MyContext) {
  const d = ctx.session.draft;
  const kind = CATEGORIES[d.cat!].objectKind;
  ctx.session.step = null; // tasdiqlash inline tugmalar orqali
  ctx.session.menuCat = undefined;

  const note = await ctx.reply("Ma'lumotlarni tekshiring 👇", { reply_markup: mainMenu() });
  ctx.session.checkMsgId = note.message_id;

  const lines = [
    `📋 <b>Ariza ma'lumotlari</b>\n`,
    `Bo'lim: ${esc(CATEGORIES[d.cat!].title)}`,
    `Xizmat: <b>${esc(d.service!)}</b>`,
  ];
  if (d.object) lines.push(`${OBJECT_LABEL[kind]}: ${esc(d.object)}`);
  if (d.phone) lines.push(`📞 Telefon: ${esc(d.phone)}`);
  if (d.comment) lines.push(`💬 Izoh: ${esc(d.comment)}`);

  await ctx.reply(lines.join("\n"), {
    parse_mode: "HTML",
    reply_markup: new InlineKeyboard()
      .text("✅ Tasdiqlash", "app:confirm")
      .text("❌ Bekor qilish", "app:cancel"),
  });
}

/** "Ma'lumotlarni tekshiring" xabarini o'chiradi (tasdiqlash yoki bekor qilingandan keyin) */
async function dropCheckNote(ctx: MyContext) {
  const id = ctx.session.checkMsgId;
  ctx.session.checkMsgId = undefined;
  if (id && ctx.chat) await ctx.api.deleteMessage(ctx.chat.id, id).catch(() => {});
}

/** Telefon raqam qabul qilingandan keyin (kontakt yoki matn) */
async function onPhone(ctx: MyContext, raw: string) {
  const phone = normalizePhone(raw.replace("📞", ""));
  if (!phone) {
    await ctx.reply("⚠️ Raqam noto'g'ri. Format: <code>+998901234567</code>", {
      parse_mode: "HTML",
    });
    return;
  }
  const step = ctx.session.step;

  if (step === "app_phone") {
    ctx.session.draft.phone = phone;
    await askNext(ctx);
  }
}

priv.on("message:contact", async (ctx) => {
  const step = ctx.session.step;
  if (step === "app_phone") {
    await onPhone(ctx, ctx.message.contact.phone_number);
  }
});

// ───────────────────────── Matnli xabarlar (menyu tugmalari ham shu yerga keladi) ─────────────────────────

priv.on("message:text", async (ctx, next) => {
  const text = ctx.message.text.trim();
  if (text.startsWith("/")) return next(); // buyruqlar alohida

  // 1) Bekor qilish
  if (text === CANCEL) {
    reset(ctx);
    await sendMainMenu(ctx, "Bekor qilindi.");
    return;
  }

  // 2) Orqaga
  if (text === BACK) {
    const inFlow = ctx.session.step !== null;
    const cat = ctx.session.menuCat;
    reset(ctx);
    if (inFlow && cat) await sendCategoryMenu(ctx, cat);
    else await sendMainMenu(ctx);
    return;
  }

  // 3) Bosh menyu tugmalari (bo'limlar) — har doim ustun turadi
  const catKey = (Object.keys(CATEGORIES) as CategoryKey[]).find(
    (k) => CATEGORIES[k].title === text
  );
  if (catKey) {
    reset(ctx);
    if (CATEGORIES[catKey].direct) {
      ctx.session.menuCat = undefined;
      await openService(ctx, catKey, CATEGORIES[catKey].services[0]);
    } else {
      await sendCategoryMenu(ctx, catKey);
    }
    return;
  }

  // 4) Bo'lim ichidagi xizmat tugmalari (faqat qadam ketmayotgan paytda)
  if (ctx.session.step === null && ctx.session.menuCat) {
    const service = CATEGORIES[ctx.session.menuCat].services.find((s) => s.label === text);
    if (service) {
      await openService(ctx, ctx.session.menuCat, service);
      return;
    }
  }

  // 5) Joriy qadamga javob
  const d = ctx.session.draft;

  switch (ctx.session.step) {
    case "app_object":
      if (text.length < 3 || text.length > 200) {
        await ctx.reply("⚠️ Iltimos, to'g'ri ma'lumot kiriting.");
        return;
      }
      d.object = text;
      await askNext(ctx);
      return;

    case "app_phone":
      if (text === SKIP) {
        d.phone = ""; // telefon kiritilmadi
        await askNext(ctx);
        return;
      }
      await onPhone(ctx, text);
      return;

    case "app_comment":
      if (text === SKIP && d.commentOptional) {
        d.comment = "";
      } else {
        if (text.length < 2) {
          await ctx.reply("⚠️ Iltimos, matn yozing.");
          return;
        }
        d.comment = text.slice(0, 1000);
      }
      await askNext(ctx);
      return;

    case "adm_names": {
      if (!(await isOwner(ctx.from.id))) {
        reset(ctx);
        return;
      }
      const count = getShifts().length;
      const names = parseNames(text, count);
      if (!names) {
        await ctx.reply(
          `⚠️ Aniq ${count} ta ism yozing. Masalan: <code>Bonu Asadbek Abdug'ani</code>\n` +
            `Ismlar bo'sh joy, vergul yoki yangi qator bilan ajratiladi.`,
          { parse_mode: "HTML" }
        );
        return;
      }
      names.forEach((name, i) => updateShift(i, { operator: name }));
      reset(ctx);
      await ctx.reply("✅ Saqlandi.", { reply_markup: mainMenu() });
      await ctx.reply(shiftsText(), { parse_mode: "HTML", reply_markup: shiftsKb() });
      return;
    }

    case "adm_name": {
      if (!(await isOwner(ctx.from.id))) {
        reset(ctx);
        return;
      }
      if (text.length < 2 || text.length > 40) {
        await ctx.reply("⚠️ Ism 2–40 belgi bo'lishi kerak.");
        return;
      }
      updateShift(d.shiftIdx!, { operator: text });
      reset(ctx);
      await ctx.reply("✅ Saqlandi.", { reply_markup: mainMenu() });
      await ctx.reply(shiftsText(), { parse_mode: "HTML", reply_markup: shiftsKb() });
      return;
    }

    case "adm_time": {
      if (!(await isOwner(ctx.from.id))) {
        reset(ctx);
        return;
      }
      const m = text.replace(/\s/g, "").match(/^(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})$/);
      if (!m || +m[1] > 23 || +m[3] > 23 || +m[2] > 59 || +m[4] > 59) {
        await ctx.reply("⚠️ Format noto'g'ri. Masalan: <code>9:00-18:00</code>", { parse_mode: "HTML" });
        return;
      }
      updateShift(d.shiftIdx!, { start: `${+m[1]}:${m[2]}`, end: `${+m[3]}:${m[4]}` });
      reset(ctx);
      await ctx.reply("✅ Saqlandi.", { reply_markup: mainMenu() });
      await ctx.reply(shiftsText(), { parse_mode: "HTML", reply_markup: shiftsKb() });
      return;
    }

    default:
      await relayToStaff(ctx); // menyu tugmasi emas — operatorga yetkaziladi
  }
});

// ───────────────────────── Ariza tasdiqlash ─────────────────────────

bot.callbackQuery("app:cancel", async (ctx) => {
  await ctx.answerCallbackQuery();
  await dropCheckNote(ctx);
  reset(ctx);
  await show(ctx, "❌ Ariza bekor qilindi.");
  await sendMainMenu(ctx);
});

bot.callbackQuery("app:confirm", async (ctx) => {
  await ctx.answerCallbackQuery();
  await dropCheckNote(ctx);
  const d = ctx.session.draft;
  if (!d.cat || !d.service) {
    await show(ctx, "⚠️ Sessiya eskirgan. Iltimos, /start bosing.");
    return;
  }

  const application = createApplication({
    userId: ctx.from.id,
    category: d.cat,
    service: d.service,
    object: d.object ?? "",
    phone: d.phone ?? "",
    comment: d.comment ?? "",
  });
  reset(ctx);
  void syncApplication(application); // Excel faylga yoziladi

  // keyingi safar tayyor tugma bo'lib chiqishi uchun eslab qolamiz
  const kind = CATEGORIES[application.category].objectKind;
  if (kind !== "other" && application.object) addUserItem(ctx.from.id, kind, application.object);

  if (GROUP_MODE) {
    await show(
      ctx,
      `✅ <b>Arizangiz qabul qilindi!</b>\n\nAriza raqami: <b>#${application.id}</b>\n\n` +
        `⏳ <b>Operator javobini kuting...</b>`
    );
    await notifyAdmin(application, ctx);
    ctx.session.menuCat = undefined;
    await ctx.reply("💬 Savolingiz bo'lsa, shu yerga yozing — xabaringiz operatorga yetkaziladi.", {
      reply_markup: mainMenu(),
    });
  } else {
    await show(
      ctx,
      `✅ <b>Arizangiz qabul qilindi!</b>\n\nAriza raqami: <b>#${application.id}</b>\n` +
        `Operator tez orada siz bilan bog'lanadi.`
    );
    await notifyAdmin(application, ctx);
    await sendMainMenu(ctx);
  }
});

// ───────────────────────── Guruhdagi topiclar (ADMIN_ID = guruh ID) ─────────────────────────

const GROUP_MODE = ADMIN_ID !== undefined && ADMIN_ID < 0;
const topicLocks = new Map<number, Promise<number | undefined>>();

interface Person {
  id: number;
  first_name: string;
  username?: string;
}

/** Foydalanuvchi uchun guruhda topic ochadi (nomida ID bor). Mavjud bo'lsa, shuni qaytaradi. */
function ensureTopic(from: Person): Promise<number | undefined> {
  if (!GROUP_MODE) return Promise.resolve(undefined);
  const user = getUser(from.id, from.username);
  if (user.topicId) return Promise.resolve(user.topicId);

  const running = topicLocks.get(from.id);
  if (running) return running;

  const task = (async () => {
    try {
      const name = `${from.first_name.slice(0, 80)} (${from.id})`;
      const topic = await bot.api.createForumTopic(ADMIN_ID!, name);
      updateUser(from.id, { topicId: topic.message_thread_id });
      const header = await bot.api.sendMessage(
        ADMIN_ID!,
        `👤 <a href="tg://user?id=${from.id}">${esc(from.first_name)}</a>` +
          (from.username ? ` (@${esc(from.username)})` : "") +
          `\n🆔 <code>${from.id}</code>`,
        { parse_mode: "HTML", message_thread_id: topic.message_thread_id }
      );
      // topic tepasida ID turishi uchun yuqoriga mahkamlaymiz (huquq bo'lmasa — jim o'tamiz)
      bot.api.pinChatMessage(ADMIN_ID!, header.message_id, { disable_notification: true }).catch(() => {});
      return topic.message_thread_id;
    } catch (e) {
      console.error(
        "Topic ochib bo'lmadi. Guruhda Topics yoqilganmi va botda 'Manage Topics' huquqi bormi?",
        e
      );
      return undefined;
    } finally {
      topicLocks.delete(from.id);
    }
  })();
  topicLocks.set(from.id, task);
  return task;
}

/** Operatorlarga xabar: guruh rejimida foydalanuvchining topiciga, aks holda ADMIN_ID ga */
async function sendToStaff(from: Person, text: string) {
  if (ADMIN_ID === undefined) {
    console.warn("ℹ️ ADMIN_ID berilmagan — ariza hech kimga yuborilmadi (test rejimi).");
    return;
  }
  const send = (threadId?: number) =>
    bot.api.sendMessage(ADMIN_ID, text, {
      parse_mode: "HTML",
      ...(threadId ? { message_thread_id: threadId } : {}),
    });
  try {
    await send(await ensureTopic(from));
  } catch (e) {
    if (GROUP_MODE) {
      // topic o'chirib yuborilgan bo'lishi mumkin — qayta ochib ko'ramiz
      updateUser(from.id, { topicId: undefined });
      try {
        await send(await ensureTopic(from));
        return;
      } catch (e2) {
        console.error("Guruhga yuborib bo'lmadi:", e2);
        return;
      }
    }
    console.error(
      "Adminga yuborib bo'lmadi. ADMIN_ID to'g'rimi va shu akkaunt botga /start bosganmi?",
      e
    );
  }
}

async function notifyAdmin(a: Application, ctx: MyContext) {
  const from = ctx.from!;
  const details = [
    a.object ? `${OBJECT_LABEL[CATEGORIES[a.category].objectKind]}: ${esc(a.object)}` : "",
    a.phone ? `📞 ${esc(a.phone)}` : "",
    a.comment ? `💬 ${esc(a.comment)}` : "",
  ]
    .filter(Boolean)
    .join("\n");
  await sendToStaff(
    from,
    `🆕 <b>Yangi ariza #${a.id}</b>\n\n` +
      `Bo'lim: ${esc(CATEGORIES[a.category].title)}\n` +
      `Xizmat: <b>${esc(a.service)}</b>\n` +
      `${details}\n\n` +
      `👤 <a href="tg://user?id=${from.id}">${esc(from.first_name)}</a>` +
      (from.username ? ` (@${esc(from.username)})` : "") +
      ` · 🆔 <code>${from.id}</code>` +
      `\n\nYopish: <code>/close ${a.id}</code>`
  );
}

// ───────────────────────── Operator ishini baholash ─────────────────────────

function ratingKb(id: number) {
  return new InlineKeyboard()
    .text("✅ Muammo hal qilindi", `rate:${id}:1`)
    .row()
    .text("❌ Muammo hal qilinmadi", `rate:${id}:0`);
}

const RATING_TEXT = (id: number) =>
  `⭐ <b>Operator ishini baholang</b>\n\nAriza <b>#${id}</b> bo'yicha muammo hal qilindimi?`;

/**
 * Arizani "yopish" — operator ishini tugatganda foydalanuvchiga baholash so'rovi ketadi.
 *  /close <id>  — admin yoki ADMIN_ID guruhidagi operatorlar uchun
 *  ADMIN_ID berilmagan bo'lsa (test rejimi) — istalgan foydalanuvchi o'zining arizasini yopib sinab ko'ra oladi
 */
bot.command("close", async (ctx) => {
  const inStaffChat = ADMIN_ID !== undefined && ctx.chat.id === ADMIN_ID;
  const isAdmin = ADMIN_ID !== undefined && (ctx.from!.id === ADMIN_ID || inStaffChat);
  const devMode = ADMIN_ID === undefined;
  if (!isAdmin && !devMode) return;

  const arg = ctx.match?.toString().trim();
  let application: Application | undefined;
  if (arg) {
    application = getApplication(Number(arg));
  } else if (inStaffChat && ctx.msg.message_thread_id) {
    // topic ichida argumentsiz /close — shu foydalanuvchining joriy arizasi
    const u = findUserByTopic(ctx.msg.message_thread_id);
    const act = u && activeApplication(u.id);
    if (act && (act.status === "new" || act.status === "unresolved")) application = act;
  } else if (devMode) {
    application = userApplications(ctx.from!.id).filter((a) => a.status === "new").pop();
  }
  if (!application) {
    await ctx.reply("Ariza topilmadi. Foydalanish: /close <id> (topic ichida — shunchaki /close)");
    return;
  }
  if (devMode && application.userId !== ctx.from!.id) {
    await ctx.reply("Test rejimida faqat o'z arizangizni yopa olasiz.");
    return;
  }

  const closed = setStatus(application.id, "done");
  if (closed) void syncApplication(closed);
  try {
    await bot.api.sendMessage(application.userId, RATING_TEXT(application.id), {
      parse_mode: "HTML",
      reply_markup: ratingKb(application.id),
    });
    await ctx.reply(`✅ #${application.id} yopildi, foydalanuvchiga baholash yuborildi.`);
  } catch {
    await ctx.reply("⚠️ Foydalanuvchiga xabar yuborib bo'lmadi.");
  }
});

// ───────────────────────── Jonli yozishma: foydalanuvchi ⇄ operatorlar guruhi ─────────────────────────

/** Foydalanuvchi xabarini (matn, rasm, ovoz, fayl...) uning topiciga ko'chiradi */
async function relayToStaff(ctx: MyContext) {
  const userId = ctx.from!.id;
  const app = GROUP_MODE ? activeApplication(userId) : undefined;

  if (!app) {
    await sendMainMenu(ctx, "Iltimos, menyudan tanlang 👇");
    return;
  }
  if (app.status === "done") {
    // operator yopgan, baho kutilmoqda
    await ctx.reply(RATING_TEXT(app.id), { parse_mode: "HTML", reply_markup: ratingKb(app.id) });
    return;
  }

  const copy = async () => {
    const topicId = await ensureTopic(ctx.from!);
    await bot.api.copyMessage(ADMIN_ID!, ctx.chat!.id, ctx.msg!.message_id, {
      ...(topicId ? { message_thread_id: topicId } : {}),
    });
  };
  try {
    await copy();
  } catch {
    // topic o'chirilgan bo'lishi mumkin — qayta ochib ko'ramiz
    updateUser(userId, { topicId: undefined });
    try {
      await copy();
    } catch (e) {
      console.error("Xabarni guruhga yuborib bo'lmadi:", e);
      await ctx.reply("⚠️ Xabar yetkazilmadi, birozdan keyin qayta urinib ko'ring.");
    }
  }
}

// Menyuda turmagan xabarlar (rasm, ovoz, fayl va h.k.) — operatorga
priv.on("message", async (ctx, next) => {
  // matn va kontaktlar yuqorida ishlangan; buyruqlar (/smena, /admin...) keyingi handlerlarga o'tishi shart
  if (ctx.message.text || ctx.message.contact) return next();
  if (ctx.session.step !== null) {
    await ctx.reply("Iltimos, so'ralgan ma'lumotni matn ko'rinishida yuboring.");
    return;
  }
  await relayToStaff(ctx);
});

// ───────────────────────── Ommaviy xabar: bitta topic → barcha foydalanuvchilar ─────────────────────────

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** Guruhda "📢 Hammaga xabar" topicini (bir marta) ochadi */
async function ensureBroadcastTopic() {
  if (!GROUP_MODE || getBroadcastTopic()) return;
  try {
    const topic = await bot.api.createForumTopic(ADMIN_ID!, "📢 Hammaga xabar");
    setBroadcastTopic(topic.message_thread_id);
    const info = await bot.api.sendMessage(
      ADMIN_ID!,
      "📢 <b>Ommaviy xabar</b>\n\nBu topicga yozilgan <b>har qanday</b> xabar (matn, rasm, video, fayl) " +
        "botning barcha foydalanuvchilariga yuboriladi.\n\n" +
        "⚠️ Yuborilgan xabarni qaytarib bo'lmaydi — yozishdan oldin tekshiring.",
      { parse_mode: "HTML", message_thread_id: topic.message_thread_id }
    );
    bot.api.pinChatMessage(ADMIN_ID!, info.message_id, { disable_notification: true }).catch(() => {});
  } catch (e) {
    console.error("Ommaviy xabar topicini ochib bo'lmadi (Manage Topics huquqi bormi?):", e);
  }
}

/** Mavjud topicni ommaviy xabar topici qilib belgilash: topic ichida /setbroadcast */
bot.command("setbroadcast", async (ctx) => {
  if (!GROUP_MODE || ctx.chat.id !== ADMIN_ID) return;
  const threadId = ctx.msg.message_thread_id;
  if (!threadId) {
    await ctx.reply("Bu buyruqni kerakli topic ichida yozing.");
    return;
  }
  if (findUserByTopic(threadId)) {
    await ctx.reply("Bu foydalanuvchining topici, ommaviy xabar uchun yaramaydi.", {
      message_thread_id: threadId,
    });
    return;
  }
  setBroadcastTopic(threadId);
  await ctx.reply("📢 Endi bu topicga yozilgan xabarlar barcha foydalanuvchilarga yuboriladi.", {
    message_thread_id: threadId,
  });
});

/** Bitta foydalanuvchiga nusxa yuboradi; Telegram "kuting" desa (429) kutib qayta urinadi */
async function copyTo(userId: number, fromChatId: number, messageId: number) {
  try {
    await bot.api.copyMessage(userId, fromChatId, messageId);
  } catch (e) {
    if (e instanceof GrammyError && e.error_code === 429) {
      const wait = (e.parameters.retry_after ?? 1) + 1;
      await sleep(wait * 1000);
      await bot.api.copyMessage(userId, fromChatId, messageId);
      return;
    }
    throw e;
  }
}

async function broadcast(messageId: number, threadId: number) {
  const ids = allUserIds();
  const say = (text: string) =>
    bot.api.sendMessage(ADMIN_ID!, text, { message_thread_id: threadId }).catch(() => {});

  if (!ids.length) {
    await say("ℹ️ Hali birorta foydalanuvchi yo'q.");
    return;
  }
  await say(`⏳ ${ids.length} ta foydalanuvchiga yuborilmoqda...`);

  let ok = 0;
  let failed = 0;
  for (const id of ids) {
    try {
      await copyTo(id, ADMIN_ID!, messageId);
      ok++;
    } catch {
      failed++; // botni bloklagan yoki o'chirib yuborgan bo'lishi mumkin
    }
    await sleep(50); // Telegram limiti (~30 xabar/soniya) dan oshmaslik uchun
  }
  await say(`✅ Yuborildi: ${ok}\n⚠️ Yetkazilmadi: ${failed}`);
}

// Operator guruhdagi topicga yozsa — foydalanuvchiga botdan yetkaziladi
const RELAYABLE = [
  "text",
  "photo",
  "video",
  "document",
  "voice",
  "audio",
  "video_note",
  "sticker",
  "animation",
  "location",
  "contact",
];

bot.on("message", async (ctx, next) => {
  if (!GROUP_MODE || ctx.chat.id !== ADMIN_ID) return next();
  const threadId = ctx.msg.message_thread_id;
  if (!threadId || ctx.from?.is_bot) return;
  if (!RELAYABLE.some((k) => k in ctx.msg)) return; // xizmat xabarlari (topic ochildi va h.k.)
  if (ctx.msg.text?.startsWith("/")) return; // buyruqlar foydalanuvchiga ketmaydi

  // Ommaviy xabar topici: hammaga yuboriladi (uzoq davom etishi mumkin — fonda ishlaydi)
  if (threadId === getBroadcastTopic()) {
    pendingBroadcasts.add(ctx.msg.message_id);
    await ctx.reply(
      `📢 Shu xabar <b>${allUserIds().length}</b> ta foydalanuvchiga yuborilsinmi?\n` +
        `<i>Yuborilgach qaytarib bo'lmaydi.</i>`,
      {
        parse_mode: "HTML",
        message_thread_id: threadId,
        reply_parameters: { message_id: ctx.msg.message_id },
        reply_markup: new InlineKeyboard()
          .text("✅ Yuborish", `bc:yes:${ctx.msg.message_id}:${threadId}`)
          .text("❌ Bekor qilish", `bc:no:${ctx.msg.message_id}:${threadId}`),
      }
    );
    return;
  }

  const user = findUserByTopic(threadId);
  if (!user) return;

  try {
    if (ctx.msg.text) {
      await bot.api.sendMessage(user.id, ctx.msg.text);
    } else {
      await bot.api.copyMessage(user.id, ctx.chat.id, ctx.msg.message_id);
    }
  } catch {
    await ctx.reply("⚠️ Foydalanuvchiga yetkazib bo'lmadi (bot bloklangan bo'lishi mumkin).", {
      message_thread_id: threadId,
    });
  }
});

// Ommaviy xabar tasdiqlash (faqat adminlar)
const pendingBroadcasts = new Set<number>();

bot.callbackQuery(/^bc:(yes|no):(\d+):(\d+)$/, async (ctx) => {
  if (!GROUP_MODE || ctx.chat?.id !== ADMIN_ID) {
    await ctx.answerCallbackQuery();
    return;
  }
  if (!(await isAdmin(ctx.from.id))) {
    await ctx.answerCallbackQuery({ text: "Faqat adminlar tasdiqlay oladi", show_alert: true });
    return;
  }
  const action = ctx.match[1];
  const messageId = Number(ctx.match[2]);
  const threadId = Number(ctx.match[3]);

  if (!pendingBroadcasts.delete(messageId)) {
    await ctx.answerCallbackQuery({ text: "Bu so'rov eskirgan yoki allaqachon bajarilgan." });
    await ctx.editMessageReplyMarkup({ reply_markup: undefined }).catch(() => {});
    return;
  }
  await ctx.answerCallbackQuery();
  if (action === "no") {
    await ctx.editMessageText("❌ Yuborish bekor qilindi.").catch(() => {});
    return;
  }
  await ctx.editMessageText("🚀 Yuborish boshlandi...").catch(() => {});
  void broadcast(messageId, threadId);
});

// ───────────────────────── Operator ishini baholash: hal qilindi/qilinmadi → 1–5 ─────────────────────────

function starsKb(id: number) {
  const kb = new InlineKeyboard();
  for (let n = 1; n <= 5; n++) kb.text(String(n), `star:${id}:${n}`);
  return kb; // bitta qatorda 1 2 3 4 5
}

bot.callbackQuery(/^rate:(\d+):([01])$/, async (ctx) => {
  await ctx.answerCallbackQuery();
  const id = Number(ctx.match[1]);
  const solved = ctx.match[2] === "1";
  const a = getApplication(id);

  if (!a || a.userId !== ctx.from.id) return;
  if (a.status !== "done") {
    await show(ctx, "ℹ️ Bu so'rovga allaqachon javob berilgan.");
    return;
  }

  const updated = setStatus(id, solved ? "resolved" : "unresolved");
  if (updated) void syncApplication(updated);

  if (solved) {
    await sendToStaff(ctx.from, `✅ Ariza #${id}: foydalanuvchi "Muammo hal qilindi" deb tasdiqladi.`);
    // ball faqat muammo hal bo'lgandan keyin so'raladi
    await show(
      ctx,
      `✅ Rahmat! Ariza <b>#${id}</b> hal qilindi deb belgilandi.\n\n` +
        `⭐ <b>Operator ishini 1 dan 5 gacha baholang</b>\n<i>1 — juda yomon, 5 — a'lo</i>`,
      starsKb(id)
    );
    return;
  }

  // Muammo hal bo'lmagan: ball so'ralmaydi, yozishma davom etadi
  await sendToStaff(
    ctx.from,
    `❌ Ariza #${id}: foydalanuvchi "Muammo hal qilinmadi" deb javob berdi. Yozishma davom etadi — muammo hal bo'lgach yana /close yozing.`
  );
  await show(
    ctx,
    `❌ Afsus, muammo hal bo'lmaganini eshitib ranjidik.\nAriza <b>#${id}</b> qayta ko'rib chiqiladi, operator siz bilan bog'lanadi.`
  );
  ctx.session.menuCat = undefined;
  await ctx.reply("⏳ Operator javobini kuting. Xabaringizni shu yerga yozing 💬", {
    reply_markup: mainMenu(),
  });
});

bot.callbackQuery(/^star:(\d+):([1-5])(?::[01])?$/, async (ctx) => {
  await ctx.answerCallbackQuery();
  const id = Number(ctx.match[1]);
  const stars = Number(ctx.match[2]);
  const a = getApplication(id);
  if (!a || a.userId !== ctx.from.id) return;
  if (a.status !== "resolved") return; // ball faqat hal bo'lgan ariza uchun
  if (a.rating) {
    await show(ctx, `⭐ Bu ariza allaqachon baholangan: <b>${a.rating}/5</b>`);
    return;
  }

  const updated = setRating(id, stars);
  if (updated) void syncApplication(updated);

  await sendToStaff(ctx.from, `⭐ Ariza #${id}: foydalanuvchi operator ishini <b>${stars}/5</b> ga baholadi.`);
  await show(ctx, `⭐ Rahmat! Bahoyingiz: <b>${stars}/5</b>`);
  await sendMainMenu(ctx);
});

// ───────────────────────── Admin panel ─────────────────────────

/** Guruh adminlari (yoki ADMIN_USER_IDS dagilar) admin hisoblanadi */
async function isAdmin(userId: number): Promise<boolean> {
  if (ADMIN_USER_IDS.includes(userId)) return true;
  if (ADMIN_ID === undefined) return false;
  if (!GROUP_MODE) return userId === ADMIN_ID;
  try {
    const m = await bot.api.getChatMember(ADMIN_ID, userId);
    return m.status === "creator" || m.status === "administrator";
  } catch {
    return false;
  }
}

/** Bot egasi: guruh egasi (creator) yoki .env dagi ADMIN_USER_IDS. Oddiy guruh adminlari — emas. */
async function isOwner(userId: number): Promise<boolean> {
  if (ADMIN_USER_IDS.includes(userId)) return true;
  if (ADMIN_ID === undefined) return false;
  if (!GROUP_MODE) return userId === ADMIN_ID;
  try {
    const m = await bot.api.getChatMember(ADMIN_ID, userId);
    return m.status === "creator";
  } catch {
    return false;
  }
}

/** "Bonu Asadbek Abdug'ani" → ["Bonu", "Asadbek", "Abdug'ani"] (vergul yoki yangi qator bilan ham bo'ladi) */
function parseNames(text: string, count: number): string[] | null {
  for (const sep of [/[\n,;]+/, /\s+/]) {
    const parts = text.split(sep).map((x) => x.trim()).filter(Boolean);
    if (parts.length === count && parts.every((p) => p.length <= 40)) return parts;
  }
  return null;
}

async function startNames(ctx: MyContext) {
  reset(ctx);
  ctx.session.step = "adm_names";
  const list = getShifts()
    .map((s, i) => `${i + 1}. ${esc(shiftLabel(s, i))} — <b>${esc(s.operator)}</b>`)
    .join("\n");
  await ctx.reply(
    `👥 <b>Smena operatorlari</b>\n\nHozir:\n${list}\n\n` +
      `Yangi operatorlarni 1, 2, 3-smena tartibida yozing.\n` +
      `Masalan: <code>Bonu Asadbek Abdug'ani</code>\n` +
      `→ 1-smena Bonu, 2-smena Asadbek, 3-smena Abdug'ani\n\n` +
      `<i>Smena vaqtlari o'zgarmaydi (ularni /admin dan o'zgartirasiz).</i>`,
    { parse_mode: "HTML", reply_markup: cancelKb() }
  );
}

priv.command("smena", async (ctx) => {
  if (!(await isOwner(ctx.from!.id))) {
    await ctx.reply("Bu buyruq faqat bot egasi uchun.");
    return;
  }
  await startNames(ctx);
});

function adminMenuKb() {
  return new InlineKeyboard()
    .text("🕐 Smenalar va operatorlar", "adm:shifts")
    .row()
    .text("📊 Excel faylni yuklab olish", "adm:export")
    .row()
    .text("📅 Oylik statistika", "adm:stats");
}

function shiftsText() {
  const list = getShifts()
    .map((s, i) => `${i + 1}. <b>${esc(shiftLabel(s, i))}</b> — ${esc(s.operator)}`)
    .join("\n");
  return `🕐 <b>Smenalar</b>\n\n${list}\n\nTahrirlash uchun smenani tanlang:`;
}

function shiftsKb() {
  const kb = new InlineKeyboard();
  getShifts().forEach((_, i) => kb.text(`✏️ ${i + 1}-smena`, `adm:shift:${i}`));
  return kb
    .row()
    .text("👥 Operatorlarni birdan o'zgartirish", "adm:names")
    .row()
    .text("⬅️ Orqaga", "adm:home");
}

async function sendExcel(ctx: MyContext) {
  await syncAll(allApplications()); // fayl yangi bo'lishi uchun
  if (!fs.existsSync(XLSX_FILE)) {
    await ctx.reply("Excel fayl hali yaratilmagan.");
    return;
  }
  await ctx.replyWithDocument(new InputFile(XLSX_FILE, "arizalar.xlsx"));
}

const fmtAvg = (v: number | null) => (v === null ? "—" : v.toFixed(2));

/** Oylik statistika: har bir operator bo'yicha jami, hal qilindi, hal qilinmadi, jarayonda, o'rtacha ball */
function monthText(ym: number) {
  const apps = allApplications();
  const rows = monthlyRows(apps, ym);
  if (!rows.length) return `📅 <b>${monthLabel(ym)}</b>\n\nBu oy uchun arizalar yo'q.`;

  const lines = rows
    .map(
      (r) =>
        `👤 <b>${esc(r.operator)}</b>\n` +
        `Jami: <b>${r.total}</b> · ✅ ${r.resolved} · ❌ ${r.unresolved} · ⏳ ${r.inProgress} · ⭐ ${fmtAvg(r.avg)}`
    )
    .join("\n\n");
  const all = summarize(apps.filter((a) => monthOf(a) === ym));
  return (
    `📅 <b>${monthLabel(ym)}</b> — oylik statistika\n\n${lines}\n\n━━━━━━━━\n` +
    `<b>Jami: ${all.total}</b> · ✅ ${all.resolved} · ❌ ${all.unresolved} · ⏳ ${all.inProgress} · ⭐ ${fmtAvg(all.avg)}\n\n` +
    `<i>✅ hal qilindi · ❌ hal qilinmadi · ⏳ jarayonda · ⭐ o'rtacha ball</i>`
  );
}

function monthKb(ym: number) {
  const prev = shiftMonth(ym, -1);
  const next = shiftMonth(ym, 1);
  const kb = new InlineKeyboard().text(`◀️ ${monthLabel(prev)}`, `adm:month:${prev}`);
  if (ym < currentMonth()) kb.text(`${monthLabel(next)} ▶️`, `adm:month:${next}`);
  return kb.row().text("⬅️ Orqaga", "adm:home");
}

priv.command("admin", async (ctx) => {
  reset(ctx);
  if (!(await isOwner(ctx.from!.id))) {
    await ctx.reply("Bu buyruq faqat bot egasi uchun.");
    return;
  }
  await ctx.reply("🛠 <b>Admin panel</b>", { parse_mode: "HTML", reply_markup: adminMenuKb() });
});

priv.command("export", async (ctx) => {
  if (!(await isOwner(ctx.from!.id))) return;
  await sendExcel(ctx);
});

bot.callbackQuery(/^adm:([a-z]+)(?::(\d+))?$/, async (ctx) => {
  if (!(await isOwner(ctx.from.id))) {
    await ctx.answerCallbackQuery({ text: "Faqat bot egasi uchun", show_alert: true });
    return;
  }
  await ctx.answerCallbackQuery();
  const action = ctx.match[1];
  const idx = ctx.match[2] !== undefined ? Number(ctx.match[2]) : undefined;
  const shift = idx !== undefined ? getShifts()[idx] : undefined;

  switch (action) {
    case "home":
      await show(ctx, "🛠 <b>Admin panel</b>", adminMenuKb());
      break;
    case "shifts":
      await show(ctx, shiftsText(), shiftsKb());
      break;
    case "shift":
      if (!shift || idx === undefined) return;
      await show(
        ctx,
        `✏️ <b>${esc(shiftLabel(shift, idx))}</b>\n\nOperator: <b>${esc(shift.operator)}</b>\nVaqt: <b>${shift.start}-${shift.end}</b>`,
        new InlineKeyboard()
          .text("👤 Operator ismi", `adm:name:${idx}`)
          .text("🕐 Vaqt", `adm:time:${idx}`)
          .row()
          .text("⬅️ Orqaga", "adm:shifts")
      );
      break;
    case "name":
      if (idx === undefined) return;
      reset(ctx);
      ctx.session.step = "adm_name";
      ctx.session.draft = { shiftIdx: idx };
      await ctx.reply(`👤 ${idx + 1}-smena uchun yangi operator ismini yozing:`, {
        reply_markup: cancelKb(),
      });
      break;
    case "time":
      if (idx === undefined) return;
      reset(ctx);
      ctx.session.step = "adm_time";
      ctx.session.draft = { shiftIdx: idx };
      await ctx.reply(
        `🕐 ${idx + 1}-smena uchun yangi vaqtni yozing.\nFormat: <code>9:00-18:00</code> (yarim tundan o'tsa ham bo'ladi: <code>18:00-2:00</code>)`,
        { parse_mode: "HTML", reply_markup: cancelKb() }
      );
      break;
    case "names":
      await startNames(ctx);
      break;
    case "export":
      await sendExcel(ctx);
      break;
    case "stats":
      await show(ctx, monthText(currentMonth()), monthKb(currentMonth()));
      break;
    case "month":
      if (idx === undefined) return;
      await show(ctx, monthText(idx), monthKb(idx));
      break;
  }
});

// ───────────────────────── Xatolar va ishga tushirish ─────────────────────────

bot.catch((err) => {
  console.error("Bot xatosi:", err.error);
});

// Buyruqlar menyusi ("/" bosilganda chiqadigan ro'yxat):
//  • shaxsiy chatda — /start, /menu (bot egasiga qo'shimcha: /smena, /admin, /export)
//  • guruhda oddiy a'zolarga — hech narsa; guruh adminlariga — faqat /close
const PRIVATE_COMMANDS = [
  { command: "start", description: "Botni boshlash" },
  { command: "menu", description: "Bosh menyu" },
];

const OWNER_COMMANDS = [
  ...PRIVATE_COMMANDS,
  { command: "smena", description: "Smena operatorlarini o'zgartirish" },
  { command: "admin", description: "Admin panel" },
  { command: "export", description: "Excel faylni yuklab olish" },
];

(async () => {
  try {
    await bot.api.deleteMyCommands(); // eski umumiy ro'yxat (guruhda /start, /menu chiqmasligi uchun)
    await bot.api.setMyCommands(PRIVATE_COMMANDS, { scope: { type: "all_private_chats" } });
    await bot.api.deleteMyCommands({ scope: { type: "all_group_chats" } });
    await bot.api.setMyCommands(
      [{ command: "close", description: "Arizani yopish (baholash so'rovi)" }],
      { scope: { type: "all_chat_administrators" } }
    );
  } catch (e) {
    console.warn("Buyruqlar menyusini yangilab bo'lmadi:", (e as Error).message);
  }
})();

bot.start({
  onStart: (me) => {
    console.log(`✅ Bot ishga tushdi: @${me.username}`);
    void ensureBroadcastTopic();
    void syncAll(allApplications()); // Excel fayl mavjud va yangi bo'lishi uchun
    if (!GROUP_MODE) {
      console.warn(
        "⚠️ Jonli yozishma (operator ⇄ foydalanuvchi) uchun ADMIN_ID guruh ID'si bo'lishi kerak (-100...) va guruhda Topics yoqilgan bo'lishi kerak."
      );
    }
    console.log(
      ADMIN_ID === undefined
        ? "ℹ️ ADMIN_ID: berilmagan (test rejimi — arizalar hech kimga yuborilmaydi, /close o'zingizda ishlaydi)"
        : `ℹ️ ADMIN_ID: ${ADMIN_ID} (${GROUP_MODE ? "guruh — har bir foydalanuvchi uchun topic" : "shaxsiy chat"})`
    );
  },
});

process.once("SIGINT", () => bot.stop());
process.once("SIGTERM", () => bot.stop());

export { bot }; // sinov uchun

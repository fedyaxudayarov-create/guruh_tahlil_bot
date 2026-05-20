import os
import json
import logging
import re
from datetime import datetime, timezone, timedelta
from collections import defaultdict
TZ = timezone(timedelta(hours=5))
from pathlib import Path

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, MessageHandler, CommandHandler,
    CallbackQueryHandler, filters, ContextTypes
)
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── SOZLAMALAR ───────────────────────────────────────────────────────────────
BOT_TOKEN = os.getenv("BOT_TOKEN", "")
OWNER_ID  = int(os.getenv("OWNER_ID", "0"))

# Railway Volume yoki lokal papka
DATA_DIR  = Path(os.getenv("DATA_DIR", "/data"))
DATA_FILE = DATA_DIR / "logs.json"

# ─── LOGGING ─────────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(message)s",
    level=logging.INFO,
)
log = logging.getLogger(__name__)

# ─── PERSISTENT STORAGE ──────────────────────────────────────────────────────
# { "guruh_id": [log_dict, ...] }
logs: dict[str, list[dict]] = defaultdict(list)

def save_logs():
    """Barcha loglarni JSON faylga saqlash"""
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(dict(logs), f, ensure_ascii=False, default=str)
    except Exception as e:
        log.error(f"save_logs xatosi: {e}")

def load_logs():
    """JSON fayldan loglarni yuklash"""
    global logs
    if DATA_FILE.exists():
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            logs = defaultdict(list, {k: v for k, v in data.items()})
            jami = sum(len(v) for v in logs.values())
            log.info(f"✅ {len(logs)} guruh, {jami} xabar yuklandi")
        except Exception as e:
            log.error(f"load_logs xatosi: {e}")
            logs = defaultdict(list)
    else:
        log.info("📂 Ma'lumot fayli yo'q — yangi boshlanadi")
        logs = defaultdict(list)

# ─── TAHLIL FUNKSIYALARI ──────────────────────────────────────────────────────

def skript(t: str) -> str:
    if not t: return "-"
    k = len(re.findall(r"[а-яёА-ЯЁ]", t))
    l = len(re.findall(r"[a-zA-ZʻʼGgQqHh]", t))
    if k > 0 and l > 0: return "Aralash"
    if k > l: return "Kirill"
    if l > k: return "Lotin"
    return "-"

SALOM_KW = [
    "ассалому алайкум","ас-саламу","салом","яхшимисиз","хайр",
    "assalomu alaykum","as-salamu","assalom","salom","yaxshimisiz","xayr",
    "добрый","привет","здравствуй",
]
TUSHUNDI_KW = [
    "тушунарли","tushunarli","тушундим","tushundim",
    "бажардим","bajardim","бажарилди","bajarildi",
    "хоп","hop","майли","mayli","ок","ok","👍","✅",
    "тайёр","tayyor","готово","готов","тугади","tugadi","қабул","qabul",
]
XATO_KW = [
    "хато","xato","нотўғри","noto'g'ri","тушунмадим","tushunmadim",
    "билмадим","bilmadim","узр","uzr","кечирасиз","kechirasiz","❌",
]
SAVOL_KW = [
    "?","савол","savol","нима","nima","қандай","qanday",
    "қачон","qachon","нега","nega","тушунтир","tushuntir",
]

def has_salom(t: str) -> bool:
    tl = t.lower()
    return any(k in tl for k in SALOM_KW)

def holat(t: str) -> str:
    tl = t.lower()
    if any(k in tl for k in TUSHUNDI_KW): return "Tushundi/Bajardi"
    if any(k in tl for k in XATO_KW):    return "Xato/Tushunmadi"
    if any(k in tl for k in SAVOL_KW):   return "Savol berdi"
    return "Oddiy izoh"

def tahlil_msg(message) -> dict:
    user    = message.from_user
    chat    = message.chat
    caption = (message.caption or "").strip()
    now = datetime.now(TZ)

    entry = {
        "ts"      : now.isoformat(),
        "sana"    : now.strftime("%Y-%m-%d"),
        "soat"    : now.strftime("%H:%M"),
        "guruh"   : chat.title or "Unknown",
        "guruh_id": str(chat.id),
        "ism"     : (f"{user.first_name or ''} {user.last_name or ''}".strip()
                     if user else "Unknown"),
        "username": (f"@{user.username}" if user and user.username
                     else str(user.id) if user else "-"),
        "user_id" : user.id if user else 0,
        "tur"     : "",
        "izoh"    : "",
    }

    if   message.photo:      entry["tur"] = "rasm";    entry["izoh"] = caption
    elif message.video:      entry["tur"] = "video";   entry["izoh"] = caption
    elif message.video_note: entry["tur"] = "dumaloq"
    elif message.voice:      entry["tur"] = "ovoz";    entry["izoh"] = caption
    elif message.audio:      entry["tur"] = "audio";   entry["izoh"] = caption
    elif message.document:   entry["tur"] = "fayl";    entry["izoh"] = caption
    elif message.sticker:    entry["tur"] = "stiker"
    elif message.text:       entry["tur"] = "matn";    entry["izoh"] = message.text.strip()
    else:                    entry["tur"] = "boshqa"

    return entry

# ─── EXCEL ───────────────────────────────────────────────────────────────────

SARLAVHA = [
    "№", "Ism", "Username", "Guruh",
    "Birinchi\nxabar", "Oxirgi\nxabar",
    "📷\nRasm", "🎥\nVideo", "⭕\nDumaloq",
    "🎤\nOvoz", "💬\nMatn", "📄\nFayl", "📦\nBoshqa",
    "Jami\nxabarlar",
    "👋\nSalomlashish",
    "✅\nTushundi", "❌\nXato", "❓\nSavol",
    "📝\nYozuv",
    "Izohlar (caption / matn)",
]
KENG = [4, 22, 16, 20, 14, 14,
        6, 6, 8, 6, 6, 6, 7,
        8,
        14,
        9, 7, 7,
        10,
        55]

C = {
    "sh_fon"  : "1F4E79", "sh_matn": "FFFFFF",
    "juft"    : "EBF3FB", "toq"    : "FFFFFF",
    "green"   : "C6EFCE", "red"    : "FFC7CE",
    "yellow"  : "FFEB9C", "blue"   : "BDD7EE",
    "salom"   : "E2EFDA", "header2": "2E4057",
    "hd2_matn": "FFFFFF",
}

def pf(hex_): return PatternFill("solid", fgColor=hex_)

chegara = Border(
    left=Side(style="thin"),  right=Side(style="thin"),
    top=Side(style="thin"),   bottom=Side(style="thin")
)

def sh(ws, row, col, val="", fon=None, bold=False, size=10,
       matn_rangi="000000", wrap=True, gorizontal="center"):
    c = ws.cell(row=row, column=col, value=val)
    if fon: c.fill = pf(fon)
    c.font = Font(bold=bold, size=size, color=matn_rangi)
    c.alignment = Alignment(horizontal=gorizontal, vertical="center", wrap_text=wrap)
    c.border = chegara
    return c

def excel_sheet(wb, guruh_nomi: str, glog: list):
    ws = wb.create_sheet(title=guruh_nomi[:28])

    for col, s in enumerate(SARLAVHA, 1):
        sh(ws, 1, col, s, fon=C["sh_fon"], bold=True, matn_rangi=C["sh_matn"])
    ws.row_dimensions[1].height = 32

    a: dict[int, dict] = {}
    for l in sorted(glog, key=lambda x: x.get("ts", "")):
        uid = l["user_id"]
        if uid not in a:
            a[uid] = {
                "ism": l["ism"], "username": l["username"], "guruh": l["guruh"],
                "birinchi": l["sana"] + " " + l["soat"],
                "oxirgi"  : l["sana"] + " " + l["soat"],
                "rasm":0,"video":0,"dumaloq":0,"ovoz":0,
                "matn":0,"fayl":0,"boshqa":0,
                "salom":0,"tushundi":0,"xato":0,"savol":0,
                "lotin":0,"kirill":0,"aralash":0,
                "izohlar": [],
            }
        d = a[uid]
        d["oxirgi"] = l["sana"] + " " + l["soat"]

        t = l["tur"]
        if   t == "rasm":    d["rasm"]    += 1
        elif t == "video":   d["video"]   += 1
        elif t == "dumaloq": d["dumaloq"] += 1
        elif t == "ovoz":    d["ovoz"]    += 1
        elif t == "matn":    d["matn"]    += 1
        elif t in ("fayl", "audio"): d["fayl"] += 1
        else:                d["boshqa"]  += 1

        izoh = l["izoh"].strip()
        if izoh:
            if has_salom(izoh):       d["salom"]    += 1
            h = holat(izoh)
            if   "Tushundi" in h: d["tushundi"] += 1
            elif "Xato"     in h: d["xato"]     += 1
            elif "Savol"    in h: d["savol"]    += 1
            s_ = skript(izoh)
            if   "Lotin"   in s_: d["lotin"]   += 1
            elif "Kirill"  in s_: d["kirill"]  += 1
            elif "Aralash" in s_: d["aralash"] += 1
            qisqa = (izoh[:120] + "…") if len(izoh) > 120 else izoh
            d["izohlar"].append(qisqa)

    for i, (uid, d) in enumerate(
        sorted(a.items(), key=lambda x: -(
            x[1]["rasm"]+x[1]["video"]+x[1]["dumaloq"]+
            x[1]["ovoz"]+x[1]["matn"]+x[1]["fayl"]+x[1]["boshqa"]
        )), 1
    ):
        row = i + 1
        fon = C["juft"] if i % 2 == 0 else C["toq"]
        jami = d["rasm"]+d["video"]+d["dumaloq"]+d["ovoz"]+d["matn"]+d["fayl"]+d["boshqa"]

        asosiy_yozuv = max(
            [("Lotin",d["lotin"]),("Kirill",d["kirill"]),("Aralash",d["aralash"])],
            key=lambda x: x[1]
        )[0] if (d["lotin"] or d["kirill"] or d["aralash"]) else "-"

        izoh_list = d["izohlar"]
        if len(izoh_list) <= 5:
            izoh_cell = "\n\n".join(f"• {x}" for x in izoh_list)
        else:
            izoh_cell = "\n\n".join(f"• {x}" for x in izoh_list[:5])
            izoh_cell += f"\n\n…va yana {len(izoh_list)-5} ta izoh"

        qiymatlar = [
            i, d["ism"], d["username"], d["guruh"],
            d["birinchi"], d["oxirgi"],
            d["rasm"] or "", d["video"] or "", d["dumaloq"] or "",
            d["ovoz"] or "", d["matn"] or "", d["fayl"] or "", d["boshqa"] or "",
            jami,
            d["salom"] or "",
            d["tushundi"] or "", d["xato"] or "", d["savol"] or "",
            asosiy_yozuv,
            izoh_cell,
        ]

        for col, q in enumerate(qiymatlar, 1):
            c = sh(ws, row, col, q, fon=fon)
            if col == 20:
                c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        if d["salom"]    > 0: ws.cell(row=row, column=15).fill = pf(C["salom"])
        if d["tushundi"] > 0: ws.cell(row=row, column=16).fill = pf(C["green"])
        if d["xato"]     > 0: ws.cell(row=row, column=17).fill = pf(C["red"])
        if d["savol"]    > 0: ws.cell(row=row, column=18).fill = pf(C["yellow"])
        for mc in [7, 8, 9, 10, 11, 12]:
            val = ws.cell(row=row, column=mc).value
            if val and val != "":
                ws.cell(row=row, column=mc).fill = pf(C["blue"])

        ws.row_dimensions[row].height = max(40, min(15 * max(1, len(izoh_list)), 200))

    for col, k in enumerate(KENG, 1):
        ws.column_dimensions[get_column_letter(col)].width = k
    ws.freeze_panes = "A2"

    jami_row = len(a) + 3
    c = ws.cell(row=jami_row, column=1, value="📊 JAMI (barcha ishtirokchilar)")
    c.font = Font(bold=True, size=10, color=C["hd2_matn"])
    c.fill = pf(C["header2"])
    ws.merge_cells(start_row=jami_row, start_column=1,
                   end_row=jami_row, end_column=len(SARLAVHA))
    ws.cell(row=jami_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[jami_row].height = 22

    jami_row2 = jami_row + 1
    jami_labels = [
        ("Ishtirokchilar", len(a)),
        ("Jami xabarlar",  sum(len([l for l in glog if l["user_id"] == uid]) for uid in a)),
        ("Rasm",    sum(d["rasm"]    for d in a.values())),
        ("Video",   sum(d["video"]   for d in a.values())),
        ("Dumaloq", sum(d["dumaloq"] for d in a.values())),
        ("Ovoz",    sum(d["ovoz"]    for d in a.values())),
        ("Matn",    sum(d["matn"]    for d in a.values())),
        ("Salomlashdi", sum(1 for d in a.values() if d["salom"] > 0)),
        ("Tushundi", sum(1 for d in a.values() if d["tushundi"] > 0)),
        ("Xato",    sum(1 for d in a.values() if d["xato"] > 0)),
    ]
    for col, (label, val) in enumerate(jami_labels, 1):
        sh(ws, jami_row2,   col, label, fon="4472C4", bold=True, matn_rangi="FFFFFF", size=9)
        sh(ws, jami_row2+1, col, val,   fon="EBF3FB", bold=True, size=10)

    ws.cell(row=jami_row2+3, column=1,
            value=f"Hisobot: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            ).font = Font(italic=True, color="808080", size=9)


# ─── EXCEL YUBORISH ───────────────────────────────────────────────────────────

async def send_excel(ctx, chat_id: int, target_gids: list[str], caption_info: str = ""):
    """Excel faylni yaratib yuborish"""
    tahlil_data = {}
    for gid in target_gids:
        if gid in logs and logs[gid]:
            tahlil_data[gid] = logs[gid]

    if not tahlil_data:
        await ctx.bot.send_message(chat_id=chat_id, text="📭 Bu guruhda hali ma'lumot yo'q.")
        return

    await ctx.bot.send_message(chat_id=chat_id, text="⏳ Excel tayyorlanmoqda…")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for gid, glog in tahlil_data.items():
        excel_sheet(wb, glog[-1]["guruh"], glog)

    fayl_nomi = f"tahlil_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    fayl_path = f"/tmp/{fayl_nomi}"
    wb.save(fayl_path)

    jami_xabar = sum(len(v) for v in tahlil_data.values())
    n_guruh    = len(tahlil_data)
    guruh_nomlari = ", ".join(
        logs[gid][-1]["guruh"] for gid in tahlil_data if logs[gid]
    )

    with open(fayl_path, "rb") as f:
        await ctx.bot.send_document(
            chat_id=chat_id,
            document=f,
            filename=fayl_nomi,
            caption=(
                f"📊 *Excel hisobot tayyor!*\n"
                f"🏷 Guruh(lar): {guruh_nomlari}\n"
                f"👥 Guruhlar: {n_guruh} | 📨 Xabarlar: {jami_xabar}"
                + (f"\n{caption_info}" if caption_info else "")
            ),
            parse_mode="Markdown",
        )
    os.remove(fayl_path)


# ─── INLINE KEYBOARD YORDAMCHISI ─────────────────────────────────────────────

def guruhlar_keyboard(prefix: str = "report") -> InlineKeyboardMarkup | None:
    """Guruhlar ro'yxatidan inline keyboard yaratish"""
    if not logs:
        return None

    klaviatura = []
    for gid, glog in logs.items():
        if not glog:
            continue
        nom  = glog[-1]["guruh"]
        uids = len(set(l["user_id"] for l in glog))
        nxbr = len(glog)
        klaviatura.append([
            InlineKeyboardButton(
                f"📁 {nom[:30]} ({nxbr} xabar, {uids} kishi)",
                callback_data=f"{prefix}:{gid}"
            )
        ])

    if prefix == "report":
        klaviatura.append([
            InlineKeyboardButton("📊 BARCHA guruhlar", callback_data=f"{prefix}:ALL")
        ])

    return InlineKeyboardMarkup(klaviatura) if klaviatura else None


# ─── HANDLERS ─────────────────────────────────────────────────────────────────

async def guruh_xabar(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Guruh xabarlarini yig'ish"""
    msg = update.message or update.channel_post
    if not msg:
        return
    entry = tahlil_msg(msg)
    gid = entry["guruh_id"]
    logs[gid].append(entry)

    # Har 20 xabardan keyin saqlash (performance uchun)
    if len(logs[gid]) % 20 == 0:
        save_logs()


async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != OWNER_ID:
        return
    await update.message.reply_text(
        "🤖 *Guruh Tahlil Boti*\n\n"
        "Bot indamay yig'adi — hech narsa yuborilmaydi.\n\n"
        "📋 *Buyruqlar:*\n"
        "/report — Guruh tanlash va Excel hisobot\n"
        "/stats — Qisqa statistika\n"
        "/guruhlar — Guruhlar ro'yxati\n"
        "/clear — Ma'lumotlarni tozalash\n"
        "/save — Qo'lda saqlash",
        parse_mode="Markdown",
    )


async def report_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Guruh tanlash uchun inline keyboard ko'rsatish"""
    if update.effective_user.id != OWNER_ID:
        return

    kb = guruhlar_keyboard("report")
    if kb is None:
        await update.message.reply_text("📭 Hozircha hech qanday guruh ma'lumoti yo'q.")
        return

    await update.message.reply_text(
        "📊 *Qaysi guruh uchun Excel hisobot tayyorlash?*\n\n"
        "Guruhni tanlang 👇",
        reply_markup=kb,
        parse_mode="Markdown",
    )


async def stats_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Statistika — guruh tanlash"""
    if update.effective_user.id != OWNER_ID:
        return

    if not logs:
        await update.message.reply_text("📭 Ma'lumot yo'q.")
        return

    # Agar bitta guruh bo'lsa, to'g'ridan statistika
    if len([g for g in logs if logs[g]]) == 1:
        gid = next(g for g in logs if logs[g])
        await show_stats(update, ctx, gid)
        return

    kb = guruhlar_keyboard("stats")
    await update.message.reply_text(
        "📊 *Qaysi guruh statistikasini ko'rish?*",
        reply_markup=kb,
        parse_mode="Markdown",
    )


async def show_stats(update_or_query, ctx, gid: str):
    """Bir guruh statistikasini ko'rsatish"""
    glog = logs.get(gid, [])
    if not glog:
        text = "📭 Bu guruhda ma'lumot yo'q."
    else:
        nom  = glog[-1]["guruh"]
        uids = len(set(l["user_id"] for l in glog))
        text = (
            f"📊 *{nom}* statistikasi\n\n"
            f"👤 Ishtirokchilar: *{uids}*\n"
            f"📨 Jami xabarlar: *{len(glog)}*\n"
            f"📷 Rasm: {sum(1 for l in glog if l['tur']=='rasm')} | "
            f"🎥 Video: {sum(1 for l in glog if l['tur']=='video')} | "
            f"⭕ Dumaloq: {sum(1 for l in glog if l['tur']=='dumaloq')}\n"
            f"🎤 Ovoz: {sum(1 for l in glog if l['tur']=='ovoz')} | "
            f"💬 Matn: {sum(1 for l in glog if l['tur']=='matn')}\n"
            f"📄 Fayl/Audio: {sum(1 for l in glog if l['tur'] in ('fayl','audio'))}\n\n"
            f"📅 Oxirgi xabar: {glog[-1]['sana']} {glog[-1]['soat']}"
        )

    if hasattr(update_or_query, "message"):
        await update_or_query.message.reply_text(text, parse_mode="Markdown")
    else:
        await update_or_query.edit_message_text(text, parse_mode="Markdown")


async def guruhlar_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Barcha guruhlar ro'yxati"""
    if update.effective_user.id != OWNER_ID:
        return
    if not any(logs.values()):
        await update.message.reply_text("📭 Hech qanday guruh yo'q.")
        return

    javob = "👥 *Guruhlar ro'yxati:*\n\n"
    for gid, glog in logs.items():
        if not glog:
            continue
        nom  = glog[-1]["guruh"]
        uids = len(set(l["user_id"] for l in glog))
        javob += (
            f"🏷 *{nom}*\n"
            f"  ID: `{gid}`\n"
            f"  📨 Xabarlar: {len(glog)} | 👤 Kishilar: {uids}\n\n"
        )
    await update.message.reply_text(javob, parse_mode="Markdown")


async def save_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Qo'lda saqlash"""
    if update.effective_user.id != OWNER_ID:
        return
    save_logs()
    await update.message.reply_text("💾 Ma'lumotlar saqlandi.")


async def clear_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Barcha ma'lumotlarni tozalash — tasdiqlash so'rash"""
    if update.effective_user.id != OWNER_ID:
        return

    kb = InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Ha, tozala", callback_data="clear:confirm"),
        InlineKeyboardButton("❌ Bekor qil",  callback_data="clear:cancel"),
    ]])
    await update.message.reply_text(
        "⚠️ *Barcha ma'lumotlar o'chiriladi!*\n\nDavom etamizmi?",
        reply_markup=kb,
        parse_mode="Markdown",
    )


# ─── CALLBACK QUERY HANDLER ───────────────────────────────────────────────────

async def callback_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.from_user.id != OWNER_ID:
        await query.answer("❌ Ruxsat yo'q", show_alert=True)
        return

    data = query.data

    # ── /report callback ───────────────────────────────────────
    if data.startswith("report:"):
        gid_val = data.split(":", 1)[1]
        await query.edit_message_text("⏳ Excel tayyorlanmoqda…")

        if gid_val == "ALL":
            target_gids = list(logs.keys())
            info = "Barcha guruhlar"
        else:
            target_gids = [gid_val]
            info = ""

        await send_excel(ctx, OWNER_ID, target_gids, info)

    # ── /stats callback ────────────────────────────────────────
    elif data.startswith("stats:"):
        gid = data.split(":", 1)[1]
        await show_stats(query, ctx, gid)

    # ── /clear callback ────────────────────────────────────────
    elif data == "clear:confirm":
        logs.clear()
        save_logs()
        await query.edit_message_text("🗑 Barcha ma'lumotlar tozalandi.")
    elif data == "clear:cancel":
        await query.edit_message_text("✅ Bekor qilindi.")


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    if not BOT_TOKEN:
        raise ValueError("❌ BOT_TOKEN o'rnatilmagan!")
    if not OWNER_ID:
        raise ValueError("❌ OWNER_ID o'rnatilmagan!")

    # Ma'lumotlarni yuklash
    load_logs()

    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start",    start))
    app.add_handler(CommandHandler("stats",    stats_cmd))
    app.add_handler(CommandHandler("guruhlar", guruhlar_cmd))
    app.add_handler(CommandHandler("report",   report_cmd))
    app.add_handler(CommandHandler("clear",    clear_cmd))
    app.add_handler(CommandHandler("save",     save_cmd))
    app.add_handler(CallbackQueryHandler(callback_handler))

    app.add_handler(MessageHandler(
        filters.ChatType.GROUPS & (
            filters.TEXT | filters.PHOTO | filters.VIDEO |
            filters.VIDEO_NOTE | filters.VOICE | filters.AUDIO |
            filters.Document.ALL | filters.Sticker.ALL
        ),
        guruh_xabar
    ))

    log.info("🤖 Bot ishga tushdi — silent mode")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()

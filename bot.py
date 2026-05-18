import os
import logging
import re
from datetime import datetime
from collections import defaultdict

from telegram import Update
from telegram.ext import (
    Application, MessageHandler, CommandHandler,
    filters, ContextTypes
)
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── SOZLAMALAR ───────────────────────────────────────────────────────────────
BOT_TOKEN = os.getenv("BOT_TOKEN", "")
OWNER_ID  = int(os.getenv("OWNER_ID", "0"))

# ─── MA'LUMOTLAR SAQLASH ──────────────────────────────────────────────────────
logs: dict[int, list[dict]] = defaultdict(list)

# ─── TAHLIL FUNKSIYALARI ─────────────────────────────────────────────────────

def skript_aniqlash(matn: str) -> str:
    if not matn:
        return "-"
    kirill = len(re.findall(r"[а-яёА-ЯЁ]", matn))
    lotin  = len(re.findall(r"[a-zA-ZʻʼGgQqHh]", matn))
    if kirill > 0 and lotin > 0:
        return "Aralash"
    if kirill > lotin:
        return "Kirill"
    if lotin > kirill:
        return "Lotin"
    return "-"

SALOM_KW = [
    "ассалому алайкум", "ас-саламу алайкум", "ас саламу алайкум",
    "ваалайкум", "ва алайкум", "салом", "яхшимисиз", "хайр",
    "assalomu alaykum", "as-salamu alaykum", "assalom",
    "va alaykum", "vaalaykum", "salom", "yaxshimisiz", "xayr",
    "добрый", "привет", "здравствуй",
]

def salom_aniqlash(matn: str) -> str:
    t = matn.lower()
    for kw in SALOM_KW:
        if kw in t:
            return "Ha"
    return "Yo'q"

TUSHUNDI_KW = [
    "тушунарли", "tushunarli", "тушундим", "tushundim",
    "бажардим", "bajardim", "бажарилди", "bajarildi",
    "хоп", "hop", "яхши", "yaxshi", "майли", "mayli",
    "ок", "ok", "👍", "✅", "тайёр", "tayyor", "готово", "готов",
    "тугади", "tugadi", "қабул", "qabul",
]
XATO_KW = [
    "хато", "xato", "нотўғри", "noto'g'ri", "тушунмадим", "tushunmadim",
    "билмадим", "bilmadim", "узр", "uzr", "кечирасиз", "kechirasiz",
    "❌", "йўқ", "yo'q", "нотугри",
]
SAVOL_KW = [
    "?", "савол", "savol", "нима", "nima", "қандай", "qanday",
    "қачон", "qachon", "нега", "nega", "ким", "kim",
    "тушунтир", "tushuntir", "қаерда", "qayerda",
]

def holat_aniqlash(matn: str) -> str:
    t = matn.lower()
    for kw in TUSHUNDI_KW:
        if kw in t:
            return "Tushundi/Bajardi"
    for kw in XATO_KW:
        if kw in t:
            return "Xato/Tushunmadi"
    for kw in SAVOL_KW:
        if kw in t:
            return "Savol berdi"
    return "Oddiy izoh"

def xabar_tahlil(message) -> dict:
    user    = message.from_user
    chat    = message.chat
    caption = (message.caption or "").strip()

    log = {
        "vaqt"        : datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "sana"        : datetime.now().strftime("%Y-%m-%d"),
        "soat"        : datetime.now().strftime("%H:%M"),
        "guruh"       : chat.title or "Unknown",
        "guruh_id"    : chat.id,
        "ism"         : (f"{user.first_name or ''} {user.last_name or ''}".strip()
                         if user else "Unknown"),
        "username"    : (f"@{user.username}" if user and user.username
                         else str(user.id) if user else "-"),
        "user_id"     : user.id if user else 0,
        "rasm"        : "",
        "video"       : "",
        "dumaloq"     : "",
        "ovoz"        : "",
        "matn_tur"    : "",
        "boshqa"      : "",
        "izoh"        : "",
        "salomlashish": "",
        "holat"       : "",
        "yozuv"       : "",
    }

    if message.photo:
        log["rasm"]  = "Ha"
        log["izoh"]  = caption
    elif message.video:
        log["video"] = "Ha"
        log["izoh"]  = caption
    elif message.video_note:
        log["dumaloq"] = "Ha"
    elif message.voice:
        log["ovoz"]  = "Ha"
        log["izoh"]  = caption
    elif message.audio:
        log["boshqa"] = "Audio"
        log["izoh"]   = caption
    elif message.document:
        log["boshqa"] = "Fayl"
        log["izoh"]   = caption
    elif message.sticker:
        log["boshqa"] = "Stiker"
    elif message.text:
        log["matn_tur"] = "Ha"
        log["izoh"]     = message.text.strip()
    else:
        log["boshqa"] = "Boshqa"

    t = log["izoh"]
    if t:
        log["salomlashish"] = salom_aniqlash(t)
        log["holat"]        = holat_aniqlash(t)
        log["yozuv"]        = skript_aniqlash(t)

    return log


# ─── HANDLER: FAQAT YIG'ADI ──────────────────────────────────────────────────

async def guruh_xabar(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = update.message or update.channel_post
    if not msg:
        return
    log = xabar_tahlil(msg)
    logs[log["guruh_id"]].append(log)
    # Hech narsa yuborilmaydi


# ─── EXCEL ───────────────────────────────────────────────────────────────────

SARLAVHALAR = [
    "№", "Sana", "Soat", "Guruh", "Ism", "Username",
    "📷 Rasm", "🎥 Video", "⭕ Dumaloq", "🎤 Ovoz", "💬 Matn", "📦 Boshqa",
    "Izoh / Caption",
    "👋 Salomlashish", "📋 Holat", "📝 Yozuv",
]
KENGLIKLARI = [5, 12, 8, 22, 22, 16,
               8, 8, 10, 8, 8, 10,
               45,
               16, 22, 10]

RANG = {
    "sarlavha_fon" : "1F4E79",
    "sarlavha_matn": "FFFFFF",
    "juft"         : "D6E4F0",
    "toq"          : "FFFFFF",
    "tushundi"     : "C6EFCE",
    "xato"         : "FFC7CE",
    "savol"        : "FFEB9C",
    "salom_ha"     : "E2EFDA",
    "media_ha"     : "BDD7EE",
}


def excel_sheet_yasash(wb, guruh_nomi: str, glog: list):
    ws = wb.create_sheet(title=guruh_nomi[:28])

    chegara        = Border(left=Side(style="thin"), right=Side(style="thin"),
                            top=Side(style="thin"),  bottom=Side(style="thin"))
    sarlavha_shrift = Font(color=RANG["sarlavha_matn"], bold=True, size=10)
    sarlavha_fon    = PatternFill("solid", fgColor=RANG["sarlavha_fon"])
    markaz          = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 1-qator sarlavhalar
    for col, s in enumerate(SARLAVHALAR, 1):
        c = ws.cell(row=1, column=col, value=s)
        c.fill = sarlavha_fon
        c.font = sarlavha_shrift
        c.alignment = markaz
        c.border = chegara
    ws.row_dimensions[1].height = 28

    # Ma'lumot qatorlari
    for i, l in enumerate(glog, 1):
        row = i + 1
        fon = RANG["juft"] if i % 2 == 0 else RANG["toq"]
        qiymatlar = [
            i, l["sana"], l["soat"], l["guruh"], l["ism"], l["username"],
            l["rasm"], l["video"], l["dumaloq"], l["ovoz"], l["matn_tur"], l["boshqa"],
            l["izoh"],
            l["salomlashish"], l["holat"], l["yozuv"],
        ]
        for col, q in enumerate(qiymatlar, 1):
            c = ws.cell(row=row, column=col, value=q)
            c.fill = PatternFill("solid", fgColor=fon)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = chegara

        # Izoh chap hizalama
        ws.cell(row=row, column=13).alignment = Alignment(
            horizontal="left", vertical="center", wrap_text=True)

        # Rang bilan belgilash
        holat = l["holat"]
        if "Tushundi" in holat:
            ws.cell(row=row, column=15).fill = PatternFill("solid", fgColor=RANG["tushundi"])
        elif "Xato" in holat:
            ws.cell(row=row, column=15).fill = PatternFill("solid", fgColor=RANG["xato"])
        elif "Savol" in holat:
            ws.cell(row=row, column=15).fill = PatternFill("solid", fgColor=RANG["savol"])

        if l["salomlashish"] == "Ha":
            ws.cell(row=row, column=14).fill = PatternFill("solid", fgColor=RANG["salom_ha"])

        for mc in [7, 8, 9, 10, 11]:
            if ws.cell(row=row, column=mc).value == "Ha":
                ws.cell(row=row, column=mc).fill = PatternFill("solid", fgColor=RANG["media_ha"])

    # Ustun kengliklari
    for col, k in enumerate(KENGLIKLARI, 1):
        ws.column_dimensions[get_column_letter(col)].width = k
    ws.freeze_panes = "A2"

    # ── ISHTIROKCHI KESIMIDA XULOSA ──────────────────────────
    boshlash = len(glog) + 3

    # Xulosa sarlavhasi
    c = ws.cell(row=boshlash, column=1, value="📊 ISHTIROKCHI KESIMIDA XULOSA")
    c.font = Font(bold=True, size=11, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="2E4057")
    ws.merge_cells(start_row=boshlash, start_column=1,
                   end_row=boshlash, end_column=len(SARLAVHALAR))
    ws.cell(row=boshlash, column=1).alignment = Alignment(horizontal="center")

    xulosa_s = [
        "Ism", "Username", "Jami",
        "📷 Rasm", "🎥 Video", "⭕ Dumaloq", "🎤 Ovoz", "💬 Matn",
        "👋 Salomlashdi", "✅ Tushundi", "❌ Xato", "❓ Savol", "📝 Asosiy yozuv",
    ]
    for col, s in enumerate(xulosa_s, 1):
        c = ws.cell(row=boshlash + 1, column=col, value=s)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="4472C4")
        c.alignment = Alignment(horizontal="center")
        c.border = chegara

    # Ishtirokchilar yig'ish
    ishtirokchilar: dict = {}
    for l in glog:
        uid = l["user_id"]
        if uid not in ishtirokchilar:
            ishtirokchilar[uid] = {
                "ism": l["ism"], "username": l["username"],
                "jami": 0, "rasm": 0, "video": 0, "dumaloq": 0, "ovoz": 0, "matn": 0,
                "salom": 0, "tushundi": 0, "xato": 0, "savol": 0,
                "lotin": 0, "kirill": 0, "aralash": 0,
            }
        d = ishtirokchilar[uid]
        d["jami"] += 1
        if l["rasm"]    == "Ha": d["rasm"]    += 1
        if l["video"]   == "Ha": d["video"]   += 1
        if l["dumaloq"] == "Ha": d["dumaloq"] += 1
        if l["ovoz"]    == "Ha": d["ovoz"]    += 1
        if l["matn_tur"]== "Ha": d["matn"]    += 1
        if l["salomlashish"] == "Ha":        d["salom"]    += 1
        if "Tushundi" in l["holat"]:         d["tushundi"] += 1
        elif "Xato"   in l["holat"]:         d["xato"]     += 1
        elif "Savol"  in l["holat"]:         d["savol"]    += 1
        if   "Lotin"   in l["yozuv"]:        d["lotin"]    += 1
        elif "Kirill"  in l["yozuv"]:        d["kirill"]   += 1
        elif "Aralash" in l["yozuv"]:        d["aralash"]  += 1

    for j, (uid, d) in enumerate(
        sorted(ishtirokchilar.items(), key=lambda x: -x[1]["jami"])
    ):
        row = boshlash + 2 + j
        asosiy = max(
            [("Lotin", d["lotin"]), ("Kirill", d["kirill"]), ("Aralash", d["aralash"])],
            key=lambda x: x[1]
        )[0] if (d["lotin"] or d["kirill"] or d["aralash"]) else "-"

        qiymatlar = [
            d["ism"], d["username"], d["jami"],
            d["rasm"], d["video"], d["dumaloq"], d["ovoz"], d["matn"],
            d["salom"], d["tushundi"], d["xato"], d["savol"], asosiy,
        ]
        fon = "EBF3FB" if j % 2 == 0 else "FFFFFF"
        for col, q in enumerate(qiymatlar, 1):
            c = ws.cell(row=row, column=col, value=q)
            c.fill = PatternFill("solid", fgColor=fon)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = chegara
        ws.cell(row=row, column=1).alignment = Alignment(horizontal="left")

        if d["salom"] > 0:
            ws.cell(row=row, column=9).fill  = PatternFill("solid", fgColor=RANG["salom_ha"])
        if d["tushundi"] > 0:
            ws.cell(row=row, column=10).fill = PatternFill("solid", fgColor=RANG["tushundi"])
        if d["xato"] > 0:
            ws.cell(row=row, column=11).fill = PatternFill("solid", fgColor=RANG["xato"])

    oxir = boshlash + 2 + len(ishtirokchilar)
    ws.cell(row=oxir + 1, column=1,
            value=f"Hisobot: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            ).font = Font(italic=True, color="808080")


async def report(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != OWNER_ID:
        return

    matn = update.message.text or ""
    target_gid = None
    if "_" in matn:
        try:
            target_gid = int(matn.split("_", 1)[1].strip())
        except ValueError:
            pass

    tahlil_logs = {target_gid: logs.get(target_gid, [])} if target_gid else dict(logs)

    if not any(tahlil_logs.values()):
        await update.message.reply_text("📭 Hozircha ma'lumot yo'q.")
        return

    await update.message.reply_text("⏳ Excel tayyorlanmoqda…")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for gid, glog in tahlil_logs.items():
        if glog:
            excel_sheet_yasash(wb, glog[-1]["guruh"], glog)

    fayl_nomi = f"/tmp/tahlil_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(fayl_nomi)

    jami = sum(len(v) for v in tahlil_logs.values())
    guruh_soni = sum(1 for v in tahlil_logs.values() if v)

    with open(fayl_nomi, "rb") as f:
        await ctx.bot.send_document(
            chat_id=OWNER_ID,
            document=f,
            filename=f"guruh_tahlil_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            caption=(
                f"📊 *Excel hisobot tayyor!*\n"
                f"👥 Guruhlar: {guruh_soni}\n"
                f"📨 Jami xabarlar: {jami}"
            ),
            parse_mode="Markdown",
        )
    os.remove(fayl_nomi)


async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != OWNER_ID:
        return
    await update.message.reply_text(
        "🤖 *Guruh Tahlil Boti*\n\n"
        "Bot guruhda *indamay* yig'adi.\n\n"
        "📋 *Buyruqlar:*\n"
        "/report — Excel hisobot\n"
        "/stats — Qisqa statistika\n"
        "/guruhlar — Guruhlar ro'yxati\n"
        "/clear — Ma'lumotlarni tozalash",
        parse_mode="Markdown",
    )

async def guruhlar(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != OWNER_ID:
        return
    if not logs:
        await update.message.reply_text("📭 Hech qanday guruh yo'q.")
        return
    javob = "👥 *Kuzatilayotgan guruhlar:*\n\n"
    for gid, glog in logs.items():
        nom = glog[-1]["guruh"] if glog else "Unknown"
        javob += f"• *{nom}*\n  ID: `{gid}` | Xabarlar: {len(glog)}\n\n"
    await update.message.reply_text(javob, parse_mode="Markdown")

async def stats(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != OWNER_ID:
        return
    if not logs:
        await update.message.reply_text("📭 Ma'lumot yo'q.")
        return
    javob = "📊 *Statistika*\n\n"
    for gid, glog in logs.items():
        if not glog:
            continue
        nom     = glog[-1]["guruh"]
        rasm    = sum(1 for l in glog if l["rasm"]    == "Ha")
        video   = sum(1 for l in glog if l["video"]   == "Ha")
        dumaloq = sum(1 for l in glog if l["dumaloq"] == "Ha")
        salom   = sum(1 for l in glog if l["salomlashish"] == "Ha")
        tush    = sum(1 for l in glog if "Tushundi" in l["holat"])
        xato    = sum(1 for l in glog if "Xato"     in l["holat"])
        ishtirokchi = len(set(l["user_id"] for l in glog))
        javob += (
            f"🏷 *{nom}*\n"
            f"  👤 Ishtirokchilar: {ishtirokchi}\n"
            f"  📨 Jami xabarlar: {len(glog)}\n"
            f"  📷 Rasm: {rasm}  🎥 Video: {video}  ⭕ Dumaloq: {dumaloq}\n"
            f"  👋 Salomlashdi: {salom}\n"
            f"  ✅ Tushundi: {tush}  ❌ Xato: {xato}\n\n"
        )
    await update.message.reply_text(javob, parse_mode="Markdown")

async def clear(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != OWNER_ID:
        return
    logs.clear()
    await update.message.reply_text("🗑 Barcha ma'lumotlar tozalandi.")


def main():
    logging.basicConfig(
        format="%(asctime)s | %(levelname)s | %(message)s",
        level=logging.INFO,
    )
    if not BOT_TOKEN:
        raise ValueError("BOT_TOKEN o'rnatilmagan!")
    if not OWNER_ID:
        raise ValueError("OWNER_ID o'rnatilmagan!")

    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start",    start))
    app.add_handler(CommandHandler("stats",    stats))
    app.add_handler(CommandHandler("guruhlar", guruhlar))
    app.add_handler(CommandHandler("report",   report))
    app.add_handler(CommandHandler("clear",    clear))

    guruh_filtri = filters.ChatType.GROUPS & (
        filters.TEXT
        | filters.PHOTO
        | filters.VIDEO
        | filters.VIDEO_NOTE
        | filters.VOICE
        | filters.AUDIO
        | filters.Document.ALL
        | filters.Sticker.ALL
    )
    app.add_handler(MessageHandler(guruh_filtri, guruh_xabar))

    logging.info("🤖 Bot ishga tushdi — silent mode")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()

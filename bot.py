import os, json, logging, re
from datetime import datetime, timezone, timedelta
from collections import defaultdict
from pathlib import Path

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import (
    Application, MessageHandler, CommandHandler,
    CallbackQueryHandler, filters, ContextTypes
)
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ═══════════════════════════════════════════════════════════════
#  SOZLAMALAR
# ═══════════════════════════════════════════════════════════════
BOT_TOKEN = os.getenv("BOT_TOKEN", "")
OWNER_ID  = int(os.getenv("OWNER_ID", "0"))
TZ        = timezone(timedelta(hours=5))

_extra = os.getenv("ADMIN_IDS", "")
ADMIN_IDS: set[int] = {OWNER_ID} | {
    int(x.strip()) for x in _extra.split(",") if x.strip().isdigit()
}
def is_admin(uid: int) -> bool: return uid in ADMIN_IDS

DATA_DIR  = Path(os.getenv("DATA_DIR", "/data"))
DATA_FILE = DATA_DIR / "logs.json"
MBRS_FILE = DATA_DIR / "members.json"

logging.basicConfig(format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO)
log = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════
#  STORAGE
# ═══════════════════════════════════════════════════════════════
logs:    dict[str, list[dict]] = defaultdict(list)
members: dict[str, dict]       = defaultdict(dict)

def save_all():
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(dict(logs), f, ensure_ascii=False, default=str)
        with open(MBRS_FILE, "w", encoding="utf-8") as f:
            json.dump(dict(members), f, ensure_ascii=False, default=str)
    except Exception as e:
        log.error(f"save_all: {e}")

def load_all():
    global logs, members
    if DATA_FILE.exists():
        try:
            with open(DATA_FILE, encoding="utf-8") as f:
                logs = defaultdict(list, json.load(f))
        except Exception as e:
            log.error(f"load logs: {e}")
    if MBRS_FILE.exists():
        try:
            with open(MBRS_FILE, encoding="utf-8") as f:
                members = defaultdict(dict, json.load(f))
        except Exception as e:
            log.error(f"load members: {e}")

def reg_member(gid: str, uid: int, ism: str, username: str):
    members[gid][str(uid)] = {"ism": ism, "username": username}

# ═══════════════════════════════════════════════════════════════
#  YORDAMCHI FUNKSIYALAR
# ═══════════════════════════════════════════════════════════════
def msg_link(chat_id_str: str, msg_id: int) -> str:
    if not msg_id: return ""
    try:
        abs_cid = str(abs(int(chat_id_str)))
        pure = abs_cid[3:] if abs_cid.startswith("100") and len(abs_cid) >= 12 else abs_cid
        return f"https://t.me/c/{pure}/{msg_id}"
    except: return ""

def skript(t: str) -> str:
    if not t: return "-"
    k = len(re.findall(r"[а-яёА-ЯЁ]", t))
    l = len(re.findall(r"[a-zA-ZʻʼGgQqHh]", t))
    if k > 0 and l > 0: return "Aralash"
    if k > l: return "Kirill"
    if l > k: return "Lotin"
    return "-"

SALOM_KW    = ["ассалому алайкум","салом","assalomu alaykum","assalom","salom"]
TUSHUNDI_KW = ["тушунарли","tushunarli","тушундим","tushundim","бажардим","bajardim",
               "хоп","hop","майли","mayli","ок","ok","👍","✅","тайёр","tayyor","қабул","qabul"]
XATO_KW     = ["хато","xato","нотўғри","noto'g'ri","тушунмадим","tushunmadim","❌"]
SAVOL_KW    = ["?","савол","savol","нима","nima","қандай","qanday","нега","nega"]

def has_salom(t): return any(k in t.lower() for k in SALOM_KW)
def holat(t):
    tl = t.lower()
    if any(k in tl for k in TUSHUNDI_KW): return "Tushundi ✅"
    if any(k in tl for k in XATO_KW):    return "Xato ❌"
    if any(k in tl for k in SAVOL_KW):   return "Savol ❓"
    return "-"

TUR_EMOJI = {
    "rasm": "📷 Rasm", "video": "🎥 Video", "dumaloq": "⭕ Dumaloq",
    "ovoz": "🎤 Ovoz", "matn": "💬 Matn", "fayl": "📄 Fayl",
    "audio": "🎵 Audio", "stiker": "🎭 Stiker", "boshqa": "📦 Boshqa",
}

def tahlil_msg(message) -> dict:
    user    = message.from_user
    chat    = message.chat
    caption = (message.caption or "").strip()
    now     = datetime.now(TZ)
    gid     = str(chat.id)
    mid     = message.message_id
    entry = {
        "ts"      : now.isoformat(),
        "sana"    : now.strftime("%Y-%m-%d"),
        "soat"    : now.strftime("%H:%M"),
        "guruh"   : chat.title or "Unknown",
        "guruh_id": gid,
        "ism"     : (f"{user.first_name or ''} {user.last_name or ''}".strip() if user else "Unknown"),
        "username": (f"@{user.username}" if user and user.username else str(user.id) if user else "-"),
        "user_id" : user.id if user else 0,
        "msg_id"  : mid,
        "link"    : msg_link(gid, mid),
        "tur": "", "izoh": "",
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

# ═══════════════════════════════════════════════════════════════
#  BOSHQARUV PANELI (Reply Keyboard)
# ═══════════════════════════════════════════════════════════════
PANEL = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton("📊 Hisobot"),    KeyboardButton("📈 Statistika")],
        [KeyboardButton("👥 A'zolar"),    KeyboardButton("🏘 Guruhlar")],
        [KeyboardButton("💾 Saqlash"),    KeyboardButton("🗑 Tozalash")],
    ],
    resize_keyboard=True,
    persistent=True,
)

# ═══════════════════════════════════════════════════════════════
#  EXCEL
# ═══════════════════════════════════════════════════════════════
C = {
    "hdr"    : "1F4E79", "hdr_t"  : "FFFFFF",
    "hdr2"   : "C00000", "hdr2_t" : "FFFFFF",
    "juft"   : "EBF3FB", "toq"    : "FFFFFF",
    "green"  : "C6EFCE", "red_bg" : "FFE0E0",
    "yellow" : "FFEB9C", "blue"   : "DDEBF7",
    "salom"  : "E2EFDA", "sum_bg" : "2E4057",
    "sum_t"  : "FFFFFF", "cont"   : "F5F5F5",
}
chegara = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)
def pf(h): return PatternFill("solid", fgColor=h)

def sc(ws, row, col, val="", fon=None, bold=False, size=10,
       color="000000", wrap=True, align="center", italic=False):
    c = ws.cell(row=row, column=col, value=val)
    if fon:  c.fill = pf(fon)
    c.font      = Font(bold=bold, size=size, color=color, italic=italic)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    c.border    = chegara
    return c

# Ustunlar: №, Ism, Username, Guruh, Sana, Vaqt, Tur, Holat, Izoh, Link
COLS = [
    ("№",        4),
    ("Ism",     24),
    ("Username",16),
    ("Guruh",   20),
    ("📅 Sana", 12),
    ("🕐 Vaqt",  8),
    ("Tur",     12),
    ("Holat",   14),
    ("Izoh",    40),
    ("🔗 Link", 42),
]
HDR = [c[0] for c in COLS]
WDT = [c[1] for c in COLS]

def sheet_yuborganlar(wb, nom: str, glog: list, sana_filter: str | None = None) -> set:
    ws  = wb.create_sheet(title=f"✅ {nom}"[:30])
    hdr = 1
    if sana_filter:
        c = ws.cell(row=1, column=1, value=f"📅 Sana bo'yicha filter: {sana_filter}")
        c.font = Font(bold=True, size=10, color="1F4E79")
        c.alignment = Alignment(horizontal="center")
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS))
        ws.row_dimensions[1].height = 16
        hdr = 2

    for col, h in enumerate(HDR, 1):
        sc(ws, hdr, col, h, fon=C["hdr"], bold=True, color=C["hdr_t"])
    ws.row_dimensions[hdr].height = 30

    if sana_filter:
        glog = [l for l in glog if l.get("sana") == sana_filter]

    # Kishilar bo'yicha guruhlash (vaqt tartibida)
    kishi_msgs: dict[int, list[dict]] = {}
    kishi_info: dict[int, dict]       = {}
    for l in sorted(glog, key=lambda x: x.get("ts", "")):
        uid = l["user_id"]
        if uid not in kishi_info:
            kishi_info[uid] = {"ism": l["ism"], "username": l["username"], "guruh": l["guruh"]}
            kishi_msgs[uid] = []
        kishi_msgs[uid].append(l)

    data_row = hdr + 1
    kishi_n  = 0

    for uid, msgs in sorted(kishi_msgs.items(),
                             key=lambda x: -len(x[1])):
        kishi_n  += 1
        info      = kishi_info[uid]
        n_rows    = len(msgs)
        start_row = data_row

        for i, l in enumerate(msgs):
            fon   = C["juft"] if kishi_n % 2 == 0 else C["toq"]
            izoh  = l.get("izoh", "").strip()
            tur   = TUR_EMOJI.get(l["tur"], l["tur"])
            h_val = holat(izoh) if izoh else "-"
            qizoh = (izoh[:120] + "…") if len(izoh) > 120 else (izoh or "-")
            link  = l.get("link", "")

            if i == 0:
                sc(ws, data_row, 1, kishi_n,       fon=fon, bold=True)
                sc(ws, data_row, 2, info["ism"],   fon=fon, bold=True, align="left")
                sc(ws, data_row, 3, info["username"], fon=fon)
                sc(ws, data_row, 4, info["guruh"], fon=fon, align="left")
            else:
                for col in [1, 2, 3, 4]:
                    sc(ws, data_row, col, "", fon=C["cont"])

            sc(ws, data_row, 5, l["sana"],  fon=fon)
            sc(ws, data_row, 6, l["soat"],  fon=fon, bold=True)
            sc(ws, data_row, 7, tur,         fon=fon)

            # Holat rangi
            h_cell = sc(ws, data_row, 8, h_val, fon=fon)
            if "✅" in h_val: h_cell.fill = pf(C["green"])
            if "❌" in h_val: h_cell.fill = pf(C["red_bg"])
            if "❓" in h_val: h_cell.fill = pf(C["yellow"])
            if has_salom(izoh): h_cell.fill = pf(C["salom"])

            # Tur rangi
            tur_cell = ws.cell(row=data_row, column=7)
            if "📷" in tur: tur_cell.fill = pf(C["blue"])
            if "🎥" in tur: tur_cell.fill = pf("FFF2CC")
            if "🎤" in tur: tur_cell.fill = pf("E2EFDA")

            sc(ws, data_row, 9, qizoh, fon=fon, align="left", size=9)

            # Link — kliklanadigan
            lc = sc(ws, data_row, 10, link if link else "-",
                    fon=fon, align="left", size=9, color="0563C1" if link else "888888")
            if link:
                lc.hyperlink = link
                lc.font = Font(size=9, color="0563C1", underline="single")

            ws.row_dimensions[data_row].height = 18
            data_row += 1

        # Kishi uchun merge: №, Ism, Username, Guruh
        if n_rows > 1:
            for col in [1, 2, 3, 4]:
                try:
                    ws.merge_cells(
                        start_row=start_row, start_column=col,
                        end_row=start_row + n_rows - 1, end_column=col
                    )
                    # Merge qilingandan keyin alignment qayta o'rnatish
                    fon = C["juft"] if kishi_n % 2 == 0 else C["toq"]
                    c = ws.cell(row=start_row, column=col)
                    c.alignment = Alignment(
                        horizontal="center" if col != 2 else "left",
                        vertical="center", wrap_text=True
                    )
                except: pass

        # Kishi yakunida yupqa chiziq
        for col in range(1, len(COLS)+1):
            c = ws.cell(row=data_row-1, column=col)
            c.border = Border(
                left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"),  bottom=Side(style="medium")
            )

    # Ustun kengliği
    for col, w in enumerate(WDT, 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = f"A{hdr+1}"

    # Yig'indi qatori
    sc(ws, data_row, 1,
       f"Jami: {kishi_n} kishi  |  {len(glog)} xabar  |  "
       f"Hisobot: {datetime.now(TZ).strftime('%Y-%m-%d %H:%M')}",
       fon=C["sum_bg"], bold=True, color=C["sum_t"], size=9)
    ws.merge_cells(start_row=data_row, start_column=1,
                   end_row=data_row, end_column=len(COLS))
    ws.row_dimensions[data_row].height = 20

    return set(kishi_msgs.keys())


def sheet_bermaganlar(wb, gid: str, nom: str, yuborgan_uids: set):
    ws = wb.create_sheet(title="❌ Munosabat bildirmaganlar"[:30])

    for col, (h, w) in enumerate([("№",4),("Ism",30),("Username",22),("Holati",28)], 1):
        sc(ws, 1, col, h, fon=C["hdr2"], bold=True, color=C["hdr2_t"])
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 24

    gmbrs      = members.get(gid, {})
    bermaganlar = {uid: inf for uid, inf in gmbrs.items()
                   if int(uid) not in yuborgan_uids}

    if not bermaganlar:
        c = ws.cell(row=2, column=1, value="✅ Barcha a'zolar munosabat bildirgan!")
        c.font = Font(bold=True, size=11, color="00AA00")
        c.alignment = Alignment(horizontal="center")
        ws.merge_cells("A2:D2")
    else:
        for i, (uid, inf) in enumerate(bermaganlar.items(), 1):
            r = i + 1
            sc(ws, r, 1, i,                            fon=C["red_bg"])
            sc(ws, r, 2, inf.get("ism","—"),           fon=C["red_bg"], align="left")
            sc(ws, r, 3, inf.get("username","—"),      fon=C["red_bg"])
            sc(ws, r, 4, "❌ Munosabat bildirmagan",   fon=C["red_bg"])
            ws.row_dimensions[r].height = 18

    jr = len(bermaganlar) + 3
    c = ws.cell(row=jr, column=1,
                value=f"Jami: {len(gmbrs)} a'zo  |  "
                      f"Munosabat bildirgan: {len(yuborgan_uids)}  |  "
                      f"Bildirmagan: {len(bermaganlar)}")
    c.font = Font(bold=True, size=10, color="FFFFFF")
    c.fill = pf(C["hdr2"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=jr, start_column=1, end_row=jr, end_column=4)
    ws.row_dimensions[jr].height = 22
    ws.freeze_panes = "A2"

    ws.cell(row=jr+1, column=1,
            value="* Bot qo'shilgandan beri ko'rilgan a'zolar asosida"
    ).font = Font(italic=True, size=8, color="888888")


# ═══════════════════════════════════════════════════════════════
#  EXCEL YUBORISH
# ═══════════════════════════════════════════════════════════════
async def send_excel(ctx, chat_id: int, target_gids: list[str], sana: str | None = None):
    tahlil = {gid: logs[gid] for gid in target_gids if gid in logs and logs[gid]}
    if not tahlil:
        await ctx.bot.send_message(chat_id=chat_id, text="📭 Ma'lumot yo'q.")
        return

    await ctx.bot.send_message(chat_id=chat_id, text="⏳ Excel tayyorlanmoqda…")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for gid, glog in tahlil.items():
        nom      = glog[-1]["guruh"]
        yuborgan = sheet_yuborganlar(wb, nom, glog, sana_filter=sana)
        sheet_bermaganlar(wb, gid, nom, yuborgan)

    fayl   = f"tahlil_{datetime.now(TZ).strftime('%Y%m%d_%H%M%S')}.xlsx"
    path   = f"/tmp/{fayl}"
    wb.save(path)

    nomlar = ", ".join(logs[g][-1]["guruh"] for g in tahlil if logs[g])
    caption = (
        f"📊 *Excel tayyor!*\n"
        f"🏷 {nomlar}\n"
        + (f"📅 Sana: {sana}\n" if sana else "")
        + f"📋 Sheet 1 — ✅ Munosabat bildirganlar\n"
          f"📋 Sheet 2 — ❌ Munosabat bildirmaganlar"
    )
    with open(path, "rb") as f:
        await ctx.bot.send_document(
            chat_id=chat_id, document=f,
            filename=fayl, caption=caption, parse_mode="Markdown",
        )
    os.remove(path)


# ═══════════════════════════════════════════════════════════════
#  INLINE KEYBOARD YORDAMCHILARI
# ═══════════════════════════════════════════════════════════════
def kb_guruhlar() -> InlineKeyboardMarkup | None:
    if not any(logs.values()): return None
    rows = []
    for gid, glog in logs.items():
        if not glog: continue
        nom  = glog[-1]["guruh"]
        nxbr = len(glog)
        uids = len(set(l["user_id"] for l in glog))
        rows.append([InlineKeyboardButton(
            f"📁 {nom[:30]} ({nxbr} xabar, {uids} kishi)",
            callback_data=f"grp:{gid}"
        )])
    rows.append([InlineKeyboardButton("📊 BARCHA guruhlar", callback_data="grp:ALL")])
    return InlineKeyboardMarkup(rows) if rows else None

def kb_sana(gid: str) -> InlineKeyboardMarkup:
    today = datetime.now(TZ).strftime("%Y-%m-%d")
    yest  = (datetime.now(TZ) - timedelta(days=1)).strftime("%Y-%m-%d")
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(f"📅 Bugun ({today})",     callback_data=f"date:{gid}:{today}")],
        [InlineKeyboardButton(f"📅 Kecha ({yest})",      callback_data=f"date:{gid}:{yest}")],
        [InlineKeyboardButton("📅 Boshqa sana...",        callback_data=f"date:{gid}:custom")],
        [InlineKeyboardButton("📊 Barchasi (filtr yo'q)", callback_data=f"date:{gid}:all")],
    ])


# ═══════════════════════════════════════════════════════════════
#  STATISTIKA MATNI
# ═══════════════════════════════════════════════════════════════
def stats_text(gid: str) -> str:
    glog = logs.get(gid, [])
    if not glog: return "📭 Bu guruhda ma'lumot yo'q."
    nom  = glog[-1]["guruh"]
    uids = len(set(l["user_id"] for l in glog))
    mbrs = len(members.get(gid, {}))
    return (
        f"📊 *{nom}*\n\n"
        f"👥 Ko'rilgan a'zolar: *{mbrs}*\n"
        f"✅ Munosabat bildirganlar: *{uids}*\n"
        f"❌ Bildirmaganlar: *{max(0, mbrs-uids)}*\n"
        f"📨 Jami xabarlar: *{len(glog)}*\n\n"
        f"📷 Rasm: {sum(1 for l in glog if l['tur']=='rasm')}  "
        f"🎥 Video: {sum(1 for l in glog if l['tur']=='video')}  "
        f"⭕ Dumaloq: {sum(1 for l in glog if l['tur']=='dumaloq')}\n"
        f"🎤 Ovoz: {sum(1 for l in glog if l['tur']=='ovoz')}  "
        f"💬 Matn: {sum(1 for l in glog if l['tur']=='matn')}\n\n"
        f"📅 Oxirgi: {glog[-1]['sana']} {glog[-1]['soat']}"
    )


# ═══════════════════════════════════════════════════════════════
#  HANDLERS
# ═══════════════════════════════════════════════════════════════

# ── Guruh xabarlarini yig'ish ────────────────────────────────
async def guruh_xabar(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = update.message or update.channel_post
    if not msg or not msg.from_user: return
    entry = tahlil_msg(msg)
    gid   = entry["guruh_id"]
    logs[gid].append(entry)
    reg_member(gid, entry["user_id"], entry["ism"], entry["username"])
    if len(logs[gid]) % 20 == 0:
        save_all()

async def yangi_azolar(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    if not msg or not msg.new_chat_members: return
    gid = str(msg.chat.id)
    for u in msg.new_chat_members:
        if u.is_bot: continue
        ism = f"{u.first_name or ''} {u.last_name or ''}".strip()
        reg_member(gid, u.id, ism, f"@{u.username}" if u.username else str(u.id))
    save_all()

# ── /start ───────────────────────────────────────────────────
async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id): return
    guruhlar_n = len([g for g in logs if logs[g]])
    jami_xabar = sum(len(v) for v in logs.values())
    await update.message.reply_text(
        f"🤖 *Guruh Tahlil Boti*\n\n"
        f"🏘 Guruhlar: *{guruhlar_n}*\n"
        f"📨 Jami xabarlar: *{jami_xabar}*\n\n"
        f"Pastdagi tugmalardan foydalaning 👇",
        parse_mode="Markdown",
        reply_markup=PANEL,
    )

# ── Panel tugmalari (Reply Keyboard) ─────────────────────────
async def panel_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id): return

    # Agar custom sana kiritilayotgan bo'lsa
    pending = ctx.user_data.get("pending_date")
    if pending:
        matn = (update.message.text or "").strip()
        try:
            datetime.strptime(matn, "%Y-%m-%d")
        except ValueError:
            await update.message.reply_text(
                "❌ Format noto'g'ri.\nMasalan: `2026-05-20`",
                parse_mode="Markdown"
            )
            return
        gid_val = pending["gid"]
        ctx.user_data.pop("pending_date")
        target  = list(logs.keys()) if gid_val == "ALL" else [gid_val]
        await send_excel(ctx, update.effective_user.id, target, sana=matn)
        return

    matn = update.message.text or ""

    if matn == "📊 Hisobot":
        kb = kb_guruhlar()
        if not kb:
            await update.message.reply_text("📭 Hozircha ma'lumot yo'q.", reply_markup=PANEL)
            return
        await update.message.reply_text(
            "📊 *1-qadam: Guruhni tanlang*",
            reply_markup=kb, parse_mode="Markdown",
        )

    elif matn == "📈 Statistika":
        aktiv = [g for g in logs if logs[g]]
        if not aktiv:
            await update.message.reply_text("📭 Ma'lumot yo'q.", reply_markup=PANEL)
            return
        if len(aktiv) == 1:
            await update.message.reply_text(stats_text(aktiv[0]),
                                             parse_mode="Markdown", reply_markup=PANEL)
            return
        rows = [[InlineKeyboardButton(f"📊 {logs[g][-1]['guruh'][:35]}",
                                       callback_data=f"stats:{g}")] for g in aktiv]
        await update.message.reply_text("📊 Qaysi guruh?",
                                         reply_markup=InlineKeyboardMarkup(rows))

    elif matn == "👥 A'zolar":
        if not any(members.values()):
            await update.message.reply_text("📭 A'zolar yo'q.", reply_markup=PANEL)
            return
        javob = "👥 *A'zolar:*\n\n"
        for gid, mbrs in members.items():
            if not mbrs: continue
            glog = logs.get(gid, [])
            nom  = glog[-1]["guruh"] if glog else f"Guruh {gid}"
            uids = len(set(l["user_id"] for l in glog))
            javob += (f"🏷 *{nom}*\n"
                      f"  👥 {len(mbrs)} a'zo  ✅ {uids} faol  "
                      f"❌ {max(0, len(mbrs)-uids)} bildirmagan\n\n")
        await update.message.reply_text(javob, parse_mode="Markdown", reply_markup=PANEL)

    elif matn == "🏘 Guruhlar":
        if not any(logs.values()):
            await update.message.reply_text("📭 Guruhlar yo'q.", reply_markup=PANEL)
            return
        javob = "🏘 *Guruhlar:*\n\n"
        for gid, glog in logs.items():
            if not glog: continue
            nom  = glog[-1]["guruh"]
            uids = len(set(l["user_id"] for l in glog))
            mbrs = len(members.get(gid, {}))
            javob += (f"🏷 *{nom}*\n"
                      f"  📨 {len(glog)} xabar  👥 {mbrs} a'zo  "
                      f"✅ {uids} faol\n"
                      f"  ID: `{gid}`\n\n")
        await update.message.reply_text(javob, parse_mode="Markdown", reply_markup=PANEL)

    elif matn == "💾 Saqlash":
        save_all()
        await update.message.reply_text("💾 Ma'lumotlar saqlandi.", reply_markup=PANEL)

    elif matn == "🗑 Tozalash":
        kb = InlineKeyboardMarkup([[
            InlineKeyboardButton("✅ Ha, tozala", callback_data="clear:yes"),
            InlineKeyboardButton("❌ Bekor",      callback_data="clear:no"),
        ]])
        await update.message.reply_text(
            "⚠️ *Barcha ma'lumotlar o'chadi!*",
            reply_markup=kb, parse_mode="Markdown",
        )

# ── Callback ─────────────────────────────────────────────────
async def callback_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    if not is_admin(q.from_user.id):
        await q.answer("❌ Ruxsat yo'q", show_alert=True); return

    d = q.data

    if d.startswith("grp:"):
        gid_val = d[4:]
        await q.edit_message_text(
            "📅 *2-qadam: Sanani tanlang*",
            reply_markup=kb_sana(gid_val), parse_mode="Markdown",
        )

    elif d.startswith("date:"):
        _, gid_val, sana_val = d.split(":", 2)
        if sana_val == "custom":
            ctx.user_data["pending_date"] = {"gid": gid_val}
            await q.edit_message_text(
                "✏️ Sanani yozing:\n`YYYY-MM-DD`  →  masalan: `2026-05-20`",
                parse_mode="Markdown",
            )
            return
        target = list(logs.keys()) if gid_val == "ALL" else [gid_val]
        sana   = None if sana_val == "all" else sana_val
        await q.edit_message_text("⏳ Excel tayyorlanmoqda…")
        await send_excel(ctx, q.from_user.id, target, sana=sana)

    elif d.startswith("stats:"):
        await q.edit_message_text(stats_text(d[6:]), parse_mode="Markdown")

    elif d == "clear:yes":
        logs.clear(); members.clear(); save_all()
        await q.edit_message_text("🗑 Tozalandi.")
    elif d == "clear:no":
        await q.edit_message_text("✅ Bekor qilindi.")


# ═══════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════
def main():
    if not BOT_TOKEN: raise ValueError("❌ BOT_TOKEN yo'q!")
    if not OWNER_ID:  raise ValueError("❌ OWNER_ID yo'q!")
    load_all()

    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start",  start))
    app.add_handler(CommandHandler("panel",  start))
    app.add_handler(CallbackQueryHandler(callback_handler))

    app.add_handler(MessageHandler(
        filters.ChatType.GROUPS & filters.StatusUpdate.NEW_CHAT_MEMBERS,
        yangi_azolar
    ))
    app.add_handler(MessageHandler(
        filters.ChatType.PRIVATE & filters.TEXT & filters.User(list(ADMIN_IDS)),
        panel_handler
    ))
    app.add_handler(MessageHandler(
        filters.ChatType.GROUPS & (
            filters.TEXT | filters.PHOTO | filters.VIDEO |
            filters.VIDEO_NOTE | filters.VOICE | filters.AUDIO |
            filters.Document.ALL | filters.Sticker.ALL
        ),
        guruh_xabar
    ))

    log.info("🤖 Bot v5 ishga tushdi")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()

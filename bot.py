import os, json, logging, re
from datetime import datetime, timezone, timedelta
from collections import defaultdict
from pathlib import Path

from telegram import (
    Update, InlineKeyboardButton, InlineKeyboardMarkup,
    ReplyKeyboardMarkup, KeyboardButton
)
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
def is_admin(uid: int) -> bool:
    return uid in ADMIN_IDS

DATA_DIR  = Path(os.getenv("DATA_DIR", "/data"))
DATA_FILE = DATA_DIR / "logs.json"
MBRS_FILE = DATA_DIR / "members.json"

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(message)s",
    level=logging.INFO
)
log = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════
#  UZUN XABAR YUBORISH
# ═══════════════════════════════════════════════════════════════
async def send_long(obj, text: str, parse_mode="Markdown",
                    reply_markup=None, max_len: int = 3800):
    parts = []
    while text:
        if len(text) <= max_len:
            parts.append(text); break
        cut = text[:max_len].rfind("\n")
        if cut < 1: cut = max_len
        parts.append(text[:cut])
        text = text[cut:].lstrip("\n")
    for i, part in enumerate(parts):
        mkp = reply_markup if i == len(parts) - 1 else None
        if hasattr(obj, "reply_text"):
            await obj.reply_text(part, parse_mode=parse_mode, reply_markup=mkp)
        else:
            await obj[0].send_message(obj[1], part, parse_mode=parse_mode, reply_markup=mkp)

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
#  GID ↔ INDEKS (callback_data 64 belgi chegarasi uchun)
# ═══════════════════════════════════════════════════════════════
def gid_list() -> list[str]:
    """Barcha aktiv guruhlar gid lari (tartiblangan)."""
    return sorted(g for g in logs if logs[g])

def gid_by_idx(idx: str) -> str | None:
    """'ALL' yoki raqamli indeks → gid."""
    if idx == "ALL":
        return "ALL"
    try:
        return gid_list()[int(idx)]
    except (IndexError, ValueError):
        return None

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

SALOM_KW    = ["ассалому алайкум","салом","assalomu alaykum","assalom","salom"]
TUSHUNDI_KW = [
    "тушунарли","tushunarli","тушундим","tushundim","бажардим","bajardim",
    "хоп","hop","майли","mayli","ок","ok","👍","✅","тайёр","tayyor","қабул","qabul"
]
XATO_KW  = ["хато","xato","нотўғри","noto'g'ri","тушунмадим","tushunmadim","❌"]
SAVOL_KW = ["?","савол","savol","нима","nima","қандай","qanday","нега","nega"]

def has_salom(t: str) -> bool:
    return any(k in t.lower() for k in SALOM_KW)

def holat(t: str) -> str:
    tl = t.lower()
    if any(k in tl for k in TUSHUNDI_KW): return "Tushundi ✅"
    if any(k in tl for k in XATO_KW):    return "Xato ❌"
    if any(k in tl for k in SAVOL_KW):   return "Savol ❓"
    return "-"

TUR_EMOJI = {
    "rasm":"📷 Rasm","video":"🎥 Video","dumaloq":"⭕ Dumaloq",
    "ovoz":"🎤 Ovoz","matn":"💬 Matn","fayl":"📄 Fayl",
    "audio":"🎵 Audio","stiker":"🎭 Stiker","boshqa":"📦 Boshqa",
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
        "soat_h"  : now.hour,          # ← soat filtri uchun
        "guruh"   : chat.title or "Unknown",
        "guruh_id": gid,
        "ism"     : (f"{user.first_name or ''} {user.last_name or ''}".strip() if user else "Unknown"),
        "username": (f"@{user.username}" if user and user.username else str(user.id) if user else "-"),
        "user_id" : user.id if user else 0,
        "msg_id"  : mid,
        "link"    : msg_link(gid, mid),
        "tur"     : "", "izoh": "",
    }
    if   message.photo:      entry["tur"]="rasm";    entry["izoh"]=caption
    elif message.video:      entry["tur"]="video";   entry["izoh"]=caption
    elif message.video_note: entry["tur"]="dumaloq"
    elif message.voice:      entry["tur"]="ovoz";    entry["izoh"]=caption
    elif message.audio:      entry["tur"]="audio";   entry["izoh"]=caption
    elif message.document:   entry["tur"]="fayl";    entry["izoh"]=caption
    elif message.sticker:    entry["tur"]="stiker"
    elif message.text:       entry["tur"]="matn";    entry["izoh"]=message.text.strip()
    else:                    entry["tur"]="boshqa"
    return entry

# ═══════════════════════════════════════════════════════════════
#  SOAT FILTRI YORDAMCHISI
# ═══════════════════════════════════════════════════════════════
# soat_filter formatlari:
#   "all"       → hamma vaqt
#   "08-12"     → 08:00 – 11:59
#   "12-18"     → 12:00 – 17:59
#   "18-23"     → 18:00 – 22:59
#   "HH-HH"     → custom oraliq (masalan "09-17")

def soat_filter_fn(entry: dict, sf: str) -> bool:
    """True → entry shu soat oralig'iga to'g'ri keladi."""
    if sf == "all":
        return True
    try:
        h = int(entry.get("soat_h", entry.get("soat","0:0").split(":")[0]))
        s, e = map(int, sf.split("-"))
        if s <= e:
            return s <= h < e
        else:                  # tungi oraliq: masalan 22-06
            return h >= s or h < e
    except:
        return True

def soat_filter_label(sf: str) -> str:
    labels = {
        "all":   "🕐 Barcha vaqt",
        "08-12": "🌅 Tong (08:00–12:00)",
        "12-18": "☀️ Kun (12:00–18:00)",
        "18-23": "🌙 Kech (18:00–23:00)",
        "00-08": "🌃 Tun (00:00–08:00)",
    }
    return labels.get(sf, f"🕐 {sf}")

# ═══════════════════════════════════════════════════════════════
#  BOSHQARUV PANELI
# ═══════════════════════════════════════════════════════════════
PANEL = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton("📊 Hisobot"),   KeyboardButton("📈 Statistika")],
        [KeyboardButton("👥 A'zolar"),   KeyboardButton("🏘 Guruhlar")],
        [KeyboardButton("📋 Bugungi"),   KeyboardButton("💾 Saqlash")],
        [KeyboardButton("🗑 Tozalash"),  KeyboardButton("ℹ️ Yordam")],
    ],
    resize_keyboard=True,
    is_persistent=True,
)

# ═══════════════════════════════════════════════════════════════
#  INLINE KEYBOARD YORDAMCHILARI
# ═══════════════════════════════════════════════════════════════

# ── 1-qadam: Guruh tanlash ───────────────────────────────────
def kb_guruhlar() -> InlineKeyboardMarkup | None:
    gl = gid_list()
    if not gl: return None
    rows = []
    for i, gid in enumerate(gl):
        glog = logs[gid]
        nom  = glog[-1]["guruh"]
        n    = len(glog)
        u    = len(set(l["user_id"] for l in glog))
        rows.append([InlineKeyboardButton(
            f"📁 {nom[:28]} ({n} xbr, {u} kishi)",
            callback_data=f"G:{i}"       # ← indeks, gid emas
        )])
    rows.append([InlineKeyboardButton("📊 BARCHA guruhlar", callback_data="G:ALL")])
    return InlineKeyboardMarkup(rows) if rows else None

# ── 2-qadam: Sana tanlash ────────────────────────────────────
def kb_sana(gi: str) -> InlineKeyboardMarkup:
    today = datetime.now(TZ).strftime("%Y-%m-%d")
    yest  = (datetime.now(TZ) - timedelta(days=1)).strftime("%Y-%m-%d")
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(f"📅 Bugun ({today})",      callback_data=f"D:{gi}:{today}")],
        [InlineKeyboardButton(f"📅 Kecha ({yest})",       callback_data=f"D:{gi}:{yest}")],
        [InlineKeyboardButton("📅 Oxirgi 7 kun",           callback_data=f"D:{gi}:week7")],
        [InlineKeyboardButton("📅 Boshqa sana…",           callback_data=f"D:{gi}:custom")],
        [InlineKeyboardButton("📊 Barchasi (filtr yo'q)",  callback_data=f"D:{gi}:all")],
        [InlineKeyboardButton("◀️ Orqaga",                 callback_data="BACK:grp")],
    ])

# ── 3-qadam: Soat oralig'i tanlash ──────────────────────────
def kb_soat(gi: str, sana_val: str) -> InlineKeyboardMarkup:
    # callback_data: "T:{gi}:{sana_val}:{soat_filter}"
    def btn(label, sf):
        return InlineKeyboardButton(label, callback_data=f"T:{gi}:{sana_val}:{sf}")
    return InlineKeyboardMarkup([
        [btn("🕐 Barcha vaqt",         "all")],
        [btn("🌅 Tong   08:00–12:00",  "08-12"),
         btn("☀️ Kun    12:00–18:00",  "12-18")],
        [btn("🌙 Kech   18:00–23:00",  "18-23"),
         btn("🌃 Tun    00:00–08:00",  "00-08")],
        [btn("✏️ Boshqa soat oralig'…","custom")],
        [InlineKeyboardButton("◀️ Orqaga", callback_data=f"BACK:sana:{gi}")],
    ])

# ═══════════════════════════════════════════════════════════════
#  EXCEL STYLE
# ═══════════════════════════════════════════════════════════════
C = {
    "hdr":"1F4E79","hdr_t":"FFFFFF","hdr2":"C00000","hdr2_t":"FFFFFF",
    "juft":"EBF3FB","toq":"FFFFFF","green":"C6EFCE","red_bg":"FFE0E0",
    "yellow":"FFEB9C","blue":"DDEBF7","salom":"E2EFDA",
    "sum_bg":"2E4057","sum_t":"FFFFFF","cont":"F5F5F5","orange":"FCE4D6",
}
chegara = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)
chegara_qalin = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="medium"),
)
def pf(h): return PatternFill("solid", fgColor=h)

def sc(ws, row, col, val="", fon=None, bold=False, size=10,
       color="000000", wrap=True, align="center", italic=False):
    c = ws.cell(row=row, column=col, value=val)
    if fon: c.fill = pf(fon)
    c.font      = Font(bold=bold, size=size, color=color, italic=italic)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    c.border    = chegara
    return c

COLS = [
    ("№",4),("Ism",24),("Username",16),("Guruh",20),
    ("📅 Sana",12),("🕐 Vaqt",8),("Tur",12),
    ("Holat",14),("Izoh",40),("🔗 Link",42),
]
HDR = [c[0] for c in COLS]
WDT = [c[1] for c in COLS]

# ═══════════════════════════════════════════════════════════════
#  EXCEL — YUBORGAN A'ZOLAR
# ═══════════════════════════════════════════════════════════════
def sheet_yuborganlar(wb, nom: str, glog: list,
                      sana_filter: str | None = None,
                      soat_fil: str = "all") -> set[str]:
    ws  = wb.create_sheet(title=f"✅ {nom}"[:30])
    hdr = 1

    # Filter sarlavhasi
    filter_parts = []
    if sana_filter: filter_parts.append(f"📅 {sana_filter}")
    if soat_fil != "all": filter_parts.append(soat_filter_label(soat_fil))
    if filter_parts:
        c = ws.cell(row=1, column=1, value="Filter: " + "  |  ".join(filter_parts))
        c.font      = Font(bold=True, size=10, color="1F4E79")
        c.alignment = Alignment(horizontal="center")
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS))
        ws.row_dimensions[1].height = 16
        hdr = 2

    for col, h in enumerate(HDR, 1):
        sc(ws, hdr, col, h, fon=C["hdr"], bold=True, color=C["hdr_t"])
    ws.row_dimensions[hdr].height = 30

    # Filter qo'llash
    filtered = glog
    if sana_filter:
        filtered = [l for l in filtered if l.get("sana") == sana_filter]
    if soat_fil != "all":
        filtered = [l for l in filtered if soat_filter_fn(l, soat_fil)]

    # Kishilar bo'yicha guruhlash
    kishi_msgs: dict[int, list[dict]] = {}
    kishi_info: dict[int, dict]       = {}
    for l in sorted(filtered, key=lambda x: x.get("ts", "")):
        uid = l["user_id"]
        if uid not in kishi_info:
            kishi_info[uid] = {"ism": l["ism"], "username": l["username"], "guruh": l["guruh"]}
            kishi_msgs[uid] = []
        kishi_msgs[uid].append(l)

    data_row = hdr + 1
    kishi_n  = 0

    for uid, msgs in sorted(kishi_msgs.items(), key=lambda x: -len(x[1])):
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
                sc(ws, data_row, 1, kishi_n,         fon=fon, bold=True)
                sc(ws, data_row, 2, info["ism"],      fon=fon, bold=True, align="left")
                sc(ws, data_row, 3, info["username"], fon=fon)
                sc(ws, data_row, 4, info["guruh"],    fon=fon, align="left")
            else:
                for col in [1,2,3,4]:
                    sc(ws, data_row, col, "", fon=C["cont"])

            sc(ws, data_row, 5, l["sana"], fon=fon)
            sc(ws, data_row, 6, l["soat"], fon=fon, bold=True)

            tur_cell = sc(ws, data_row, 7, tur, fon=fon)
            if "📷" in tur: tur_cell.fill = pf(C["blue"])
            if "🎥" in tur: tur_cell.fill = pf("FFF2CC")
            if "🎤" in tur: tur_cell.fill = pf("E2EFDA")
            if "📄" in tur: tur_cell.fill = pf(C["orange"])

            h_cell = sc(ws, data_row, 8, h_val, fon=fon)
            if "✅" in h_val:     h_cell.fill = pf(C["green"])
            elif "❌" in h_val:   h_cell.fill = pf(C["red_bg"])
            elif "❓" in h_val:   h_cell.fill = pf(C["yellow"])
            elif has_salom(izoh): h_cell.fill = pf(C["salom"])

            sc(ws, data_row, 9, qizoh, fon=fon, align="left", size=9)

            lc = sc(ws, data_row, 10, link if link else "-",
                    fon=fon, align="left", size=9,
                    color="0563C1" if link else "888888")
            if link:
                lc.hyperlink = link
                lc.font = Font(size=9, color="0563C1", underline="single")

            ws.row_dimensions[data_row].height = 18
            data_row += 1

        if n_rows > 1:
            for col in [1,2,3,4]:
                try:
                    ws.merge_cells(
                        start_row=start_row, start_column=col,
                        end_row=start_row+n_rows-1, end_column=col
                    )
                    c = ws.cell(row=start_row, column=col)
                    c.alignment = Alignment(
                        horizontal="center" if col != 2 else "left",
                        vertical="center", wrap_text=True
                    )
                except: pass

        for col in range(1, len(COLS)+1):
            ws.cell(row=data_row-1, column=col).border = chegara_qalin

    for col, w in enumerate(WDT, 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = f"A{hdr+1}"

    sc(ws, data_row, 1,
       f"Jami: {kishi_n} kishi  |  {len(filtered)} xabar  |  "
       f"Hisobot: {datetime.now(TZ).strftime('%Y-%m-%d %H:%M')}",
       fon=C["sum_bg"], bold=True, color=C["sum_t"], size=9)
    ws.merge_cells(
        start_row=data_row, start_column=1,
        end_row=data_row, end_column=len(COLS)
    )
    ws.row_dimensions[data_row].height = 20

    return {str(uid) for uid in kishi_msgs.keys()}


# ═══════════════════════════════════════════════════════════════
#  EXCEL — BERMAGAN A'ZOLAR
# ═══════════════════════════════════════════════════════════════
def sheet_bermaganlar(wb, gid: str, nom: str, yuborgan_uids: set[str]):
    ws = wb.create_sheet(title="❌ Bildirmaganlar"[:30])
    col_conf = [("№",4),("Ism",30),("Username",22),("Holati",28)]
    for col,(h,w) in enumerate(col_conf, 1):
        sc(ws, 1, col, h, fon=C["hdr2"], bold=True, color=C["hdr2_t"])
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 24

    gmbrs       = members.get(gid, {})
    bermaganlar = {uid: inf for uid, inf in gmbrs.items() if uid not in yuborgan_uids}

    if not bermaganlar:
        c = ws.cell(row=2, column=1, value="✅ Barcha a'zolar munosabat bildirgan!")
        c.font      = Font(bold=True, size=11, color="00AA00")
        c.alignment = Alignment(horizontal="center")
        ws.merge_cells("A2:D2")
        ws.row_dimensions[2].height = 24
    else:
        for i,(uid,inf) in enumerate(
            sorted(bermaganlar.items(), key=lambda x: x[1].get("ism","")), start=1
        ):
            r = i+1
            sc(ws, r, 1, i,                          fon=C["red_bg"])
            sc(ws, r, 2, inf.get("ism","—"),         fon=C["red_bg"], align="left")
            sc(ws, r, 3, inf.get("username","—"),    fon=C["red_bg"])
            sc(ws, r, 4, "❌ Munosabat bildirmagan", fon=C["red_bg"])
            ws.row_dimensions[r].height = 18

    jr = len(bermaganlar)+3
    c  = ws.cell(row=jr, column=1,
                 value=f"Jami: {len(gmbrs)} a'zo  |  "
                       f"Munosabat bildirgan: {len(yuborgan_uids)}  |  "
                       f"Bildirmagan: {len(bermaganlar)}")
    c.font      = Font(bold=True, size=10, color="FFFFFF")
    c.fill      = pf(C["hdr2"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=jr, start_column=1, end_row=jr, end_column=4)
    ws.row_dimensions[jr].height = 22
    ws.freeze_panes = "A2"
    ws.cell(row=jr+1, column=1,
            value="* Bot qo'shilgandan beri ko'rilgan a'zolar asosida"
    ).font = Font(italic=True, size=8, color="888888")


# ═══════════════════════════════════════════════════════════════
#  EXCEL — UMUMIY STATISTIKA SHEET
# ═══════════════════════════════════════════════════════════════
def sheet_statistika(wb, tahlil: dict, soat_fil: str = "all"):
    ws = wb.create_sheet(title="📊 Statistika", index=0)
    c  = ws.cell(row=1, column=1, value="📊 GURUH TAHLIL BOTI — UMUMIY STATISTIKA")
    c.font      = Font(bold=True, size=14, color="FFFFFF")
    c.fill      = pf(C["sum_bg"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("A1:J1")
    ws.row_dimensions[1].height = 36

    info = f"Hisobot sanasi: {datetime.now(TZ).strftime('%Y-%m-%d %H:%M')}"
    if soat_fil != "all":
        info += f"  |  Soat filter: {soat_filter_label(soat_fil)}"
    c2 = ws.cell(row=2, column=1, value=info)
    c2.font      = Font(italic=True, size=9, color="555555")
    c2.alignment = Alignment(horizontal="center")
    ws.merge_cells("A2:J2")

    stat_hdr = ["Guruh nomi","Jami a'zo","Faol","Bildirmagan","Jami xabar",
                "📷 Rasm","🎥 Video","🎤 Ovoz","💬 Matn","📄 Fayl"]
    stat_wdt = [28,12,10,14,12,8,8,8,8,8]
    for col,h in enumerate(stat_hdr, 1):
        sc(ws, 3, col, h, fon=C["hdr"], bold=True, color=C["hdr_t"])
    ws.row_dimensions[3].height = 24

    row=4; jami_xabar=0; jami_faol=0
    for gid, glog in tahlil.items():
        if not glog: continue
        filtered   = [l for l in glog if soat_filter_fn(l, soat_fil)]
        nom        = glog[-1]["guruh"]
        faol       = len(set(l["user_id"] for l in filtered))
        mbrs_n     = len(members.get(gid, {}))
        bildirmagan= max(0, mbrs_n-faol)
        jami_xabar+= len(filtered); jami_faol+=faol
        fon        = C["juft"] if row%2==0 else C["toq"]
        vals = [nom, mbrs_n, faol, bildirmagan, len(filtered),
                sum(1 for l in filtered if l["tur"]=="rasm"),
                sum(1 for l in filtered if l["tur"]=="video"),
                sum(1 for l in filtered if l["tur"]=="ovoz"),
                sum(1 for l in filtered if l["tur"]=="matn"),
                sum(1 for l in filtered if l["tur"]=="fayl")]
        for col,v in enumerate(vals, 1):
            cell = sc(ws, row, col, v, fon=fon, align="left" if col==1 else "center")
            if col==4 and isinstance(v,int) and v>0:
                cell.fill = pf(C["red_bg"])
        ws.row_dimensions[row].height = 18
        row+=1

    sc(ws, row, 1, "JAMI", fon=C["sum_bg"], bold=True, color=C["sum_t"])
    sc(ws, row, 3, jami_faol,  fon=C["sum_bg"], bold=True, color=C["sum_t"])
    sc(ws, row, 5, jami_xabar, fon=C["sum_bg"], bold=True, color=C["sum_t"])
    for col in [2,4,6,7,8,9,10]:
        sc(ws, row, col, "", fon=C["sum_bg"])
    ws.row_dimensions[row].height = 22
    for col,w in enumerate(stat_wdt, 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A4"


# ═══════════════════════════════════════════════════════════════
#  EXCEL YUBORISH
# ═══════════════════════════════════════════════════════════════
async def send_excel(ctx, chat_id: int, target_gids: list[str],
                     sana: str | None = None, soat_fil: str = "all"):
    tahlil = {gid: logs[gid] for gid in target_gids if gid in logs and logs[gid]}
    if not tahlil:
        await ctx.bot.send_message(chat_id=chat_id, text="📭 Ma'lumot yo'q.")
        return

    await ctx.bot.send_message(chat_id=chat_id, text="⏳ Excel tayyorlanmoqda…")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    sheet_statistika(wb, tahlil, soat_fil=soat_fil)

    for gid, glog in tahlil.items():
        nom      = glog[-1]["guruh"]
        yuborgan = sheet_yuborganlar(wb, nom, glog,
                                     sana_filter=sana, soat_fil=soat_fil)
        sheet_bermaganlar(wb, gid, nom, yuborgan)

    fayl  = f"tahlil_{datetime.now(TZ).strftime('%Y%m%d_%H%M%S')}.xlsx"
    path  = f"/tmp/{fayl}"
    wb.save(path)

    nomlar = ", ".join(logs[g][-1]["guruh"] for g in tahlil if logs[g])
    filter_info = ""
    if sana:       filter_info += f"📅 Sana: {sana}\n"
    if soat_fil != "all": filter_info += f"🕐 Soat: {soat_filter_label(soat_fil)}\n"

    caption = (
        f"📊 *Excel tayyor!*\n"
        f"🏷 {nomlar}\n"
        f"{filter_info}"
        f"📋 Sheet 1 — 📊 Umumiy statistika\n"
        f"📋 Sheet 2 — ✅ Munosabat bildirganlar\n"
        f"📋 Sheet 3 — ❌ Munosabat bildirmaganlar"
    )
    with open(path, "rb") as f:
        await ctx.bot.send_document(
            chat_id=chat_id, document=f,
            filename=fayl, caption=caption, parse_mode="Markdown",
        )
    os.remove(path)


# ═══════════════════════════════════════════════════════════════
#  STATISTIKA MATNI
# ═══════════════════════════════════════════════════════════════
def stats_text(gid: str) -> str:
    glog = logs.get(gid, [])
    if not glog: return "📭 Bu guruhda ma'lumot yo'q."
    nom  = glog[-1]["guruh"]
    uids = len(set(l["user_id"] for l in glog))
    mbrs = len(members.get(gid, {}))
    hafta= [(datetime.now(TZ)-timedelta(days=i)).strftime("%Y-%m-%d") for i in range(7)]
    hafta_n = sum(1 for l in glog if l.get("sana") in hafta)
    return (
        f"📊 *{nom}*\n\n"
        f"👥 Ko'rilgan a'zolar: *{mbrs}*\n"
        f"✅ Munosabat bildirganlar: *{uids}*\n"
        f"❌ Bildirmaganlar: *{max(0,mbrs-uids)}*\n"
        f"📨 Jami xabarlar: *{len(glog)}*\n"
        f"📅 Oxirgi 7 kun: *{hafta_n}*\n\n"
        f"📷 {sum(1 for l in glog if l['tur']=='rasm')}  "
        f"🎥 {sum(1 for l in glog if l['tur']=='video')}  "
        f"⭕ {sum(1 for l in glog if l['tur']=='dumaloq')}  "
        f"🎤 {sum(1 for l in glog if l['tur']=='ovoz')}  "
        f"💬 {sum(1 for l in glog if l['tur']=='matn')}  "
        f"📄 {sum(1 for l in glog if l['tur']=='fayl')}\n\n"
        f"📅 Oxirgi: {glog[-1]['sana']} {glog[-1]['soat']}"
    )

def yordam_text() -> str:
    return (
        "ℹ️ *Guruh Tahlil Boti — Yordam*\n\n"
        "📊 *Hisobot* — 3 qadam: guruh → sana → soat → Excel\n"
        "📈 *Statistika* — guruh faollik ko'rsatkichlari\n"
        "👥 *A'zolar* — a'zolar soni va faollik\n"
        "🏘 *Guruhlar* — kuzatilayotgan guruhlar\n"
        "📋 *Bugungi* — bugungi barcha guruhlar hisoboti\n"
        "💾 *Saqlash* — ma'lumotlarni saqlash\n"
        "🗑 *Tozalash* — barcha ma'lumotlarni o'chirish\n\n"
        "📌 *Hisobot oqimi:*\n"
        "  1️⃣ Guruhni tanlang (alohida yoki barchasi)\n"
        "  2️⃣ Sanani tanlang (bugun/kecha/7kun/custom/barchasi)\n"
        "  3️⃣ Soat oralig'ini tanlang\n"
        "  4️⃣ Excel fayl yuklanadi\n\n"
        "🕐 *Soat filtrlari:*\n"
        "  🌅 Tong 08–12  ☀️ Kun 12–18  🌙 Kech 18–23  🌃 Tun 00–08\n"
        "  ✏️ Boshqa — ixtiyoriy oraliq kiritish"
    )


# ═══════════════════════════════════════════════════════════════
#  HANDLERS
# ═══════════════════════════════════════════════════════════════

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

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id): return
    guruhlar_n  = len([g for g in logs if logs[g]])
    jami_xabar  = sum(len(v) for v in logs.values())
    jami_azolar = sum(len(v) for v in members.values())
    await update.message.reply_text(
        f"🤖 *Guruh Tahlil Boti v7*\n\n"
        f"🏘 Guruhlar: *{guruhlar_n}*\n"
        f"👥 Jami a'zolar: *{jami_azolar}*\n"
        f"📨 Jami xabarlar: *{jami_xabar}*\n\n"
        f"Pastdagi tugmalardan foydalaning 👇",
        parse_mode="Markdown",
        reply_markup=PANEL,
    )

# ── Panel (Reply Keyboard) ────────────────────────────────────
async def panel_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id): return

    # Custom sana kutish holati
    if ctx.user_data.get("pending") == "date":
        matn = (update.message.text or "").strip()
        try:
            datetime.strptime(matn, "%Y-%m-%d")
        except ValueError:
            await update.message.reply_text(
                "❌ Format noto'g'ri.\nMasalan: `2026-05-20`", parse_mode="Markdown"
            )
            return
        gi       = ctx.user_data.get("gi", "ALL")
        soat_fil = ctx.user_data.get("soat_fil", "all")
        ctx.user_data.pop("pending", None)
        gid_val  = gid_by_idx(gi)
        target   = gid_list() if gid_val == "ALL" else ([gid_val] if gid_val else [])
        await send_excel(ctx, update.effective_user.id, target, sana=matn, soat_fil=soat_fil)
        return

    # Custom soat kutish holati
    if ctx.user_data.get("pending") == "soat":
        matn = (update.message.text or "").strip()
        # "HH-HH" format tekshirish
        import re as _re
        if not _re.match(r"^\d{1,2}-\d{1,2}$", matn):
            await update.message.reply_text(
                "❌ Format noto'g'ri.\nMasalan: `09-17` yoki `22-06`",
                parse_mode="Markdown"
            )
            return
        s, e = map(int, matn.split("-"))
        if not (0 <= s <= 23 and 0 <= e <= 23):
            await update.message.reply_text("❌ Soat 0–23 oralig'ida bo'lishi kerak.")
            return
        gi       = ctx.user_data.get("gi", "ALL")
        sana_val = ctx.user_data.get("sana_val", "all")
        ctx.user_data.pop("pending", None)
        gid_val  = gid_by_idx(gi)
        target   = gid_list() if gid_val == "ALL" else ([gid_val] if gid_val else [])
        sana     = None if sana_val in ("all", "week7") else sana_val
        await send_excel(ctx, update.effective_user.id, target, sana=sana, soat_fil=matn)
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

    elif matn == "📋 Bugungi":
        aktiv = gid_list()
        if not aktiv:
            await update.message.reply_text("📭 Ma'lumot yo'q.", reply_markup=PANEL)
            return
        bugun = datetime.now(TZ).strftime("%Y-%m-%d")
        await send_excel(ctx, update.effective_user.id, aktiv, sana=bugun)

    elif matn == "📈 Statistika":
        aktiv = gid_list()
        if not aktiv:
            await update.message.reply_text("📭 Ma'lumot yo'q.", reply_markup=PANEL)
            return
        if len(aktiv) == 1:
            await send_long(update.message, stats_text(aktiv[0]), reply_markup=PANEL)
            return
        rows = [[InlineKeyboardButton(
            f"📊 {logs[g][-1]['guruh'][:35]}",
            callback_data=f"stats:{i}"
        )] for i, g in enumerate(aktiv)]
        rows.append([InlineKeyboardButton("📊 Barcha guruhlar", callback_data="stats:ALL")])
        await update.message.reply_text("📊 Qaysi guruh?", reply_markup=InlineKeyboardMarkup(rows))

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
                      f"❌ {max(0,len(mbrs)-uids)} bildirmagan\n\n")
        await send_long(update.message, javob, reply_markup=PANEL)

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
                      f"  📨 {len(glog)} xbr  👥 {mbrs} a'zo  ✅ {uids} faol\n"
                      f"  ID: `{gid}`\n\n")
        await send_long(update.message, javob, reply_markup=PANEL)

    elif matn == "💾 Saqlash":
        save_all()
        await update.message.reply_text("💾 Ma'lumotlar saqlandi.", reply_markup=PANEL)

    elif matn == "ℹ️ Yordam":
        await update.message.reply_text(yordam_text(), parse_mode="Markdown", reply_markup=PANEL)

    elif matn == "🗑 Tozalash":
        kb = InlineKeyboardMarkup([[
            InlineKeyboardButton("✅ Ha, tozala", callback_data="clear:yes"),
            InlineKeyboardButton("❌ Bekor",      callback_data="clear:no"),
        ]])
        await update.message.reply_text(
            "⚠️ *Barcha ma'lumotlar o'chadi!\nIshonchingiz komilmi?*",
            reply_markup=kb, parse_mode="Markdown",
        )


# ── Callback Handler ─────────────────────────────────────────
async def callback_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    if not is_admin(q.from_user.id):
        await q.answer("❌ Ruxsat yo'q", show_alert=True); return

    d = q.data

    # ── Orqaga tugmalari ────────────────────────────────────
    if d == "BACK:grp":
        kb = kb_guruhlar()
        await q.edit_message_text(
            "📊 *1-qadam: Guruhni tanlang*",
            reply_markup=kb, parse_mode="Markdown",
        )
        return

    if d.startswith("BACK:sana:"):
        gi = d[len("BACK:sana:"):]
        await q.edit_message_text(
            "📅 *2-qadam: Sanani tanlang*",
            reply_markup=kb_sana(gi), parse_mode="Markdown",
        )
        return

    # ── 1-qadam: Guruh tanlash (G:indeks) ──────────────────
    if d.startswith("G:"):
        gi = d[2:]
        if gid_by_idx(gi) is None:
            await q.answer("Guruh topilmadi", show_alert=True); return
        await q.edit_message_text(
            "📅 *2-qadam: Sanani tanlang*",
            reply_markup=kb_sana(gi), parse_mode="Markdown",
        )

    # ── 2-qadam: Sana tanlash (D:gi:sana_val) ──────────────
    elif d.startswith("D:"):
        _, gi, sana_val = d.split(":", 2)

        if sana_val == "custom":
            ctx.user_data["pending"]  = "date"
            ctx.user_data["gi"]       = gi
            ctx.user_data["soat_fil"] = "all"
            await q.edit_message_text(
                "✏️ *Sanani yozing:*\n`YYYY-MM-DD`  →  masalan: `2026-05-20`",
                parse_mode="Markdown",
            )
            return

        # week7 → sana_val ni saqlash (keyingi qadamda filter qilinadi)
        await q.edit_message_text(
            "🕐 *3-qadam: Soat oralig'ini tanlang*",
            reply_markup=kb_soat(gi, sana_val), parse_mode="Markdown",
        )

    # ── 3-qadam: Soat tanlash (T:gi:sana_val:soat_fil) ─────
    elif d.startswith("T:"):
        _, gi, sana_val, soat_fil = d.split(":", 3)

        if soat_fil == "custom":
            ctx.user_data["pending"]  = "soat"
            ctx.user_data["gi"]       = gi
            ctx.user_data["sana_val"] = sana_val
            await q.edit_message_text(
                "✏️ *Soat oralig'ini yozing:*\n"
                "`HH-HH`  →  masalan: `09-17` yoki tungi `22-06`",
                parse_mode="Markdown",
            )
            return

        gid_val = gid_by_idx(gi)
        if gid_val is None:
            await q.answer("Guruh topilmadi", show_alert=True); return

        await q.edit_message_text("⏳ Excel tayyorlanmoqda…")

        if sana_val == "week7":
            hafta  = [(datetime.now(TZ)-timedelta(days=i)).strftime("%Y-%m-%d") for i in range(7)]
            target = gid_list() if gid_val=="ALL" else [gid_val]
            orig   = {g: list(logs[g]) for g in target}
            for g in target:
                logs[g] = [l for l in orig[g] if l.get("sana") in hafta]
            await send_excel(ctx, q.from_user.id, target, sana=None, soat_fil=soat_fil)
            for g in target:
                logs[g] = orig[g]
            return

        target = gid_list() if gid_val=="ALL" else [gid_val]
        sana   = None if sana_val=="all" else sana_val
        await send_excel(ctx, q.from_user.id, target, sana=sana, soat_fil=soat_fil)

    # ── Statistika ──────────────────────────────────────────
    elif d.startswith("stats:"):
        val = d[6:]
        if val == "ALL":
            aktiv = gid_list()
            txt   = "\n\n".join(stats_text(g) for g in aktiv)
        else:
            gid_val = gid_by_idx(val)
            txt     = stats_text(gid_val) if gid_val else "❌ Guruh topilmadi."
        if len(txt) <= 3800:
            await q.edit_message_text(txt, parse_mode="Markdown")
        else:
            await q.edit_message_text("📊 Statistika yuborilmoqda…")
            await send_long((ctx.bot, q.from_user.id), txt)

    # ── Tozalash ────────────────────────────────────────────
    elif d == "clear:yes":
        logs.clear(); members.clear(); save_all()
        await q.edit_message_text("🗑 Barcha ma'lumotlar tozalandi.")
    elif d == "clear:no":
        await q.edit_message_text("✅ Bekor qilindi.")


# ═══════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════
def main():
    if not BOT_TOKEN:
        raise ValueError("❌ BOT_TOKEN yo'q!")
    if not OWNER_ID:
        raise ValueError("❌ OWNER_ID yo'q!")

    load_all()
    log.info(f"✅ Yuklandi: {sum(len(v) for v in logs.values())} xabar, "
             f"{sum(len(v) for v in members.values())} a'zo")

    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("panel", start))
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

    log.info("🤖 Guruh Tahlil Boti v7 ishga tushdi")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()

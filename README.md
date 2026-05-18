# 🤖 Guruh Tahlil Boti

Telegram guruhlarida rasm, video, dumaloq video va izohlarni avtomatik tahlil qiluvchi bot.

---

## ✅ Imkoniyatlar

| Xususiyat | Tavsif |
|-----------|--------|
| 📷 Rasm / 🎥 Video / ⭕ Dumaloq video | Barcha media turlarini aniqlaydi |
| 👋 Salomlashish | "Assalomu alaykum", "Salom" va boshqalarni aniqlaydi |
| ✅ Topshiriq holati | "Tushunarli", "Bajardim", "Xato" kabialarni aniqlaydi |
| 🟢 Lotin / 🔵 Kirill | Yozuv turini farqlaydi |
| 📊 Excel hisobot | Vaqt, guruh, foydalanuvchi bo'yicha batafsil |
| 👥 Ko'p guruh | 3-4 guruhni alohida kuzatadi |
| 🔒 Faqat owner | Tahlillar faqat sizga ko'rinadi |

---

## 🚀 O'rnatish (Railway)

### 1. Bot yaratish
1. [@BotFather](https://t.me/BotFather) ga boring
2. `/newbot` yozing
3. Bot nomini kiriting (masalan: `Tahlil Bot`)
4. Token oling: `1234567890:AAF...`

### 2. OWNER_ID olish
1. [@userinfobot](https://t.me/userinfobot) ga boring
2. `/start` bosing
3. ID raqamingizni oling (masalan: `987654321`)

### 3. Railway ga yuklash
1. [railway.app](https://railway.app) ga kiring
2. **New Project → GitHub Repo** (yoki **Deploy from template**)
3. `.env` o'zgaruvchilarni kiriting:
   ```
   BOT_TOKEN=sizning_token
   OWNER_ID=sizning_id
   ```
4. Deploy!

---

## 💬 Buyruqlar (faqat owner uchun)

| Buyruq | Tavsif |
|--------|--------|
| `/start` | Botni ishga tushirish |
| `/stats` | Barcha guruhlar statistikasi |
| `/guruhlar` | Kuzatilayotgan guruhlar ro'yxati |
| `/report` | Barcha guruhlar Excel hisobot |
| `/report_-100123456` | Bitta guruh hisoboti (ID bilan) |
| `/clear` | Barcha ma'lumotlarni tozalash |

---

## 📊 Excel hisobot ustunlari

- **Vaqt** — Xabar yuborilgan sana va soat
- **Guruh** — Guruh nomi
- **Ism** — Foydalanuvchi ismi
- **Username** — @username yoki ID
- **Xabar turi** — Rasm / Video / Matn / Dumaloq video…
- **Matn/Izoh** — Yozilgan matn yoki caption
- **Yozuv** — Lotin / Kirill / Aralash
- **Salom** — Salomlashgan yoki yo'q
- **Holat** — Tushundi / Xato / Savol berdi / Oddiy izoh
- **Tahlil** — Qisqa xulosa

---

## 🔧 Guruhga qo'shish

1. Botni guruhga qo'shing
2. Botni **admin** qiling (xabarlarni o'qishi uchun)
3. Tayyor! Har bir xabar avtomatik tahlil qilinib, sizga yuboriladi.

---

## 📌 Eslatma

- Ma'lumotlar **bot qayta ishga tushganda** o'chadi (RAM da saqlanadi)
- Doimiy saqlash uchun Railway Volume yoki PostgreSQL ulash kerak
- Bir vaqtda 3-4 guruhga qo'shilishi mumkin

---

*Ishlab chiquvchi: Farruxbek | Toshkent hokimiyati*

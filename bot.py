import os
import io
import sqlite3
import logging
import threading
from datetime import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer

from openpyxl import Workbook
from openpyxl.styles import Font
from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    InputFile,
)
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)

logging.basicConfig(
    format="%(asctime)s %(levelname)s %(name)s %(message)s",
    level=logging.INFO,
)
log = logging.getLogger("sarpo-bot")

# ---- settings from environment ----
TOKEN = os.environ["BOT_TOKEN"]
ADMIN_ID = int(os.environ["ADMIN_ID"])
DB_PATH = os.environ.get("DB_PATH", "bot.db")

BRANCHES = ["Sarpo Sergeli", "Sarpo Yunusobod", "Sarpo General"]
SIZES = ["XL", "L", "M", "S"]


# =========================================================
#  DATABASE
# =========================================================
def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = db()
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            position INTEGER,
            article TEXT,
            name TEXT,
            file_id TEXT
        );
        CREATE TABLE IF NOT EXISTS counts (
            branch TEXT,
            product_id INTEGER,
            xl INTEGER, l INTEGER, m INTEGER, s INTEGER,
            user_id INTEGER,
            username TEXT,
            updated_at TEXT,
            PRIMARY KEY (branch, product_id)
        );
        CREATE TABLE IF NOT EXISTS sessions (
            user_id INTEGER PRIMARY KEY,
            branch TEXT,
            idx INTEGER
        );
        """
    )
    conn.commit()
    conn.close()


def products_all():
    conn = db()
    rows = conn.execute("SELECT * FROM products ORDER BY position").fetchall()
    conn.close()
    return rows


def products_count():
    conn = db()
    n = conn.execute("SELECT COUNT(*) c FROM products").fetchone()["c"]
    conn.close()
    return n


def add_product(article, name, file_id):
    conn = db()
    pos = conn.execute("SELECT COALESCE(MAX(position),0) m FROM products").fetchone()["m"] + 1
    conn.execute(
        "INSERT INTO products(position, article, name, file_id) VALUES (?,?,?,?)",
        (pos, article, name, file_id),
    )
    conn.commit()
    conn.close()
    return pos


def delete_last_product():
    conn = db()
    row = conn.execute("SELECT id, article FROM products ORDER BY position DESC LIMIT 1").fetchone()
    if row:
        conn.execute("DELETE FROM products WHERE id=?", (row["id"],))
        conn.commit()
    conn.close()
    return row


def clear_products():
    conn = db()
    conn.execute("DELETE FROM products")
    conn.commit()
    conn.close()


def branch_counted(branch):
    conn = db()
    n = conn.execute("SELECT COUNT(*) c FROM counts WHERE branch=?", (branch,)).fetchone()["c"]
    conn.close()
    return n


def branch_last_user(branch):
    conn = db()
    row = conn.execute(
        "SELECT username FROM counts WHERE branch=? ORDER BY updated_at DESC LIMIT 1",
        (branch,),
    ).fetchone()
    conn.close()
    return row["username"] if row else None


def save_count(branch, product_id, nums, user_id, username):
    conn = db()
    conn.execute(
        """INSERT INTO counts(branch, product_id, xl,l,m,s, user_id, username, updated_at)
           VALUES(?,?,?,?,?,?,?,?,?,?)
           ON CONFLICT(branch, product_id) DO UPDATE SET
             xl=excluded.xl, l=excluded.l, m=excluded.m, s=excluded.s,
             user_id=excluded.user_id, username=excluded.username, updated_at=excluded.updated_at""",
        (branch, product_id, *nums, user_id, username,
         datetime.now().isoformat(timespec="seconds")),
    )
    conn.commit()
    conn.close()


def get_session(user_id):
    conn = db()
    row = conn.execute("SELECT * FROM sessions WHERE user_id=?", (user_id,)).fetchone()
    conn.close()
    return row


def set_session(user_id, branch, idx):
    conn = db()
    conn.execute(
        """INSERT INTO sessions(user_id, branch, idx) VALUES(?,?,?)
           ON CONFLICT(user_id) DO UPDATE SET branch=excluded.branch, idx=excluded.idx""",
        (user_id, branch, idx),
    )
    conn.commit()
    conn.close()


def clear_counts():
    conn = db()
    conn.execute("DELETE FROM counts")
    conn.execute("DELETE FROM sessions")
    conn.commit()
    conn.close()


def who_label(user):
    return f"@{user.username}" if user.username else user.full_name


# =========================================================
#  ADMIN: katalog yuklash
# =========================================================
async def on_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if user.id != ADMIN_ID:
        await update.message.reply_text("Mahsulot yuklash faqat administrator uchun.")
        return
    caption = update.message.caption
    if not caption or not caption.strip():
        await update.message.reply_text(
            "Rasmga izoh qo'shing — artikul va nomi. Masalan:\n22101 Anor gullari - oq rang"
        )
        return
    file_id = update.message.photo[-1].file_id
    parts = caption.strip().split(None, 1)          # birinchi so'z = artikul, qolgani = nomi
    article = parts[0]
    name = parts[1].strip() if len(parts) > 1 else ""
    pos = add_product(article, name, file_id)
    await update.message.reply_text(f"✅ {pos}-mahsulot qo'shildi: {article} — {name}")


async def cmd_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return
    prods = products_all()
    if not prods:
        await update.message.reply_text("Katalog bo'sh.")
        return
    tail = prods[-30:]
    lines = [f"{p['position']}. {p['article']} — {p['name']}" for p in tail]
    head = f"Jami mahsulotlar: {len(prods)} (oxirgi {len(tail)} tasi)\n"
    await update.message.reply_text(head + "\n".join(lines))


async def cmd_undo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return
    row = delete_last_product()
    if row:
        await update.message.reply_text(f"Oxirgi mahsulot o'chirildi: {row['article']}")
    else:
        await update.message.reply_text("O'chiradigan narsa yo'q.")


async def cmd_clear(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return
    kb = InlineKeyboardMarkup([[InlineKeyboardButton("⚠️ Ha, butun katalogni o'chirish", callback_data="clearcat")]])
    await update.message.reply_text("Butun mahsulotlar katalogi o'chirilsinmi?", reply_markup=kb)


async def cmd_done(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return
    n = products_count()
    await update.message.reply_text(f"Katalog tayyor: {n} ta mahsulot. Xodimlar boshlashi mumkin: /start")


async def cmd_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return
    total = products_count()
    lines = []
    for b in BRANCHES:
        done = branch_counted(b)
        who = branch_last_user(b)
        who_txt = f"  @{who}" if (done and who) else ""
        lines.append(f"{b}: {done}/{total}{who_txt}")
    await update.message.reply_text("Hisob holati:\n" + "\n".join(lines))


async def cmd_newcount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return
    kb = InlineKeyboardMarkup([[InlineKeyboardButton("Ha, yangi hisobni boshlash", callback_data="newcount")]])
    await update.message.reply_text(
        "Barcha qoldiqlar nollanadi va yangi hisob boshlanadi. Mahsulotlar katalogi qoladi. Davom etamizmi?",
        reply_markup=kb,
    )


async def cmd_export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Qoldiqlar"
    headers = ["Filial", "Artikul", "Nomi", "XL", "L", "M", "S", "Sana"]
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True)

    conn = db()
    rows = conn.execute(
        """SELECT c.branch, p.article, p.name, c.xl, c.l, c.m, c.s, c.updated_at
           FROM counts c JOIN products p ON p.id = c.product_id
           ORDER BY c.branch, p.position"""
    ).fetchall()
    conn.close()

    for r in rows:
        ws.append([r["branch"], r["article"], r["name"],
                   r["xl"], r["l"], r["m"], r["s"], r["updated_at"]])

    widths = [16, 10, 28, 5, 5, 5, 5, 5, 20]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    fname = f"sarpo_qoldiqlar_{datetime.now():%Y-%m-%d}.xlsx"
    await update.message.reply_document(document=InputFile(bio, filename=fname))
    if not rows:
        await update.message.reply_text("Hozircha birorta ham to'ldirilgan mahsulot yo'q — fayl bo'sh.")


# =========================================================
#  XODIM: sanash jarayoni
# =========================================================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if products_count() == 0:
        await update.message.reply_text("Katalog hali bo'sh. Administrator mahsulotlarni yuklashini kuting.")
        return
    kb = [[InlineKeyboardButton(b, callback_data=f"b:{i}")] for i, b in enumerate(BRANCHES)]
    await update.message.reply_text("Filialni tanlang:", reply_markup=InlineKeyboardMarkup(kb))


async def send_product(context: ContextTypes.DEFAULT_TYPE, user_id: int):
    s = get_session(user_id)
    if not s:
        return
    prods = products_all()
    total = len(prods)
    idx = s["idx"]
    if idx >= total:
        branch = s["branch"]
        await context.bot.send_message(
            user_id, f"✅ Tayyor! {branch}: {total} tadan {total} tasi sanaldi. Rahmat!"
        )
        await notify_admin_done(context, branch)
        return
    p = prods[idx]
    caption = (
        f"{p['article']} · {p['name']}\n"
        f"Mahsulot {idx + 1} / {total}\n\n"
        f"Qoldiqni tartib bilan yozing: {' '.join(SIZES)}\n"
        f"Masalan: 2 3 1 0 4"
    )
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("∅ Qoldiq yo'q", callback_data="zero")],
        [InlineKeyboardButton("← Orqaga", callback_data="back")],
    ])
    await context.bot.send_photo(user_id, photo=p["file_id"], caption=caption, reply_markup=kb)


async def on_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    s = get_session(user.id)
    if not s:
        return  # foydalanuvchi sanash rejimida emas — e'tiborsiz qoldiramiz
    text = update.message.text.replace(",", " ")
    parts = text.split()
    if len(parts) != 5 or not all(x.isdigit() for x in parts):
        await update.message.reply_text(
            f"5 ta son kerak, tartib bilan: {' '.join(SIZES)}. Masalan: 2 3 1 0"
        )
        return
    nums = [int(x) for x in parts]
    context.user_data["pending"] = nums
    summary = " · ".join(f"{sz}: {n}" for sz, n in zip(SIZES, nums))
    kb = InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Saqlash va keyingisi", callback_data="save"),
        InlineKeyboardButton("✏️ Qaytadan", callback_data="redo"),
    ]])
    await update.message.reply_text(summary, reply_markup=kb)


async def handle_save(update: Update, context: ContextTypes.DEFAULT_TYPE, zeros: bool):
    query = update.callback_query
    user = query.from_user
    s = get_session(user.id)
    if not s:
        await query.edit_message_text("Sessiya topilmadi, /start ni bosing.")
        return
    prods = products_all()
    total = len(prods)
    idx = s["idx"]
    if idx >= total:
        await query.edit_message_text("Bu filial allaqachon yakunlangan.")
        return
    if zeros:
        nums = [0, 0, 0, 0]
    else:
        nums = context.user_data.get("pending")
        if nums is None:
            await query.edit_message_text("Avval 5 ta son kiriting.")
            return
    p = prods[idx]
    save_count(s["branch"], p["id"], nums, user.id, user.username or user.full_name)
    context.user_data.pop("pending", None)
    set_session(user.id, s["branch"], idx + 1)
    try:
        await query.edit_message_reply_markup(reply_markup=None)
    except Exception:
        pass
    # administratorga har bir mahsulot bo'yicha xabar
    await notify_admin_item(context, s["branch"], p, nums, who_label(user), idx + 1, total)
    await send_product(context, user.id)


async def handle_back(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    user = query.from_user
    s = get_session(user.id)
    if not s:
        return
    idx = max(0, s["idx"] - 1)
    set_session(user.id, s["branch"], idx)
    context.user_data.pop("pending", None)
    try:
        await query.edit_message_reply_markup(reply_markup=None)
    except Exception:
        pass
    await send_product(context, user.id)


async def notify_admin_item(context, branch, product, nums, who, num, total):
    xl, l, m, s_, xs = nums
    txt = (
        f"📝 {branch} · {who}\n"
        f"{product['article']} {product['name']}  ({num}/{total})\n"
        f"XL:{xl}  L:{l}  M:{m}  S:{s_}"
    )
    try:
        await context.bot.send_message(ADMIN_ID, txt)
    except Exception as e:
        log.warning("cannot notify admin item: %s", e)


async def notify_admin_done(context: ContextTypes.DEFAULT_TYPE, branch: str):
    total = products_count()
    who = branch_last_user(branch)
    who_txt = f"@{who}" if who else "xodim"
    try:
        await context.bot.send_message(
            ADMIN_ID, f"✅ {branch} tayyor — {total} ta mahsulot, sanadi {who_txt}."
        )
    except Exception as e:
        log.warning("cannot notify admin: %s", e)


async def on_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    data = query.data
    user = query.from_user
    await query.answer()

    if data.startswith("b:"):
        idx_b = int(data.split(":")[1])
        branch = BRANCHES[idx_b]
        total = products_count()
        done = branch_counted(branch)
        if total > 0 and done >= total:
            kb = InlineKeyboardMarkup([[InlineKeyboardButton("🔄 Qaytadan sanash", callback_data=f"restart:{idx_b}")]])
            await query.edit_message_text(f"{branch} allaqachon yakunlangan: {total} / {total}.", reply_markup=kb)
            return
        if 0 < done < total:
            kb = InlineKeyboardMarkup([
                [InlineKeyboardButton(f"▶️ Davom etish ({done + 1}-dan)", callback_data=f"cont:{idx_b}")],
                [InlineKeyboardButton("🔄 Boshidan boshlash", callback_data=f"restart:{idx_b}")],
            ])
            await query.edit_message_text(f"{branch}: {done} / {total} sanalgan.", reply_markup=kb)
            return
        set_session(user.id, branch, 0)
        await query.edit_message_text(f"Filial: {branch}")
        await send_product(context, user.id)
        return

    if data.startswith("cont:"):
        idx_b = int(data.split(":")[1])
        branch = BRANCHES[idx_b]
        set_session(user.id, branch, branch_counted(branch))
        await query.edit_message_text(f"Davom etamiz: {branch}")
        await send_product(context, user.id)
        return

    if data.startswith("restart:"):
        idx_b = int(data.split(":")[1])
        branch = BRANCHES[idx_b]
        set_session(user.id, branch, 0)
        await query.edit_message_text(f"Boshidan boshlaymiz: {branch}")
        await send_product(context, user.id)
        return

    if data == "zero":
        await handle_save(update, context, zeros=True)
        return
    if data == "save":
        await handle_save(update, context, zeros=False)
        return
    if data == "redo":
        context.user_data.pop("pending", None)
        await query.edit_message_text("Mayli, 5 ta sonni qaytadan kiriting: " + " ".join(SIZES))
        return
    if data == "back":
        await handle_back(update, context)
        return

    if data == "clearcat":
        if user.id != ADMIN_ID:
            return
        clear_products()
        await query.edit_message_text("Katalog tozalandi.")
        return
    if data == "newcount":
        if user.id != ADMIN_ID:
            return
        clear_counts()
        await query.edit_message_text("Yangi hisob boshlandi. Qoldiqlar nollandi, katalog joyida.")
        return


# =========================================================
#  Health server (Railway uchun) + startup
# =========================================================
class Health(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"ok")

    def log_message(self, *args):
        pass


def run_health():
    port = int(os.environ.get("PORT", "8080"))
    HTTPServer(("0.0.0.0", port), Health).serve_forever()


def main():
    init_db()
    threading.Thread(target=run_health, daemon=True).start()

    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("list", cmd_list))
    app.add_handler(CommandHandler("undo", cmd_undo))
    app.add_handler(CommandHandler("clear", cmd_clear))
    app.add_handler(CommandHandler("done", cmd_done))
    app.add_handler(CommandHandler("export", cmd_export))
    app.add_handler(CommandHandler("report", cmd_report))
    app.add_handler(CommandHandler("newcount", cmd_newcount))
    app.add_handler(CallbackQueryHandler(on_callback))
    app.add_handler(MessageHandler(filters.PHOTO, on_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))

    log.info("Bot started")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()

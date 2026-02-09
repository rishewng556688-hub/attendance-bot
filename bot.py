import os
import sqlite3
from datetime import datetime
import asyncio

from aiogram import Bot, Dispatcher
from aiogram.types import (
    Message,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    CallbackQuery
)
from aiogram.filters import Command
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode

from openpyxl import Workbook

TOKEN = os.getenv("TG_BOT_TOKEN")

ADMIN_IDS = {
    8114765174  # 改成你的 user_id
}

DB_FILE = "attendance.db"

bot = Bot(
    token=TOKEN,
    default=DefaultBotProperties(parse_mode=ParseMode.HTML)
)
dp = Dispatcher()


# ================= 数据库 =================

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
    CREATE TABLE IF NOT EXISTS records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        chat_id INTEGER,
        user_id INTEGER,
        name TEXT,
        action TEXT,
        timestamp TEXT
    )
    """)
    conn.commit()
    conn.close()


def save_record(chat_id, user_id, name, action):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute(
        "INSERT INTO records (chat_id, user_id, name, action, timestamp) VALUES (?, ?, ?, ?, ?)",
        (chat_id, user_id, name, action, now)
    )
    conn.commit()
    conn.close()


def get_today_records(chat_id, user_id):
    today = datetime.now().strftime("%Y-%m-%d")
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
        SELECT action, timestamp
        FROM records
        WHERE chat_id = ? AND user_id = ? AND date(timestamp) = ?
        ORDER BY timestamp ASC
    """, (chat_id, user_id, today))
    rows = c.fetchall()
    conn.close()
    return rows


def get_all_today_records(chat_id):
    today = datetime.now().strftime("%Y-%m-%d")
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
        SELECT user_id, name, action, timestamp
        FROM records
        WHERE chat_id = ? AND date(timestamp) = ?
        ORDER BY user_id, timestamp ASC
    """, (chat_id, today))
    rows = c.fetchall()
    conn.close()
    return rows


def get_month_records(chat_id):
    month_prefix = datetime.now().strftime("%Y-%m")
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
        SELECT user_id, name, action, timestamp
        FROM records
        WHERE chat_id = ?
        AND strftime('%Y-%m', timestamp) = ?
        ORDER BY user_id, timestamp ASC
    """, (chat_id, month_prefix))
    rows = c.fetchall()
    conn.close()
    return rows


# ================= 计算逻辑 =================

def calculate_work_time(records):
    total_seconds = 0
    last_start = None
    working = False
    pause_actions = {"抽烟", "上厕所", "吃饭", "离开"}

    for action, ts in records:
        t = datetime.strptime(ts, "%Y-%m-%d %H:%M:%S")

        if action == "上班" and not working:
            last_start = t
            working = True

        elif action in pause_actions and working:
            total_seconds += (t - last_start).seconds
            working = False

        elif action == "回坐" and not working:
            last_start = t
            working = True

        elif action == "下班" and working:
            total_seconds += (t - last_start).seconds
            working = False

    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    return f"{hours}小时{minutes}分钟"


def count_actions(records):
    counts = {
        "抽烟": 0,
        "吃饭": 0,
        "上厕所": 0,
        "离开": 0
    }
    for action, _ in records:
        if action in counts:
            counts[action] += 1
    return counts


# ================= Excel =================

def export_today_excel(chat_id):
    rows = get_all_today_records(chat_id)
    users = {}

    for user_id, name, action, ts in rows:
        users.setdefault(user_id, {"name": name, "records": []})
        users[user_id]["records"].append((action, ts))

    wb = Workbook()
    ws = wb.active
    ws.title = "今日考勤"

    ws.append(["姓名", "工作时长", "抽烟", "吃饭", "上厕所", "离开"])

    for data in users.values():
        records = data["records"]
        counts = count_actions(records)

        ws.append([
            data["name"],
            calculate_work_time(records),
            counts["抽烟"],
            counts["吃饭"],
            counts["上厕所"],
            counts["离开"],
        ])

    filename = f"attendance_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    wb.save(filename)
    return filename


def export_month_excel(chat_id):
    rows = get_month_records(chat_id)
    users = {}

    for user_id, name, action, ts in rows:
        users.setdefault(user_id, {"name": name, "records": []})
        users[user_id]["records"].append((action, ts))

    wb = Workbook()
    ws = wb.active
    ws.title = "本月考勤"

    ws.append(["姓名", "工作时长", "抽烟", "吃饭", "上厕所", "离开"])

    for data in users.values():
        records = data["records"]
        counts = count_actions(records)

        ws.append([
            data["name"],
            calculate_work_time(records),
            counts["抽烟"],
            counts["吃饭"],
            counts["上厕所"],
            counts["离开"],
        ])

    filename = f"attendance_{datetime.now().strftime('%Y-%m')}.xlsx"
    wb.save(filename)
    return filename


# ================= UI =================

def keyboard():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="🟢 上班", callback_data="上班"),
                InlineKeyboardButton(text="🔴 下班", callback_data="下班")
            ],
            [
                InlineKeyboardButton(text="🚬 抽烟", callback_data="抽烟"),
                InlineKeyboardButton(text="🚻 上厕所", callback_data="上厕所")
            ],
            [
                InlineKeyboardButton(text="🍚 吃饭", callback_data="吃饭"),
                InlineKeyboardButton(text="🚶 离开", callback_data="离开")
            ],
            [
                InlineKeyboardButton(text="🪑 回坐", callback_data="回坐")
            ]
        ]
    )


# ================= 处理器 =================

@dp.message(Command("start"))
async def start(message: Message):
    await message.reply("请选择打卡：", reply_markup=keyboard())


@dp.callback_query()
async def handle_callback(callback: CallbackQuery):
    save_record(
        callback.message.chat.id,
        callback.from_user.id,
        callback.from_user.first_name,
        callback.data
    )

    await callback.answer("已记录")
    await callback.message.reply(
        f"{callback.from_user.first_name} 已打卡：{callback.data}"
    )


@dp.message(Command("today"))
async def today(message: Message):
    records = get_today_records(message.chat.id, message.from_user.id)

    if not records:
        await message.reply("今天还没有打卡记录。")
        return

    text = "📋 今日记录：\n\n"

    for action, ts in records:
        text += f"{ts[11:]} - {action}\n"

    text += f"\n⏱ 实际工作时间：{calculate_work_time(records)}\n"

    counts = count_actions(records)

    text += (
        f"🚬 抽烟：{counts['抽烟']} 次\n"
        f"🍚 吃饭：{counts['吃饭']} 次\n"
        f"🚻 上厕所：{counts['上厕所']} 次\n"
        f"🚶 离开：{counts['离开']} 次"
    )

    await message.reply(text)


@dp.message(Command("admin_excel"))
async def admin_excel(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        await message.reply("⛔ 无权限")
        return

    filename = export_today_excel(message.chat.id)
    await message.reply_document(open(filename, "rb"), caption="📤 今日考勤 Excel")


@dp.message(Command("admin_month_excel"))
async def admin_month_excel(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        await message.reply("⛔ 无权限")
        return

    filename = export_month_excel(message.chat.id)
    await message.reply_document(open(filename, "rb"), caption="📅 本月考勤 Excel")


# ================= 启动 =================

async def main():
    init_db()
    print("Bot started...")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())

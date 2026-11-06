import asyncio
import json
import logging
import os
import re
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import aiohttp
import pandas as pd
import pytz
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    FSInputFile,
    KeyboardButton,
    ReplyKeyboardMarkup,
)
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from bs4 import BeautifulSoup

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
RAW_GROUP_ID = os.getenv("TARGET_GROUP_ID", "").strip()
RAW_THREAD_ID = os.getenv("TARGET_THREAD_ID", "").strip()

TIMEZONE = pytz.timezone("Asia/Tbilisi")
EXCEL_SOURCE_PATH = "Брони.xlsx"
STATE_DB_PATH = "fines_state.json"
REPORT_EXCEL_PATH = "Отчет_штрафы_Батуми.xlsx"

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())
scheduler = AsyncIOScheduler(timezone=TIMEZONE)

is_checking_now = False

def get_target_group_id() -> Optional[int | str]:
    if not RAW_GROUP_ID:
        return None
    if RAW_GROUP_ID.startswith("-") and RAW_GROUP_ID[1:].isdigit():
        return int(RAW_GROUP_ID)
    if RAW_GROUP_ID.isdigit():
        return int(RAW_GROUP_ID)
    return RAW_GROUP_ID

def get_target_thread_id() -> Optional[int]:
    if RAW_THREAD_ID and RAW_THREAD_ID.isdigit():
        return int(RAW_THREAD_ID)
    return None

def clean_plate(val: any) -> str:
    if pd.isna(val):
        return ""
    s = re.sub(r"[^A-Za-z0-9]", "", str(val)).upper()
    if len(s) == 7 and re.match(r"^[A-Z]{2}\d{3}[A-Z]{2}$", s):
        return s
    return ""

def clean_tp(val: any) -> str:
    if pd.isna(val):
        return ""
    s = re.sub(r"[^A-Za-z0-9]", "", str(val)).upper()
    if re.match(r"^[A-Z]{2,3}\d{6,8}$", s):
        return s
    return ""

def load_fleet_from_excel(filepath: str = EXCEL_SOURCE_PATH) -> Dict[str, Dict]:
    if not os.path.exists(filepath):
        return {}
    try:
        df = pd.read_excel(filepath, sheet_name="Номера")
    except Exception:
        return {}

    cars: Dict[str, Dict] = {}
    for idx, row in df.iterrows():
        ncols = len(row)
        raw_name = str(row.iloc[4]).strip() if ncols > 4 and pd.notna(row.iloc[4]) else ""
        plate_main = clean_plate(row.iloc[5]) if ncols > 5 else ""

        if plate_main:
            tps_ordered = []
            for c_idx in [6, 7, 8]:
                if ncols > c_idx:
                    tp = clean_tp(row.iloc[c_idx])
                    if tp and tp not in tps_ordered:
                        tps_ordered.append(tp)
            for c_idx in [0, 1]:
                if ncols > c_idx:
                    tp = clean_tp(row.iloc[c_idx])
                    if tp and tp not in tps_ordered:
                        tps_ordered.append(tp)

            if plate_main not in cars:
                cars[plate_main] = {
                    "name": raw_name if raw_name and raw_name.lower() != "nan" else f"Авто {plate_main}",
                    "tps": tps_ordered,
                }
            else:
                for tp in tps_ordered:
                    if tp not in cars[plate_main]["tps"]:
                        cars[plate_main]["tps"].append(tp)

    for idx, row in df.iterrows():
        ncols = len(row)
        plate_left = clean_plate(row.iloc[0]) if ncols > 0 else ""
        tp_left = clean_tp(row.iloc[1]) if ncols > 1 else ""
        if plate_left and tp_left:
            if plate_left not in cars:
                cars[plate_left] = {"name": f"Авто {plate_left}", "tps": [tp_left]}
            else:
                if tp_left not in cars[plate_left]["tps"]:
                    cars[plate_left]["tps"].append(tp_left)
    return cars

def load_state() -> dict:
    if os.path.exists(STATE_DB_PATH):
        try:
            with open(STATE_DB_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"last_check_time": None, "last_total_debt": 0.0, "known_fines": {}, "cars_summary": {}}

def save_state(state: dict):
    try:
        with open(STATE_DB_PATH, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def calculate_days_and_deadline(issue_date_str: str) -> Tuple[int, str, str]:
    now = datetime.now(TIMEZONE).replace(tzinfo=None)
    parsed_issue = None
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            parsed_issue = datetime.strptime(issue_date_str.strip()[:10], fmt)
            break
        except Exception:
            pass

    if parsed_issue:
        days_passed = (now - parsed_issue).days
        days_left_30 = 30 - days_passed
        days_left_10 = 10 - days_passed
        dl_30_str = (parsed_issue + pd.Timedelta(days=30)).strftime("%d.%m.%Y")
        if days_left_30 < 0:
            return days_left_30, dl_30_str, "OVERDUE"
        elif days_left_30 <= 5:
            return days_left_30, dl_30_str, "URGENT_PENALTY"
        elif 0 <= days_left_10 <= 3:
            return days_left_30, dl_30_str, "URGENT_DISCOUNT"
        return days_left_30, dl_30_str, "NORMAL"
    return 30, "Не указан", "NORMAL"

async def fetch_fines_single_tp(session: aiohttp.ClientSession, plate: str, tp: str) -> List[Dict]:
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
        "Referer": "https://protocols.ge/",
        "Origin": "https://protocols.ge/",
    }
    
    # Имитируем переключение на вкладку "Автомобиль" (передаем параметры поиска по машине и техпаспорту)
    payload_variants = [
        {"type": "car", "plate": plate, "tech_passport": tp},
        {"search_type": "vehicle", "car_number": plate, "document_number": tp},
        {"car": plate, "cert": tp},
        {"plate": plate, "doc_number": tp, "search_target": "vehicle"}
    ]

    found_fines = []
    req_timeout = aiohttp.ClientTimeout(total=5)

    for payload in payload_variants:
        try:
            async with session.post("https://protocols.ge/", data=payload, headers=headers, timeout=req_timeout) as resp:
                if resp.status != 200:
                    continue
                body_text = await resp.text()
                soup = BeautifulSoup(body_text, "html.parser")
                rows = soup.find_all("tr")
                for tr in rows:
                    cols = [td.get_text(" ", strip=True) for td in tr.find_all(["td", "th"])]
                    if len(cols) < 2:
                        continue
                    row_joined = " | ".join(cols)
                    if any(w in row_joined.lower() for w in ["paid", "გადახდილი", "оплачен"]):
                        continue
                    
                    amt_match = re.search(r"(\d{1,4}(?:[.,]\d{1,2})?)\s*(?:GEL|ლ|лари)", row_joined, re.I)
                    prot_match = re.search(r"\b([A-ZА-Я]{0,3}\d{6,12})\b", row_joined, re.I)
                    date_match = re.search(r"\b(\d{2}\.\d{2}\.\d{4}|\d{4}-\d{2}-\d{2})\b", row_joined)

                    if amt_match and prot_match:
                        amount = float(amt_match.group(1).replace(",", "."))
                        if amount <= 0:
                            continue
                        prot_id = prot_match.group(1).upper()
                        issue_date = date_match.group(1) if date_match else datetime.now(TIMEZONE).strftime("%d.%m.%Y")
                        days_left, dl_str, urg = calculate_days_and_deadline(issue_date)
                        found_fines.append({
                            "protocol_id": prot_id, "plate": plate, "tp_used": tp,
                            "amount": amount, "issue_date": issue_date, "deadline": dl_str,
                            "days_left": days_left, "urgency": urg, "description": "Штраф"
                        })
                if found_fines:
                    return found_fines
        except Exception:
            pass
    return found_fines

async def check_car_all_tps(session: aiohttp.ClientSession, plate: str, tps: List[str]) -> Tuple[List[Dict], List[str]]:
    unique_fines = {}
    checked_tps = []
    for tp in tps:
        checked_tps.append(tp)
        fines = await fetch_fines_single_tp(session, plate, tp)
        for f in fines:
            unique_fines[f["protocol_id"]] = f
        if unique_fines:
            break
    return list(unique_fines.values()), checked_tps

def generate_clean_excel_report(cars_summary: Dict[str, Dict], total_debt: float, diff_debt: float) -> str:
    rows = []
    for plate, info in cars_summary.items():
        rows.append({
            "Госномер": plate,
            "Автомобиль": info["name"],
            "Кол-во штрафов (шт)": info["fines_count"],
            "Сумма долга (GEL)": info["total_amount"],
            "Ближайший срок": info["nearest_deadline"],
            "Статус": info["status_label"],
            "Квитанции": ", ".join(info["protocol_ids"]) if info["protocol_ids"] else "—",
        })
    df = pd.DataFrame(rows).sort_values(by=["Сумма долга (GEL)"], ascending=False)
    with pd.ExcelWriter(REPORT_EXCEL_PATH, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Штрафы", index=False)
    return REPORT_EXCEL_PATH

async def run_full_fleet_check(notify_chat_id=None, notify_thread_id=None, progress_msg=None):
    global is_checking_now
    if is_checking_now:
        return
    is_checking_now = True
    try:
        fleet = load_fleet_from_excel()
        if not fleet:
            if notify_chat_id:
                await bot.send_message(chat_id=notify_chat_id, message_thread_id=notify_thread_id, text="⚠️ Сначала отправьте файл `Брони.xlsx` в чат!")
            return

        state = load_state()
        prev_total = float(state.get("last_total_debt", 0.0))
        
        cars_summary = {}
        total_debt = 0.0
        plates_list = list(fleet.keys())
        total_cars = len(plates_list)
        completed_cars = 0

        sem = asyncio.Semaphore(10)
        async with aiohttp.ClientSession() as session:
            async def worker(plate: str):
                nonlocal completed_cars
                car_data = fleet[plate]
                async with sem:
                    fines, checked_tps = await check_car_all_tps(session, plate, car_data["tps"])
                    completed_cars += 1
                    if progress_msg and (completed_cars % 20 == 0 or completed_cars == total_cars):
                        try:
                            await progress_msg.edit_text(f"⏳ Проверка автопарка: **{completed_cars} из {total_cars}** машин...", parse_mode="Markdown")
                        except Exception:
                            pass
                    return plate, car_data, fines, checked_tps

            results = await asyncio.gather(*(worker(p) for p in plates_list))

        for plate, car_data, fines, checked_tps in results:
            car_amount = sum(f["amount"] for f in fines)
            total_debt += car_amount
            cars_summary[plate] = {
                "name": car_data["name"],
                "tps": checked_tps,
                "fines_count": len(fines),
                "total_amount": car_amount,
                "nearest_deadline": fines[0]["deadline"] if fines else "—",
                "status_label": "🟡 Есть штраф" if fines else "✅ Чисто",
                "protocol_ids": [f["protocol_id"] for f in fines],
            }

        diff = total_debt - prev_total
        now_str = datetime.now(TIMEZONE).strftime("%d.%m.%Y %H:%M")
        save_state({"last_check_time": now_str, "last_total_debt": total_debt, "cars_summary": cars_summary})

        report_path = generate_clean_excel_report(cars_summary, total_debt, diff)
        text = (
            f"📊 **СВОДКА ПО ШТРАФАМ ({now_str})**\n"
            f"🚗 Проверено машин: `{len(fleet)}`\n"
            f"💰 **Общий долг:** `{total_debt:.0f} GEL` (Динамика: `{'+' if diff>=0 else ''}{diff:.0f} GEL`)\n"
            f"🔴 Машин со штрафами: `{sum(1 for c in cars_summary.values() if c['total_amount']>0)}`"
        )

        if notify_chat_id:
            await bot.send_document(chat_id=notify_chat_id, message_thread_id=notify_thread_id, document=FSInputFile(report_path), caption=text, parse_mode="Markdown")
    finally:
        is_checking_now = False

def main_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="🔍 Проверить штрафы сейчас")],
            [KeyboardButton(text="📊 Скачать таблицу штрафов"), KeyboardButton(text="📈 Общая сводка и баланс")],
            [KeyboardButton(text="🚗 Список машин и ТП"), KeyboardButton(text="ℹ️ Инструкция и ID чата")],
        ],
        resize_keyboard=True
    )

@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    fleet = load_fleet_from_excel()
    await message.answer(f"👋 Бот запущен! Машин в базе: `{len(fleet)}`.", reply_markup=main_keyboard(), parse_mode="Markdown")

@dp.message(Command("debug"))
async def debug_protocols_site(message: types.Message):
    parts = message.text.split(maxsplit=1)
    plate = clean_plate(parts[1]) if len(parts) > 1 else "PG397PP"
    fleet = load_fleet_from_excel()
    tps = fleet.get(plate, {}).get("tps", ["BVA1627475"])
    tp = tps[0] if tps else "BVA1627475"

    wait_m = await message.answer(f"🛠 Тестируем поиск автомобиля `{plate}` (ТП `{tp}`)...", parse_mode="Markdown")
    async with aiohttp.ClientSession() as session:
        fines, checked = await check_car_all_tps(session, plate, [tp])
        await wait_m.edit_text(f"📡 **Результат теста для `{plate}`:**\n• Проверены ТП: `{checked}`\n• Найдено штрафов: `{len(fines)}`\n• Детали: `{fines}`", parse_mode="Markdown")

@dp.message(F.text == "🔍 Проверить штрафы сейчас")
@dp.message(Command("check"))
async def manual_check_cmd(message: types.Message):
    global is_checking_now
    if is_checking_now:
        await message.answer("⏳ Проверка уже идет...")
        return
    prog_msg = await message.answer("🚀 Запуск проверки с переключением на вкладку автомобиля...\n⏳ Проверено: 0 из 166", parse_mode="Markdown")
    asyncio.create_task(run_full_fleet_check(message.chat.id, message.message_thread_id, prog_msg))

@dp.message(F.text == "📊 Скачать таблицу штрафов")
async def send_excel_report_cmd(message: types.Message):
    if os.path.exists(REPORT_EXCEL_PATH):
        await message.answer_document(document=FSInputFile(REPORT_EXCEL_PATH))
    else:
        await message.answer("ℹ️ Отчет еще не сформирован.")

@dp.message(F.text == "📈 Общая сводка и баланс")
async def show_current_balance(message: types.Message):
    state = load_state()
    await message.answer(f"💰 Общий долг: `{state.get('last_total_debt', 0):.0f} GEL`", parse_mode="Markdown")

@dp.message(F.text == "🚗 Список машин и ТП")
async def show_fleet_stats(message: types.Message):
    fleet = load_fleet_from_excel()
    await message.answer(f"🚗 Загружено машин: `{len(fleet)}`", parse_mode="Markdown")

@dp.message(F.text == "ℹ️ Инструкция и ID чата")
@dp.message(Command("chatid"))
async def show_chat_info(message: types.Message):
    await message.answer(f"💬 Chat ID: `{message.chat.id}`\n🧵 Thread ID: `{message.message_thread_id}`", parse_mode="Markdown")

@dp.message(F.document)
async def handle_new_excel_upload(message: types.Message):
    doc = message.document
    if doc.file_name.lower().endswith(".xlsx"):
        file_info = await bot.get_file(doc.file_id)
        await bot.download_file(file_info.file_path, EXCEL_SOURCE_PATH)
        fleet = load_fleet_from_excel(EXCEL_SOURCE_PATH)
        await message.answer(f"✅ Файл обновлен! Машин: `{len(fleet)}`", parse_mode="Markdown")

async def main():
    scheduler.add_job(run_full_fleet_check, "cron", hour=10, minute=0)
    scheduler.start()
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())

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

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)

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


# =====================================================================
# 1. ЧТЕНИЕ И ОЧИСТКА БАЗЫ МАШИН ИЗ ВКЛАДКИ "Номера" (Брони.xlsx)
# =====================================================================
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
        logging.warning(f"Файл {filepath} не найден.")
        return {}

    try:
        df = pd.read_excel(filepath, sheet_name="Номера")
    except Exception as e:
        logging.error(f"Ошибка при чтении листа 'Номера' из {filepath}: {e}")
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
                if raw_name and cars[plate_main]["name"].startswith("Авто "):
                    cars[plate_main]["name"] = raw_name
                for tp in tps_ordered:
                    if tp not in cars[plate_main]["tps"]:
                        cars[plate_main]["tps"].append(tp)

    for idx, row in df.iterrows():
        ncols = len(row)
        plate_left = clean_plate(row.iloc[0]) if ncols > 0 else ""
        tp_left = clean_tp(row.iloc[1]) if ncols > 1 else ""

        if plate_left and tp_left:
            if plate_left not in cars:
                cars[plate_left] = {
                    "name": f"Авто {plate_left}",
                    "tps": [tp_left],
                }
            else:
                if tp_left not in cars[plate_left]["tps"]:
                    cars[plate_left]["tps"].append(tp_left)

    return cars


# =====================================================================
# 2. БАЗА СОСТОЯНИЯ ШТРАФОВ (fines_state.json)
# =====================================================================
def load_state() -> dict:
    if os.path.exists(STATE_DB_PATH):
        try:
            with open(STATE_DB_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logging.error(f"Ошибка чтения {STATE_DB_PATH}: {e}")
    return {
        "last_check_time": None,
        "last_total_debt": 0.0,
        "known_fines": {},
        "cars_summary": {},
    }


def save_state(state: dict):
    try:
        with open(STATE_DB_PATH, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logging.error(f"Ошибка сохранения {STATE_DB_PATH}: {e}")


# =====================================================================
# 3. ПРОВЕРКА ШТРАФОВ С ЗАЩИТОЙ ОТ ЗАВИСАНИЙ (ТАЙМАУТ 4 СЕК)
# =====================================================================
def is_driver_or_inactive_fine(raw_item: dict | str) -> bool:
    text_repr = str(raw_item).lower()
    if any(w in text_repr for w in ["paid", "გადახდილი", "оплачен", "status_paid", "inactive", "canceled"]):
        return True
    if any(w in text_repr for w in ["driver_only", "personal_fine", "disabled", "greyed"]):
        return True
    return False


def calculate_days_and_deadline(issue_date_str: str, explicit_deadline_str: str = "") -> Tuple[int, str, str]:
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


async def fetch_fines_single_tp(
    session: aiohttp.ClientSession, plate: str, tp: str
) -> List[Dict]:
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
        "Accept": "application/json, text/html, */*",
        "Referer": "https://protocols.ge/",
    }
    payload = {"plate": plate, "tech_passport": tp}
    found_fines: List[Dict] = []

    req_timeout = aiohttp.ClientTimeout(total=4)

    for method, url in [
        ("POST", "https://protocols.ge/"),
        ("GET", "https://protocols.ge/"),
    ]:
        try:
            if method == "GET":
                req_ctx = session.get(url, params=payload, headers=headers, timeout=req_timeout)
            else:
                req_ctx = session.post(url, data=payload, headers=headers, timeout=req_timeout)

            async with req_ctx as resp:
                if resp.status != 200:
                    continue

                body_text = await resp.text()
                soup = BeautifulSoup(body_text, "html.parser")
                rows = soup.find_all("tr")
                for tr in rows:
                    tr_classes = " ".join(tr.get("class", [])).lower()
                    tr_style = (tr.get("style") or "").lower()
                    if any(k in tr_classes or k in tr_style for k in ["disabled", "gray", "grey", "muted", "line-through", "driver", "paid"]):
                        continue

                    cols = [td.get_text(" ", strip=True) for td in tr.find_all(["td", "th"])]
                    if len(cols) < 2:
                        continue
                    row_joined = " | ".join(cols)
                    if is_driver_or_inactive_fine(row_joined):
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
                            "protocol_id": prot_id,
                            "plate": plate,
                            "tp_used": tp,
                            "amount": amount,
                            "issue_date": issue_date,
                            "deadline": dl_str,
                            "days_left": days_left,
                            "urgency": urg,
                            "description": "Активный протокол",
                        })
                if found_fines:
                    return found_fines
        except Exception:
            pass

    return found_fines


async def check_car_all_tps(
    session: aiohttp.ClientSession, plate: str, tps: List[str]
) -> Tuple[List[Dict], List[str]]:
    unique_fines: Dict[str, Dict] = {}
    checked_tps: List[str] = []

    for tp in tps:
        checked_tps.append(tp)
        fines = await fetch_fines_single_tp(session, plate, tp)
        for f in fines:
            pid = f["protocol_id"]
            if pid not in unique_fines:
                unique_fines[pid] = f
        if unique_fines:
            break  # Если нашли штраф по актуальному ТП, старые ТП можно не дергать

    return list(unique_fines.values()), checked_tps


# =====================================================================
# 4. ГЕНЕРАЦИЯ ЧИСТОЙ EXCEL-ТАБЛИЦЫ
# =====================================================================
def generate_clean_excel_report(cars_summary: Dict[str, Dict], total_debt: float, diff_debt: float) -> str:
    rows_all = []
    rows_fines_only = []

    for plate, info in cars_summary.items():
        row_dict = {
            "Госномер": plate,
            "Автомобиль": info["name"],
            "Кол-во штрафов (шт)": info["fines_count"],
            "Сумма долга (GEL)": info["total_amount"],
            "Ближайший срок оплаты": info["nearest_deadline"],
            "Статус срочности": info["status_label"],
            "Номера квитанций": ", ".join(info["protocol_ids"]) if info["protocol_ids"] else "—",
            "Проверенные техпаспорта": ", ".join(info["tps"]),
            "Активный ТП": info.get("active_tp") or "—",
        }
        rows_all.append(row_dict)
        if info["total_amount"] > 0:
            rows_fines_only.append(row_dict)

    df_all = pd.DataFrame(rows_all).sort_values(by=["Сумма долга (GEL)", "Госномер"], ascending=[False, True])
    if rows_fines_only:
        df_fines = pd.DataFrame(rows_fines_only).sort_values(by=["Сумма долга (GEL)"], ascending=False)
    else:
        df_fines = pd.DataFrame(columns=[
            "Госномер", "Автомобиль", "Кол-во штрафов (шт)", "Сумма долга (GEL)",
            "Ближайший срок оплаты", "Статус срочности", "Номера квитанций",
            "Проверенные техпаспорта", "Активный ТП"
        ])

    summary_row = pd.DataFrame([{
        "Госномер": "ИТОГО ПО ПАРКУ",
        "Автомобиль": f"Динамика: {'+' if diff_debt >= 0 else ''}{diff_debt:.0f} GEL",
        "Кол-во штрафов (шт)": int(df_all["Кол-во штрафов (шт)"].sum()) if not df_all.empty else 0,
        "Сумма долга (GEL)": total_debt,
        "Ближайший срок оплаты": datetime.now(TIMEZONE).strftime("%d.%m.%Y %H:%M"),
        "Статус срочности": "",
        "Номера квитанций": "",
        "Проверенные техпаспорта": "",
        "Активный ТП": "",
    }])

    df_fines_out = pd.concat([summary_row, df_fines], ignore_index=True)
    df_all_out = pd.concat([summary_row, df_all], ignore_index=True)

    with pd.ExcelWriter(REPORT_EXCEL_PATH, engine="openpyxl") as writer:
        df_fines_out.to_excel(writer, sheet_name="Только со штрафами", index=False)
        df_all_out.to_excel(writer, sheet_name="Весь автопарк", index=False)

    return REPORT_EXCEL_PATH


# =====================================================================
# 5. БЫСТРАЯ ПАРАЛЛЕЛЬНАЯ ПРОВЕРКА С ЖИВЫМ ПРОГРЕСС-БАРОМ В ЧАТЕ
# =====================================================================
async def run_full_fleet_check(
    notify_chat_id: Optional[int | str] = None,
    notify_thread_id: Optional[int] = None,
    progress_msg: Optional[types.Message] = None,
):
    global is_checking_now
    if is_checking_now:
        return

    is_checking_now = True
    try:
        fleet = load_fleet_from_excel()
        if not fleet:
            if notify_chat_id:
                await bot.send_message(
                    chat_id=notify_chat_id,
                    message_thread_id=notify_thread_id,
                    text="⚠️ Файл `Брони.xlsx` не найден! Просто перешлите файл `Брони.xlsx` сюда в чат.",
                    parse_mode="Markdown",
                )
            return

        prev_state = load_state()
        prev_known_fines: Dict[str, Dict] = prev_state.get("known_fines", {})
        prev_total_debt: float = float(prev_state.get("last_total_debt", 0.0))

        new_known_fines: Dict[str, Dict] = {}
        cars_summary: Dict[str, Dict] = {}
        newly_appeared_fines: List[Dict] = []
        urgent_fines: List[Dict] = []

        total_debt = 0.0
        total_fines_count = 0

        plates_list = list(fleet.keys())
        total_cars = len(plates_list)
        completed_cars = 0

        # Проверяем параллельно по 10 машин сразу, чтобы весь парк проходил за 20-30 секунд
        sem = asyncio.Semaphore(10)

        async with aiohttp.ClientSession() as session:
            async def worker(plate: str):
                nonlocal completed_cars
                car_data = fleet[plate]
                async with sem:
                    fines, checked_tps = await check_car_all_tps(session, plate, car_data["tps"])
                    completed_cars += 1
                    # Каждые 20 машин обновляем статус в чате, чтобы было видно движение
                    if progress_msg and (completed_cars % 20 == 0 or completed_cars == total_cars):
                        try:
                            await progress_msg.edit_text(
                                f"⏳ Идет быстрая проверка автопарка: **{completed_cars} из {total_cars}** машин проверено...",
                                parse_mode="Markdown",
                            )
                        except Exception:
                            pass
                    return plate, car_data, fines, checked_tps

            results = await asyncio.gather(*(worker(p) for p in plates_list))

        for plate, car_data, fines, checked_tps in results:
            car_name = car_data["name"]
            car_amount = sum(f["amount"] for f in fines)
            total_debt += car_amount
            total_fines_count += len(fines)

            nearest_dl = "—"
            status_label = "✅ Чисто"
            active_tp_used = ""

            if fines:
                fines.sort(key=lambda x: x["days_left"])
                nearest_dl = f"{fines[0]['deadline']} ({fines[0]['days_left']} дн.)"
                active_tp_used = fines[0]["tp_used"]

                if any(f["urgency"] == "OVERDUE" for f in fines):
                    status_label = "🔴 ПРОСРОЧЕН!"
                elif any(f["urgency"] == "URGENT_PENALTY" for f in fines):
                    status_label = "🔥 Срочно (скоро удвоение!)"
                elif any(f["urgency"] == "URGENT_DISCOUNT" for f in fines):
                    status_label = "⚠️ Сгорает скидка 20%"
                else:
                    status_label = "🟡 Есть штраф"

            for f in fines:
                pid = f["protocol_id"]
                f_record = {**f, "car_name": car_name}
                new_known_fines[pid] = f_record

                if pid not in prev_known_fines:
                    newly_appeared_fines.append(f_record)

                if f["urgency"] in ("OVERDUE", "URGENT_PENALTY", "URGENT_DISCOUNT"):
                    urgent_fines.append(f_record)

            cars_summary[plate] = {
                "name": car_name,
                "tps": checked_tps,
                "active_tp": active_tp_used,
                "fines_count": len(fines),
                "total_amount": car_amount,
                "nearest_deadline": nearest_dl,
                "status_label": status_label,
                "protocol_ids": [f["protocol_id"] for f in fines],
            }

        diff_debt = total_debt - prev_total_debt
        now_str = datetime.now(TIMEZONE).strftime("%d.%m.%Y %H:%M")

        save_state({
            "last_check_time": now_str,
            "last_total_debt": total_debt,
            "known_fines": new_known_fines,
            "cars_summary": cars_summary,
        })

        report_path = generate_clean_excel_report(cars_summary, total_debt, diff_debt)

        cars_with_fines_cnt = sum(1 for c in cars_summary.values() if c["total_amount"] > 0)
        diff_sign = "+" if diff_debt > 0 else ""
        diff_emoji = "📈" if diff_debt > 0 else ("📉" if diff_debt < 0 else "➖")

        summary_lines = [
            f"📊 **СВОДКА ПО ШТРАФАМ АВТОПАРКА ({now_str})**",
            f"🚗 Проверено машин: `{len(fleet)}` (связок ТП: `{sum(len(x['tps']) for x in fleet.values())}`)",
            f"💰 **Общая сумма штрафов:** `{total_debt:.0f} GEL`",
            f"{diff_emoji} **Динамика:** `{diff_sign}{diff_debt:.0f} GEL`",
            f"🔴 **Машин со штрафами:** `{cars_with_fines_cnt}` из `{len(fleet)}` (протоколов: `{total_fines_count}`)",
        ]

        if newly_appeared_fines:
            summary_lines.append("\n🚨 **НОВЫЕ ШТРАФЫ:**")
            for nf in newly_appeared_fines[:15]:
                summary_lines.append(
                    f"• `{nf['plate']}` ({nf['car_name']}) — **{nf['amount']:.0f} GEL** (до `{nf['deadline']}`)"
                )

        if urgent_fines:
            summary_lines.append("\n⏳ **ГОРЯЩИЕ СРОКИ:**")
            for uf in urgent_fines[:15]:
                summary_lines.append(
                    f"• `{uf['plate']}` ({uf['car_name']}) — **{uf['amount']:.0f} GEL** (осталось `{uf['days_left']} дн.`)"
                )

        summary_lines.append("\n📎 *Полная таблица по всем 166 авто прикреплена ниже.*")
        final_text = "\n".join(summary_lines)

        target_chat = notify_chat_id or get_target_group_id()
        target_thread = notify_thread_id if notify_chat_id else get_target_thread_id()

        if target_chat:
            doc_file = FSInputFile(report_path)
            await bot.send_document(
                chat_id=target_chat,
                message_thread_id=target_thread,
                document=doc_file,
                caption=final_text,
                parse_mode="Markdown",
            )

    except Exception as e:
        logging.error(f"Ошибка в цикле проверки: {e}")
        if notify_chat_id:
            await bot.send_message(
                chat_id=notify_chat_id,
                message_thread_id=notify_thread_id,
                text=f"⚠️ Ошибка при проверке: `{e}`",
                parse_mode="Markdown",
            )
    finally:
        is_checking_now = False


# =====================================================================
# 6. КНОПКИ, КОМАНДЫ И ДИАГНОСТИКА САЙТА (/debug)
# =====================================================================
def main_keyboard() -> ReplyKeyboardMarkup:
    kb = [
        [KeyboardButton(text="🔍 Проверить штрафы сейчас")],
        [KeyboardButton(text="📊 Скачать таблицу штрафов"), KeyboardButton(text="📈 Общая сводка и баланс")],
        [KeyboardButton(text="🚗 Список машин и ТП"), KeyboardButton(text="ℹ️ Инструкция и ID чата")],
    ]
    return ReplyKeyboardMarkup(keyboard=kb, resize_keyboard=True)


@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    fleet = load_fleet_from_excel()
    total_tps = sum(len(v["tps"]) for v in fleet.values())
    await message.answer(
        "👋 **Бот мониторинга штрафов автопарка (Батуми) готов!**\n\n"
        f"🚗 **Машин в базе (`Брони.xlsx`):** `{len(fleet)}`\n"
        f"🪪 **Привязанных техпаспортов:** `{total_tps}`\n\n"
        "🔧 **Полезные команды:**\n"
        "• `/car PG397PP` — проверить одну конкретную машину\n"
        "• `/debug PG397PP` — техническая проверка ответа сайта `protocols.ge`",
        reply_markup=main_keyboard(),
        parse_mode="Markdown",
    )


@dp.message(Command("debug"))
async def debug_protocols_site(message: types.Message):
    """Показывает реальный технический ответ от protocols.ge для диагностики капчи/блокировки."""
    parts = message.text.split(maxsplit=1)
    plate = clean_plate(parts[1]) if len(parts) > 1 else "PG397PP"
    fleet = load_fleet_from_excel()
    tps = fleet.get(plate, {}).get("tps", ["BVA1627475"])
    tp = tps[0] if tps else "BVA1627475"

    wait_m = await message.answer(f"🛠 Тестирую прямое соединение с `protocols.ge` для `{plate}` (ТП `{tp}`)...", parse_mode="Markdown")
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/130.0.0.0"}

    try:
        async with aiohttp.ClientSession() as session:
            async with session.get("https://protocols.ge/", params={"plate": plate, "tech_passport": tp}, headers=headers, timeout=8) as resp:
                status = resp.status
                text = await resp.text()
                has_captcha = "captcha" in text.lower() or "recaptcha" in text.lower()
                snippet = re.sub(r"\s+", " ", BeautifulSoup(text, "html.parser").get_text())[:400]

                await wait_m.edit_text(
                    f"📡 **Результат диагностики `protocols.ge`:**\n\n"
                    f"• **HTTP Статус:** `{status}`\n"
                    f"• **Защита Google reCAPTCHA на странице:** `{'ДА ⚠️ (требуется обход/API)' if has_captcha else 'НЕТ ✅'}`\n"
                    f"• **Фрагмент текста страницы:**\n`{snippet}`",
                    parse_mode="Markdown",
                )
    except Exception as e:
        await wait_m.edit_text(f"❌ **Сайт `protocols.ge` не отвечает с сервера Railway:**\n`{e}`", parse_mode="Markdown")


@dp.message(F.text == "🔍 Проверить штрафы сейчас")
@dp.message(Command("check"))
async def manual_check_cmd(message: types.Message):
    global is_checking_now
    if is_checking_now:
        await message.answer("⏳ Проверка уже идет! Подождите пару секунд, я обновляю прогресс-бар в сообщении выше.")
        return

    prog_msg = await message.answer(
        "🚀 **Запускаю быструю параллельную проверку 166 машин...**\n⏳ Проверено: **0 из 166**",
        parse_mode="Markdown",
    )
    # Запускаем в фоне через create_task, чтобы соединение с Telegram НИКОГДА не рвалось!
    asyncio.create_task(
        run_full_fleet_check(
            notify_chat_id=message.chat.id,
            notify_thread_id=message.message_thread_id,
            progress_msg=prog_msg,
        )
    )


@dp.message(F.text == "ℹ️ Инструкция и ID чата")
@dp.message(Command("chatid"))
async def show_chat_info(message: types.Message):
    await message.answer(
        f"• **ID этого чата (`TARGET_GROUP_ID`):** `{message.chat.id}`\n"
        f"• **ID вкладки (`TARGET_THREAD_ID`):** `{message.message_thread_id or 'Нет (общий чат)'}`",
        parse_mode="Markdown",
    )


@dp.message(F.text == "🚗 Список машин и ТП")
async def show_fleet_stats(message: types.Message):
    fleet = load_fleet_from_excel()
    if not fleet:
        await message.answer("⚠️ База машин пуста. Перешлите файл `Брони.xlsx` сюда в чат!")
        return

    multi_tp_cars = [f"`{p}` ({d['name']}): {', '.join(d['tps'])}" for p, d in fleet.items() if len(d["tps"]) > 1]
    text = (
        f"🚗 **Всего автомобилей в базе:** `{len(fleet)}`\n"
        f"🪪 **Машин с несколькими ТП (старый + новый):** `{len(multi_tp_cars)}`\n\n"
        f"**Примеры связок (первые 15):**\n" + "\n".join(multi_tp_cars[:15])
    )
    await message.answer(text, parse_mode="Markdown")


@dp.message(F.text == "📈 Общая сводка и баланс")
async def show_current_balance(message: types.Message):
    state = load_state()
    if not state.get("last_check_time"):
        await message.answer("ℹ️ Проверка еще не завершалась. Нажмите **«🔍 Проверить штрафы сейчас»**.")
        return

    cars_sum = state.get("cars_summary", {})
    fines_cars = [(p, d) for p, d in cars_sum.items() if d.get("total_amount", 0) > 0]
    fines_cars.sort(key=lambda x: x[1]["total_amount"], reverse=True)

    lines = [
        f"📈 **ТЕКУЩИЙ БАЛАНС ШТРАФОВ АВТОПАРКА**",
        f"🕒 Последнее обновление: `{state['last_check_time']}`",
        f"💰 **Общая задолженность:** `{state.get('last_total_debt', 0):.0f} GEL`",
        f"🚗 **Машин со штрафами:** `{len(fines_cars)}` из `{len(cars_sum)}`\n",
    ]
    for p, d in fines_cars[:20]:
        lines.append(f"• `{p}` ({d['name']}): **{d['total_amount']:.0f} GEL** ({d['fines_count']} шт.)")

    await message.answer("\n".join(lines), parse_mode="Markdown")


@dp.message(F.text == "📊 Скачать таблицу штрафов")
async def send_excel_report_cmd(message: types.Message):
    state = load_state()
    if not os.path.exists(REPORT_EXCEL_PATH):
        fleet = load_fleet_from_excel()
        if not fleet:
            await message.answer("⚠️ Сначала скиньте файл `Брони.xlsx` в чат!")
            return
        empty_sum = {
            p: {
                "name": d["name"],
                "tps": d["tps"],
                "active_tp": "",
                "fines_count": 0,
                "total_amount": 0.0,
                "nearest_deadline": "—",
                "status_label": "Ожидает проверки",
                "protocol_ids": [],
            }
            for p, d in fleet.items()
        }
        generate_clean_excel_report(empty_sum, 0.0, 0.0)

    doc = FSInputFile(REPORT_EXCEL_PATH)
    await message.answer_document(
        document=doc,
        caption=f"📊 Таблица штрафов.\nОбновление: `{state.get('last_check_time') or 'База номеров'}`",
        parse_mode="Markdown",
    )


@dp.message(F.document)
async def handle_new_excel_upload(message: types.Message):
    doc = message.document
    if not doc.file_name.lower().endswith(".xlsx"):
        return

    wait_msg = await message.answer("⏳ Загружаю и проверяю файл `Брони.xlsx`...")
    file_info = await bot.get_file(doc.file_id)
    await bot.download_file(file_info.file_path, EXCEL_SOURCE_PATH)

    fleet = load_fleet_from_excel(EXCEL_SOURCE_PATH)
    total_tps = sum(len(v["tps"]) for v in fleet.values())
    await wait_msg.edit_text(
        f"✅ **Файл `Брони.xlsx` успешно загружен!**\n"
        f"• Найдено машин во вкладке `Номера`: `{len(fleet)}`\n"
        f"• Всего связок техпаспортов: `{total_tps}`",
        parse_mode="Markdown",
    )


async def scheduled_daily_job():
    group_id = get_target_group_id()
    thread_id = get_target_thread_id()
    await run_full_fleet_check(notify_chat_id=group_id, notify_thread_id=thread_id)


async def main():
    scheduler.add_job(scheduled_daily_job, "cron", hour=10, minute=0)
    scheduler.start()
    logging.info("Бот мониторинга штрафов запущен! Расписание активно.")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
